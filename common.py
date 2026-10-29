"""Shared model definitions and helpers for the replication package (v4).
Imported by 03_glm_gam_gbm.py, 04_proxy_capacity.py and 05_bootstrap.py."""
import os, numpy as np, pandas as pd
from sklearn.preprocessing import OneHotEncoder, StandardScaler, SplineTransformer
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline
from sklearn.linear_model import PoissonRegressor
from sklearn.ensemble import HistGradientBoostingRegressor
from sklearn.metrics import roc_auc_score, mean_poisson_deviance

HERE = os.path.dirname(os.path.abspath(__file__))
SEED = 2026
CAT = ["VehAge","MariStat","SocioGrp","VehUsage","VehBody","VehPrice","VehEngine","VehEnergy","VehMaxSpeed","VehClass","Garage"]
NUM = ["LicAge","DrivAge","HasKmLimit","BonusMalus","RiskVar"]
NUM_SPLINE = ["LicAge","DrivAge","BonusMalus","RiskVar"]
NUM_LINEAR = ["HasKmLimit"]
X_COLS = CAT + NUM
FAMS = ["GLM", "GAM", "GBM"]

def load_prepared():
    return pd.read_csv(f"{HERE}/data/freMPL1_prepared.csv", dtype={"VehAge": str})

def make_pre(cols_cat, cols_num):
    return ColumnTransformer([("c", OneHotEncoder(handle_unknown="ignore"), cols_cat), ("n", StandardScaler(), cols_num)])

def glm(cols_cat, cols_num):
    return Pipeline([("pre", make_pre(cols_cat, cols_num)), ("m", PoissonRegressor(alpha=1e-4, max_iter=3000))])

def gam(cols_cat):
    pre = ColumnTransformer([("c", OneHotEncoder(handle_unknown="ignore"), cols_cat),
                             ("s", Pipeline([("sc", StandardScaler()), ("sp", SplineTransformer(n_knots=8, degree=3, include_bias=False))]), NUM_SPLINE),
                             ("l", StandardScaler(), NUM_LINEAR)])
    return Pipeline([("pre", pre), ("m", PoissonRegressor(alpha=1e-3, max_iter=5000))])

def gbm(seed=SEED):
    return HistGradientBoostingRegressor(loss="poisson", max_iter=300, learning_rate=0.05, max_leaf_nodes=15,
                                         min_samples_leaf=50, l2_regularization=1.0, random_state=seed)

_CATS = None
def _categories():
    """Fixed category levels for every categorical column, taken from the full prepared data once, so that a frame in which a
    column is constant (e.g. Gender set to 'Female' for every policy in the counterfactual prediction) is encoded with the same
    integer codes as in training. (Encoding per-frame with .cat.codes would map a constant column to code 0 regardless of its value.)"""
    global _CATS
    if _CATS is None:
        d = load_prepared(); _CATS = {c: sorted(d[c].astype(str).unique()) for c in CAT + ["Gender"]}
    return _CATS

def gbm_frame(d, cols):
    Z = d[cols].copy(); cats = _categories()
    for c in Z.columns:
        if c in cats: Z[c] = pd.Categorical(Z[c].astype(str), categories=cats[c]).codes
    return Z

def fit_pair(fam, tr, seed=SEED):
    """Fit best-estimate (with Gender) and unawareness (without) models of one class. Returns (m_be, m_un, predict_fn)."""
    cols_be, cols_un = X_COLS + ["Gender"], X_COLS
    w = tr.Exposure.values
    if fam == "GLM":
        m_be = glm(CAT + ["Gender"], NUM).fit(tr[cols_be], tr.freq, m__sample_weight=w)
        m_un = glm(CAT, NUM).fit(tr[cols_un], tr.freq, m__sample_weight=w)
        pred = lambda m, d, cols: m.predict(d[cols])
    elif fam == "GAM":
        m_be = gam(CAT + ["Gender"]).fit(tr[cols_be], tr.freq, m__sample_weight=w)
        m_un = gam(CAT).fit(tr[cols_un], tr.freq, m__sample_weight=w)
        pred = lambda m, d, cols: m.predict(d[cols])
    else:
        m_be = gbm(seed).fit(gbm_frame(tr, cols_be), tr.freq, sample_weight=w)
        m_un = gbm(seed).fit(gbm_frame(tr, cols_un), tr.freq, sample_weight=w)
        pred = lambda m, d, cols: m.predict(gbm_frame(d, cols))
    return m_be, m_un, pred

def three_prices(fam, m_be, m_un, pred, te, p_female):
    cols_be, cols_un = X_COLS + ["Gender"], X_COLS
    mu_be = pred(m_be, te, cols_be); mu_un = pred(m_un, te, cols_un)
    te_f, te_m = te.copy(), te.copy(); te_f["Gender"] = "Female"; te_m["Gender"] = "Male"
    mu_df = p_female * pred(m_be, te_f, cols_be) + (1 - p_female) * pred(m_be, te_m, cols_be)
    return mu_be, mu_un, mu_df

def indices(mu_be, mu_un, mu_df, w):
    idi = np.average(np.abs(mu_un - mu_df), weights=w) / np.average(mu_df, weights=w)
    ddi = np.average(np.abs(mu_be - mu_df), weights=w) / np.average(mu_df, weights=w)
    return float(ddi), float(idi)

def evaluate(te, mu):
    w = te.Exposure.values; f, m = te.Female.values == 1, te.Female.values == 0
    return dict(deviance=float(mean_poisson_deviance(te.freq, mu, sample_weight=w)),
                auc=float(roc_auc_score(te.ClaimInd, mu * te.Exposure)),
                prem_f=float(np.average(mu[f], weights=w[f])), prem_m=float(np.average(mu[m], weights=w[m])),
                ratio_fm=float(np.average(mu[f], weights=w[f]) / np.average(mu[m], weights=w[m])),
                ae_f=float(te.ClaimInd.values[f].sum() / (mu[f] * w[f]).sum()),
                ae_m=float(te.ClaimInd.values[m].sum() / (mu[m] * w[m]).sum()))

def recalibrate(mu_cal, y_cal, w_cal):
    """Two-parameter log-linear Poisson recalibration: E[y] = exp(a + b log mu). Returns a function mu -> recalibrated mu."""
    X = np.log(np.clip(mu_cal, 1e-9, None)).reshape(-1, 1)
    m = PoissonRegressor(alpha=0.0, max_iter=5000).fit(X, y_cal, sample_weight=w_cal)
    a, b = float(m.intercept_), float(m.coef_[0])
    return (lambda mu: np.exp(a + b * np.log(np.clip(mu, 1e-9, None)))), a, b
