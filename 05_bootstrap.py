"""Step 5 — inference and robustness.
 (a) Hold-out bootstrap: resample the 6,180 test policies 1,000 times with fitted models held fixed -> percentile 95% CIs.
 (b) Refit robustness: 30 repeated stratified 70/30 splits; every model class refitted on each split -> distribution of IDI
     and of the GBM/GLM ratio under model-refitting variation.
 (c) Calibration robustness: refit each class on 80% of the training set, recalibrate BE and UN predictions on the
     remaining 20% with a two-parameter log-linear Poisson map, recompute IDI and A/E on the hold-out.
Output: results_bootstrap.json"""
import os, json, numpy as np, pandas as pd, warnings
from sklearn.model_selection import train_test_split
from common import *
warnings.filterwarnings("ignore")
N_BOOT, N_REFIT = 1000, 30
df = load_prepared(); te = pd.read_csv(f"{HERE}/data/test_predictions.csv", dtype={"VehAge": str})
w_te = te.Exposure.values; fem = te.Female.values == 1; n = len(te)
out = {}

# ---------- (a) hold-out bootstrap, models fixed ----------
rs = np.random.RandomState(SEED)
boot = {f: {"idi": [], "ddi": [], "share5_f": [], "share5_m": []} for f in FAMS}
ratios = {"ratio_GBM_GLM": [], "ratio_GAM_GLM": [], "ratio_GBM_GAM": []}
store = {f: (te[f"mu_be_{f}"].values, te[f"mu_un_{f}"].values, te[f"mu_df_{f}"].values) for f in FAMS}
for _ in range(N_BOOT):
    ix = rs.choice(n, n, replace=True); w = w_te[ix]; fx = fem[ix]; idis = {}
    for f in FAMS:
        be, un, dfp = (a[ix] for a in store[f]); ddi, idi = indices(be, un, dfp, w); r = un / dfp
        boot[f]["idi"].append(idi); boot[f]["ddi"].append(ddi)
        boot[f]["share5_f"].append(float((np.abs(r[fx] - 1) > 0.05).mean())); boot[f]["share5_m"].append(float((np.abs(r[~fx] - 1) > 0.05).mean()))
        idis[f] = idi
    ratios["ratio_GBM_GLM"].append(idis["GBM"] / idis["GLM"]); ratios["ratio_GAM_GLM"].append(idis["GAM"] / idis["GLM"]); ratios["ratio_GBM_GAM"].append(idis["GBM"] / idis["GAM"])
ci = lambda v: [float(np.percentile(v, 2.5)), float(np.percentile(v, 97.5))]
pt = {f: indices(*store[f], w_te)[1] for f in FAMS}
out["holdout"] = {"n_boot": N_BOOT, "design": "hold-out policies resampled with replacement; fitted models held fixed; percentile 95% CI"}
for f in FAMS: out["holdout"][f] = {k: ci(v) for k, v in boot[f].items()}
out["holdout"]["ratio_GBM_GLM"] = {"point": pt["GBM"] / pt["GLM"], "ci": ci(ratios["ratio_GBM_GLM"])}
out["holdout"]["ratio_GAM_GLM"] = {"point": pt["GAM"] / pt["GLM"], "ci": ci(ratios["ratio_GAM_GLM"])}
out["holdout"]["ratio_GBM_GAM"] = {"point": pt["GBM"] / pt["GAM"], "ci": ci(ratios["ratio_GBM_GAM"])}

# ---------- (b) refit robustness: repeated splits ----------
rec = {f: {"idi": [], "ddi": [], "deviance_un": [], "auc_un": [], "gender_used": []} for f in FAMS}; rr = []
for k in range(N_REFIT):
    seed = SEED + 1 + k
    tr_k, te_k = train_test_split(df, test_size=0.3, random_state=seed, stratify=df.ClaimInd)
    pf = tr_k.Female.mean(); wk = te_k.Exposure.values; idis = {}
    for f in FAMS:
        m_be, m_un, pred = fit_pair(f, tr_k, seed)
        be, un, dfp = three_prices(f, m_be, m_un, pred, te_k, pf); ddi, idi = indices(be, un, dfp, wk)
        # does the best-estimate model respond to gender at all? (flip gender for every policy and compare)
        tf, tm = te_k.copy(), te_k.copy(); tf["Gender"] = "Female"; tm["Gender"] = "Male"
        used = bool(np.max(np.abs(pred(m_be, tf, X_COLS + ["Gender"]) - pred(m_be, tm, X_COLS + ["Gender"]))) > 1e-9)
        ev = evaluate(te_k, un)
        rec[f]["idi"].append(idi); rec[f]["ddi"].append(ddi); rec[f]["deviance_un"].append(ev["deviance"]); rec[f]["auc_un"].append(ev["auc"]); rec[f]["gender_used"].append(used); idis[f] = idi
    rr.append(idis["GBM"] / idis["GLM"]); print(f"refit {k+1}/{N_REFIT}: IDI GLM {idis['GLM']:.4f} GAM {idis['GAM']:.4f} GBM {idis['GBM']:.4f} ratio {rr[-1]:.2f} gender_used GBM={rec['GBM']['gender_used'][-1]}")
summ = lambda v: {"mean": float(np.mean(v)), "sd": float(np.std(v, ddof=1)), "median": float(np.median(v)), "min": float(np.min(v)), "max": float(np.max(v)), "p2.5": float(np.percentile(v, 2.5)), "p97.5": float(np.percentile(v, 97.5))}
out["refit"] = {"n_splits": N_REFIT, "design": "repeated stratified 70/30 splits with all models refitted; seeds 2027-2056"}
for f in FAMS:
    out["refit"][f] = {k: summ(v) for k, v in rec[f].items() if k != "gender_used"}
    out["refit"][f]["values_idi"] = [float(x) for x in rec[f]["idi"]]
    out["refit"][f]["n_gender_unused"] = int(sum(1 for u in rec[f]["gender_used"] if not u))
g_idi, b_idi = np.array(rec["GLM"]["idi"]), np.array(rec["GBM"]["idi"])
act = b_idi > 0
out["refit"]["ratio_GBM_GLM"] = summ(rr); out["refit"]["ratio_GBM_GLM_values"] = [float(x) for x in rr]
out["refit"]["n_GBM_above_GLM"] = int((b_idi > g_idi).sum())
out["refit"]["n_GBM_zero"] = int((~act).sum())
out["refit"]["GBM_idi_when_gender_used"] = summ(b_idi[act]) if act.any() else None
out["refit"]["ratio_when_gender_used"] = summ(np.array(rr)[act]) if act.any() else None
out["refit"]["n_ratio_above_2_when_used"] = int((np.array(rr)[act] > 2).sum())

# ---------- (c) calibration robustness ----------
# Recalibration maps are estimated from 5-fold cross-fitted (out-of-fold) predictions on the TRAINING set, so that the
# maps are learned on ~14,400 policies never seen by the model producing the prediction; they are then applied to the
# main models' hold-out predictions. Two maps are reported: intercept-only (level) and two-parameter log-linear (level + slope).
from sklearn.model_selection import StratifiedKFold
tr, te0 = train_test_split(df, test_size=0.3, random_state=SEED, stratify=df.ClaimInd)
pf = tr.Female.mean(); w0 = te0.Exposure.values; calib = {}
kf = StratifiedKFold(5, shuffle=True, random_state=SEED)
for f in FAMS:
    oof_be, oof_un = np.zeros(len(tr)), np.zeros(len(tr))
    for a, b in kf.split(tr, tr.ClaimInd):
        m_be_k, m_un_k, pred_k = fit_pair(f, tr.iloc[a])
        oof_be[b] = pred_k(m_be_k, tr.iloc[b], X_COLS + ["Gender"]); oof_un[b] = pred_k(m_un_k, tr.iloc[b], X_COLS)
    m_be, m_un, pred = fit_pair(f, tr)                      # main models (identical to step 3)
    be_raw, un_raw, df_raw = three_prices(f, m_be, m_un, pred, te0, pf)
    cols_be = X_COLS + ["Gender"]; tf, tm = te0.copy(), te0.copy(); tf["Gender"] = "Female"; tm["Gender"] = "Male"
    be_f, be_m = pred(m_be, tf, cols_be), pred(m_be, tm, cols_be)
    res_f = {"raw": {}, "level": {}, "loglinear": {}}
    ddi, idi = indices(be_raw, un_raw, df_raw, w0); e = evaluate(te0, un_raw)
    res_f["raw"] = {"idi": idi, "ddi": ddi, "ae_f": e["ae_f"], "ae_m": e["ae_m"], "deviance_un": e["deviance"]}
    # level (intercept-only) map: multiply by observed/expected on OOF predictions
    c_be = float(tr.ClaimInd.sum() / (oof_be * tr.Exposure).sum()); c_un = float(tr.ClaimInd.sum() / (oof_un * tr.Exposure).sum())
    be_l, un_l, df_l = c_be * be_raw, c_un * un_raw, c_be * (pf * be_f + (1 - pf) * be_m)
    ddi, idi = indices(be_l, un_l, df_l, w0); e = evaluate(te0, un_l)
    res_f["level"] = {"idi": idi, "ddi": ddi, "ae_f": e["ae_f"], "ae_m": e["ae_m"], "deviance_un": e["deviance"], "factor_BE": c_be, "factor_UN": c_un}
    # two-parameter log-linear map
    g_be, a1, b1 = recalibrate(oof_be, tr.freq.values, tr.Exposure.values); g_un, a2, b2 = recalibrate(oof_un, tr.freq.values, tr.Exposure.values)
    be_g, un_g, df_g = g_be(be_raw), g_un(un_raw), pf * g_be(be_f) + (1 - pf) * g_be(be_m)
    ddi, idi = indices(be_g, un_g, df_g, w0); e = evaluate(te0, un_g)
    res_f["loglinear"] = {"idi": idi, "ddi": ddi, "ae_f": e["ae_f"], "ae_m": e["ae_m"], "deviance_un": e["deviance"], "params_BE": [a1, b1], "params_UN": [a2, b2]}
    calib[f] = res_f
    print(f"calib {f}: IDI raw {res_f['raw']['idi']:.4f} level {res_f['level']['idi']:.4f} loglinear {res_f['loglinear']['idi']:.4f} (slope UN {b2:.3f}); "
          f"A/E raw {res_f['raw']['ae_f']:.3f}/{res_f['raw']['ae_m']:.3f} -> loglinear {res_f['loglinear']['ae_f']:.3f}/{res_f['loglinear']['ae_m']:.3f}")
for kind in ("raw", "level", "loglinear"):
    calib[f"ratio_GBM_GLM_{kind}"] = calib["GBM"][kind]["idi"] / calib["GLM"][kind]["idi"]
out["calibration"] = {"design": "5-fold cross-fitted out-of-fold training predictions used to estimate level-only and two-parameter log-linear Poisson recalibration maps; maps applied to main models' hold-out predictions; DF price rebuilt from recalibrated best-estimate model", **calib}
json.dump(out, open(f"{HERE}/results_bootstrap.json", "w"), indent=2)
print(json.dumps({k: out["refit"][k] for k in FAMS + ["ratio_GBM_GLM"]}, indent=1))
