"""Step 6 — additional checks requested in the October 2026 review (v7).
 (a) A/E by gender: hold-out bootstrap 95% CIs, training-sample A/E, and mean/sd across the 30 refitted splits.
 (b) Decomposing the GBM indirect-discrimination index:
     - implied unawareness price  mu_UN*(x) = sum_d mu_BE(x,d) P(D=d|X=x), with P(D|X) from a gender classifier fitted on the
       training set (no refitting of the pricing model -> isolates the proxy effect given the best-estimate model);
     - placebo: refit the best-estimate model with an independent random binary variable in place of gender (10 draws) and
       compute the same index -> measures the refitting/tree-structure component that is unrelated to any proxy;
     - interaction-free GBM (interaction_cst: one feature per constraint set) -> isolates interaction-learning capacity.
 (c) Hyper-parameter sensitivity: ridge alpha grid for the GLM (5-fold CV deviance) and a GBM grid (5-fold CV deviance),
     each with the resulting IDI, plus early-stopping GBM; log-linear recalibration slope by alpha.
 (d) Regulatory Stringency Index: dominance statement and a re-scoring (coding perturbation) analysis.
 (e) Wall-clock time of the fairness routine (steps 3 and 4) on this machine.
Output: results_review.json"""
import os, json, time, itertools, platform, numpy as np, pandas as pd, warnings
from sklearn.model_selection import train_test_split, StratifiedKFold
from sklearn.ensemble import HistGradientBoostingRegressor, HistGradientBoostingClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.pipeline import Pipeline
from sklearn.metrics import roc_auc_score
from common import *
warnings.filterwarnings("ignore")
T0 = time.time()
df = load_prepared()
tr, te = train_test_split(df, test_size=0.3, random_state=SEED, stratify=df.ClaimInd)
tr, te = tr.copy(), te.copy()
pf = tr.Female.mean(); w_te = te.Exposure.values; fem = te.Female.values == 1
cols_be, cols_un = X_COLS + ["Gender"], X_COLS
out = {"machine": {"platform": platform.platform(), "cpu": "Apple M1 Pro (10 cores)", "python": platform.python_version()}}

def ae_by_gender(d, mu):
    w = d.Exposure.values; f = d.Female.values == 1
    return float(d.ClaimInd.values[f].sum() / (mu[f] * w[f]).sum()), float(d.ClaimInd.values[~f].sum() / (mu[~f] * w[~f]).sum())

# ------------------------------------------------------------------ (e) timing of the routine (steps 3 + 4 equivalents)
t = time.time()
models = {}; prices = {}
for fam in FAMS:
    m_be, m_un, pred = fit_pair(fam, tr); models[fam] = (m_be, m_un, pred)
    prices[fam] = three_prices(fam, m_be, m_un, pred, te, pf)
t_models = time.time() - t
t = time.time()
cv = StratifiedKFold(5, shuffle=True, random_state=SEED); aucs = []
Zall = gbm_frame(df, X_COLS)
for a, b in cv.split(Zall, df.Female):
    c = HistGradientBoostingClassifier(max_iter=300, learning_rate=0.05, max_leaf_nodes=15, random_state=SEED).fit(Zall.iloc[a], df.Female.iloc[a])
    aucs.append(roc_auc_score(df.Female.iloc[b], c.predict_proba(Zall.iloc[b])[:, 1]))
t_proxy = time.time() - t
out["runtime_seconds"] = {"three_model_pairs_and_prices": round(t_models, 1), "proxy_capacity_cv": round(t_proxy, 1), "routine_total": round(t_models + t_proxy, 1)}
print("runtime", out["runtime_seconds"])

# ------------------------------------------------------------------ (a) A/E uncertainty
rs = np.random.RandomState(SEED); n = len(te); NB = 1000
ae = {}
for fam in FAMS:
    for k, mu in zip(["BE", "UN", "DF"], prices[fam]):
        f0, m0 = ae_by_gender(te, mu); bf, bm = [], []
        for _ in range(NB):
            ix = rs.choice(n, n, replace=True); d = te.iloc[ix]; bf_, bm_ = ae_by_gender(d, mu[ix]); bf.append(bf_); bm.append(bm_)
        ae[f"{fam}_{k}"] = {"ae_f": f0, "ae_m": m0, "ci_f": [float(np.percentile(bf, 2.5)), float(np.percentile(bf, 97.5))],
                            "ci_m": [float(np.percentile(bm, 2.5)), float(np.percentile(bm, 97.5))],
                            "gap_m_minus_f": m0 - f0, "ci_gap": [float(np.percentile(np.array(bm) - np.array(bf), 2.5)), float(np.percentile(np.array(bm) - np.array(bf), 97.5))]}
    # training-sample A/E for BE and UN
    m_be, m_un, pred = models[fam]
    for k, m, cols in [("BE", m_be, cols_be), ("UN", m_un, cols_un)]:
        f1, m1 = ae_by_gender(tr, pred(m, tr, cols)); ae[f"{fam}_{k}"]["train_ae_f"] = f1; ae[f"{fam}_{k}"]["train_ae_m"] = m1
out["ae"] = {"n_boot": NB, **ae}
print("A/E GLM UN hold-out", ae["GLM_UN"]["ae_f"], ae["GLM_UN"]["ci_f"], ae["GLM_UN"]["ae_m"], ae["GLM_UN"]["ci_m"], "train", ae["GLM_UN"]["train_ae_f"], ae["GLM_UN"]["train_ae_m"])

# ------------------------------------------------------------------ (b1) implied unawareness price via P(D|X)
Ztr, Zte = gbm_frame(tr, X_COLS), gbm_frame(te, X_COLS)
clf_g = HistGradientBoostingClassifier(max_iter=300, learning_rate=0.05, max_leaf_nodes=15, random_state=SEED).fit(Ztr, tr.Female)
clf_l = Pipeline([("pre", make_pre(CAT, NUM)), ("m", LogisticRegression(max_iter=3000))]).fit(tr[X_COLS], tr.Female)
p_g, p_l = clf_g.predict_proba(Zte)[:, 1], clf_l.predict_proba(te[X_COLS])[:, 1]
out["proxy_auc_holdout"] = {"gbm": float(roc_auc_score(te.Female, p_g)), "logit": float(roc_auc_score(te.Female, p_l))}
te_f, te_m = te.copy(), te.copy(); te_f["Gender"] = "Female"; te_m["Gender"] = "Male"
dec = {}
for fam in FAMS:
    m_be, m_un, pred = models[fam]; mu_be, mu_un, mu_df = prices[fam]
    be_f, be_m = pred(m_be, te_f, cols_be), pred(m_be, te_m, cols_be)
    r = {"idi_refit": indices(mu_be, mu_un, mu_df, w_te)[1]}
    for nm, p in [("gbm", p_g), ("logit", p_l)]:
        mu_imp = p * be_f + (1 - p) * be_m
        r[f"idi_implied_{nm}"] = float(np.average(np.abs(mu_imp - mu_df), weights=w_te) / np.average(mu_df, weights=w_te))
        r[f"corr_implied_vs_un_{nm}"] = float(np.corrcoef(mu_imp, mu_un)[0, 1])
    # pure refitting component: distance between the refitted unawareness model and the implied unawareness price
    mu_imp = p_g * be_f + (1 - p_g) * be_m
    r["dist_un_vs_implied"] = float(np.average(np.abs(mu_un - mu_imp), weights=w_te) / np.average(mu_df, weights=w_te))
    r["gender_effect_be"] = float(np.average(np.abs(be_f - be_m), weights=w_te) / np.average(mu_df, weights=w_te))
    dec[fam] = r
    print(fam, {k: round(v, 5) for k, v in r.items()})
out["decomposition"] = dec

# ------------------------------------------------------------------ (b2) placebo: random binary variable in place of gender
plc = {fam: [] for fam in FAMS}
for j in range(10):
    rsj = np.random.RandomState(1000 + j)
    trp, tep = tr.copy(), te.copy()
    trp["Gender"] = np.where(rsj.rand(len(trp)) < pf, "Female", "Male"); tep["Gender"] = np.where(rsj.rand(len(tep)) < pf, "Female", "Male")
    ppf = (trp.Gender == "Female").mean()
    for fam in FAMS:
        m_be, m_un, pred = fit_pair(fam, trp)
        be, un, dfp = three_prices(fam, m_be, m_un, pred, tep, ppf)
        plc[fam].append(indices(be, un, dfp, w_te)[1])
    print("placebo draw", j + 1, {f: round(plc[f][-1], 5) for f in FAMS})
out["placebo"] = {fam: {"mean": float(np.mean(v)), "median": float(np.median(v)), "min": float(np.min(v)), "max": float(np.max(v)), "values": [float(x) for x in v]} for fam, v in plc.items()}

# ------------------------------------------------------------------ (b3) interaction-free GBM
def gbm_add(seed=SEED):
    return HistGradientBoostingRegressor(loss="poisson", max_iter=300, learning_rate=0.05, max_leaf_nodes=15, min_samples_leaf=50,
                                         l2_regularization=1.0, random_state=seed, interaction_cst=[[i] for i in range(len(cols_be))])
def gbm_add_un(seed=SEED):
    return HistGradientBoostingRegressor(loss="poisson", max_iter=300, learning_rate=0.05, max_leaf_nodes=15, min_samples_leaf=50,
                                         l2_regularization=1.0, random_state=seed, interaction_cst=[[i] for i in range(len(cols_un))])
w_tr = tr.Exposure.values
m_be_a = gbm_add().fit(gbm_frame(tr, cols_be), tr.freq, sample_weight=w_tr)
m_un_a = gbm_add_un().fit(gbm_frame(tr, cols_un), tr.freq, sample_weight=w_tr)
pred_a = lambda m, d, cols: m.predict(gbm_frame(d, cols))
be_a, un_a, df_a = three_prices("GBM", m_be_a, m_un_a, pred_a, te, pf)
ev_a = evaluate(te, un_a); ddi_a, idi_a = indices(be_a, un_a, df_a, w_te)
ra = un_a / df_a
out["gbm_no_interactions"] = {"idi": idi_a, "ddi": ddi_a, "deviance_un": ev_a["deviance"], "auc_un": ev_a["auc"], "ae_f": ev_a["ae_f"], "ae_m": ev_a["ae_m"],
                              "share5_f": float((np.abs(ra[fem] - 1) > 0.05).mean()), "share5_m": float((np.abs(ra[~fem] - 1) > 0.05).mean()),
                              "ratio_to_GLM": idi_a / dec["GLM"]["idi_refit"], "ratio_full_GBM_to_additive_GBM": dec["GBM"]["idi_refit"] / idi_a if idi_a > 0 else None}
print("GBM no interactions", out["gbm_no_interactions"])

# ------------------------------------------------------------------ (c) hyper-parameter sensitivity
kf = StratifiedKFold(5, shuffle=True, random_state=SEED)
def cv_dev(make_be, make_un, frame):
    dv = []
    for a, b in kf.split(tr, tr.ClaimInd):
        A, B = tr.iloc[a], tr.iloc[b]
        m = make_un().fit(frame(A, cols_un), A.freq, **({"sample_weight": A.Exposure.values} if frame is gbm_frame else {"m__sample_weight": A.Exposure.values}))
        mu = m.predict(frame(B, cols_un)); dv.append(mean_poisson_deviance(B.freq, mu, sample_weight=B.Exposure.values))
    return float(np.mean(dv))
ident = lambda d, cols: d[cols]
# GLM ridge grid
glm_grid = {}
for alpha in [0.0, 1e-5, 1e-4, 1e-3, 1e-2, 1e-1]:
    mk_be = lambda a=alpha: Pipeline([("pre", make_pre(CAT + ["Gender"], NUM)), ("m", PoissonRegressor(alpha=a, max_iter=5000))])
    mk_un = lambda a=alpha: Pipeline([("pre", make_pre(CAT, NUM)), ("m", PoissonRegressor(alpha=a, max_iter=5000))])
    cvd = cv_dev(mk_be, mk_un, ident)
    m_be = mk_be().fit(tr[cols_be], tr.freq, m__sample_weight=w_tr); m_un = mk_un().fit(tr[cols_un], tr.freq, m__sample_weight=w_tr)
    pred = lambda m, d, cols: m.predict(d[cols])
    be, un, dfp = three_prices("GLM", m_be, m_un, pred, te, pf); ddi, idi = indices(be, un, dfp, w_te); ev = evaluate(te, un)
    # out-of-fold recalibration slope for the UN model
    oof = np.zeros(len(tr))
    for a, b in kf.split(tr, tr.ClaimInd):
        mm = mk_un().fit(tr.iloc[a][cols_un], tr.iloc[a].freq, m__sample_weight=tr.iloc[a].Exposure.values); oof[b] = mm.predict(tr.iloc[b][cols_un])
    _, a0, b0 = recalibrate(oof, tr.freq.values, w_tr)
    glm_grid[str(alpha)] = {"cv_deviance_un": cvd, "holdout_deviance_un": ev["deviance"], "auc_un": ev["auc"], "idi": idi, "ddi": ddi, "recal_slope_un": b0,
                            "train_deviance_un": float(mean_poisson_deviance(tr.freq, m_un.predict(tr[cols_un]), sample_weight=w_tr))}
    print("GLM alpha", alpha, {k: round(v, 5) for k, v in glm_grid[str(alpha)].items()})
best_alpha = min(glm_grid, key=lambda k: glm_grid[k]["cv_deviance_un"])
out["glm_alpha_grid"] = {"grid": glm_grid, "cv_best_alpha": best_alpha, "main_alpha": "0.0001"}
# GBM grid
gbm_grid = {}
for it, lr, leaves in itertools.product([100, 300, 600], [0.03, 0.05, 0.1], [7, 15, 31]):
    mk = lambda it=it, lr=lr, lv=leaves: HistGradientBoostingRegressor(loss="poisson", max_iter=it, learning_rate=lr, max_leaf_nodes=lv, min_samples_leaf=50, l2_regularization=1.0, random_state=SEED)
    cvd = cv_dev(mk, mk, gbm_frame)
    m_be = mk().fit(gbm_frame(tr, cols_be), tr.freq, sample_weight=w_tr); m_un = mk().fit(gbm_frame(tr, cols_un), tr.freq, sample_weight=w_tr)
    be, un, dfp = three_prices("GBM", m_be, m_un, pred_a, te, pf); ddi, idi = indices(be, un, dfp, w_te); ev = evaluate(te, un)
    key = f"{it}|{lr}|{leaves}"
    gbm_grid[key] = {"max_iter": it, "learning_rate": lr, "max_leaf_nodes": leaves, "cv_deviance_un": cvd, "holdout_deviance_un": ev["deviance"], "auc_un": ev["auc"], "idi": idi, "ddi": ddi}
    print("GBM", key, round(cvd, 5), round(ev["deviance"], 5), round(idi, 5))
best_gbm = min(gbm_grid, key=lambda k: gbm_grid[k]["cv_deviance_un"])
# early stopping variant
mk_es = lambda: HistGradientBoostingRegressor(loss="poisson", max_iter=1000, learning_rate=0.05, max_leaf_nodes=15, min_samples_leaf=50, l2_regularization=1.0, random_state=SEED, early_stopping=True, validation_fraction=0.2, n_iter_no_change=20)
m_be = mk_es().fit(gbm_frame(tr, cols_be), tr.freq, sample_weight=w_tr); m_un = mk_es().fit(gbm_frame(tr, cols_un), tr.freq, sample_weight=w_tr)
be, un, dfp = three_prices("GBM", m_be, m_un, pred_a, te, pf); ddi, idi = indices(be, un, dfp, w_te); ev = evaluate(te, un)
out["gbm_grid"] = {"grid": gbm_grid, "cv_best": best_gbm, "main": "300|0.05|15",
                   "idi_range": [float(min(v["idi"] for v in gbm_grid.values())), float(max(v["idi"] for v in gbm_grid.values()))],
                   "n_configs": len(gbm_grid), "n_idi_above_glm_main": int(sum(v["idi"] > dec["GLM"]["idi_refit"] for v in gbm_grid.values())),
                   "early_stopping": {"n_iter_be": int(m_be.n_iter_), "n_iter_un": int(m_un.n_iter_), "idi": idi, "ddi": ddi, "holdout_deviance_un": ev["deviance"], "auc_un": ev["auc"]}}
print("GBM cv best", best_gbm, gbm_grid[best_gbm]); print("early stopping", out["gbm_grid"]["early_stopping"])

# ------------------------------------------------------------------ (a2) A/E across the 30 refitted splits
aes = {f: {"ae_f": [], "ae_m": []} for f in FAMS}
for k in range(30):
    seed = SEED + 1 + k
    tr_k, te_k = train_test_split(df, test_size=0.3, random_state=seed, stratify=df.ClaimInd)
    for f in FAMS:
        m_be, m_un, pred = fit_pair(f, tr_k, seed); un = pred(m_un, te_k, cols_un); a_f, a_m = ae_by_gender(te_k, un)
        aes[f]["ae_f"].append(a_f); aes[f]["ae_m"].append(a_m)
    print("refit A/E split", k + 1, {f: (round(aes[f]["ae_f"][-1], 3), round(aes[f]["ae_m"][-1], 3)) for f in FAMS})
out["ae_refit"] = {f: {"ae_f_mean": float(np.mean(v["ae_f"])), "ae_f_sd": float(np.std(v["ae_f"], ddof=1)), "ae_m_mean": float(np.mean(v["ae_m"])), "ae_m_sd": float(np.std(v["ae_m"], ddof=1)),
                       "n_f_below_1": int(sum(x < 1 for x in v["ae_f"])), "n_m_above_1": int(sum(x > 1 for x in v["ae_m"]))} for f, v in aes.items()}

# ------------------------------------------------------------------ (d) RSI: dominance and re-scoring perturbations
scores = {"IAIS": [1, 2, 2, 2, 2, 1], "NAIC": [2, 2, 3, 2, 2, 2], "EU": [3, 3, 3, 3, 3, 3]}
dims = ["Legal nature", "Scope & risk tiering", "Accountability & governance", "Fairness & non-discrimination", "Transparency & explainability", "Oversight & assurance"]
dom = {"NAIC_ge_IAIS_all_dims": all(n >= i for n, i in zip(scores["NAIC"], scores["IAIS"])), "EU_ge_NAIC_all_dims": all(e >= n for e, n in zip(scores["EU"], scores["NAIC"])),
       "gap_NAIC_IAIS": sum(scores["NAIC"]) - sum(scores["IAIS"]), "gap_EU_NAIC": sum(scores["EU"]) - sum(scores["NAIC"])}
cells = [(reg, d) for reg in scores for d in range(6)]
def total(sc): return {k: sum(v) for k, v in sc.items()}
def rank_ok(sc): t = total(sc); return t["EU"] > t["NAIC"] > t["IAIS"]
def perturb(sc, changes):
    s = {k: list(v) for k, v in sc.items()}
    for (reg, d, delta) in changes:
        s[reg][d] = s[reg][d] + delta
        if not 0 <= s[reg][d] <= 3: return None
    return s
moves = [(reg, d, delta) for (reg, d) in cells for delta in (-1, 1)]
res = {}
for k in (1, 2, 3):
    n_valid = n_ok = n_tie = 0
    for combo in itertools.combinations(moves, k):
        if len({(r, d) for r, d, _ in combo}) < k: continue
        s = perturb(scores, combo)
        if s is None: continue
        n_valid += 1; t = total(s)
        if rank_ok(s): n_ok += 1
        elif t["EU"] >= t["NAIC"] >= t["IAIS"]: n_tie += 1
    res[f"k{k}"] = {"n_valid": n_valid, "n_strict_order_preserved": n_ok, "share_preserved": n_ok / n_valid, "n_tie_only": n_tie, "n_reversal": n_valid - n_ok - n_tie}
    print("perturb", k, res[f"k{k}"])
# adversarial scenario: every contestable score moved one point against the ranking
adv = {"IAIS": [2, 2, 3, 3, 2, 2], "NAIC": [1, 2, 2, 2, 2, 2], "EU": [3, 2, 3, 3, 3, 3]}
adv_t = total(adv)
out["rsi_rescoring"] = {"dominance": dom, "perturbations": res, "adversarial_scenario": {"scores": adv, "totals": adv_t, "note": "IAIS +1 on legal nature (binding via ICP adoption), accountability, fairness (paras 86-91) and oversight; NAIC -1 on legal nature and accountability; EU -1 on scope (horizontal, not insurance-specific). Ranking under this scenario: " + " > ".join(sorted(adv_t, key=lambda k: -adv_t[k]))},
                        "min_one_point_changes_to_tie_NAIC_IAIS": dom["gap_NAIC_IAIS"], "min_one_point_changes_to_reverse_NAIC_IAIS": dom["gap_NAIC_IAIS"] + 1,
                        "min_one_point_changes_to_tie_EU_NAIC": dom["gap_EU_NAIC"]}

# ------------------------------------------------------------------ (b4) uncertainty of the implied index: hold-out bootstrap (models fixed) and 30 refitted splits
rs = np.random.RandomState(SEED + 7); imp_boot = {f: [] for f in FAMS}; imp_store = {}
for fam in FAMS:
    m_be, m_un, pred = models[fam]; mu_df = prices[fam][2]
    be_f, be_m = pred(m_be, te_f, cols_be), pred(m_be, te_m, cols_be); imp_store[fam] = (p_g * be_f + (1 - p_g) * be_m, mu_df)
for _ in range(NB):
    ix = rs.choice(n, n, replace=True); w = w_te[ix]
    for fam in FAMS:
        mi, md = imp_store[fam]; imp_boot[fam].append(float(np.average(np.abs(mi[ix] - md[ix]), weights=w) / np.average(md[ix], weights=w)))
for fam in FAMS: dec[fam]["idi_implied_gbm_ci"] = [float(np.percentile(imp_boot[fam], 2.5)), float(np.percentile(imp_boot[fam], 97.5))]
imp_refit = {f: [] for f in FAMS}; ref_refit = {f: [] for f in FAMS}
for k in range(30):
    seed = SEED + 1 + k
    tr_k, te_k = train_test_split(df, test_size=0.3, random_state=seed, stratify=df.ClaimInd); pfk = tr_k.Female.mean(); wk = te_k.Exposure.values
    ck = HistGradientBoostingClassifier(max_iter=300, learning_rate=0.05, max_leaf_nodes=15, random_state=seed).fit(gbm_frame(tr_k, X_COLS), tr_k.Female)
    pk = ck.predict_proba(gbm_frame(te_k, X_COLS))[:, 1]
    tkf, tkm = te_k.copy(), te_k.copy(); tkf["Gender"] = "Female"; tkm["Gender"] = "Male"
    for f in FAMS:
        m_be, m_un, pred = fit_pair(f, tr_k, seed)
        bf, bm = pred(m_be, tkf, cols_be), pred(m_be, tkm, cols_be); mdf = pfk * bf + (1 - pfk) * bm; mun = pred(m_un, te_k, cols_un)
        mimp = pk * bf + (1 - pk) * bm
        imp_refit[f].append(float(np.average(np.abs(mimp - mdf), weights=wk) / np.average(mdf, weights=wk)))
        ref_refit[f].append(float(np.average(np.abs(mun - mdf), weights=wk) / np.average(mdf, weights=wk)))
    print("implied refit split", k + 1, {f: (round(imp_refit[f][-1], 4), round(ref_refit[f][-1], 4)) for f in FAMS})
summ = lambda v: {"mean": float(np.mean(v)), "sd": float(np.std(v, ddof=1)), "median": float(np.median(v)), "min": float(np.min(v)), "max": float(np.max(v))}
out["implied_refit"] = {f: {"implied": summ(imp_refit[f]), "refit_based": summ(ref_refit[f]), "values_implied": imp_refit[f], "values_refit": ref_refit[f],
                            "n_splits_implied_GBM_below_GLM": None} for f in FAMS}
gi, bi = np.array(imp_refit["GLM"]), np.array(imp_refit["GBM"])
out["implied_refit"]["n_splits_GBM_implied_below_GLM_implied"] = int((bi < gi).sum())
out["implied_refit"]["ratio_GBM_GLM_implied"] = summ(bi / gi)
out["implied_refit"]["n_splits_refit_GBM_exceeds_implied_GBM_by_factor2"] = int((np.array(ref_refit["GBM"]) > 2 * bi).sum())
out["total_runtime_seconds_step6"] = round(time.time() - T0, 1)
json.dump(out, open(f"{HERE}/results_review.json", "w"), indent=2)
print("implied refit", {f: out["implied_refit"][f]["implied"] for f in FAMS}, out["implied_refit"]["ratio_GBM_GLM_implied"])
print("done in", out["total_runtime_seconds_step6"], "s")
