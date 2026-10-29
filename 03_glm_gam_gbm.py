"""Step 3 — fit GLM / additive spline GLM (GAM) / GBM Poisson frequency models; best-estimate, unawareness and
discrimination-free prices on the 30% hold-out; performance, calibration and discrimination indices.
Output: results_models.json, data/test_predictions.csv"""
import os, json, numpy as np, pandas as pd, warnings
from sklearn.model_selection import train_test_split
from common import *
warnings.filterwarnings("ignore")
df = load_prepared()
tr, te = train_test_split(df, test_size=0.3, random_state=SEED, stratify=df.ClaimInd)
p_female = tr.Female.mean(); w_te = te.Exposure.values
res = {"n_train": int(len(tr)), "n_test": int(len(te)), "p_female_train": float(p_female)}
te = te.copy()
for fam in FAMS:
    m_be, m_un, pred = fit_pair(fam, tr)
    mu_be, mu_un, mu_df = three_prices(fam, m_be, m_un, pred, te, p_female)
    res[f"{fam}_BE"] = evaluate(te, mu_be); res[f"{fam}_UN"] = evaluate(te, mu_un); res[f"{fam}_DF"] = evaluate(te, mu_df)
    ddi, idi = indices(mu_be, mu_un, mu_df, w_te)
    res[f"{fam}_BE"]["direct_disc_index"] = ddi
    res[f"{fam}_UN"]["indirect_disc_index"] = idi
    res[f"{fam}_UN"]["corr_un_df"] = float(np.corrcoef(mu_un, mu_df)[0, 1])
    te[f"ratio_{fam}"] = mu_un / mu_df; te[f"mu_be_{fam}"] = mu_be; te[f"mu_un_{fam}"] = mu_un; te[f"mu_df_{fam}"] = mu_df
    r = te[f"ratio_{fam}"]
    res[f"{fam}_ratio_stats"] = {g: dict(mean=float(r[te.Gender == g].mean()), p10=float(r[te.Gender == g].quantile(.1)),
                                         p90=float(r[te.Gender == g].quantile(.9)),
                                         share_over_5pct=float((abs(r[te.Gender == g] - 1) > 0.05).mean())) for g in ["Female", "Male"]}
json.dump(res, open(f"{HERE}/results_models.json", "w"), indent=2)
te.to_csv(f"{HERE}/data/test_predictions.csv", index=False)
pd.set_option("display.width", 220)
print(pd.DataFrame({k: v for k, v in res.items() if isinstance(v, dict) and "deviance" in v}).T.round(4))
