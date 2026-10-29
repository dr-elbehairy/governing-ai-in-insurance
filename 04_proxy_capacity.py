"""Step 4 — proxy capacity: how well do the non-protected rating factors predict gender? 5-fold CV AUC (logistic, GBM)
and permutation importance. Output: results_proxy.json"""
import os, json, numpy as np, pandas as pd, warnings
from sklearn.model_selection import StratifiedKFold
from sklearn.pipeline import Pipeline
from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import HistGradientBoostingClassifier
from sklearn.metrics import roc_auc_score
from sklearn.inspection import permutation_importance
from common import *
warnings.filterwarnings("ignore")
df = load_prepared()
cv = StratifiedKFold(5, shuffle=True, random_state=SEED); proxy = {}
for nm, mdl, Z in [("logit", Pipeline([("pre", make_pre(CAT, NUM)), ("m", LogisticRegression(max_iter=3000))]), df[X_COLS]),
                   ("gbm", HistGradientBoostingClassifier(max_iter=300, learning_rate=0.05, max_leaf_nodes=15, random_state=SEED), gbm_frame(df, X_COLS))]:
    aucs = []
    for a, b in cv.split(Z, df.Female):
        mdl.fit(Z.iloc[a], df.Female.iloc[a]); aucs.append(roc_auc_score(df.Female.iloc[b], mdl.predict_proba(Z.iloc[b])[:, 1]))
    proxy[nm] = (float(np.mean(aucs)), float(np.std(aucs)))
Zall = gbm_frame(df, X_COLS)
clf = HistGradientBoostingClassifier(max_iter=300, learning_rate=0.05, max_leaf_nodes=15, random_state=SEED).fit(Zall, df.Female)
pi = permutation_importance(clf, Zall, df.Female, scoring="roc_auc", n_repeats=5, random_state=SEED)
imp = pd.Series(pi.importances_mean, index=X_COLS).sort_values(ascending=False)
json.dump({"proxy_auc": proxy, "proxy_importance": imp.round(4).to_dict()}, open(f"{HERE}/results_proxy.json", "w"), indent=2)
print(proxy); print(imp.head(8))
