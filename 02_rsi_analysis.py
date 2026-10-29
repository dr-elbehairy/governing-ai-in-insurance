"""Step 2 — Regulatory Stringency Index: reconciled scores, weight-sensitivity analysis, and (when available) inter-rater kappa.
Output: results_rsi.json. Kappa is read from rsi_kappa.json produced by rsi_kappa.py once the second coder's sheet is filled."""
import os, json, numpy as np
HERE = os.path.dirname(os.path.abspath(__file__))
dims = ["Legal nature","Scope & risk tiering","Accountability & governance","Fairness & non-discrimination",
        "Transparency & explainability","Oversight & assurance"]
scores = {"IAIS": [1,2,2,2,2,1], "NAIC": [2,2,3,2,2,2], "EU": [3,3,3,3,3,3]}
schemes = {"Equal weights": [1,1,1,1,1,1], "Substantive x1.5": [1,1,1.5,1.5,1.5,1.5], "Form x1.5": [1.5,1.5,1,1,1,1],
           "Fairness & oversight x2": [1,1,1,2,1,2], "Legal nature excluded": [0,1,1,1,1,1]}
sens = {}
for nm, w in schemes.items():
    w = np.array(w, float); mx = 3 * w.sum()
    tot = {k: float(np.dot(w, v)) for k, v in scores.items()}
    norm = {k: float(100 * tot[k] / mx) for k in tot}
    sens[nm] = {"weights": list(w), "max": float(mx), "totals": tot, "pct_of_max": norm,
                "ranking": sorted(tot, key=lambda k: -tot[k]),
                "gap_NAIC_IAIS_pct": float(norm["NAIC"] - norm["IAIS"]), "gap_EU_NAIC_pct": float(norm["EU"] - norm["NAIC"])}
out = {"dimensions": dims, "scores": scores, "schemes": sens}
kp = f"{HERE}/rsi_kappa.json"  # optional; not used in v5
out["kappa"] = json.load(open(kp)) if os.path.exists(kp) else None
json.dump(out, open(f"{HERE}/results_rsi.json", "w"), indent=2)
for nm, s in sens.items(): print(nm, s["ranking"], {k: round(v, 1) for k, v in s["pct_of_max"].items()})
print("kappa:", "available" if out["kappa"] else "not yet computed (second coder sheet pending)")
