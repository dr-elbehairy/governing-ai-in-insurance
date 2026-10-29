"""Runs the numbered pipeline in order and merges the partial result files into results.json (consumed by figures.py and build_manuscript.py)."""
import os, json, subprocess, sys
HERE = os.path.dirname(os.path.abspath(__file__)); PY = sys.executable
steps = ["01_data_preparation.py", "02_rsi_analysis.py", "03_glm_gam_gbm.py", "04_proxy_capacity.py", "05_bootstrap.py", "06_reviewer_checks.py"]
if "--only" in sys.argv: steps = [s for s in steps if s.startswith(sys.argv[sys.argv.index("--only") + 1])]
if "--merge-only" in sys.argv: steps = []
for s in steps:
    print("=" * 20, s); subprocess.run([PY, f"{HERE}/{s}"], check=True, cwd=HERE)
R = {}
R.update(json.load(open(f"{HERE}/results_data.json")))
R.update(json.load(open(f"{HERE}/results_models.json")))
R.update(json.load(open(f"{HERE}/results_proxy.json")))
B = json.load(open(f"{HERE}/results_bootstrap.json"))
R["bootstrap"] = B["holdout"]; R["refit"] = B["refit"]; R["calibration"] = B["calibration"]
R["rsi_sensitivity"] = json.load(open(f"{HERE}/results_rsi.json"))
R["review"] = json.load(open(f"{HERE}/results_review.json"))
json.dump(R, open(f"{HERE}/results.json", "w"), indent=2); print("merged -> results.json")
