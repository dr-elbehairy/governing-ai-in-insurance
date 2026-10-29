# Governing AI in Insurance — replication package (v1.0.0)

Code, derived data and results to reproduce the empirical component of

> Elbehairy, M., & Sharkawy, M. (2026). *Governing AI in Insurance: Comparing the IAIS, NAIC and EU AI Act Frameworks and Testing Proxy Discrimination in Motor Pricing.* Manuscript, Department of Statistics and Insurance, Faculty of Commerce, Suez Canal University.

## What the package reproduces

| Script | Purpose | Output |
|---|---|---|
| `01_data_preparation.py` | Loads the CASdatasets `freMPL1` export, drops zero-exposure rows, collapses the 37 socio-professional codes to 6 groups, builds the analysis frame. | `data/freMPL1_prepared.csv`, `results_data.json` |
| `02_rsi_analysis.py` | Regulatory Stringency Index: reconciled scores and the (descriptive) alternative weighting schemes. | `results_rsi.json` |
| `03_glm_gam_gbm.py` | Poisson GLM / additive spline GLM (GAM) / histogram GBM; best-estimate, unawareness and discrimination-free prices on the 30 % hold-out; deviance, AUC, A/E by gender, DDI and IDI (Table 6). | `results_models.json`, `data/test_predictions.csv` |
| `04_proxy_capacity.py` | 5-fold CV AUC for predicting gender from the 16 non-protected factors and permutation importance (Figure 3). | `results_proxy.json` |
| `05_bootstrap.py` | 1,000-resample hold-out bootstrap; 30 repeated stratified splits with all models refitted; recalibration check (Table 9). | `results_bootstrap.json` |
| `06_reviewer_checks.py` | Decomposition of the unawareness test: implied unawareness price with the best-estimate model held fixed (IDI*), placebo attribute, interaction-free GBM (Table 11); GLM penalty and GBM hyper-parameter grids (Table 10); A/E bootstrap intervals, training A/E and refit means; re-scoring analysis of the index (Table 8, Figure 5); runtime. | `results_review.json` |
| `run_all.py` | Runs 01–06 in order and merges the partial files into `results.json`. `--only 06` reruns one step; `--merge-only` just merges. | `results.json` |
| `figures.py` | Figures 1–5 (300 dpi PNG). | `figures/` |
| `build_manuscript.py` | Regenerates the manuscript .docx from `results.json`; no number in the paper is typed by hand. | — |
| `common.py` | Shared model definitions, fixed-level categorical encoding, price and index functions. | — |

`results/` holds the JSON outputs that the manuscript was built from, so every number in the paper can be traced without rerunning anything. `RSI_scoring_sheet.xlsx` holds the six-dimension coding of the three regulatory instruments with the anchoring text passages (Table 4).

## Data

`freMPL1` (French motor third-party liability, 2004; 20,600 policies) from the open **CASdatasets** collection (Dutang & Charpentier, 2020, http://cas.uqam.ca). The CSV export is included in `data/` for convenience and is redistributed with attribution under the CASdatasets licence; the authors of this package claim no rights in it.

## Method summary

* Protected attribute D = gender; X = 16 non-protected rating factors. 70/30 stratified split (seed 2026). Poisson deviance with exposure offset.
* GLM: one-hot + linear numerics, ridge α = 1e-4 (grid 0–0.1 in step 06). GAM: cubic B-splines (8 knots) on LicAge, DrivAge, BonusMalus, RiskVar; no interactions. GBM: histogram gradient boosting, 300 trees, lr 0.05, ≤ 15 leaves, ≥ 50 obs/leaf, L2 = 1 (27-point grid and early stopping in step 06).
* Categorical factors, including the counterfactual gender assignments, are encoded with fixed category levels drawn from the full data (`common.gbm_frame`).
* Discrimination-free price per Lindholm, Richman, Tsanakas & Wüthrich (2022): μ_DF(x) = Σ_d μ_BE(x,d) P(D = d).
* Implied unawareness price (step 06): μ_UN*(x) = Σ_d μ_BE(x,d) P̂(D = d | X = x), with P̂ from a gradient-boosting gender classifier fitted on the training set; IDI* = E|μ_UN* − μ_DF| / E[μ_DF]. Placebo: gender replaced by an independent Bernoulli variable of the same prevalence and the full drop-and-refit procedure repeated (10 seeds).

## Environment

```bash
pip install -r requirements.txt   # Python 3.11, scikit-learn 1.3.0, pandas 2.0.3, numpy 1.24
python run_all.py                 # ≈ 8 min on an Apple M1 Pro laptop (steps 01–06)
python figures.py
```

## Citation and licence

Code: MIT (see `LICENSE`). Data: CASdatasets licence, see above. Please cite the manuscript above when using this package.
