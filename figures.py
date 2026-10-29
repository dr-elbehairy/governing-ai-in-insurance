import os, json, numpy as np, pandas as pd, matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import rcParams
rcParams["font.family"] = "serif"; rcParams["font.serif"] = ["Times New Roman", "DejaVu Serif"]; rcParams["font.size"] = 10
HERE = os.path.dirname(os.path.abspath(__file__)); FIG = os.path.join(os.path.dirname(HERE), "figures"); os.makedirs(FIG, exist_ok=True)
NAVY, PETROL, GOLD, GREY = "#1C4076", "#007A88", "#B8860B", "#555555"
r = json.load(open(f"{HERE}/results.json"))

# ---------- Figure 2 (in text): Regulatory Stringency Index radar ----------
dims = r["rsi_sensitivity"]["dimensions"]; sc = r["rsi_sensitivity"]["scores"]
scores = {"IAIS (2025)": sc["IAIS"], "NAIC (2023)": sc["NAIC"], "EU AI Act (2024)": sc["EU"]}
ang = np.linspace(0, 2*np.pi, len(dims), endpoint=False).tolist(); ang += ang[:1]
fig = plt.figure(figsize=(6.2, 4.6)); ax = plt.subplot(111, polar=True)
for (k, v), c, ls in zip(scores.items(), [PETROL, GOLD, NAVY], ["-", "--", "-."]):
    vv = v + v[:1]; ax.plot(ang, vv, ls, color=c, lw=1.8, label=k); ax.fill(ang, vv, color=c, alpha=0.08)
ax.set_xticks(ang[:-1]); ax.set_xticklabels([d.replace(" & ", " &\n") for d in dims], fontsize=8.5)
ax.set_yticks([0, 1, 2, 3]); ax.set_yticklabels(["0", "1", "2", "3"], fontsize=7); ax.set_ylim(0, 3)
ax.legend(loc="lower center", bbox_to_anchor=(0.5, -0.22), ncol=3, fontsize=8, frameon=False)
plt.tight_layout(); plt.savefig(f"{FIG}/fig1_rsi_radar.png", dpi=300, bbox_inches="tight"); plt.close()

# ---------- Figure 1 (in text): timeline of instruments ----------
events = [("2023-12", "NAIC Model\nBulletin", GOLD), ("2023-11", "Colorado\nReg. 10-1-1", GREY), ("2024-07", "NYDFS\nCircular 7", GREY),
          ("2024-08", "EU AI Act\nin force", NAVY), ("2024-11", "IAIS draft\nAP", PETROL), ("2025-07", "IAIS final AP\n(2 Jul)", PETROL),
          ("2025-08", "EIOPA\nOpinion", NAVY), ("2025-10", "   Colorado\n   auto/health", GREY), ("2026-06", "Texas\nB-0003-26", GREY),
          ("2026-07", "Digital\nOmnibus", NAVY), ("2027-12", "Annex III\nhigh-risk", NAVY), ("2028-08", "Annex I\nhigh-risk", NAVY)]
fig, ax = plt.subplots(figsize=(7.4, 3.6))
xs = [pd.Timestamp(e[0] + "-01") for e in events]
ax.hlines(0, pd.Timestamp("2023-09-01"), pd.Timestamp("2028-10-01"), color=GREY, lw=1)
for i, (e, x) in enumerate(zip(events, xs)):
    lev = [1.0, -1.0, 0.55, -0.55][i % 4]
    ax.vlines(x, 0, lev*0.6, color=e[2], lw=1.4); ax.plot(x, 0, "o", color=e[2], ms=6)
    ax.text(x, lev*0.7, e[1], ha="center", va="bottom" if lev > 0 else "top", fontsize=9, color=e[2], fontweight="bold")
ax.axvline(pd.Timestamp("2026-10-29"), color="k", ls=":", lw=1.0); ax.text(pd.Timestamp("2026-10-29"), 1.55, "date of writing", ha="center", fontsize=9)
ax.set_ylim(-1.7, 1.8); ax.set_yticks([]); ax.spines[["left", "right", "top"]].set_visible(False); ax.tick_params(axis="x", labelsize=9)
ax.set_xlim(pd.Timestamp("2023-09-01"), pd.Timestamp("2028-10-01"))
plt.tight_layout(); plt.savefig(f"{FIG}/fig2_timeline.png", dpi=300, bbox_inches="tight"); plt.close()

# ---------- Figure 3: proxy recovery ----------
imp = pd.Series(r["proxy_importance"]).sort_values(ascending=True)
fig, axes = plt.subplots(1, 2, figsize=(7.2, 3.2), gridspec_kw={"width_ratios": [1, 1.6]})
a = axes[0]; names = ["Logistic", "Gradient\nboosting"]; vals = [r["proxy_auc"]["logit"][0], r["proxy_auc"]["gbm"][0]]
errs = [r["proxy_auc"]["logit"][1], r["proxy_auc"]["gbm"][1]]
a.bar(names, vals, yerr=errs, color=[PETROL, NAVY], width=0.55, capsize=3); a.axhline(0.5, color=GREY, ls="--", lw=1)
a.set_ylim(0.4, 0.85); a.set_ylabel("AUC for predicting gender\nfrom non-protected features")
for i, v in enumerate(vals): a.text(i, v + 0.01, f"{v:.3f}", ha="center", fontsize=8)
a.set_title("(a) Proxy recovery", fontsize=9)
b = axes[1]; imp.tail(10).plot.barh(ax=b, color=NAVY); b.set_xlabel("Permutation importance (ΔAUC)"); b.set_title("(b) Features carrying the proxy signal", fontsize=9)
plt.tight_layout(); plt.savefig(f"{FIG}/fig3_proxy.png", dpi=300, bbox_inches="tight"); plt.close()

# ---------- Figure 4: unawareness / discrimination-free ratio distributions, three model classes ----------
te = pd.read_csv(f"{HERE}/data/test_predictions.csv")
fig, axes = plt.subplots(1, 3, figsize=(7.4, 2.9), sharey=True)
for ax, fam, t in zip(axes, ["GLM", "GAM", "GBM"], ["(a) Poisson GLM", "(b) Spline GLM (GAM)", "(c) Gradient boosting"]):
    for g, c in [("Female", GOLD), ("Male", NAVY)]:
        ax.hist(te.loc[te.Gender == g, f"ratio_{fam}"], bins=60, range=(0.85, 1.15), density=True, alpha=0.55, color=c, label=g)
    ax.axvline(1, color="k", lw=0.8); ax.set_title(t, fontsize=9)
    idi = r[f"{fam}_UN"]["indirect_disc_index"]; lo, hi = r["bootstrap"][fam]["idi"]
    ax.text(0.03, 0.95, f"IDI {100*idi:.2f}%\n[{100*lo:.2f}, {100*hi:.2f}]", transform=ax.transAxes, va="top", fontsize=7.5)
axes[1].set_xlabel("Unawareness price ÷ discrimination-free price")
axes[0].set_ylabel("Density"); axes[0].legend(frameon=False, fontsize=8, loc="upper right")
plt.tight_layout(); plt.savefig(f"{FIG}/fig4_ratio.png", dpi=300, bbox_inches="tight"); plt.close()

# ---------- Figure 5: RSI re-scoring (coding perturbation) analysis ----------
import itertools
RV = json.load(open(f"{HERE}/results_review.json"))
scores = {"IAIS": [1,2,2,2,2,1], "NAIC": [2,2,3,2,2,2], "EU": [3,3,3,3,3,3]}
cells = [(reg, d) for reg in scores for d in range(6)]; moves = [(reg, d, dl) for (reg, d) in cells for dl in (-1, 1)]
def tot_after(combo):
    s_ = {k: list(v) for k, v in scores.items()}
    for reg, d, dl in combo:
        s_[reg][d] += dl
        if not 0 <= s_[reg][d] <= 3: return None
    return {k: sum(v) for k, v in s_.items()}
gaps = {1: [], 2: [], 3: []}
for k in (1, 2, 3):
    for combo in itertools.combinations(moves, k):
        if len({(r, d) for r, d, _ in combo}) < k: continue
        t = tot_after(combo)
        if t: gaps[k].append((t["NAIC"] - t["IAIS"], t["EU"] - t["NAIC"]))
fig, axes = plt.subplots(1, 2, figsize=(7.2, 2.9))
for ax, j, ttl, base in [(axes[0], 0, "(a) NAIC minus IAIS total", 3), (axes[1], 1, "(b) EU minus NAIC total", 5)]:
    for k, c, off in [(1, PETROL, -0.25), (2, GOLD, 0.0), (3, NAVY, 0.25)]:
        v = np.array([g[j] for g in gaps[k]]); vals, cnt = np.unique(v, return_counts=True)
        ax.bar(vals + off, 100 * cnt / len(v), 0.25, color=c, label=f"{k} re-scored cell{'s' if k > 1 else ''}")
    ax.axvline(0, color="k", lw=0.8, ls="--"); ax.axvline(base, color=GREY, lw=0.8, ls=":")
    ax.set_title(ttl, fontsize=9); ax.set_xlabel("Point gap after re-scoring"); ax.spines[["right", "top"]].set_visible(False)
axes[0].set_ylabel("Share of perturbations, %"); axes[0].legend(frameon=False, fontsize=7.5)
plt.tight_layout(); plt.savefig(f"{FIG}/fig5_rsi_rescoring.png", dpi=300, bbox_inches="tight"); plt.close()
print("figures done ->", FIG)
