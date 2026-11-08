"""
Northbridge Lending: README / site charts.

Input : data_clean/powerbi/*.csv (scripts/export_powerbi.py), data_clean/northbridge.db, model outputs
Output: images/chart_*.png and a smaller copy in site/img/
Run   : python scripts/make_charts.py   (from the project root)
"""
from pathlib import Path
import json
import sqlite3
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import PercentFormatter
import numpy as np
import pandas as pd
from sklearn.metrics import roc_curve

ROOT = Path(__file__).resolve().parents[1]
P = ROOT / "data_clean" / "powerbi"
IMG, SITE = ROOT / "images", ROOT / "site" / "img"
IMG.mkdir(exist_ok=True)
SITE.mkdir(parents=True, exist_ok=True)
NAVY, PRIMARY, COMPARE, GOLD, GOOD, WARN, BAD, SLATE = ("#0B1F3A", "#1F4E79", "#B8C2CF", "#C8A24A", "#0E7C86",
                                                        "#D08C2E", "#B5523B", "#64748B")
plt.rcParams.update({"font.family": "Segoe UI", "font.size": 10, "axes.spines.top": False,
                     "axes.spines.right": False, "axes.titleweight": "bold", "axes.titlesize": 12,
                     "axes.titlelocation": "left", "axes.edgecolor": "#CBD5E1"})
pct = PercentFormatter(1.0, decimals=0)


def save(fig, name):
    fig.tight_layout()
    fig.savefig(IMG / name, dpi=150)
    fig.savefig(SITE / name, dpi=110)
    plt.close(fig)
    print(f"wrote images/{name}")


loan = pd.read_csv(P / "dim_loan.csv")
fnd = pd.read_csv(P / "findings.csv").set_index("key").value
MM = json.loads((ROOT / "data_clean" / "model_metrics.json").read_text())

# 1. grade: default rate and price -------------------------------------------------------------
old = loan[loan.issue_month <= "2023-12"]
g = old.groupby("grade").agg(d24=("default_24m", "mean"), rate=("int_rate", "mean"), n=("loan_id", "size"))
net = old.groupby("grade").apply(lambda d: d.net_loss.sum() / d.loan_amnt.sum(), include_groups=False)
fig, ax = plt.subplots(figsize=(9, 4))
x = np.arange(len(g))
ax.bar(x - 0.2, g.d24, 0.4, color=PRIMARY, label="Defaulted within 24 months")
ax.bar(x + 0.2, net, 0.4, color=COMPARE, label="Net loss so far (% of amount lent)")
for i, (d, r) in enumerate(zip(g.d24, g.rate)):
    ax.text(i - 0.2, d + 0.004, f"{d:.1%}", ha="center", fontsize=8.5, color=NAVY)
ax.set_xticks(x, [f"{gr}\n{r:.1f}% rate" for gr, r in zip(g.index, g.rate)])
ax.yaxis.set_major_formatter(pct)
ax.legend(frameon=False, loc="upper left")
ax.set_title(f"The lender's grade ranks risk: {g.d24['A']:.1%} (A) to {g.d24['G']:.1%} (G) default within 24 months")
ax.text(0, -0.25, "Loans issued 2021-2023 (all have 24 months of history). Average interest rate under each grade.",
        transform=ax.transAxes, fontsize=8, color=SLATE)
save(fig, "chart_grade.png")

# 2. vintages and unemployment -------------------------------------------------------------------
vin = pd.read_csv(P / "fact_vintage.csv")
# cut each year where its youngest quarter stops being observed, so the mix of loans stays constant along the curve
cap = vin[vin.observed == 1].groupby(["vintage_year", "vintage_q"]).mob.max().groupby("vintage_year").min()
vy = vin[vin.mob <= vin.vintage_year.map(cap)].groupby(["vintage_year", "mob"]).agg(d=("cum_defaults", "sum"), n=("loans", "sum")).reset_index()
vy["rate"] = vy.d / vy.n
flows = pd.read_csv(P / "fact_month_flows.csv").groupby("month")[["new_defaults", "performing_loans"]].sum()
mon = pd.read_csv(P / "dim_month.csv").set_index("month")
q = flows.join(mon).loc["2021-07":]
q["q"] = q.quarter
qq = q.groupby("q").agg(d=("new_defaults", "sum"), n=("performing_loans", "sum"), u=("unemployment_rate", "mean"))
fig, (a1, a2) = plt.subplots(1, 2, figsize=(11.5, 4.1), gridspec_kw={"width_ratios": [1, 1.25]})
for (y_, c) in zip([2021, 2022, 2023, 2024], [COMPARE, "#7A8CA5", BAD, GOLD]):
    s = vy[vy.vintage_year == y_]
    a1.plot(s.mob, s.rate, color=c, lw=2.4, label=str(y_))
    a1.text(s.mob.iloc[-1] + 0.5, s.rate.iloc[-1], f" {y_}", va="center", color=c, fontsize=9, fontweight="bold")
a1.yaxis.set_major_formatter(pct)
a1.set_xlabel("Months on book")
a1.set_xlim(0, 52)
a1.set_title("Cumulative default rate by year issued")
a2.bar(range(len(qq)), 1000 * qq.d / qq.n, color=PRIMARY, width=0.7)
a2.set_ylabel("New defaults per 1,000 performing loans", color=PRIMARY, fontsize=9)
a3 = a2.twinx()
a3.plot(range(len(qq)), qq.u, color=GOLD, lw=2.6, marker="o", ms=4)
a3.set_ylabel("Unemployment %", color=GOLD, fontsize=9)
a3.spines["right"].set_visible(True)
a3.set_ylim(3, 6.5)
a2.set_xticks(range(len(qq)), [s.replace(" Q", "\nQ") for s in qq.index], fontsize=7.5)
a2.set_title("Defaults follow unemployment")
fig.text(0.01, 0.01, f"2023 loans met the 2024 unemployment peak early in life, when default risk is highest. "
         f"Estimated: +1 pt unemployment = +{np.exp(fnd['unemployment_beta']) - 1:.0%} default starts (true: +35%).",
         fontsize=8, color=SLATE)
save(fig, "chart_vintage.png")

# 3. scorecard: ROC and calibration --------------------------------------------------------------
oot = loan[loan.vintage_year == 2023]
y = oot.default_24m.astype(int)
fig, (a1, a2) = plt.subplots(1, 2, figsize=(11.5, 4.2))
for name, s, c in [("Scorecard", oot.pd_24m, PRIMARY), ("Lender grade", oot.grade.map({g_: i for i, g_ in enumerate("ABCDEFG")}), GOLD)]:
    fpr, tpr, _ = roc_curve(y, s)
    key = "scorecard" if name == "Scorecard" else "lender_grade"
    a1.plot(fpr, tpr, color=c, lw=2.4, label=f"{name}: Gini {MM[key]['oot']['gini']:.2f}")
a1.plot([0, 1], [0, 1], color="#CBD5E1", ls="--", lw=1)
a1.set_xlabel("Share of good loans flagged")
a1.set_ylabel("Share of defaults caught")
a1.legend(frameon=False, loc="lower right")
a1.set_title("Out-of-time test (2023 loans): ROC")
b = pd.read_csv(P / "score_bands.csv")
xb = np.arange(len(b))
a2.bar(xb - 0.2, b.predicted_pd_pct / 100, 0.4, color=COMPARE, label="Predicted")
a2.bar(xb + 0.2, b.actual_default_pct / 100, 0.4, color=PRIMARY, label="Actual")
a2.set_xticks(xb, [f"{i}\n{r}" for i, r in zip(b.score_band, b.score_range)], fontsize=7.5)
a2.yaxis.set_major_formatter(pct)
a2.legend(frameon=False, loc="upper left")
a2.set_title(f"Default within 24 months by score band (1 = best): {MM['band_spread_oot']:.1f}x spread")
fig.text(0.01, 0.01, f"Trained on 2021-22 loans. On 2023 loans it predicted {MM['calibration_oot']['predicted_pct']:.1f}% and "
         f"{MM['calibration_oot']['actual_pct']:.1f}% defaulted: unemployment rose after they were issued, which an "
         "application score cannot know (the IFRS 9 model adds it).", fontsize=8, color=SLATE)
save(fig, "chart_scorecard.png")

# 4. IFRS 9: ECL by stage and the backtest --------------------------------------------------------
ecl = pd.read_csv(P / "fact_ecl.csv")
s25 = ecl[ecl["asof"] == "2025-12"].groupby("stage").agg(ead=("ead", "sum"), ecl=("ecl", "sum"), n=("loan_id", "size"))
bt = pd.read_csv(P / "ecl_backtest.csv")
fig, (a1, a2) = plt.subplots(1, 2, figsize=(11.5, 4.1), gridspec_kw={"width_ratios": [1, 1.3]})
cols = [PRIMARY, WARN, BAD]
a1.bar(s25.index, s25.ecl / 1e6, color=cols)
for i, (e, d, n) in enumerate(zip(s25.ecl, s25.ead, s25.n)):
    a1.text(i, e / 1e6 + 0.05, f"${e / 1e6:.2f}M\n{e / d:.1%} of ${d / 1e6:.0f}M\n{n:,} loans", ha="center", fontsize=8.5, color=NAVY)
a1.set_ylim(0, s25.ecl.max() / 1e6 * 1.45)
a1.set_ylabel("Expected credit loss ($M)")
a1.set_title(f"ECL at 31 Dec 2025: ${s25.ecl.sum() / 1e6:.2f}M ({s25.ecl.sum() / s25.ead.sum():.2%} of book)")
xb = np.arange(len(bt))
a2.bar(xb - 0.2, bt.predicted_weighted, 0.4, color=COMPARE, label=f"Expected at Dec 2024 ({bt.predicted_weighted.sum():,.0f})")
a2.bar(xb + 0.2, bt.actual, 0.4, color=PRIMARY, label=f"Actual in 2025 ({bt.actual.sum():,})")
a2.set_xticks(xb, bt.score_band)
a2.set_xlabel("Score band (1 = best)")
a2.legend(frameon=False, loc="upper left")
a2.set_title("Backtest: 12-month PDs set at Dec 2024 vs 2025 defaults")
fig.text(0.01, 0.01, f"Stage 1: 12-month PD. Stage 2 (30+ DPD or recently cured): lifetime PD. Stage 3 (90+ DPD): defaulted. "
         f"With 2025's actual unemployment the Dec 2024 model would have expected {fnd['backtest_actual_u']:,.0f}.",
         fontsize=8, color=SLATE)
save(fig, "chart_ecl.png")

# 5. policy change ---------------------------------------------------------------------------------
pg = loan[loan.policy_group.notna()].groupby("policy_group").agg(d=("default_12m", "mean"), r=("int_rate", "mean"), n=("loan_id", "size"))
pg = pg.loc[["Before · DTI ≤ 40%", "After · DTI ≤ 40%", "After · DTI 40-50% (new)"]]
fig, ax = plt.subplots(figsize=(9, 3.6))
ax.barh(range(3), pg.d, color=[COMPARE, COMPARE, BAD])
for i, (d, r, n) in enumerate(zip(pg.d, pg.r, pg.n)):
    ax.text(d + 0.001, i, f"  {d:.1%}  ·  rate {r:.1f}%  ·  {n:,} loans", va="center", fontsize=9, color=NAVY)
ax.set_yticks(range(3), pg.index)
ax.invert_yaxis()
ax.xaxis.set_major_formatter(pct)
ax.set_xlim(0, pg.d.max() * 1.9)
ax.set_title(f"Online DTI limit raised to 50% (Jul 2023): new loans default {fnd['policy_multiple']:.1f}x as often")
ax.text(0, -0.22, f"12-month default rate, online loans issued Jul 2022 - Dec 2024. The extra {fnd['policy_rate_gap']:.1f} pts of interest "
        f"covers the ~{(fnd['policy_new_12m'] - fnd['policy_rest_12m']) * fnd['lgd_all'] * 100:.1f} pts of extra yearly loss.",
        transform=ax.transAxes, fontsize=8, color=SLATE)
save(fig, "chart_policy.png")
