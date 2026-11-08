"""
Build the Power BI tables (data_clean/powerbi/*.csv) and the headline findings from the database, the scorecard
and the ECL results. Run after clean.py, run_sql.py, model.py and ecl.py.

Tables: dim_loan, fact_loan_month (gzip), fact_recovery, fact_month_status, fact_month_flows, dim_month, fact_vintage, fact_roll, fact_ecl,
        fact_stage_migration, ecl_backtest, score_bands, scorecard, model_metrics, truth_check, dq_log, findings
Run: python scripts/export_powerbi.py
"""
from pathlib import Path
import json
import sqlite3
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
CLEAN = ROOT / "data_clean"
OUT = CLEAN / "powerbi"
OUT.mkdir(exist_ok=True)
truth = json.loads((ROOT / "data_raw" / "true_effects.json").read_text())
MM = json.loads((CLEAN / "model_metrics.json").read_text())
EP = json.loads((CLEAN / "ecl_parameters.json").read_text())

con = sqlite3.connect(CLEAN / "northbridge.db")
L = pd.read_sql("SELECT * FROM loans", con)
O = pd.read_sql("SELECT * FROM loan_outcomes", con)
S = pd.read_sql("SELECT * FROM loan_scores", con)
LM = pd.read_sql("SELECT loan_id, month, months_on_book, status, days_past_due, amount_paid, balance_end, "
                 "charge_off_amount FROM loan_months", con)
R = pd.read_sql("SELECT * FROM recoveries", con)
macro = pd.read_sql("SELECT * FROM macro", con)
E = pd.read_csv(CLEAN / "ecl_loans.csv")

# ---------------------------------------------------------------- dim_loan
D = L.merge(O, on="loan_id").merge(S, on="loan_id")
PURPOSE = {"debt_consolidation": "Debt consolidation", "credit_card": "Credit card refinance",
           "home_improvement": "Home improvement", "major_purchase": "Major purchase", "car": "Car",
           "medical": "Medical", "small_business": "Small business", "other": "Other"}
D["purpose_label"] = D.purpose.map(PURPOSE)
fico_edges = [-np.inf, 660, 700, 740, 780, np.inf]
fico_lab = ["620-659", "660-699", "700-739", "740-779", "780+"]
band = lambda x, edges, labs: pd.cut(x, edges, labels=labs, right=False).astype(object).where(x.notna(), "Missing")
D["fico_band"] = band(D.fico_score, fico_edges, fico_lab)
D["dti_band"] = band(D.dti, [-np.inf, 10, 20, 30, 40, np.inf], ["0-10%", "10-20%", "20-30%", "30-40%", "40-50%"])
D["income_band"] = band(D.annual_inc, [-np.inf, 40e3, 60e3, 80e3, 120e3, np.inf], ["< $40k", "$40-60k", "$60-80k", "$80-120k", "$120k+"])
D["emp_band"] = np.select([D.emp_years.isna(), D.emp_years < 1, D.emp_years < 5, D.emp_years < 10],
                          ["Unknown", "< 1 year", "1-4 years", "5-9 years"], "10+ years")
D["term"] = D.term_months.astype(str) + " months"
online = (D.channel == "Online") & D.issue_month.between("2022-07", "2024-12")
POL_BEFORE, POL_AFTER, POL_NEW = "Before · DTI ≤ 40%", "After · DTI ≤ 40%", "After · DTI 40-50% (new)"
D["policy_group"] = np.where(~online, None, np.where(D.issue_month < "2023-07", POL_BEFORE,
                             np.where(D.dti > 40, POL_NEW, POL_AFTER)))
D["net_loss"] = (D.charge_off_amount - D.recoveries).round(2)
D["status_order"] = D.status_now.map({"Current": 1, "30 DPD": 2, "60 DPD": 3, "90+ DPD": 4, "Charged Off": 5,
                                      "Prepaid": 6, "Fully Paid": 7})
e25 = E[E["asof"] == "2025-12"].set_index("loan_id")
D["stage_2025"] = D.loan_id.map(e25.stage).map({1: "Stage 1", 2: "Stage 2", 3: "Stage 3"})
dim_loan = D[["loan_id", "issue_date", "issue_month", "vintage_q", "vintage_year", "channel", "region", "loan_amnt",
              "term", "int_rate", "installment", "grade", "sub_grade", "purpose_label", "home_ownership",
              "verification_status", "emp_band", "fico_score", "fico_band", "dti", "dti_band", "annual_inc",
              "income_band", "score", "score_band", "score_range", "pd_24m", "status_now", "status_order",
              "months_observed", "months_available", "defaulted", "default_month", "default_mob", "default_12m", "default_24m",
              "charge_off_amount", "recoveries", "net_loss", "balance_now", "is_open", "policy_group", "stage_2025"]]
dim_loan = dim_loan.rename(columns={"purpose_label": "purpose"})

# ---------------------------------------------------------------- loan x month detail (for Power BI; gzip)
STATUS_ORDER = {"Current": 1, "30 DPD": 2, "60 DPD": 3, "90+ DPD": 4, "Charged Off": 5, "Prepaid": 6, "Fully Paid": 7}
LM = LM.sort_values(["loan_id", "months_on_book"])
LM["next_status"] = LM.groupby("loan_id").status.shift(-1)
fact_loan_month = LM.rename(columns={"months_on_book": "mob"})[
    ["loan_id", "month", "mob", "status", "days_past_due", "amount_paid", "balance_end", "charge_off_amount", "next_status"]].copy()
fact_loan_month["status_order"] = fact_loan_month.status.map(STATUS_ORDER)
fact_loan_month.to_csv(OUT / "fact_loan_month.csv.gz", index=False, compression="gzip")
print(f"fact_loan_month        {len(fact_loan_month):>7,} rows (gzip)")
fact_recovery = R.copy()

# ---------------------------------------------------------------- monthly book composition and flows
G = LM.merge(L[["loan_id", "grade"]], on="loan_id")
book = G[G.status.isin(["Current", "30 DPD", "60 DPD", "90+ DPD"])]
month_status = (book.groupby(["month", "grade", "status"]).agg(loans=("loan_id", "size"), balance=("balance_end", "sum"))
                .reset_index().round(2))
orig = L.groupby(["issue_month", "grade"]).agg(originations=("loan_id", "size"), originated_amount=("loan_amnt", "sum"))
orig.index.names = ["month", "grade"]
newdef = G[(G.status == "90+ DPD") & (G.days_past_due == 90)].groupby(["month", "grade"]).size().rename("new_defaults")
co = G[G.status == "Charged Off"].groupby(["month", "grade"]).charge_off_amount.sum().rename("charge_offs")
rec = R.merge(L[["loan_id", "grade"]], on="loan_id").groupby(["month", "grade"]).recovery_amount.sum().rename("recoveries")
perf = G[G.status.isin(["Current", "30 DPD", "60 DPD"])].groupby(["month", "grade"]).size().rename("performing_loans")
paid = G.groupby(["month", "grade"]).amount_paid.sum().rename("collections")
months = pd.Index(macro.month, name="month")
idx = pd.MultiIndex.from_product([months, list("ABCDEFG")], names=["month", "grade"])
flows = pd.concat([orig, newdef, co, rec, perf, paid], axis=1).reindex(idx).fillna(0).reset_index().round(2)
dim_month = macro.copy()
dim_month["month_date"] = pd.to_datetime(dim_month.month + "-01").dt.strftime("%Y-%m-%d")
dim_month["year"] = dim_month.month.str[:4].astype(int)
dim_month["quarter"] = dim_month.month.str[:4] + " Q" + ((dim_month.month.str[5:].astype(int) + 2) // 3).astype(str)

# ---------------------------------------------------------------- vintage curves (vintage quarter x grade x MOB)
V = D[["loan_id", "vintage_q", "grade", "default_mob", "months_available"]]
vin = []
for (vq, g), d in V.groupby(["vintage_q", "grade"]):
    max_mob = int(d.months_available.min())
    dm = d.default_mob.dropna().astype(int)
    for mob in range(1, 49):
        vin.append((vq, g, mob, len(d), int((dm <= mob).sum()), int(mob <= max_mob)))
fact_vintage = pd.DataFrame(vin, columns=["vintage_q", "grade", "mob", "loans", "cum_defaults", "observed"])
fact_vintage["vintage_year"] = fact_vintage.vintage_q.str[:4].astype(int)

# ---------------------------------------------------------------- roll rates (status this month -> next month)
rr = LM[LM.status.isin(["Current", "30 DPD", "60 DPD"]) & LM.next_status.notna()]
fact_roll = rr.groupby(["month", "status", "next_status"]).size().rename("loan_months").reset_index()
fact_roll.columns = ["month", "from_status", "to_status", "loan_months"]
order = {"Current": 1, "30 DPD": 2, "60 DPD": 3, "90+ DPD": 4, "Charged Off": 5, "Prepaid": 6, "Fully Paid": 7}
fact_roll["from_order"] = fact_roll.from_status.map(order)
fact_roll["to_order"] = fact_roll.to_status.map(order)

# ---------------------------------------------------------------- ECL tables
fact_ecl = E.copy()
fact_ecl["stage"] = "Stage " + fact_ecl.stage.astype(str)
fact_ecl["asof_label"] = fact_ecl["asof"].map({"2024-12": "31 Dec 2024", "2025-12": "31 Dec 2025"})
e24 = E[E["asof"] == "2024-12"][["loan_id", "stage", "ead"]].rename(columns={"stage": "stage_2024", "ead": "ead_2024"})
e25s = E[E["asof"] == "2025-12"][["loan_id", "stage"]].rename(columns={"stage": "stage_2025"})
mig = e24.merge(e25s, on="loan_id", how="left").merge(O[["loan_id", "status_now"]], on="loan_id")
mig["to_state"] = np.where(mig.stage_2025.notna(), "Stage " + mig.stage_2025.fillna(0).astype(int).astype(str),
                           mig.status_now.map({"Charged Off": "Written off", "Prepaid": "Repaid early",
                                               "Fully Paid": "Matured"}))
fact_mig = (mig.groupby(["stage_2024", "to_state"]).agg(loans=("loan_id", "size"), ead_2024=("ead_2024", "sum"))
            .reset_index().round(2))
fact_mig["stage_2024"] = "Stage " + fact_mig.stage_2024.astype(str)
fact_mig["to_order"] = fact_mig.to_state.map({"Stage 1": 1, "Stage 2": 2, "Stage 3": 3, "Written off": 4,
                                              "Repaid early": 5, "Matured": 6})
back = pd.read_csv(CLEAN / "ecl_backtest.csv")
back = back[back.score_band != "All"].copy()
back["score_band"] = back.score_band.astype(int)

# ---------------------------------------------------------------- model tables
bands = pd.read_csv(CLEAN / "score_bands.csv")
card = pd.read_csv(CLEAN / "scorecard.csv")
FEAT = {"fico_score": "FICO score", "dti": "Debt-to-income %", "revol_util": "Revolving utilisation %",
        "log_income": "Income (log)", "inq_last_6mths": "Credit inquiries (6m)", "delinq_2yrs": "Delinquencies (2y)",
        "term_60": "60-month term", "emp_band": "Employment length", "home_ownership": "Home ownership",
        "verification_status": "Income verification", "purpose": "Purpose", "channel": "Channel"}
card["feature_label"] = card.feature.map(FEAT)


def nice_bin(f, b):
    if b == "missing":
        return "Missing"
    if f == "term_60":
        return {"0": "36 months", "1": "60 months"}[b]
    if f == "log_income" and b.startswith("["):
        lo, hi = b.strip("[)").split(", ")
        fmt = lambda v: "" if v in ("-inf", "inf") else f"${np.exp(float(v)) / 1000:,.0f}k"
        return f"< {fmt(hi)}" if lo == "-inf" else (f"{fmt(lo)}+" if hi == "inf" else f"{fmt(lo)} - {fmt(hi)}")
    if b.startswith("["):
        lo, hi = b.strip("[)").split(", ")
        g = lambda v: f"{float(v):g}"
        if f in ("inq_last_6mths", "delinq_2yrs"):
            return {"-inf": "0", "0.5": "1", "1.5": "2", "2.5": "3+"}[lo]
        return f"< {g(hi)}" if lo == "-inf" else (f"{g(lo)}+" if hi == "inf" else f"{g(lo)} - {g(hi)}")
    return PURPOSE.get(b, b)


card["bin_label"] = [nice_bin(f, b) for f, b in zip(card.feature, card.bin)]
card["bin_order"] = card.groupby("feature").cumcount() + 1
card["sort_key"] = card.feature.map({f: i for i, f in enumerate(FEAT, 1)}) * 100 + card.bin_order
card = card.sort_values("sort_key")
# labels like "Missing" repeat across characteristics; make each unique (invisible zero-width spaces) so the label can
# sort by sort_key in Power BI
card["bin_label"] = card.bin_label + card.groupby("bin_label").cumcount().map(lambda k: "​" * k)
card = card[["feature", "feature_label", "bin", "bin_label", "bin_order", "sort_key", "loans", "bad_rate_pct", "woe", "points",
             "iv_feature"]]
metrics = []
for model, key in [("Scorecard", "scorecard"), ("Lender grade", "lender_grade"), ("Gradient boosting", "gradient_boosting")]:
    for sample, sk in [("Dev 2021-22", "dev"), ("Test 2023", "oot")]:
        metrics.append({"model": model, "sample": sample, **MM[key][sk]})
model_metrics = pd.DataFrame(metrics)
truth_check = pd.read_csv(CLEAN / "truth_check.csv")
truth_check["effect_label"] = truth_check.effect.map({
    "fico_per_10pts_above_700": "FICO, per +10 points", "dti_per_pt": "DTI, per +1 pt", "revol_util_per_10pts": "Revolving use, per +10 pts",
    "inquiry_each": "Each credit inquiry", "delinq_2y_each": "Each past delinquency", "log_income_vs_66k": "Income (log)",
    "term_60": "60-month term", "emp_lt_1y": "Employed < 1 year", "not_verified": "Income not verified",
    "home_MORTGAGE": "Has a mortgage", "purpose_small_business": "Purpose: small business", "purpose_medical": "Purpose: medical",
    "purpose_car": "Purpose: car", "channel_Broker": "Channel: broker", "channel_Branch": "Channel: branch"})
truth_check["same_sign"] = truth_check.same_sign.map({True: "✓", False: "✗"})
dq = pd.read_csv(CLEAN / "dq_log.csv")

# ---------------------------------------------------------------- findings
q = lambda sql: pd.read_sql(sql, con)
old = D[D.issue_month <= "2023-12"]
g24 = old.groupby("grade").default_24m.mean()
v24 = D[D.vintage_year <= 2023].groupby("vintage_year").default_mob.apply(lambda s: (s <= 24).mean())
pol = D[D.policy_group.notna()].groupby("policy_group").default_12m.mean()
new_pol = pol[POL_NEW]
rest_pol = pol[POL_AFTER]
rate_gap = D[D.policy_group == POL_NEW].int_rate.mean() - D[D.policy_group == POL_AFTER].int_rate.mean()
lgd = EP["2025-12"]["lgd"]
s25 = pd.read_csv(CLEAN / "ecl_summary.csv")
s25 = s25[s25["asof"] == "2025-12"].set_index("stage")
bt = EP["backtest_2025"]
F = [
    ("loans", len(D), f"{len(D):,} loans, ${L.loan_amnt.sum() / 1e6:,.0f}M originated Jan 2021 - Dec 2024, average rate {L.int_rate.mean():.1f}%"),
    ("default_rate_24m", old.default_24m.mean(), f"{old.default_24m.mean():.1%} of loans issued 2021-2023 defaulted within 24 months"),
    ("grade_g_vs_a", g24["G"] / g24["A"], f"the grade ranks risk: {g24['A']:.1%} (A) to {g24['G']:.1%} (G) default within 24 months, {g24['G'] / g24['A']:.1f}x"),
    ("vintage_2023_24m", v24[2023], f"the 2023 vintage reached {v24[2023]:.1%} defaults by month 24 vs {v24[2021]:.1%} for 2021: it met the 2024 unemployment rise early in life"),
    ("vintage_2021_24m", v24[2021], ""),
    ("policy_new_12m", new_pol, f"online loans with DTI 40-50% (allowed from Jul 2023) default at {new_pol:.1%} in 12 months vs "
                                f"{rest_pol:.1%} for the rest ({new_pol / rest_pol:.1f}x); they pay {rate_gap:.1f} pts more interest, "
                                f"more than the ~{(new_pol - rest_pol) * lgd['ALL'] * 100:.1f} pts a year of extra expected loss"),
    ("policy_rest_12m", rest_pol, ""),
    ("policy_multiple", new_pol / rest_pol, ""),
    ("policy_rate_gap", rate_gap, ""),
    ("lgd_all", lgd["ALL"], f"loss given default {lgd['ALL']:.1%}: {lgd['MORTGAGE']:.1%} for homeowners with a mortgage vs {lgd['RENT']:.1%} for renters"),
    ("lgd_mortgage", lgd["MORTGAGE"], ""), ("lgd_rent", lgd["RENT"], ""),
    ("gini_scorecard_oot", MM["scorecard"]["oot"]["gini"], f"the scorecard's Gini on 2023 loans (out of time) is {MM['scorecard']['oot']['gini']:.2f} vs {MM['lender_grade']['oot']['gini']:.2f} for the lender's grade"),
    ("gini_grade_oot", MM["lender_grade"]["oot"]["gini"], ""),
    ("gini_gb_oot", MM["gradient_boosting"]["oot"]["gini"], ""),
    ("ks_scorecard_oot", MM["scorecard"]["oot"]["ks"], ""),
    ("band_spread", MM["band_spread_oot"], f"the worst score band defaults {MM['band_spread_oot']:.1f}x as often as the best"),
    ("truth_corr", MM["truth_check"]["correlation"], f"all {MM['truth_check']['effects']} estimated risk effects have the true sign (correlation with the truth {MM['truth_check']['correlation']:.2f})"),
    ("unemployment_beta", EP["2025-12"]["unemployment_beta"], f"each +1 pt of unemployment raises monthly default starts by {np.exp(EP['2025-12']['unemployment_beta']) - 1:.0%} (true: {np.exp(0.30) - 1:.0%})"),
    ("ecl_total", float(s25.loc["Total", "ecl"]), f"IFRS 9 ECL at 31 Dec 2025: ${s25.loc['Total', 'ecl'] / 1e6:.2f}M on a ${s25.loc['Total', 'ead'] / 1e6:.1f}M open book ({s25.loc['Total', 'coverage_pct']:.2f}% coverage)"),
    ("ead_total", float(s25.loc["Total", "ead"]), ""),
    ("coverage_total", float(s25.loc["Total", "coverage_pct"]) / 100, ""),
    ("coverage_s1", float(s25.loc["1", "coverage_pct"]) / 100, ""), ("coverage_s2", float(s25.loc["2", "coverage_pct"]) / 100, ""),
    ("coverage_s3", float(s25.loc["3", "coverage_pct"]) / 100, ""),
    ("ecl_downside", float(s25.loc["Total", "ecl_downside"]), ""), ("ecl_upside", float(s25.loc["Total", "ecl_upside"]), ""),
    ("backtest_predicted", bt["predicted_weighted"], f"backtest: at Dec 2024 the model expected {bt['predicted_weighted']:,.0f} defaults in 2025 (scenario-weighted); {bt['actual']:,} happened. With 2025's actual unemployment it would have said {bt['predicted_with_actual_2025_unemployment']:,.0f}"),
    ("backtest_actual", bt["actual"], ""), ("backtest_actual_u", bt["predicted_with_actual_2025_unemployment"], ""),
    ("net_loss", float(D.net_loss.sum()), f"${D.charge_off_amount.sum() / 1e6:.1f}M charged off, ${D.recoveries.sum() / 1e6:.1f}M recovered, net loss ${D.net_loss.sum() / 1e6:.1f}M"),
]
findings = pd.DataFrame(F, columns=["key", "value", "finding"])

for name, df in [("dim_loan", dim_loan), ("fact_month_status", month_status), ("fact_month_flows", flows),
                 ("dim_month", dim_month), ("fact_vintage", fact_vintage), ("fact_roll", fact_roll), ("fact_ecl", fact_ecl),
                 ("fact_stage_migration", fact_mig), ("ecl_backtest", back), ("score_bands", bands), ("scorecard", card),
                 ("model_metrics", model_metrics), ("truth_check", truth_check), ("dq_log", dq), ("findings", findings),
                 ("fact_recovery", fact_recovery)]:
    df.to_csv(OUT / f"{name}.csv", index=False)
    print(f"{name:<22} {len(df):>7,} rows")
for k, v, f in F:
    if f:
        print(" -", f)
