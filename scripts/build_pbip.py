"""
Northbridge Lending credit risk: generate the Power BI dashboard as a Power BI Project (PBIP).

Writes powerbi/Northbridge.pbip:
  Northbridge.SemanticModel/  loan table + loan-month history (1.16M rows) + recoveries + IFRS 9 ECL per loan,
                              scorecard, backtest and data-quality tables (data_clean/powerbi/*.csv); measures in TMDL
  Northbridge.Report/         custom theme, sidebar navigation, 6 report pages, a Loan Detail drill-through page
                              and a month tooltip page (PBIR JSON)

Open the .pbip in Power BI Desktop (Nov 2025 release or later) and click Refresh.
Run: python scripts/build_pbip.py   (from the project root, after scripts/export_powerbi.py)
"""
import csv
import gzip
import json
import shutil
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PBI = ROOT / "powerbi"
NAME = "Northbridge"
SM = PBI / f"{NAME}.SemanticModel"
RP = PBI / f"{NAME}.Report"
DATA_FOLDER = str(ROOT / "data_clean" / "powerbi") + "\\"
BASE_THEME = "CY26SU04"
THEME_SRC = Path(r"C:\Program Files\Microsoft Power BI Desktop\bin\WebView2Resources"
                 r"\minerva\sharedresources\BaseThemes") / f"{BASE_THEME}.json"
CUSTOM_THEME = "Northbridge_Executive.json"
SCHEMA = "https://developer.microsoft.com/json-schemas/fabric"
V_VISUAL, V_PAGE, V_REPORT = "2.4.0", "2.0.0", "3.0.0"

# =========================================================================
# Design system
# =========================================================================
NAVY = "#0B1F3A"        # sidebar, primary text on light
NAVY_2 = "#16314F"      # active nav item
INK = "#0F172A"
SLATE = "#64748B"
SLATE_2 = "#94A3B8"
MIST = "#CBD5E1"
LINE = "#E6E9F0"
CANVAS = "#F3F5F9"
WHITE = "#FFFFFF"
GOLD = "#C8A24A"
PRIMARY = "#1F4E79"     # main series
COMPARE = "#B8C2CF"     # comparison series
GOOD, WARN, BAD, NEUTRAL = "#0E7C86", "#D08C2E", "#B5523B", "#94A3B8"
AMBER = "#C2703A"
DATA_COLORS = [PRIMARY, GOOD, GOLD, "#7A8CA5", BAD, "#5B8E7D", "#8C6BB1", WARN]
FONT, FONT_SB, FONT_LIGHT = "Segoe UI", "Segoe UI Semibold", "Segoe UI Light"


def gid(*parts):
    return str(uuid.uuid5(uuid.NAMESPACE_URL, "northbridge/" + "/".join(parts)))


def hid(*parts):
    return gid(*parts).replace("-", "")[:20]


def write(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def write_json(path, obj):
    write(path, json.dumps(obj, indent=2, ensure_ascii=False) + "\n")


# =========================================================================
# Semantic model
# =========================================================================
S, I, D, DT, B = "string", "int64", "double", "dateTime", "boolean"
MTYPE = {S: "type text", I: "Int64.Type", D: "type number", DT: "type date", B: "logical"}


def cols(*spec):
    """(name, type[, hidden[, extra]]) -> column tuples. Every column summarizeBy none."""
    out = []
    for name, dtype, *rest in spec:
        hidden = bool(rest and rest[0])
        extra = dict(rest[1]) if len(rest) > 1 else {}
        if dtype == DT:
            extra.setdefault("formatString", "dd mmm yyyy")
        out.append((name, dtype, hidden, extra))
    return out


def order_col(new, src, values):
    """M step: position of [src] in a fixed list (1-based), used as a sort-by column."""
    lst = "{{" + ", ".join(f'"{v}"' for v in values) + "}}"           # doubled: the step goes through str.format
    return (f"Ord_{new}", f'Table.AddColumn({{prev}}, "{new}", each if [{src}] = null then null else '
                          f'List.PositionOf({lst}, [{src}]) + 1, Int64.Type)')


STATUSES = ["Current", "30 DPD", "60 DPD", "90+ DPD", "Charged Off", "Prepaid", "Fully Paid"]
POLICY = ["Before · DTI ≤ 40%", "After · DTI ≤ 40%", "After · DTI 40-50% (new)"]

# table -> (csv file or None, columns loaded, extra M steps (list of (name, expr using {prev})), derived columns)
TABLES = {
    "dim_loan": ("dim_loan.csv", cols(
        ("loan_id", S, False, {"isKey": None}), ("issue_date", DT), ("issue_month", S, True), ("vintage_q", S),
        ("vintage_year", I, False, {"formatString": "0"}), ("channel", S), ("region", S), ("loan_amnt", D, True),
        ("term", S), ("int_rate", D, True), ("installment", D, True), ("grade", S), ("sub_grade", S), ("purpose", S),
        ("home_ownership", S), ("verification_status", S), ("emp_band", S, False, {"sortByColumn": "emp_order"}),
        ("fico_score", D, True), ("fico_band", S), ("dti", D, True), ("dti_band", S), ("annual_inc", D, True),
        ("income_band", S, False, {"sortByColumn": "income_order"}), ("score", I, True), ("score_band", I),
        ("score_range", S), ("pd_24m", D, True), ("status_now", S, False, {"sortByColumn": "status_order"}),
        ("status_order", I, True), ("months_observed", I, True), ("months_available", I, True), ("defaulted", I, True),
        ("default_month", S, True), ("default_mob", I, True), ("default_12m", D, True), ("default_24m", D, True),
        ("charge_off_amount", D, True), ("recoveries", D, True), ("net_loss", D, True), ("balance_now", D, True),
        ("is_open", I, True), ("policy_group", S, False, {"sortByColumn": "policy_order"}), ("stage_2025", S)),
        [order_col("emp_order", "emp_band", ["< 1 year", "1-4 years", "5-9 years", "10+ years", "Unknown"]),
         order_col("income_order", "income_band", ["< $40k", "$40-60k", "$60-80k", "$80-120k", "$120k+", "Missing"]),
         order_col("policy_order", "policy_group", POLICY)],
        cols(("emp_order", I, True), ("income_order", I, True), ("policy_order", I, True))),
    "fact_loan_month": ("fact_loan_month.csv.gz", cols(
        ("loan_id", S, True), ("month", S, True), ("mob", I, True), ("status", S, False, {"sortByColumn": "status_order"}),
        ("days_past_due", I, True), ("amount_paid", D, True), ("balance_end", D, True), ("charge_off_amount", D, True),
        ("next_status", S, False, {"sortByColumn": "next_order"}), ("status_order", I, True)),
        [order_col("next_order", "next_status", STATUSES)], cols(("next_order", I, True))),
    "fact_recovery": ("fact_recovery.csv", cols(("loan_id", S, True), ("month", S, True), ("recovery_amount", D, True)), [], []),
    "fact_ecl": ("fact_ecl.csv", cols(
        ("asof", S, True), ("loan_id", S, True), ("stage", S), ("mob", I, True), ("ead", D, True), ("lgd", D, True),
        ("pd12", D, True), ("pd_life", D, True), ("ecl", D, True), ("ecl_Base", D, True), ("ecl_Upside", D, True),
        ("ecl_Downside", D, True), ("asof_label", S, True)), [], []),
    "dim_month": ("dim_month.csv", cols(
        ("month", S, True, {"isKey": None}), ("unemployment_rate", D, True), ("base_rate", D, True), ("month_date", DT, False,
                                                                                                      {"formatString": "mmm yyyy"}),
        ("year", I, False, {"formatString": "0"}), ("quarter", S)),
        [("Label", 'Table.AddColumn({prev}, "month_label", each Date.ToText([month_date], "MMM yyyy", "en-US"), type text)')],
        cols(("month_label", S, False, {"sortByColumn": "month_date"}))),
    "dim_mob": (None, cols(("mob", I, False, {"formatString": "0"})), [], []),
    "dim_scenario": (None, cols(("scenario", S, False, {"sortByColumn": "scenario_order"}), ("scenario_order", I, True),
                                ("weight", D, True)), [], []),
    "fact_stage_migration": ("fact_stage_migration.csv", cols(
        ("stage_2024", S), ("to_state", S, False, {"sortByColumn": "to_order"}), ("loans", I, True), ("ead_2024", D, True),
        ("to_order", I, True)), [], []),
    "ecl_backtest": ("ecl_backtest.csv", cols(
        ("score_band", I), ("loans", I, True), ("predicted_base", D, True), ("predicted_weighted", D, True),
        ("predicted_actual_unemployment", D, True), ("actual", I, True)), [], []),
    "score_bands": ("score_bands.csv", cols(
        ("score_band", I), ("score_range", S), ("loans", I, True), ("predicted_pd_pct", D, True), ("actual_default_pct", D, True)),
        [("BandLabel", 'Table.AddColumn({prev}, "band_label", each Text.From([score_band]) & " · " & [score_range], type text)')],
        cols(("band_label", S, False, {"sortByColumn": "score_band"}))),
    "scorecard": ("scorecard.csv", cols(
        ("feature", S, True), ("feature_label", S), ("bin", S, True), ("bin_label", S, False, {"sortByColumn": "sort_key"}),
        ("bin_order", I, True),
        ("sort_key", I, True), ("loans", I, True), ("bad_rate_pct", D, True), ("woe", D, True), ("points", I, True),
        ("iv_feature", D, True)), [], []),
    "model_metrics": ("model_metrics.csv", cols(("model", S), ("sample", S), ("auc", D, True), ("gini", D, True), ("ks", D, True)), [], []),
    "truth_check": ("truth_check.csv", cols(("effect", S, True), ("effect_label", S), ("true", D, True), ("estimated", D, True),
                                            ("same_sign", S)), [], []),
    "dq_log": ("dq_log.csv", cols(
        ("step", I, True), ("table_name", S), ("issue", S, False, {"sortByColumn": "step"}), ("rows_affected", I, True),
        ("action", S), ("injected", I, True)), [], []),
    "findings": ("findings.csv", cols(("key", S), ("value", D, True), ("finding", S)), [], []),
}

CUSTOM_M = {
    "dim_mob": ['Table.FromList({1..48}, Splitter.SplitByNothing(), type table [mob = Int64.Type])'],
    "dim_scenario": ['#table(type table [scenario = text, scenario_order = Int64.Type, weight = number], '
                     '{{"Upside", 1, 0.2}, {"Base", 2, 0.6}, {"Downside", 3, 0.2}, {"Weighted", 4, 1}})'],
}

# (from table, from column, to table, to column[, active])
RELATIONSHIPS = [
    ("fact_loan_month", "loan_id", "dim_loan", "loan_id"), ("fact_loan_month", "month", "dim_month", "month"),
    ("fact_recovery", "loan_id", "dim_loan", "loan_id"), ("fact_recovery", "month", "dim_month", "month"),
    ("fact_ecl", "loan_id", "dim_loan", "loan_id"),
    ("dim_loan", "issue_month", "dim_month", "month", False),
]

# ---------------------------------------------------------------- measures
CUR, CURM, PCT, PCT2, INT, DEC1, DEC2, DEC3 = (r"\$#,0;(\$#,0);\$#,0", r"\$#,0.0,,\M;(\$#,0.0,,\M);\$#,0.0,,\M",
                                               "0.0%;-0.0%;0.0%", "0.00%;-0.00%;0.00%", "#,0", "#,0.0", "#,0.00", "0.000")
MULT = '0.0"x"'
SPCT = "+0%;-0%;0%"
MEASURES = []   # (name, dax, format, folder, hidden)


def m(name, dax, fmt=None, folder="", hidden=False):
    MEASURES.append((name, dax, fmt, folder, hidden))


def finding(key):
    return f'CALCULATE ( MAX ( findings[value] ), findings[key] = "{key}" )'


OPEN = '{ "Current", "30 DPD", "60 DPD", "90+ DPD" }'
LAST = '"2025-12"'
# 1 Portfolio ---------------------------------------------------------------------------
m("Loans", "COUNTROWS ( dim_loan )", INT, "1 Portfolio")
m("Amount Lent", "SUM ( dim_loan[loan_amnt] )", CUR, "1 Portfolio")
m("Avg Loan", "DIVIDE ( [Amount Lent], [Loans] )", CUR, "1 Portfolio")
m("Avg Rate", "DIVIDE ( SUMX ( dim_loan, dim_loan[int_rate] * dim_loan[loan_amnt] ), [Amount Lent] ) / 100", PCT, "1 Portfolio")
m("Defaults", "SUM ( dim_loan[defaulted] )", INT, "1 Portfolio")
m("Default Rate 24m", "DIVIDE ( SUM ( dim_loan[default_24m] ), COUNT ( dim_loan[default_24m] ) )", PCT, "1 Portfolio")
m("Default Rate 12m", "DIVIDE ( SUM ( dim_loan[default_12m] ), COUNT ( dim_loan[default_12m] ) )", PCT, "1 Portfolio")
m("Charged Off", "SUM ( dim_loan[charge_off_amount] )", CUR, "1 Portfolio")
m("Recovered", "SUM ( dim_loan[recoveries] )", CUR, "1 Portfolio")
m("Net Loss", "[Charged Off] - [Recovered]", CUR, "1 Portfolio")
m("Net Loss Rate", "DIVIDE ( [Net Loss], [Amount Lent] )", PCT, "1 Portfolio")
m("Recovery Rate", "DIVIDE ( [Recovered], [Charged Off] )", PCT, "1 Portfolio")
m("Book Balance", f"CALCULATE ( SUM ( fact_loan_month[balance_end] ), KEEPFILTERS ( fact_loan_month[status] IN {OPEN} ) )", CUR, "1 Portfolio")
m("Open Balance", f"CALCULATE ( [Book Balance], fact_loan_month[month] = {LAST} )", CUR, "1 Portfolio")
m("Open Loans", f"CALCULATE ( COUNTROWS ( fact_loan_month ), fact_loan_month[month] = {LAST}, fact_loan_month[status] IN {OPEN} )",
  INT, "1 Portfolio")
m("Arrears Balance", 'CALCULATE ( [Book Balance], KEEPFILTERS ( fact_loan_month[status] IN { "30 DPD", "60 DPD", "90+ DPD" } ) )', CUR, "1 Portfolio")
m("Arrears Share", "DIVIDE ( [Arrears Balance], [Book Balance] )", PCT, "1 Portfolio")
m("Arrears Share Now", f"CALCULATE ( [Arrears Share], fact_loan_month[month] = {LAST} )", PCT, "1 Portfolio")
m("New Defaults", 'CALCULATE ( COUNTROWS ( fact_loan_month ), fact_loan_month[status] = "90+ DPD", fact_loan_month[days_past_due] = 90 )',
  INT, "1 Portfolio")
m("Performing Loans", 'CALCULATE ( COUNTROWS ( fact_loan_month ), fact_loan_month[status] IN { "Current", "30 DPD", "60 DPD" } )',
  INT, "1 Portfolio")
m("Defaults per 1,000", "DIVIDE ( [New Defaults], [Performing Loans] ) * 1000", DEC2, "1 Portfolio")
m("Unemployment", "DIVIDE ( AVERAGE ( dim_month[unemployment_rate] ), 100 )", PCT, "1 Portfolio")
m("Originations", "CALCULATE ( [Amount Lent], USERELATIONSHIP ( dim_loan[issue_month], dim_month[month] ) )", CUR, "1 Portfolio")
m("Collections", "SUM ( fact_loan_month[amount_paid] )", CUR, "1 Portfolio")
m("Recoveries in Month", "SUM ( fact_recovery[recovery_amount] )", CUR, "1 Portfolio")
# 2 Risk drivers ------------------------------------------------------------------------
m("Avg FICO", "AVERAGE ( dim_loan[fico_score] )", INT, "2 Risk")
m("Avg DTI", "DIVIDE ( AVERAGE ( dim_loan[dti] ), 100 )", PCT, "2 Risk")
m("Avg Income", "AVERAGE ( dim_loan[annual_inc] )", CUR, "2 Risk")
m("FICO Multiple", 'DIVIDE ( CALCULATE ( [Default Rate 24m], dim_loan[fico_band] = "620-659" ), '
  'CALCULATE ( [Default Rate 24m], dim_loan[fico_band] = "780+" ) )', MULT, "2 Risk")
m("Riskiest Purpose", 'VAR t = ADDCOLUMNS ( VALUES ( dim_loan[purpose] ), "@v", [Default Rate 24m] ) '
  'RETURN MAXX ( TOPN ( 1, t, [@v], DESC ), dim_loan[purpose] )', None, "2 Risk")
m("Riskiest Purpose Rate", 'MAXX ( VALUES ( dim_loan[purpose] ), [Default Rate 24m] )', PCT, "2 Risk")
m("Policy Multiple", finding("policy_multiple"), MULT, "2 Risk")
m("Policy New 12m", finding("policy_new_12m"), PCT, "2 Risk")
m("Policy Rest 12m", finding("policy_rest_12m"), PCT, "2 Risk")
m("Policy Rate Gap", finding("policy_rate_gap"), '+0.0" pts"', "2 Risk")
m("LGD", finding("lgd_all"), PCT, "2 Risk")
m("LGD Mortgage", finding("lgd_mortgage"), PCT, "2 Risk")
m("LGD Rent", finding("lgd_rent"), PCT, "2 Risk")
# 3 Vintages ----------------------------------------------------------------------------
m("Cum Default Rate",
  "VAR m = MAX ( dim_mob[mob] ) VAR avail = MIN ( dim_loan[months_available] ) RETURN "
  "IF ( m > avail, BLANK (), DIVIDE ( CALCULATE ( COUNTROWS ( dim_loan ), NOT ISBLANK ( dim_loan[default_mob] ), "
  "dim_loan[default_mob] <= m ), [Loans] ) )", PCT, "3 Vintages")
m("Vintage 2023 at 24m", finding("vintage_2023_24m"), PCT, "3 Vintages")
m("Vintage 2021 at 24m", finding("vintage_2021_24m"), PCT, "3 Vintages")
m("Roll Rate", 'IF ( SELECTEDVALUE ( fact_loan_month[status] ) IN { "Current", "30 DPD", "60 DPD" } '
  '&& NOT ISBLANK ( SELECTEDVALUE ( fact_loan_month[next_status] ) ), '
  "DIVIDE ( COUNTROWS ( fact_loan_month ), CALCULATE ( COUNTROWS ( fact_loan_month ), "
  "REMOVEFILTERS ( fact_loan_month[next_status], fact_loan_month[next_order] ), NOT ISBLANK ( fact_loan_month[next_status] ) ) ) )",
  PCT, "3 Vintages")


def roll(frm, to):
    return (f'CALCULATE ( DIVIDE ( CALCULATE ( COUNTROWS ( fact_loan_month ), fact_loan_month[next_status] = "{to}" ), '
            f'CALCULATE ( COUNTROWS ( fact_loan_month ), NOT ISBLANK ( fact_loan_month[next_status] ) ) ), '
            f'fact_loan_month[status] = "{frm}" )')


m("Roll 30 to 60", roll("30 DPD", "60 DPD"), PCT, "3 Vintages")
m("Roll 60 to 90", roll("60 DPD", "90+ DPD"), PCT, "3 Vintages")
m("Cure from 30", roll("30 DPD", "Current"), PCT, "3 Vintages")
m("Roll Current to 30", roll("Current", "30 DPD"), PCT2, "3 Vintages")
m("Unemployment Sensitivity", f"EXP ( {finding('unemployment_beta')} ) - 1", SPCT, "3 Vintages")
m("Roll Heat", 'VAR r = [Roll Rate] RETURN SWITCH ( TRUE (), ISBLANK ( r ), BLANK (), r > 0.6, "#1F4E79", r > 0.3, "#4F7BA6", '
  'r > 0.1, "#8FB0CF", r > 0.01, "#C9D8E8", "#F3F6FA" )', None, "8 Formatting", hidden=True)
m("Roll Heat Font", 'IF ( [Roll Rate] > 0.3, "#FFFFFF", "#1E293B" )', None, "8 Formatting", hidden=True)
# 4 Scorecard ---------------------------------------------------------------------------
m("Gini Scorecard", finding("gini_scorecard_oot"), DEC2, "4 Scorecard")
m("Gini Grade", finding("gini_grade_oot"), DEC2, "4 Scorecard")
m("KS Scorecard", finding("ks_scorecard_oot"), DEC2, "4 Scorecard")
m("Band Spread", finding("band_spread"), MULT, "4 Scorecard")
m("Truth Correlation", finding("truth_corr"), DEC2, "4 Scorecard")
m("AUC", "SUM ( model_metrics[auc] )", DEC3, "4 Scorecard")
m("Gini", "SUM ( model_metrics[gini] )", DEC3, "4 Scorecard")
m("KS", "SUM ( model_metrics[ks] )", DEC3, "4 Scorecard")
m("Predicted PD", "DIVIDE ( SUM ( score_bands[predicted_pd_pct] ), 100 )", PCT, "4 Scorecard")
m("Actual Default", "DIVIDE ( SUM ( score_bands[actual_default_pct] ), 100 )", PCT, "4 Scorecard")
m("Points", "SUM ( scorecard[points] )", INT, "4 Scorecard")
m("Bin Bad Rate", "DIVIDE ( SUM ( scorecard[bad_rate_pct] ), 100 )", PCT, "4 Scorecard")
m("WoE", "SUM ( scorecard[woe] )", DEC2, "4 Scorecard")
m("Bin Loans", "SUM ( scorecard[loans] )", INT, "4 Scorecard")
m("Sort Key", "MIN ( scorecard[sort_key] )", INT, "9 Helpers", hidden=True)
m("True Effect", "SUM ( truth_check[true] )", DEC3, "4 Scorecard")
m("Estimated Effect", "SUM ( truth_check[estimated] )", DEC3, "4 Scorecard")
m("Avg Score", "AVERAGE ( dim_loan[score] )", INT, "4 Scorecard")
m("Avg PD 24m", "AVERAGE ( dim_loan[pd_24m] )", PCT, "4 Scorecard")
# 5 IFRS 9 ------------------------------------------------------------------------------


def at(expr, asof="2025-12"):
    return f'CALCULATE ( {expr}, fact_ecl[asof] = "{asof}" )'


m("ECL", at("SUM ( fact_ecl[ecl] )"), CUR, "5 IFRS 9")
m("EAD", at("SUM ( fact_ecl[ead] )"), CUR, "5 IFRS 9")
m("Coverage", "DIVIDE ( [ECL], [EAD] )", PCT2, "5 IFRS 9")
m("ECL Loans", at("COUNTROWS ( fact_ecl )"), INT, "5 IFRS 9")
m("ECL Base", at("SUM ( fact_ecl[ecl_Base] )"), CUR, "5 IFRS 9")
m("ECL Upside", at("SUM ( fact_ecl[ecl_Upside] )"), CUR, "5 IFRS 9")
m("ECL Downside", at("SUM ( fact_ecl[ecl_Downside] )"), CUR, "5 IFRS 9")
m("ECL Dec 2024", at("SUM ( fact_ecl[ecl] )", "2024-12"), CUR, "5 IFRS 9")
m("EAD Dec 2024", at("SUM ( fact_ecl[ead] )", "2024-12"), CUR, "5 IFRS 9")
m("Coverage Dec 2024", "DIVIDE ( [ECL Dec 2024], [EAD Dec 2024] )", PCT2, "5 IFRS 9")
m("PD 12m (EAD-weighted)", at("DIVIDE ( SUMX ( fact_ecl, fact_ecl[pd12] * fact_ecl[ead] ), SUM ( fact_ecl[ead] ) )"), PCT2, "5 IFRS 9")
m("Lifetime PD (EAD-weighted)", at("DIVIDE ( SUMX ( fact_ecl, fact_ecl[pd_life] * fact_ecl[ead] ), SUM ( fact_ecl[ead] ) )"), PCT2, "5 IFRS 9")
m("LGD (EAD-weighted)", at("DIVIDE ( SUMX ( fact_ecl, fact_ecl[lgd] * fact_ecl[ead] ), SUM ( fact_ecl[ead] ) )"), PCT, "5 IFRS 9")
for s_ in (1, 2, 3):
    m(f"ECL Stage {s_}", f'CALCULATE ( [ECL], fact_ecl[stage] = "Stage {s_}" )', CUR, "5 IFRS 9")
    m(f"Coverage Stage {s_}", f'CALCULATE ( [Coverage], fact_ecl[stage] = "Stage {s_}" )', PCT2, "5 IFRS 9")
    m(f"EAD Stage {s_}", f'CALCULATE ( [EAD], fact_ecl[stage] = "Stage {s_}" )', CUR, "5 IFRS 9")
    m(f"Loans Stage {s_}", f'CALCULATE ( [ECL Loans], fact_ecl[stage] = "Stage {s_}" )', INT, "5 IFRS 9")
m("Scenario ECL", 'SWITCH ( SELECTEDVALUE ( dim_scenario[scenario] ), "Upside", [ECL Upside], "Base", [ECL Base], '
  '"Downside", [ECL Downside], "Weighted", [ECL] )', CUR, "5 IFRS 9")
m("Scenario Color", f'SWITCH ( SELECTEDVALUE ( dim_scenario[scenario] ), "Upside", "{GOOD}", "Downside", "{BAD}", '
  f'"Weighted", "{NAVY}", "{COMPARE}" )', None, "8 Formatting", hidden=True)
m("Backtest Expected", "SUM ( ecl_backtest[predicted_weighted] )", INT, "5 IFRS 9")
m("Backtest Actual", "SUM ( ecl_backtest[actual] )", INT, "5 IFRS 9")
m("Backtest Actual Unemployment", "SUM ( ecl_backtest[predicted_actual_unemployment] )", INT, "5 IFRS 9")
m("Backtest Ratio", "DIVIDE ( [Backtest Expected], [Backtest Actual] )", '0.00"x"', "5 IFRS 9")
m("Migration Loans", "SUM ( fact_stage_migration[loans] )", INT, "5 IFRS 9")
m("Migration Share", "DIVIDE ( [Migration Loans], CALCULATE ( [Migration Loans], REMOVEFILTERS ( fact_stage_migration[to_state], "
  "fact_stage_migration[to_order] ) ) )", PCT, "5 IFRS 9")
m("Migration Heat", 'VAR r = [Migration Share] RETURN SWITCH ( TRUE (), ISBLANK ( r ), BLANK (), r > 0.5, "#1F4E79", r > 0.2, "#4F7BA6", '
  'r > 0.08, "#8FB0CF", r > 0.02, "#C9D8E8", "#F3F6FA" )', None, "8 Formatting", hidden=True)
m("Migration Heat Font", 'IF ( [Migration Share] > 0.2, "#FFFFFF", "#1E293B" )', None, "8 Formatting", hidden=True)
m("ECL Rank", "IF ( ISINSCOPE ( dim_loan[loan_id] ), RANKX ( ALLSELECTED ( dim_loan[loan_id] ), [ECL], , DESC, DENSE ) )", INT, "5 IFRS 9")
# 6 Data quality ------------------------------------------------------------------------
m("DQ Rows", "SUM ( dq_log[rows_affected] )", INT, "6 Data quality")
m("DQ Injected", "SUM ( dq_log[injected] )", INT, "6 Data quality")
m("DQ Match", 'IF ( ISBLANK ( [DQ Injected] ), BLANK (), IF ( [DQ Rows] = [DQ Injected], "✓", "✗" ) )', None, "6 Data quality")
m("Match Color", f'IF ( [DQ Match] = "✓", "{GOOD}", "{BAD}" )', None, "8 Formatting", hidden=True)
m("Rows Received", 'CALCULATE ( [DQ Rows], dq_log[issue] = "Rows received" )', INT, "6 Data quality")
m("Duplicates Removed", 'CALCULATE ( [DQ Rows], dq_log[issue] = "Exact duplicate rows" )', INT, "6 Data quality")
m("Test Loans Removed", 'CALCULATE ( [DQ Rows], dq_log[issue] = "Test loans (TEST- id, amount 1.00)" )', INT, "6 Data quality")
m("Clean Loans", 'CALCULATE ( [DQ Rows], dq_log[issue] = "Clean loans" )', INT, "6 Data quality")
m("Fixes", "CALCULATE ( [DQ Rows], NOT ISBLANK ( dq_log[injected] ) )", INT, "6 Data quality")
m("Field Fixes", 'CALCULATE ( [Fixes], NOT dq_log[issue] IN { "Exact duplicate rows", "Test loans (TEST- id, amount 1.00)", '
  '"Duplicate month rows" } )', INT, "6 Data quality")
m("Checks Matched", 'VAR t = FILTER ( dq_log, NOT ISBLANK ( dq_log[injected] ) ) RETURN COUNTROWS ( FILTER ( t, dq_log[rows_affected] = '
  'dq_log[injected] ) ) & " of " & COUNTROWS ( t ) & " fixes = injected count"', None, "6 Data quality")
m("Fix Color", f'IF ( SELECTEDVALUE ( dq_log[issue] ) IN {{ "Exact duplicate rows", "Test loans (TEST- id, amount 1.00)", '
  f'"Duplicate month rows" }}, "{BAD}", "{WARN}" )', None, "8 Formatting", hidden=True)
# 7 Dynamic text ------------------------------------------------------------------------


def _sel(colref, all_text):
    return (f'IF ( ISFILTERED ( {colref} ), CONCATENATEX ( VALUES ( {colref} ), {colref}, ", ", {colref}, ASC ), '
            f'"{all_text}" )')


m("Filter Context",
  f'{_sel("dim_loan[vintage_year]", "Issued 2021 – 2024")} & "   ·   " & {_sel("dim_loan[grade]", "All grades")} & '
  f'"   ·   " & {_sel("dim_loan[channel]", "All channels")} & "   ·   performance to 31 Dec 2025"', None, "7 Dynamic text")
m("Blank Text", '""', None, "9 Helpers", hidden=True)
m("Title Loan", '"Loan " & SELECTEDVALUE ( dim_loan[loan_id] ) & "  ·  " & SELECTEDVALUE ( dim_loan[sub_grade] ) & "  ·  " & '
  'SELECTEDVALUE ( dim_loan[purpose] ) & "  ·  issued " & FORMAT ( MAX ( dim_loan[issue_date] ), "mmm yyyy" )', None, "7 Dynamic text")
m("Title Month", 'FORMAT ( MAX ( dim_month[month_date] ), "mmmm yyyy" ) & "  ·  unemployment " & FORMAT ( [Unemployment], "0.0%" )',
  None, "7 Dynamic text")
m("Title Vintage", '"Cumulative default rate by months on book  ·  2023 loans: " & FORMAT ( [Vintage 2023 at 24m], "0.0%" ) & '
  '" by month 24 vs " & FORMAT ( [Vintage 2021 at 24m], "0.0%" ) & " for 2021"', None, "7 Dynamic text")
m("Loan Status", 'SELECTEDVALUE ( dim_loan[status_now] ) & IF ( SELECTEDVALUE ( dim_loan[stage_2025] ) <> "", "  ·  " & '
  'SELECTEDVALUE ( dim_loan[stage_2025] ), "" )', None, "7 Dynamic text")
m("Loan Rate", "DIVIDE ( AVERAGE ( dim_loan[int_rate] ), 100 )", "0.00%", "7 Dynamic text")
m("Loan Score", "MAX ( dim_loan[score] )", INT, "7 Dynamic text")
m("Loan Balance", "SUM ( dim_loan[balance_now] )", CUR, "7 Dynamic text")
m("Month Balance", "SUM ( fact_loan_month[balance_end] )", CUR, "7 Dynamic text")
m("Month DPD", "MAX ( fact_loan_month[days_past_due] )", INT, "7 Dynamic text")
m("Month Paid", "SUM ( fact_loan_month[amount_paid] )", CUR, "7 Dynamic text")
# Notes under the KPI values -------------------------------------------------------------
NOTES = {
    "Note Lent": 'FORMAT ( [Amount Lent] / 1e6, "$#,0" ) & "M lent · avg " & FORMAT ( [Avg Loan], "$#,0" )',
    "Note Open": 'FORMAT ( [Open Loans], "#,0" ) & " open loans · " & FORMAT ( [Arrears Share Now], "0.0%" ) & " in arrears"',
    "Note D24": '"Loans issued 2021 – 2023"',
    "Note Loss": '"Charged off " & FORMAT ( [Charged Off] / 1e6, "$#,0.0" ) & "M − recovered " & FORMAT ( [Recovered] / 1e6, "$#,0.0" ) & "M"',
    "Note Cov": '"ECL " & FORMAT ( [ECL] / 1e6, "$#,0.00" ) & "M · 31 Dec 2025"',
    "Note Rate": '"Weighted by amount lent"',
    "Note FICO": '"Default 620-659 vs 780+"',
    "Note Purpose": '[Riskiest Purpose] & ", within 24 months"',
    "Note Policy": 'FORMAT ( [Policy New 12m], "0.0%" ) & " vs " & FORMAT ( [Policy Rest 12m], "0.0%" ) & " (12-month default)"',
    "Note LGD": '"Mortgage " & FORMAT ( [LGD Mortgage], "0%" ) & " · renters " & FORMAT ( [LGD Rent], "0%" )',
    "Note V23": '"2023 loans · month 24"',
    "Note V21": '"2021 loans · month 24"',
    "Note R3060": 'FORMAT ( [Cure from 30], "0%" ) & " cure back to current"',
    "Note R6090": '"60 days late → default next month"',
    "Note Unemp": '"Default starts, per +1 pt"',
    "Note Gini": '"Out of time: 2023 loans"',
    "Note GiniG": '"The lender\'s A-G grade, same loans"',
    "Note KS": '"Max gap, defaults vs goods"',
    "Note Spread": '"Worst vs best of 10 bands"',
    "Note Truth": '"15 of 15 effects with the true sign"',
    "Note ECL": 'FORMAT ( [Coverage], "0.00%" ) & " of " & FORMAT ( [EAD] / 1e6, "$#,0.0" ) & "M open balance"',
    "Note S1": 'FORMAT ( [Coverage Stage 1], "0.00%" ) & " · 12-month PD · " & FORMAT ( [Loans Stage 1], "#,0" ) & " loans"',
    "Note S2": 'FORMAT ( [Coverage Stage 2], "0.0%" ) & " · lifetime PD · " & FORMAT ( [Loans Stage 2], "#,0" ) & " loans"',
    "Note S3": 'FORMAT ( [Coverage Stage 3], "0.0%" ) & " · defaulted · " & FORMAT ( [Loans Stage 3], "#,0" ) & " loans"',
    "Note BT": '"Actual in 2025: " & FORMAT ( [Backtest Actual], "#,0" )',
    "Note Rows": '"Loan extract rows"',
    "Note Dups": '"+ " & FORMAT ( [Test Loans Removed], "0" ) & " test loans removed"',
    "Note Fixes": '"Text income, dates, typos, codes"',
    "Note Clean": '"One row per loan"',
    "Note Checks": "[Checks Matched]",
    "Note Amount": 'SELECTEDVALUE ( dim_loan[term] ) & " · " & FORMAT ( SUM ( dim_loan[installment] ), "$#,0" ) & "/month"',
    "Note LRate": '"Grade " & SELECTEDVALUE ( dim_loan[grade] )',
    "Note LScore": '"Band " & SELECTEDVALUE ( dim_loan[score_band] ) & " · PD 24m " & FORMAT ( AVERAGE ( dim_loan[pd_24m] ), "0.0%" )',
    "Note LBal": '"At 31 Dec 2025"',
    "Note LECL": 'IF ( ISBLANK ( [ECL] ), "Closed loan", "PD 12m " & FORMAT ( [PD 12m (EAD-weighted)], "0.0%" ) & " · LGD " & FORMAT ( [LGD (EAD-weighted)], "0%" ) )',
}
for k_, v_ in NOTES.items():
    m(k_, v_, None, "9 Helpers", hidden=True)
m("Status Color", f'SWITCH ( SELECTEDVALUE ( dim_loan[status_now] ), "Current", "{PRIMARY}", "30 DPD", "{WARN}", "60 DPD", "{AMBER}", '
  f'"90+ DPD", "{BAD}", "Charged Off", "{BAD}", "{GOOD}" )', None, "8 Formatting", hidden=True)
m("Stage Color", f'SWITCH ( SELECTEDVALUE ( fact_ecl[stage] ), "Stage 1", "{PRIMARY}", "Stage 2", "{WARN}", "{BAD}" )',
  None, "8 Formatting", hidden=True)
m("Default Rate 12m (Policy)", "IF ( ISBLANK ( SELECTEDVALUE ( dim_loan[policy_group] ) ), BLANK (), [Default Rate 12m] )",
  PCT, "2 Risk")
m("Policy Color", f'IF ( SELECTEDVALUE ( dim_loan[policy_group] ) = "{POLICY[2]}", "{BAD}", "{COMPARE}" )', None, "8 Formatting", hidden=True)
m("DPD Color", f'VAR d = [Month DPD] RETURN SWITCH ( TRUE (), d >= 90, "{BAD}", d >= 60, "{AMBER}", d >= 30, "{WARN}", "{PRIMARY}" )',
  None, "8 Formatting", hidden=True)


def q(name):
    return name if name.replace("_", "").isalnum() else f"'{name}'"


def col_tmdl(table, name, dtype, hidden, extra):
    out = [f"\tcolumn {q(name)}", f"\t\tdataType: {dtype}"]
    if hidden:
        out.append("\t\tisHidden")
    for k, v in extra.items():
        out.append(f"\t\t{k}" if v is None else f"\t\t{k}: {v}")
    if dtype == I and "formatString" not in extra:
        out.append("\t\tformatString: 0")
    out += [f"\t\tlineageTag: {gid(table, name)}", "\t\tsummarizeBy: none", f"\t\tsourceColumn: {name}",
            "", "\t\tannotation SummarizationSetBy = User", ""]
    return out


def m_partition(table):
    csv, columns, steps_extra, _ = TABLES[table]
    if csv is None:
        return CUSTOM_M[table]
    types = ", ".join(f'{{"{n}", {MTYPE[t]}}}' for n, t, _, _ in columns if t != B)
    content = (f'Binary.Decompress(File.Contents(DataFolder & "{csv}"), Compression.GZip)' if csv.endswith(".gz")
               else f'File.Contents(DataFolder & "{csv}")')
    steps = [
        ("Source", f'Csv.Document({content}, [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv])'),
        ("Promoted", "Table.PromoteHeaders(Source, [PromoteAllScalars = true])"),
        ("Kept", "Table.SelectColumns(Promoted, {" + ", ".join(f'"{n}"' for n, *_ in columns) + "})"),
        ("Blanks", 'Table.TransformColumns(Kept, {}, each if _ = "" then null else _)'),
        ("Typed", f'Table.TransformColumnTypes(Blanks, {{{types}}}, "en-US")'),
    ]
    prev = "Typed"
    for name, expr in steps_extra:
        steps.append((name, expr.format(prev=prev)))
        prev = name
    lines = [f"    {n} = {e}," for n, e in steps]
    lines[-1] = lines[-1].rstrip(",")
    return ["let"] + lines + ["in", f"    {prev}"]


def tmdl_table(table):
    _, columns, _, derived = TABLES[table]
    out = [f"table {table}", f"\tlineageTag: {gid(table)}", ""]
    for c in columns + derived:
        out += col_tmdl(table, *c)
    out += [f"\tpartition {table} = m", "\t\tmode: import", "\t\tsource ="]
    out += ["\t\t\t\t" + line for line in m_partition(table)]
    out += ["", "\tannotation PBI_ResultType = Table", ""]
    return "\n".join(out)


def tmdl_measures():
    out = ["table _Measures", f"\tlineageTag: {gid('_Measures')}", ""]
    for name, dax, fmt, folder, hidden in MEASURES:
        out.append(f"\tmeasure {q(name)} = {dax}")
        if fmt:
            out.append(f"\t\tformatString: {fmt}")
        if hidden:
            out.append("\t\tisHidden")
        out += [f"\t\tdisplayFolder: {folder}", f"\t\tlineageTag: {gid('m', name)}", ""]
    out += ["\tcolumn Value", "\t\tdataType: string", "\t\tisHidden", f"\t\tlineageTag: {gid('_Measures', 'Value')}",
            "\t\tsummarizeBy: none", "\t\tsourceColumn: Value", "",
            "\t\tannotation SummarizationSetBy = Automatic", "",
            "\tpartition _Measures = m", "\t\tmode: import", "\t\tsource =",
            '\t\t\t\t#table(type table [Value = text], {})', "",
            "\tannotation PBI_ResultType = Table", ""]
    return "\n".join(out)


def build_semantic_model():
    if SM.exists():
        shutil.rmtree(SM)
    write_json(SM / "definition.pbism", {
        "$schema": f"{SCHEMA}/item/semanticModel/definitionProperties/1.0.0/schema.json",
        "version": "4.2", "settings": {}})
    write_json(SM / ".platform", {
        "$schema": f"{SCHEMA}/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "SemanticModel", "displayName": NAME},
        "config": {"version": "2.0", "logicalId": gid("sm")}})
    d = SM / "definition"
    write(d / "database.tmdl", "database\n\tcompatibilityLevel: 1600\n\n")
    tables = ["_Measures", *TABLES]
    write(d / "model.tmdl", "\n".join([
        "model Model", "\tculture: en-US", "\tdefaultPowerBIDataSourceVersion: powerBI_V3",
        "\tdiscourageImplicitMeasures", "\tsourceQueryCulture: en-US", "\tdataAccessOptions",
        "\t\tlegacyRedirects", "\t\treturnErrorValuesAsNull", "",
        f"annotation PBI_QueryOrder = {json.dumps(['DataFolder', *tables])}", "",
        "annotation __PBI_TimeIntelligenceEnabled = 0", "",
        *[f"ref table {t}" for t in tables], ""]))
    write(d / "expressions.tmdl", "\n".join([
        f'expression DataFolder = "{DATA_FOLDER}" meta [IsParameterQuery = true, Type = "Text", '
        "IsParameterQueryRequired = true]",
        f"\tlineageTag: {gid('DataFolder')}", "", "\tannotation PBI_ResultType = Text", ""]))
    rels = []
    for ft, fc, tt, tc, *active in RELATIONSHIPS:
        rels += [f"relationship {gid('rel', ft, fc, tt)}"]
        if active and not active[0]:
            rels += ["\tisActive: false"]
        rels += [f"\tfromColumn: {ft}.{fc}", f"\ttoColumn: {tt}.{tc}", ""]
    write(d / "relationships.tmdl", "\n".join(rels))
    write(d / "tables" / "_Measures.tmdl", tmdl_measures())
    for t in TABLES:
        write(d / "tables" / f"{t}.tmdl", tmdl_table(t))


# =========================================================================
# Theme
# =========================================================================
def sc(c):
    return {"solid": {"color": c}}


THEME = {
    "name": "Northbridge Executive",
    "dataColors": DATA_COLORS,
    "foreground": INK, "foregroundNeutralSecondary": SLATE, "foregroundNeutralTertiary": SLATE_2,
    "background": WHITE, "backgroundLight": "#F6F8FB", "backgroundNeutral": LINE,
    "tableAccent": PRIMARY, "good": GOOD, "neutral": WARN, "bad": BAD,
    "maximum": PRIMARY, "center": "#7FA3C7", "minimum": "#EAF0F7", "null": "#F1F3F7",
    "hyperlink": PRIMARY, "visitedHyperlink": PRIMARY,
    "textClasses": {
        "callout": {"fontSize": 24, "fontFace": FONT_SB, "color": INK},
        "title": {"fontSize": 11, "fontFace": FONT_SB, "color": INK},
        "header": {"fontSize": 10, "fontFace": FONT_SB, "color": "#334155"},
        "label": {"fontSize": 9, "fontFace": FONT, "color": "#475569"},
    },
    "visualStyles": {
        "*": {"*": {
            "background": [{"show": True, "color": sc(WHITE), "transparency": 0}],
            "border": [{"show": True, "color": sc(LINE), "radius": 12, "width": 1}],
            "dropShadow": [{"show": True, "color": sc(NAVY), "position": "Outer", "preset": "Custom",
                            "transparency": 90, "shadowBlur": 14, "shadowDistance": 3, "shadowSpread": 0,
                            "angle": 90}],
            "padding": [{"top": 12, "bottom": 10, "left": 14, "right": 14}],
            "title": [{"show": True, "fontColor": sc(INK), "fontSize": 11, "fontFamily": FONT_SB,
                       "titleWrap": True}],
            "subTitle": [{"fontColor": sc(SLATE), "fontSize": 9, "fontFamily": FONT}],
            "categoryAxis": [{"showAxisTitle": False, "gridlineStyle": "dotted", "labelColor": sc(SLATE),
                              "fontSize": 9, "concatenateLabels": False}],
            "valueAxis": [{"showAxisTitle": False, "gridlineStyle": "dotted", "gridlineColor": sc(LINE),
                           "labelColor": sc(SLATE), "fontSize": 9}],
            "legend": [{"show": True, "position": "TopLeft", "labelColor": sc("#475569"), "fontSize": 9,
                        "showTitle": False}],
            "labels": [{"color": sc("#334155"), "fontSize": 9}],
            "lineStyles": [{"strokeWidth": 3}],
            "outspacePane": [{"backgroundColor": sc(WHITE), "foregroundColor": sc(INK), "transparency": 0,
                              "border": True, "borderColor": sc(LINE)}],
        }},
        "page": {"*": {"background": [{"color": sc(CANVAS), "transparency": 0}],
                       "outspace": [{"color": sc("#E7EBF1"), "transparency": 0}]}},
        "card": {"*": {"labels": [{"color": sc(INK), "fontSize": 24, "fontFamily": FONT_SB}],
                       "categoryLabels": [{"show": False}],
                       "title": [{"alignment": "center", "fontSize": 9, "fontColor": sc(SLATE),
                                  "fontFamily": FONT_SB}]}},
        "tableEx": {"*": {
            "grid": [{"gridVertical": False, "gridHorizontal": True, "gridHorizontalColor": sc("#EEF1F6"),
                      "outlineColor": sc(LINE), "rowPadding": 5}],
            "columnHeaders": [{"fontColor": sc("#475569"), "backColor": sc("#F6F8FB"), "fontFamily": FONT_SB,
                               "fontSize": 9, "wordWrap": True}],
            "values": [{"fontColorPrimary": sc("#1E293B"), "backColorPrimary": sc(WHITE),
                        "fontColorSecondary": sc("#1E293B"), "backColorSecondary": sc("#FAFBFD"), "fontSize": 9}],
            "total": [{"fontFamily": FONT_SB, "backColor": sc("#F6F8FB"), "fontColor": sc(INK)}]}},
        "pivotTable": {"*": {
            "grid": [{"gridVertical": False, "gridHorizontal": True, "gridHorizontalColor": sc("#EEF1F6"),
                      "outlineColor": sc(LINE), "rowPadding": 5}],
            "columnHeaders": [{"fontColor": sc("#475569"), "backColor": sc("#F6F8FB"), "fontFamily": FONT_SB,
                               "fontSize": 9}],
            "rowHeaders": [{"fontColor": sc("#1E293B"), "fontSize": 9, "showExpandCollapseButtons": False,
                            "legacyStyleDisabled": True}],
            "values": [{"fontColorPrimary": sc("#1E293B"), "backColorPrimary": sc(WHITE),
                        "backColorSecondary": sc(WHITE), "fontSize": 9}],
            "subTotals": [{"fontFamily": FONT_SB, "backgroundColor": sc("#F6F8FB")}]}},
        "slicer": {"*": {
            "header": [{"show": True, "fontColor": sc(MIST), "textSize": 8, "fontFamily": FONT_SB,
                        "outline": "None"}],
            "items": [{"fontColor": sc(INK), "background": sc(WHITE), "textSize": 9, "outline": "None"}],
            "background": [{"show": False}], "border": [{"show": False}], "dropShadow": [{"show": False}],
            "padding": [{"top": 0, "bottom": 0, "left": 0, "right": 0}]}},
        "textbox": {"*": {"background": [{"show": False}], "border": [{"show": False}],
                          "dropShadow": [{"show": False}], "padding": [{"top": 0, "bottom": 0, "left": 0,
                                                                         "right": 0}]}},
        "actionButton": {"*": {"background": [{"show": False}], "border": [{"show": False}],
                               "dropShadow": [{"show": False}], "visualHeader": [{"show": False}]}},
    },
}


# =========================================================================
# Report helpers (PBIR)
# =========================================================================
def lit(v):
    if isinstance(v, bool):
        s = "true" if v else "false"
    elif isinstance(v, (int, float)):
        s = f"{v}D"
    else:
        s = "'" + str(v).replace("'", "''") + "'"
    return {"expr": {"Literal": {"Value": s}}}


def lit_int(v):
    return {"expr": {"Literal": {"Value": f"{v}L"}}}


def col(table, column):
    return {"Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": column}}


def mea(name):
    assert any(x[0] == name for x in MEASURES), f"unknown measure {name}"
    return {"Measure": {"Expression": {"SourceRef": {"Entity": "_Measures"}}, "Property": name}}


def ref(field):
    k = next(iter(field))
    return f"{field[k]['Expression']['SourceRef']['Entity']}.{field[k]['Property']}"


def proj(field, display=None):
    p = {"field": field, "queryRef": ref(field), "nativeQueryRef": field[next(iter(field))]["Property"]}
    if display:
        p["displayName"] = display
    return p


def solid(color):
    """Colour from a hex string or from a measure (conditional formatting)."""
    return {"solid": {"color": lit(color) if isinstance(color, str) else {"expr": color}}}


def text_or_measure(v):
    return lit(v) if isinstance(v, str) else {"expr": v}


def plain_container(**over):
    """No card chrome: for elements that sit on the sidebar or overlay a tile."""
    o = {"background": [{"properties": {"show": lit(False)}}],
         "border": [{"properties": {"show": lit(False)}}],
         "dropShadow": [{"properties": {"show": lit(False)}}],
         "visualHeader": [{"properties": {"show": lit(False)}}],
         "title": [{"properties": {"show": lit(False)}}],
         "padding": [{"properties": {"top": lit(0), "bottom": lit(0), "left": lit(0), "right": lit(0)}}]}
    o.update(over)
    return o


class Page:
    def __init__(self, key, title, kind=None, width=1280, height=720):
        self.key, self.name, self.title = key, hid("page", key), title
        self.kind, self.width, self.height = kind, width, height
        self.visuals = []

    def add(self, vid, x, y, w, h, visual, extra=None):
        v = {"$schema": f"{SCHEMA}/item/report/definition/visualContainer/{V_VISUAL}/schema.json",
             "name": hid(self.key, vid),
             "position": {"x": round(x, 2), "y": round(y, 2), "z": len(self.visuals) * 100,
                          "height": round(h, 2), "width": round(w, 2), "tabOrder": len(self.visuals) * 100},
             "visual": visual}
        if extra:
            v.update(extra)
        self.visuals.append(v)


def visual(vtype, roles=None, title=None, subtitle=None, sort=None, objects=None, container=None,
           tooltip_page=None):
    v = {"visualType": vtype, "drillFilterOtherVisuals": True}
    if roles:
        v["query"] = {"queryState": {r: {"projections": [proj(f) if "Column" in f or "Measure" in f
                                                         else proj(*f) for f in fs]}
                                     for r, fs in roles.items()}}
        if sort:
            v["query"]["sortDefinition"] = {"sort": [{"field": f, "direction": d} for f, d in sort],
                                            "isDefaultSort": False}
    if objects:
        v["objects"] = objects
    vco = dict(container or {})
    if title is not None:
        vco.setdefault("title", [{"properties": {"show": lit(True), "text": text_or_measure(title)}}])
    if subtitle:
        vco["subTitle"] = [{"properties": {"show": lit(True), "text": lit(subtitle)}}]
    if tooltip_page:
        vco["visualTooltip"] = [{"properties": {"show": lit(True), "type": lit("ReportPage"),
                                                "section": lit(tooltip_page)}}]
    if vco:
        v["visualContainerObjects"] = vco
    return v


def textbox(paragraphs, container=None, align="left"):
    """paragraphs: list of lists of (text, size, color, bold) runs."""
    paras = []
    for runs in paragraphs:
        paras.append({"horizontalTextAlignment": align, "textRuns": [{"value": t, "textStyle": {
            "fontSize": f"{size}pt", "color": color, "fontFamily": FONT_SB if bold else FONT,
            **({"fontWeight": "bold"} if bold else {})}} for t, size, color, bold in runs]})
    return {"visualType": "textbox", "drillFilterOtherVisuals": True,
            "objects": {"general": [{"properties": {"paragraphs": paras}}]},
            "visualContainerObjects": container or plain_container()}


def panel(color, radius=0):
    """A coloured rectangle (sidebar, accents) made from an empty textbox container."""
    return {"visualType": "textbox", "drillFilterOtherVisuals": True,
            "objects": {"general": [{"properties": {"paragraphs": [{"textRuns": [{"value": ""}]}]}}]},
            "visualContainerObjects": plain_container(
                background=[{"properties": {"show": lit(True), "color": solid(color), "transparency": lit(0)}}],
                border=[{"properties": {"show": lit(radius > 0), "color": solid(color), "radius": lit(radius)}}])}


def card(measure, *, size=24, color=INK, display=None, millions=False, exact=False, title=None,
         container=None, font=FONT_SB):
    lab = {"color": solid(color), "fontSize": lit(size), "fontFamily": lit(font)}
    if millions:
        lab.update(labelDisplayUnits=lit(1000000), labelPrecision=lit_int(2))
    elif exact:
        lab.update(labelDisplayUnits=lit(1))
    return visual("card", {"Values": [(mea(measure), display)] if display else [mea(measure)]}, title=title,
                  objects={"labels": [{"properties": lab}],
                           "categoryLabels": [{"properties": {"show": lit(False)}}]},
                  container=container)


def series_colors(colors):
    return [{"properties": {"fill": solid(c)}, "selector": {"metadata": ref(mea(n))}} for n, c in colors]


def measure_fill(name):
    return [{"properties": {"fill": solid(mea(name))},
             "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}]}}]


def cf(measure_ref, *, back=None, font=None):
    """Conditional formatting on one value column of a table/matrix (field-value rule)."""
    props = {}
    if back:
        props["backColor"] = solid(mea(back))
    if font:
        props["fontColor"] = solid(mea(font))
    return {"properties": props,
            "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}], "metadata": measure_ref}}


def labels(show=True, **extra):
    extra.setdefault("labelPosition", "OutsideEnd")
    return [{"properties": {"show": lit(show), **{k: (lit(v) if not isinstance(v, dict) else v)
                                                   for k, v in extra.items()}}}]


# =========================================================================
# Layout
# =========================================================================
W, H = 1280, 720
SIDEBAR_W = 208
X0, X1 = 224, 1264          # content area
GAP = 12
NAV = [("portfolio", "Portfolio"), ("risk", "Risk Drivers"), ("vintage", "Vintages & Economy"),
       ("scorecard", "Scorecard"), ("ecl", "IFRS 9 ECL"), ("quality", "Data Quality")]
SLICERS = [(col("dim_loan", "vintage_year"), "YEAR ISSUED", "year"),
           (col("dim_loan", "grade"), "GRADE", "grade"),
           (col("dim_loan", "channel"), "CHANNEL", "channel")]
PAGE_NAMES = {k: hid("page", k) for k, _ in NAV}
PAGE_NAMES["loan"] = hid("page", "loan")
PAGE_NAMES["tooltip"] = hid("page", "tooltip")


def sidebar(p, active):
    p.add("sb_bg", 0, 0, SIDEBAR_W, H, panel(NAVY))
    p.add("sb_brand", 20, 22, 180, 52, textbox([
        [("NORTHBRIDGE", 13, GOLD, True)],
        [("Credit Risk Analytics", 9, SLATE_2, False)]]))
    p.add("sb_rule", 20, 82, 168, 1, panel("#22395A"))
    p.add("sb_navlabel", 20, 94, 168, 16, textbox([[("DASHBOARDS", 7, "#7C8DA6", True)]]))
    for i, (key, label) in enumerate(NAV):
        y = 114 + i * 40
        is_active = key == active
        if is_active:
            p.add(f"nav_bg_{key}", 12, y, 184, 34, panel(NAVY_2, radius=8))
            p.add(f"nav_mark_{key}", 12, y + 8, 3, 18, panel(GOLD, radius=1))
        p.add(f"nav_label_{key}", 28, y + 8, 160, 20, textbox(
            [[(label, 10, WHITE if is_active else "#A9B6C9", is_active)]]))
        p.add(f"nav_{key}", 12, y, 184, 34, {
            "visualType": "actionButton", "drillFilterOtherVisuals": True,
            "objects": {
                "icon": [{"properties": {"show": lit(False)}}, {"properties": {"show": lit(False)}, "selector": {"id": "default"}}],
                "text": [{"properties": {"show": lit(False)}}, {"properties": {"show": lit(False)}, "selector": {"id": "default"}}],
                "fill": [{"properties": {"show": lit(False)}}, {"properties": {"show": lit(False)}, "selector": {"id": "default"}}],
                "outline": [{"properties": {"show": lit(False)}}, {"properties": {"show": lit(False)}, "selector": {"id": "default"}}]},
            "visualContainerObjects": {
                **plain_container(),
                "visualLink": [{"properties": {"show": lit(True), "type": lit("PageNavigation"),
                                               "navigationSection": lit(PAGE_NAMES[key])}}]}})
    y2 = 114 + len(NAV) * 40 + 10
    p.add("sb_rule2", 20, y2, 168, 1, panel("#22395A"))
    p.add("sb_filterlabel", 20, y2 + 12, 168, 16, textbox([[("FILTERS", 7, "#7C8DA6", True)]]))
    for i, (f, label, g) in enumerate(SLICERS):
        y = y2 + 30 + i * 58
        v = visual("slicer", {"Values": [(f, label)]}, objects={
            "data": [{"properties": {"mode": lit("Dropdown")}}],
            "header": [{"properties": {"show": lit(True), "fontColor": solid(MIST), "textSize": lit(8),
                                       "fontFamily": lit(FONT_SB)}}],
            "items": [{"properties": {"fontColor": solid(INK), "background": solid(WHITE), "textSize": lit(9)}}]},
            container=plain_container())
        v["syncGroup"] = {"groupName": g, "fieldChanges": True, "filterChanges": True}
        p.add(f"slicer_{g}", 20, y, 168, 54, v)
    p.add("sb_foot", 20, 640, 184, 66, textbox([
        [("Synthetic data · 50,000 loans", 7, "#7C8DA6", False)],
        [("Ctrl + click a menu item in Desktop", 7, "#7C8DA6", False)],
        [("Right-click a loan to drill through", 7, "#7C8DA6", False)]]))


def header(p, title):
    p.add("hdr_title", X0, 14, 620, 34, textbox([[(title, 18, INK, True)]]))
    p.add("hdr_context", X0, 46, 900, 22, visual(
        "card", {"Values": [mea("Blank Text")]},
        objects={"labels": [{"properties": {"fontSize": lit(1), "color": solid(CANVAS)}}],
                 "categoryLabels": [{"properties": {"show": lit(False)}}]},
        container=plain_container(title=[{"properties": {
            "show": lit(True), "text": {"expr": mea("Filter Context")}, "fontColor": solid(SLATE),
            "fontSize": lit(9), "fontFamily": lit(FONT), "alignment": lit("left")}}])))


def frame():
    """White rounded card with shadow (the tile background)."""
    return {"visualType": "textbox", "drillFilterOtherVisuals": True,
            "objects": {"general": [{"properties": {"paragraphs": [{"textRuns": [{"value": ""}]}]}}]},
            "visualContainerObjects": plain_container(
                background=[{"properties": {"show": lit(True), "color": solid(WHITE), "transparency": lit(0)}}],
                border=[{"properties": {"show": lit(True), "color": solid(LINE), "radius": lit(12)}}],
                dropShadow=[{"properties": {"show": lit(True), "color": solid(NAVY), "position": lit("Outer"),
                                            "preset": lit("Custom"), "transparency": lit(90), "shadowBlur": lit(14),
                                            "shadowDistance": lit(3), "shadowSpread": lit(0), "angle": lit(90)}}])}


def kpi_tiles(p, tiles, y=76, h=104):
    """tiles: (value measure, label, sub-measure, sub-colour measure or hex, number style)."""
    n = len(tiles)
    w = (X1 - X0 - (n - 1) * GAP) / n
    for i, (val, label, sub, sub_color, style) in enumerate(tiles):
        x = X0 + i * (w + GAP)
        p.add(f"kpi_{i}_frame", x, y, w, h, frame())
        p.add(f"kpi_{i}_rule", x + w / 2 - 14, y + 12, 28, 3, panel(GOLD))
        p.add(f"kpi_{i}_label", x + 8, y + 20, w - 16, 18,
              textbox([[(label.upper(), 7.5, SLATE, True)]], align="center"))
        p.add(f"kpi_{i}_value", x + 6, y + 38, w - 12, 40, card(
            val, size=22, millions=style == "M", exact=style == "N", container=plain_container()))
        is_hex = sub_color.startswith("#")
        sub_card = card(sub, size=8.5, font=FONT_SB, color=sub_color if is_hex else SLATE,
                        container=plain_container())
        if not is_hex:          # colour from a measure: teal = better, terracotta = worse
            sub_card["objects"]["labels"][0]["properties"]["color"] = solid(mea(sub_color))
        p.add(f"kpi_{i}_sub", x + 6, y + 78, w - 12, 20, sub_card)



# =========================================================================
# Pages
# =========================================================================
CW = X1 - X0                      # content width
MONTH = (col("dim_month", "month_date"), "Month")


def value_colors(column, mapping, literal=lambda v: f"'{v}'"):
    """Series colours by category value (e.g. status -> colour)."""
    return [{"properties": {"fill": solid(c)},
             "selector": {"data": [{"scopeId": {"Comparison": {"ComparisonKind": 0, "Left": column,
                                                                "Right": {"Literal": {"Value": literal(v)}}}}}]}}
            for v, c in mapping.items()]


def heat_cf(value_measure, back, font):
    return [cf(ref(mea(value_measure)), back=back, font=font)]


NO_TOTALS = {"total": [{"properties": {"totals": lit(False)}}]}
NO_SUBTOTALS = {"subTotals": [{"properties": {"rowSubtotals": lit(False), "columnSubtotals": lit(False)}}]}
STATUS_COLORS = {"Current": PRIMARY, "30 DPD": WARN, "60 DPD": AMBER, "90+ DPD": BAD}


def page_portfolio():
    p = Page("portfolio", "Portfolio")
    sidebar(p, "portfolio")
    header(p, "Northbridge Lending  ·  Consumer Loan Book")
    kpi_tiles(p, [
        ("Loans", "Loans issued", "Note Lent", SLATE, "N"),
        ("Open Balance", "Open balance, Dec 2025", "Note Open", SLATE, "M"),
        ("Default Rate 24m", "Default within 24 months", "Note D24", SLATE, ""),
        ("Net Loss", "Net credit loss to date", "Note Loss", SLATE, "M"),
        ("Coverage", "IFRS 9 loss allowance", "Note Cov", SLATE, "")])
    STATUS = (col("fact_loan_month", "status"), "Status")
    p.add("balance", X0, 192, 640, 226, visual(
        "columnChart", {"Category": [MONTH], "Series": [STATUS], "Y": [(mea("Book Balance"), "Balance")]},
        title="Open balance by payment status", subtitle="The book peaked in 2024; arrears (amber/red) grew with unemployment",
        objects={"dataPoint": value_colors(col("fact_loan_month", "status"), STATUS_COLORS),
                 "categoryAxis": [{"properties": {"axisType": lit("Scalar")}}],
                 "valueAxis": [{"properties": {"labelDisplayUnits": lit(1000000)}}]},
        tooltip_page=PAGE_NAMES["tooltip"]))
    p.add("orig", X0 + 640 + GAP, 192, CW - 640 - GAP, 226, visual(
        "clusteredColumnChart", {"Category": [MONTH], "Y": [(mea("Originations"), "Amount lent")]},
        title="New lending by month", subtitle="Grew ~1% a month, with a spring peak, until Dec 2024",
        objects={"dataPoint": [{"properties": {"fill": solid(PRIMARY)}}],
                 "valueAxis": [{"properties": {"labelDisplayUnits": lit(1000000)}}],
                 "categoryAxis": [{"properties": {"axisType": lit("Scalar")}}]}))
    p.add("grades", X0, 430, CW, 274, visual(
        "tableEx", {"Values": [(col("dim_loan", "grade"), "Grade"), (mea("Loans"), "Loans"), (mea("Amount Lent"), "Amount lent"),
                               (mea("Avg Rate"), "Avg rate"), (mea("Default Rate 24m"), "Default 24m"),
                               (mea("Net Loss Rate"), "Net loss % lent"), (mea("Open Balance"), "Open balance"),
                               (mea("ECL"), "ECL"), (mea("Coverage"), "ECL coverage")]},
        title="The book by grade", subtitle="Price, realised loss and the IFRS 9 allowance rise together from A to G",
        sort=[(col("dim_loan", "grade"), "Ascending")]))
    return p


def page_risk():
    p = Page("risk", "Risk Drivers")
    sidebar(p, "risk")
    header(p, "What Drives Default")
    kpi_tiles(p, [
        ("Default Rate 24m", "Default within 24 months", "Note D24", SLATE, ""),
        ("Avg Rate", "Average interest rate", "Note Rate", SLATE, ""),
        ("FICO Multiple", "Low vs high credit score", "Note FICO", BAD, ""),
        ("Riskiest Purpose Rate", "Riskiest loan purpose", "Note Purpose", BAD, ""),
        ("Policy Multiple", "DTI 40-50% policy (Jul 2023)", "Note Policy", BAD, "")])
    tips = [(mea("Loans"), "Loans"), (mea("Avg Rate"), "Avg rate")]
    PURPOSE = (col("dim_loan", "purpose"), "Purpose")
    p.add("purpose", X0, 192, 514, 250, visual(
        "clusteredBarChart", {"Category": [PURPOSE], "Y": [(mea("Default Rate 24m"), "Default within 24 months")], "Tooltips": tips},
        title="Default rate by loan purpose", subtitle="Small-business and medical loans are the riskiest; car loans the safest",
        sort=[(mea("Default Rate 24m"), "Descending")],
        objects={"dataPoint": [{"properties": {"fill": solid(PRIMARY)}}], "labels": labels(True, labelPrecision=lit_int(1)),
                 "categoryAxis": [{"properties": {"innerPadding": lit(12), "fontSize": lit(8)}}]}))
    p.add("policy", X0 + 514 + GAP, 192, CW - 514 - GAP, 250, visual(
        "clusteredBarChart", {"Category": [(col("dim_loan", "policy_group"), "Online loans")],
                              "Y": [(mea("Default Rate 12m (Policy)"), "Default within 12 months")], "Tooltips": tips},
        title="Online DTI limit raised from 40% to 50% in July 2023",
        subtitle="The newly allowed loans default 1.7x as often; their extra interest covers the extra loss",
        sort=[(col("dim_loan", "policy_group"), "Ascending")],
        objects={"dataPoint": measure_fill("Policy Color"), "labels": labels(True, labelPrecision=lit_int(1))}))
    w = (CW - 2 * GAP) / 3
    for i, (c, t, s) in enumerate([("fico_band", "FICO score", "Default falls steadily as the credit score rises"),
                                   ("dti_band", "Debt-to-income", "More of income already owed = more defaults"),
                                   ("income_band", "Annual income", "Higher income, lower risk")]):
        p.add(f"band_{c}", X0 + i * (w + GAP), 454, w, 250, visual(
            "clusteredColumnChart", {"Category": [(col("dim_loan", c), t)], "Y": [(mea("Default Rate 24m"), "Default within 24 months")],
                                     "Tooltips": tips},
            title=f"Default rate by {t.lower() if c != 'fico_band' else t}", subtitle=s,
            sort=[(col("dim_loan", c), "Ascending")],
            objects={"dataPoint": [{"properties": {"fill": solid(PRIMARY)}}], "labels": labels(True, labelPrecision=lit_int(1))}))
    return p


def page_vintage():
    p = Page("vintage", "Vintages & Economy")
    sidebar(p, "vintage")
    header(p, "Vintages, Roll Rates and the Economy")
    kpi_tiles(p, [
        ("Vintage 2023 at 24m", "2023 loans defaulted", "Note V23", BAD, ""),
        ("Vintage 2021 at 24m", "2021 loans defaulted", "Note V21", SLATE, ""),
        ("Roll 30 to 60", "30 → 60 days late", "Note R3060", SLATE, ""),
        ("Roll 60 to 90", "60 → 90+ days (default)", "Note R6090", SLATE, ""),
        ("Unemployment Sensitivity", "Unemployment effect", "Note Unemp", BAD, "")])
    YEAR = col("dim_loan", "vintage_year")
    p.add("vintage", X0, 192, 600, 512, visual(
        "lineChart", {"Category": [(col("dim_mob", "mob"), "Months on book")], "Series": [(YEAR, "Year issued")],
                      "Y": [(mea("Cum Default Rate"), "Cumulative default rate")]},
        title=mea("Title Vintage"),
        subtitle="Each line stops at the age its youngest loans have reached · 2023 loans met the 2024 unemployment peak early in life",
        sort=[(col("dim_mob", "mob"), "Ascending")],
        objects={"categoryAxis": [{"properties": {"axisType": lit("Scalar")}}],
                 "dataPoint": value_colors(YEAR, {2021: COMPARE, 2022: "#7A8CA5", 2023: BAD, 2024: GOLD}, literal=lambda v: f"{v}L")}))
    p.add("roll", X0 + 600 + GAP, 192, CW - 600 - GAP, 240, visual(
        "pivotTable", {"Rows": [(col("fact_loan_month", "status"), "This month")],
                       "Columns": [(col("fact_loan_month", "next_status"), "Next month")],
                       "Values": [(mea("Roll Rate"), "Share")]},
        title="Roll rates: where loans are next month", subtitle="Half of 30-day-late loans cure; half roll on toward default",
        objects={"values": heat_cf("Roll Rate", "Roll Heat", "Roll Heat Font"), **NO_SUBTOTALS,
                 "columnHeaders": [{"properties": {"fontSize": lit(8)}}]}))
    p.add("macro", X0 + 600 + GAP, 444, CW - 600 - GAP, 260, visual(
        "lineClusteredColumnComboChart", {"Category": [MONTH], "Y": [(mea("Defaults per 1,000"), "New defaults per 1,000 performing loans")],
                                          "Y2": [(mea("Unemployment"), "Unemployment")]},
        title="Default starts follow unemployment", subtitle="Bars: new defaults per 1,000 performing loans · line: unemployment",
        objects={"dataPoint": [{"properties": {"fill": solid(PRIMARY)}, "selector": {"metadata": ref(mea("Defaults per 1,000"))}},
                               {"properties": {"fill": solid(GOLD)}, "selector": {"metadata": ref(mea("Unemployment"))}}],
                 "categoryAxis": [{"properties": {"axisType": lit("Scalar")}}],
                 "legend": [{"properties": {"show": lit(False)}}]},
        tooltip_page=PAGE_NAMES["tooltip"]))
    return p


def page_scorecard():
    p = Page("scorecard", "Scorecard")
    sidebar(p, "scorecard")
    header(p, "Application Scorecard")
    kpi_tiles(p, [
        ("Gini Scorecard", "Scorecard Gini", "Note Gini", GOOD, ""),
        ("Gini Grade", "Lender grade Gini", "Note GiniG", SLATE, ""),
        ("KS Scorecard", "Kolmogorov-Smirnov", "Note KS", SLATE, ""),
        ("Band Spread", "Score band spread", "Note Spread", SLATE, ""),
        ("Truth Correlation", "Match with true effects", "Note Truth", GOOD, "")])
    p.add("calib", X0, 192, 600, 250, visual(
        "clusteredColumnChart", {"Category": [(col("score_bands", "band_label"), "Score band (1 = best)")],
                                 "Y": [(mea("Predicted PD"), "Predicted"), (mea("Actual Default"), "Actual")]},
        title="Default within 24 months by score band, 2023 loans",
        subtitle="Ranks well; under-predicts because unemployment rose after these loans were issued",
        sort=[(col("score_bands", "band_label"), "Ascending")],
        objects={"dataPoint": series_colors([("Predicted PD", COMPARE), ("Actual Default", PRIMARY)]),
                 "categoryAxis": [{"properties": {"fontSize": lit(7)}}]}))
    p.add("models", X0 + 600 + GAP, 192, CW - 600 - GAP, 250, visual(
        "tableEx", {"Values": [(col("model_metrics", "model"), "Model"), (col("model_metrics", "sample"), "Sample"),
                               (mea("AUC"), "AUC"), (mea("Gini"), "Gini"), (mea("KS"), "KS")]},
        title="Gradient boosting overfits: best in development, not on 2023 loans",
        sort=[(col("model_metrics", "model"), "Descending")], objects=NO_TOTALS))
    p.add("points", X0, 454, 600, 250, visual(
        "tableEx", {"Values": [(col("scorecard", "feature_label"), "Characteristic"), (col("scorecard", "bin_label"), "Value"),
                               (mea("Bin Loans"), "Loans (dev)"), (mea("Bin Bad Rate"), "Bad rate"), (mea("WoE"), "WoE"),
                               (mea("Points"), "Points")]},
        title="Scorecard points (600 = 50:1 odds, +20 points doubles the odds)",
        subtitle="Add up one row per characteristic to score an applicant",
        sort=[(col("scorecard", "bin_label"), "Ascending")], objects=NO_TOTALS))
    p.add("truth", X0 + 600 + GAP, 454, CW - 600 - GAP, 250, visual(
        "tableEx", {"Values": [(col("truth_check", "effect_label"), "Effect (log-odds)"), (mea("True Effect"), "True"),
                               (mea("Estimated Effect"), "Estimated"), (col("truth_check", "same_sign"), "Sign")]},
        title="Truth check: estimated vs the generator's true effects",
        subtitle="Plain logistic regression on the raw fields, 2021-23 loans", objects=NO_TOTALS))
    return p


def page_ecl():
    p = Page("ecl", "IFRS 9 ECL")
    sidebar(p, "ecl")
    header(p, "IFRS 9 Expected Credit Loss  ·  31 Dec 2025")
    kpi_tiles(p, [
        ("ECL", "Expected credit loss", "Note ECL", SLATE, "M"),
        ("ECL Stage 1", "Stage 1 · performing", "Note S1", SLATE, "M"),
        ("ECL Stage 2", "Stage 2 · credit risk up", "Note S2", WARN, "M"),
        ("ECL Stage 3", "Stage 3 · defaulted", "Note S3", BAD, "M"),
        ("Backtest Expected", "Backtest: defaults expected", "Note BT", SLATE, "N")])
    STAGE = (col("fact_ecl", "stage"), "Stage")
    p.add("stages", X0, 192, 612, 250, visual(
        "tableEx", {"Values": [STAGE, (mea("ECL Loans"), "Loans"), (mea("EAD"), "Exposure (EAD)"),
                               (mea("PD 12m (EAD-weighted)"), "PD 12m"), (mea("Lifetime PD (EAD-weighted)"), "Lifetime PD"),
                               (mea("LGD (EAD-weighted)"), "LGD"), (mea("ECL"), "ECL"), (mea("Coverage"), "Coverage")]},
        title="ECL = PD × LGD × EAD, by stage",
        subtitle="Stage 1: 12-month PD · Stage 2 (30+ days late or cured < 6 months): lifetime PD · Stage 3: 90+ days late",
        sort=[(col("fact_ecl", "stage"), "Ascending")]))
    p.add("scen", X0 + 612 + GAP, 192, CW - 612 - GAP, 250, visual(
        "clusteredColumnChart", {"Category": [(col("dim_scenario", "scenario"), "Scenario")], "Y": [(mea("Scenario ECL"), "ECL")]},
        title="Economic scenarios", subtitle="Unemployment −0.7 / flat / +1.5 pts, weighted 20 / 60 / 20%",
        sort=[(col("dim_scenario", "scenario"), "Ascending")],
        objects={"dataPoint": measure_fill("Scenario Color"),
                 "labels": labels(True, labelDisplayUnits=1000000, labelPrecision=lit_int(2)),
                 "categoryAxis": [{"properties": {"fontSize": lit(8)}}]}))
    p.add("migration", X0, 454, 500, 250, visual(
        "pivotTable", {"Rows": [(col("fact_stage_migration", "stage_2024"), "Dec 2024")],
                       "Columns": [(col("fact_stage_migration", "to_state"), "Dec 2025")],
                       "Values": [(mea("Migration Share"), "Share of loans")]},
        title="Stage migration, Dec 2024 → Dec 2025", subtitle="Where last year's loans ended up",
        objects={"values": heat_cf("Migration Share", "Migration Heat", "Migration Heat Font"), **NO_SUBTOTALS,
                 "columnHeaders": [{"properties": {"fontSize": lit(8), "wordWrap": lit(True)}}]}))
    p.add("top", X0 + 500 + GAP, 454, CW - 500 - GAP, 250, visual(
        "tableEx", {"Values": [(mea("ECL Rank"), "#"), (col("dim_loan", "loan_id"), "Loan"), (col("dim_loan", "sub_grade"), "Grade"),
                               (col("dim_loan", "stage_2025"), "Stage"), (mea("EAD"), "EAD"),
                               (mea("PD 12m (EAD-weighted)"), "PD 12m"), (mea("ECL"), "ECL")]},
        title="Largest expected losses", subtitle="Right-click a loan → Drill through → Loan Detail",
        sort=[(mea("ECL"), "Descending")], objects=NO_TOTALS))
    return p


def page_quality():
    p = Page("quality", "Data Quality")
    sidebar(p, "quality")
    header(p, "Data Quality  ·  Cleaning Log")
    kpi_tiles(p, [
        ("Rows Received", "Rows received", "Note Rows", SLATE, "N"),
        ("Duplicates Removed", "Duplicates removed", "Note Dups", BAD, "N"),
        ("Field Fixes", "Fields repaired", "Note Fixes", WARN, "N"),
        ("Clean Loans", "Clean loans", "Note Clean", SLATE, "N"),
        ("Fixes", "Rows touched by a fix", "Note Checks", GOOD, "N")])
    ISSUE = (col("dq_log", "issue"), "Issue")
    p.add("fixes", X0, 192, 480, 512, visual(
        "clusteredBarChart", {"Category": [ISSUE], "Y": [(mea("Fixes"), "Rows")], "Tooltips": [(mea("DQ Injected"), "Injected")]},
        title="What was wrong in the extract", subtitle="Red = rows removed · orange = values repaired or flagged",
        sort=[(mea("Fixes"), "Descending")],
        objects={"dataPoint": measure_fill("Fix Color"), "labels": labels(True, labelDisplayUnits=1),
                 "categoryAxis": [{"properties": {"fontSize": lit(8), "maxMarginFactor": lit(50)}}]}))
    p.add("log", X0 + 480 + GAP, 192, CW - 480 - GAP, 512, visual(
        "tableEx", {"Values": [ISSUE, (mea("DQ Rows"), "Rows"),
                               (mea("DQ Injected"), "Injected"), (mea("DQ Match"), "Match"), (col("dq_log", "action"), "Action")]},
        title="Every fix, checked against the number injected", subtitle="scripts/clean.py → dq_log · synthetic data: the truth is known",
        sort=[(col("dq_log", "issue"), "Ascending")],
        objects={"values": [cf(ref(mea("DQ Match")), font="Match Color")], **NO_TOTALS}))
    return p


def page_loan():
    p = Page("loan", "Loan Detail", kind="Drillthrough")
    sidebar(p, None)
    p.add("hdr_title", X0 + 48, 14, 900, 34, card("Title Loan", size=1, container=plain_container(
        title=[{"properties": {"show": lit(True), "text": {"expr": mea("Title Loan")}, "fontColor": solid(INK),
                               "fontSize": lit(18), "fontFamily": lit(FONT_SB), "alignment": lit("left")}}])))
    p.visuals[-1]["visual"]["objects"]["labels"][0]["properties"]["color"] = solid(CANVAS)
    p.add("back", X0, 16, 36, 30, {
        "visualType": "actionButton", "drillFilterOtherVisuals": True,
        "objects": {"icon": [{"properties": {"shapeType": lit("back"), "lineColor": solid(PRIMARY),
                                             "lineWeight": lit(2)}, "selector": {"id": "default"}}],
                    "fill": [{"properties": {"show": lit(True), "fillColor": solid(WHITE)}, "selector": {"id": "default"}}],
                    "outline": [{"properties": {"show": lit(False)}},
                                {"properties": {"show": lit(False)}, "selector": {"id": "default"}}]},
        "visualContainerObjects": {**plain_container(
            border=[{"properties": {"show": lit(True), "color": solid(LINE), "radius": lit(8)}}]),
            "visualLink": [{"properties": {"show": lit(True), "type": lit("Back")}}]}})
    tmp = Page("tmp", "")
    header(tmp, "")
    ctx = tmp.visuals[1]
    ctx["name"] = hid("loan", "hdr_context")
    p.visuals.append(ctx)
    kpi_tiles(p, [
        ("Amount Lent", "Loan amount", "Note Amount", SLATE, "N"),
        ("Loan Rate", "Interest rate", "Note LRate", SLATE, ""),
        ("Loan Score", "Application score", "Note LScore", SLATE, "N"),
        ("Loan Balance", "Balance at Dec 2025", "Loan Status", SLATE, "N"),
        ("ECL", "IFRS 9 ECL", "Note LECL", SLATE, "N")])
    p.add("history", X0, 192, CW, 262, visual(
        "lineClusteredColumnComboChart", {"Category": [MONTH], "Y": [(mea("Month DPD"), "Days past due")],
                                          "Y2": [(mea("Month Balance"), "Balance")]},
        title="Payment history", subtitle="Line: balance owed · bars: days past due",
        objects={"dataPoint": [{"properties": {"fill": solid(mea("DPD Color"))},
                                "selector": {"data": [{"dataViewWildcard": {"matchingOption": 1}}],
                                             "metadata": ref(mea("Month DPD"))}},
                               {"properties": {"fill": solid(PRIMARY)}, "selector": {"metadata": ref(mea("Month Balance"))}}],
                 "categoryAxis": [{"properties": {"axisType": lit("Scalar")}}],
                 "legend": [{"properties": {"show": lit(False)}}]}))
    p.add("profile", X0, 466, 430, 238, visual(
        "tableEx", {"Values": [(col("dim_loan", "channel"), "Channel"), (col("dim_loan", "home_ownership"), "Home"),
                               (col("dim_loan", "fico_band"), "FICO"), (col("dim_loan", "dti_band"), "DTI"),
                               (col("dim_loan", "income_band"), "Income")]},
        title="Applicant profile", subtitle="At application", objects=NO_TOTALS))
    p.add("months", X0 + 430 + GAP, 466, CW - 430 - GAP, 238, visual(
        "tableEx", {"Values": [(col("dim_month", "month_label"), "Month"), (col("fact_loan_month", "status"), "Status"),
                               (mea("Month DPD"), "Days late"), (mea("Month Paid"), "Paid"), (mea("Month Balance"), "Balance")]},
        title="Month by month", subtitle="From the loan_months table",
        sort=[(col("dim_month", "month_label"), "Ascending")], objects=NO_TOTALS))
    return p


def page_tooltip():
    p = Page("tooltip", "Tooltip · Month", kind="Tooltip", width=320, height=220)
    p.add("bg", 0, 0, 320, 220, panel(WHITE))
    p.add("title", 14, 10, 292, 30, card("Title Month", size=1, container=plain_container(
        title=[{"properties": {"show": lit(True), "text": {"expr": mea("Title Month")}, "fontColor": solid(INK),
                               "fontSize": lit(11), "fontFamily": lit(FONT_SB), "alignment": lit("left")}}])))
    p.visuals[-1]["visual"]["objects"]["labels"][0]["properties"]["color"] = solid(WHITE)
    for i, (mname, label) in enumerate([("Book Balance", "OPEN BALANCE"), ("Arrears Share", "IN ARREARS (30+)"),
                                        ("New Defaults", "NEW DEFAULTS"), ("Collections", "CASH COLLECTED")]):
        x, y = 14 + (i % 2) * 150, 46 + (i // 2) * 84
        p.add(f"t{i}", x, y, 142, 76, card(mname, size=15, title=label, exact=mname == "New Defaults"))
    return p


PAGES = [page_portfolio, page_risk, page_vintage, page_scorecard, page_ecl, page_quality, page_loan, page_tooltip]
DRILL_FIELD = col("dim_loan", "loan_id")


def top_loan():
    """Loan with the largest ECL at Dec 2025: the default for the drill-through page when opened directly."""
    best, val = None, -1.0
    with open(ROOT / "data_clean" / "powerbi" / "fact_ecl.csv", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            if r["asof"] == "2025-12" and float(r["ecl"]) > val:
                best, val = r["loan_id"], float(r["ecl"])
    return best


DEFAULT_LOAN = top_loan()


def build_report():
    if RP.exists():
        shutil.rmtree(RP)
    write_json(RP / "definition.pbir", {
        "$schema": f"{SCHEMA}/item/report/definitionProperties/2.0.0/schema.json",
        "version": "4.0", "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}}})
    write_json(RP / ".platform", {
        "$schema": f"{SCHEMA}/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "Report", "displayName": NAME},
        "config": {"version": "2.0", "logicalId": gid("report")}})
    base = RP / "StaticResources" / "SharedResources" / "BaseThemes" / f"{BASE_THEME}.json"
    base.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(THEME_SRC, base)
    write_json(RP / "StaticResources" / "RegisteredResources" / CUSTOM_THEME, THEME)
    write_json(PBI / "theme" / CUSTOM_THEME, THEME)        # standalone copy, importable in any report
    version = {"visual": V_VISUAL, "report": V_REPORT, "page": V_PAGE}
    d = RP / "definition"
    write_json(d / "version.json", {
        "$schema": f"{SCHEMA}/item/report/definition/versionMetadata/1.0.0/schema.json", "version": "2.0.0"})
    write_json(d / "report.json", {
        "$schema": f"{SCHEMA}/item/report/definition/report/{V_REPORT}/schema.json",
        "themeCollection": {
            "baseTheme": {"name": BASE_THEME, "reportVersionAtImport": version, "type": "SharedResources"},
            "customTheme": {"name": CUSTOM_THEME, "reportVersionAtImport": version, "type": "RegisteredResources"}},
        "resourcePackages": [
            {"name": "SharedResources", "type": "SharedResources",
             "items": [{"name": BASE_THEME, "path": f"BaseThemes/{BASE_THEME}.json", "type": "BaseTheme"}]},
            {"name": "RegisteredResources", "type": "RegisteredResources",
             "items": [{"name": CUSTOM_THEME, "path": CUSTOM_THEME, "type": "CustomTheme"}]}],
        "objects": {"outspacePane": [{"properties": {"expanded": lit(False)}}]},
        "settings": {"useStylableVisualContainerHeader": True, "defaultDrillFilterOtherVisuals": True,
                     "allowChangeFilterTypes": True, "useEnhancedTooltips": True,
                     "exportDataMode": "AllowSummarized"}})
    pages = [f() for f in PAGES]
    write_json(d / "pages" / "pages.json", {
        "$schema": f"{SCHEMA}/item/report/definition/pagesMetadata/1.0.0/schema.json",
        "pageOrder": [p.name for p in pages], "activePageName": pages[0].name})
    for p in pages:
        page = {"$schema": f"{SCHEMA}/item/report/definition/page/{V_PAGE}/schema.json",
                "name": p.name, "displayName": p.title, "displayOption": "FitToPage",
                "height": p.height, "width": p.width}
        if p.kind == "Drillthrough":
            fname, pname = hid("filter", "loan"), hid("param", "loan")
            page.update({
                "type": "Drillthrough", "visibility": "HiddenInViewMode",
                "pageBinding": {"name": hid("binding", "loan"), "type": "Drillthrough",
                                "parameters": [{"name": pname, "boundFilter": fname,
                                                "fieldExpr": DRILL_FIELD}]},
                "filterConfig": {"filters": [{"name": fname, "field": DRILL_FIELD, "type": "Categorical",
                                              # default loan when the page is opened directly;
                                              # a drill-through replaces it
                                              "filter": {"Version": 2,
                                                         "From": [{"Name": "s", "Entity": "dim_loan", "Type": 0}],
                                                         "Where": [{"Condition": {"In": {
                                                             "Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "s"}},
                                                                                         "Property": "loan_id"}}],
                                                             "Values": [[{"Literal": {"Value": "'" + DEFAULT_LOAN + "'"}}]]}}}]}}]}})
        elif p.kind == "Tooltip":
            page.update({"type": "Tooltip", "visibility": "HiddenInViewMode",
                         "pageBinding": {"name": hid("binding", "tooltip"), "type": "Tooltip",
                                         "parameters": []}})
        write_json(d / "pages" / p.name / "page.json", page)
        for v in p.visuals:
            write_json(d / "pages" / p.name / "visuals" / v["name"] / "visual.json", v)
    return pages


if __name__ == "__main__":
    build_semantic_model()
    pages = build_report()
    write_json(PBI / f"{NAME}.pbip", {
        "$schema": f"{SCHEMA}/pbip/pbipProperties/1.0.0/schema.json", "version": "1.0",
        "artifacts": [{"report": {"path": f"{NAME}.Report"}}], "settings": {"enableAutoRecovery": True}})
    write(PBI / ".gitignore", "**/.pbi/localSettings.json\n**/.pbi/cache.abf\n")
    print(f"wrote {PBI / (NAME + '.pbip')}")
    print(f"  model : {len(TABLES) + 1} tables, {len(MEASURES)} measures "
          f"({sum(1 for x in MEASURES if not x[4])} visible), {len(RELATIONSHIPS)} relationships")
    print(f"  report: {len(pages)} pages, {sum(len(p.visuals) for p in pages)} visuals")

