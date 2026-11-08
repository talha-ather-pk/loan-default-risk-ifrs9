"""
Clean the raw extracts and load them into SQLite (data_clean/northbridge.db).

Every fix is logged in the table dq_log with the number of rows it touched, next to the number the generator
injected (data_raw/true_effects.json), so the cleaning can be audited.

Tables: loans, loan_months, recoveries, macro, loan_outcomes, dq_log
Run: python scripts/clean.py
"""
from pathlib import Path
import json
import sqlite3
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RAW, CLEAN = ROOT / "data_raw", ROOT / "data_clean"
CLEAN.mkdir(exist_ok=True)
DB = CLEAN / "northbridge.db"
OBS_END = pd.Period("2025-12", "M")
inj = json.loads((RAW / "true_effects.json").read_text())["injected"]
log = []


def note(issue, rows, action, injected_key=None):
    log.append({"step": len(log) + 1, "table_name": "loans" if "month" not in action else "loan_months",
                "issue": issue, "rows_affected": int(rows), "action": action,
                "injected": inj.get(injected_key) if injected_key else None})
    print(f"{len(log):>2}. {issue:<45} {rows:>7,}  {action}")


# ------------------------------------------------------------------ loans
L = pd.read_csv(RAW / "loans.csv", dtype=str, keep_default_na=False, na_values=[""])
note("Rows received", len(L), "loaded")
n = len(L); L = L.drop_duplicates(); note("Exact duplicate rows", n - len(L), "removed", "duplicate_rows")
test = L.loan_id.str.startswith("TEST") | (L.loan_amnt.astype(float) <= 1)
note("Test loans (TEST- id, amount 1.00)", test.sum(), "removed", "test_loans"); L = L[~test].copy()
assert L.loan_id.is_unique, "loan_id must be unique after removing duplicates"

iso = pd.to_datetime(L.issue_d, format="%Y-%m-%d", errors="coerce")
txt = pd.to_datetime(L.issue_d, format="%b-%Y", errors="coerce")
note("Issue date as 'Apr-2023' (broker feed)", iso.isna().sum(), "parsed; day set to 1st, flagged",
     "issue_date_month_text_format (broker feed, day lost)")
assert (iso.notna() | txt.notna()).all()
L["issue_date"] = iso.fillna(txt)
L["issue_day_known"] = iso.notna().astype(int)

L["term_months"] = L.term.str.extract(r"(\d+)")[0].astype(int)
L["int_rate"] = L.int_rate.str.rstrip("%").astype(float)
L["revol_util"] = L.revol_util.str.rstrip("%").astype(float)

txt_inc = L.annual_inc.str.contains(r"[$,]", na=False)
note("Income stored as text ('$72,000.00')", txt_inc.sum(), "symbols stripped", "income_as_text_with_symbols")
L["annual_inc"] = pd.to_numeric(L.annual_inc.str.replace(r"[$,]", "", regex=True), errors="coerce")
note("Income missing", L.annual_inc.isna().sum(), "kept, flagged (imputed only inside the model)", "income_missing")

note("Employment length 'n/a'", (L.emp_length == "n/a").sum(), "set to null", "emp_length_na")
L["emp_years"] = L.emp_length.replace({"n/a": None, "< 1 year": "0", "10+ years": "10"}).str.extract(r"(\d+)")[0].astype(float)

clean_purpose = L.purpose.str.strip().str.lower().str.replace(" ", "_")
note("Purpose with odd case or spaces", (clean_purpose != L.purpose).sum(), "normalised", "purpose_case_or_spacing")
L["purpose"] = clean_purpose

TYPO = {"Nrth": "North", "Sout": "South", "Esat": "East", "Wets": "West", "Cental": "Central", "Coastl": "Coastal"}
note("Region typos", L.region.isin(TYPO).sum(), "mapped to the region", "region_typo")
L["region"] = L.region.replace(TYPO)

L["dti"] = pd.to_numeric(L.dti)
bad_dti = (L.dti < 0) | (L.dti > 100)
note("DTI outside 0-100 (-1 / 999 codes)", bad_dti.sum(), "set to null", "dti_invalid")
L.loc[bad_dti, "dti"] = np.nan
L["fico_score"] = pd.to_numeric(L.fico_score)
note("FICO score missing", L.fico_score.isna().sum(), "kept, flagged", "fico_missing")
odd_home = L.home_ownership.isin(["ANY", "NONE"])
note("Home ownership 'ANY' / 'NONE'", odd_home.sum(), "grouped as OTHER", "home_ownership_other")
L.loc[odd_home, "home_ownership"] = "OTHER"

for c in ["loan_amnt", "installment"]:
    L[c] = L[c].astype(float)
for c in ["inq_last_6mths", "delinq_2yrs"]:
    L[c] = L[c].astype(int)
L["issue_month"] = L.issue_date.dt.to_period("M").astype(str)
L["vintage_q"] = L.issue_date.dt.to_period("Q").astype(str)
L["vintage_year"] = L.issue_date.dt.year
loans = L[["loan_id", "issue_date", "issue_month", "vintage_q", "vintage_year", "issue_day_known", "channel", "region",
           "loan_amnt", "term_months", "int_rate", "installment", "grade", "sub_grade", "emp_years", "home_ownership",
           "annual_inc", "verification_status", "purpose", "dti", "fico_score", "revol_util", "inq_last_6mths",
           "delinq_2yrs"]].sort_values("loan_id").reset_index(drop=True)
loans["issue_date"] = loans.issue_date.dt.strftime("%Y-%m-%d")
note("Clean loans", len(loans), "loaded")

# ------------------------------------------------------------------ monthly performance
M = pd.read_csv(RAW / "loan_months.csv")
n = len(M); M = M.drop_duplicates(); note("Duplicate month rows", n - len(M), "removed (loan_months)", "loan_months_duplicate_rows")
assert not M.duplicated(["loan_id", "months_on_book"]).any()
orphan = ~M.loan_id.isin(loans.loan_id)
assert orphan.sum() == 0, f"{orphan.sum()} month rows without a loan"
R = pd.read_csv(RAW / "recoveries.csv")
macro = pd.read_csv(RAW / "macro.csv")

# ------------------------------------------------------------------ one outcome row per loan
g = M.groupby("loan_id")
last = M.loc[g.months_on_book.idxmax(), ["loan_id", "month", "months_on_book", "status", "days_past_due", "balance_end"]]
last.columns = ["loan_id", "last_month", "months_observed", "status_now", "dpd_now", "balance_now"]
first90 = M[M.status == "90+ DPD"].groupby("loan_id").agg(default_mob=("months_on_book", "min"), default_month=("month", "min"))
ever30 = M[M.days_past_due >= 30].groupby("loan_id").size().rename("months_delinquent")
co = M[M.status == "Charged Off"].set_index("loan_id")[["month", "charge_off_amount"]].rename(columns={"month": "charge_off_month"})
rec = R.groupby("loan_id").agg(recoveries=("recovery_amount", "sum"))
paid = g.amount_paid.sum().rename("total_paid")
O = (loans[["loan_id", "issue_month", "term_months", "loan_amnt"]].merge(last, on="loan_id", how="left")
     .merge(first90, on="loan_id", how="left").merge(ever30, on="loan_id", how="left")
     .merge(co, on="loan_id", how="left").merge(rec, on="loan_id", how="left").merge(paid, on="loan_id", how="left"))
O["months_delinquent"] = O.months_delinquent.fillna(0).astype(int)
O[["charge_off_amount", "recoveries"]] = O[["charge_off_amount", "recoveries"]].fillna(0.0)
O["defaulted"] = O.default_mob.notna().astype(int)
issue_p = pd.PeriodIndex(O.issue_month, freq="M")
months_available = (OBS_END - issue_p).map(lambda x: x.n)
closed = O.status_now.isin(["Fully Paid", "Prepaid", "Charged Off"])
for h in (12, 24):
    # an h-month outcome is defined only for loans issued at least h months before the data ends; counting early
    # defaults of younger loans (but not their non-defaults) would bias the rate upward
    O[f"default_{h}m"] = np.where(months_available >= h, (O.default_mob <= h).astype(int), np.nan)
O["months_available"] = months_available
O["is_open"] = (~closed).astype(int)
O = O.drop(columns=["issue_month", "term_months", "loan_amnt"])

with sqlite3.connect(DB) as con:
    for name, df in [("loans", loans), ("loan_months", M), ("recoveries", R), ("macro", macro),
                     ("loan_outcomes", O), ("dq_log", pd.DataFrame(log))]:
        df.to_sql(name, con, index=False, if_exists="replace")
    con.execute("CREATE INDEX IF NOT EXISTS ix_lm ON loan_months(loan_id, months_on_book)")
    con.execute("CREATE INDEX IF NOT EXISTS ix_lm_month ON loan_months(month)")
pd.DataFrame(log).to_csv(CLEAN / "dq_log.csv", index=False)
mismatch = [r for r in log if r["injected"] is not None and r["injected"] != r["rows_affected"]]
print(f"\n{DB.name}: {len(loans):,} loans, {len(M):,} loan-months, {len(R):,} recoveries; "
      f"{O.defaulted.sum():,} defaults. Fixes not matching the injected count: {len(mismatch)}")
assert not mismatch, mismatch
