"""
Synthetic loan book for "Northbridge Lending" (fictional consumer lender).

50,000 unsecured personal loans issued Jan 2021 - Dec 2024, followed month by month until 31 Dec 2025.
The true default process is known (logistic risk score + age curve + unemployment), so every finding,
the scorecard and the IFRS 9 backtest can be checked against the truth.

Outputs (data_raw/):
  loans.csv          application/loan extract, with the mess real extracts have (see INJECT below)
  loan_months.csv    one row per loan per month on book: status, days past due, payment, balance
  recoveries.csv     cash recovered after charge-off
  macro.csv          monthly unemployment and policy base rate
  true_effects.json  the true coefficients, policy change, macro sensitivity and injected-problem counts

Run: python scripts/generate_data.py   (seeded; ~1 min)
"""
from pathlib import Path
import json
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data_raw"
RAW.mkdir(exist_ok=True)
rng = np.random.default_rng(5)

N = 50_000
START = pd.Period("2021-01", "M")
N_ISSUE_MONTHS = 48                      # Jan 2021 - Dec 2024
OBS_END = 59                             # calendar index of Dec 2025 (Jan 2021 = 0)
cal = pd.period_range(START, periods=OBS_END + 1, freq="M")

# ---------------------------------------------------------------- macro
t = np.arange(OBS_END + 1)
unemp = np.interp(t, [0, 12, 26, 35, 41, 47, 53, 59], [5.6, 4.1, 4.0, 5.4, 6.0, 4.8, 4.4, 4.3])
unemp = np.round(unemp + rng.normal(0, 0.05, len(t)), 1)
base_rate = np.interp(t, [0, 14, 30, 47, 59], [0.25, 0.5, 5.25, 4.5, 3.75])
MACRO_BETA = 0.30                        # hazard multiplier = exp(beta * (unemployment - 4.5))
macro_mult = np.exp(MACRO_BETA * (unemp - 4.5))

# ---------------------------------------------------------------- originations
w = 1.012 ** np.arange(N_ISSUE_MONTHS)
season = {1: .92, 2: .95, 3: 1.05, 4: 1.06, 5: 1.06, 6: 1.02, 7: 1.0, 8: 1.0, 9: 1.01, 10: 1.02, 11: .98, 12: .88}
w = w * np.array([season[(START + i).month] for i in range(N_ISSUE_MONTHS)])
k = np.sort(rng.choice(N_ISSUE_MONTHS, N, p=w / w.sum()))          # issue month index
issue_day = rng.integers(1, 29, N)

channel = rng.choice(["Online", "Branch", "Broker"], N, p=[.55, .30, .15])
POLICY_IDX = 30                          # Jul 2023: online max DTI raised from 40% to 50%
dti_cap = np.where((channel == "Online") & (k >= POLICY_IDX), 50.0, 40.0)


def draw_until(fn, ok):
    x = fn(N)
    bad = ~ok(x)
    while bad.any():
        x[bad] = fn(bad.sum())
        bad = ~ok(x)
    return x


fico = draw_until(lambda n: np.round(rng.normal(708, 42, n)), lambda x: (x >= 620) & (x <= 850))
income = np.round(np.clip(np.exp(rng.normal(np.log(66_000), 0.5, N)), 18_000, 450_000), -2)
dti = rng.gamma(3.4, 6.2, N)
for _ in range(50):
    bad = dti > dti_cap
    if not bad.any():
        break
    dti[bad] = rng.gamma(3.4, 6.2, bad.sum())
dti = np.round(dti, 2)
revol = np.round(np.clip(rng.beta(2.2, 2.4, N) * 100, 0, 100), 1)
inq = rng.poisson(0.7, N)
delinq = rng.poisson(0.22, N)
emp = rng.choice(["< 1 year", "1 year", "2 years", "3 years", "4 years", "5 years", "6 years", "7 years",
                  "8 years", "9 years", "10+ years"], N,
                 p=[.08, .07, .09, .08, .07, .07, .06, .05, .05, .04, .34])
home = rng.choice(["RENT", "MORTGAGE", "OWN"], N, p=[.45, .42, .13])
PURPOSES = {"debt_consolidation": .52, "credit_card": .20, "home_improvement": .08, "major_purchase": .04,
            "car": .05, "medical": .03, "small_business": .03, "other": .05}
purpose = rng.choice(list(PURPOSES), N, p=list(PURPOSES.values()))
term = rng.choice([36, 60], N, p=[.72, .28])
verif = rng.choice(["Verified", "Source Verified", "Not Verified"], N, p=[.35, .37, .28])
region = rng.choice(["North", "South", "East", "West", "Central", "Coastal"], N, p=[.18, .2, .17, .16, .14, .15])
amount = np.round(np.clip(np.exp(rng.normal(np.log(14_000) + .4 * np.log(income / 66_000), .5)), 1_000, 40_000) / 25) * 25

# ---------------------------------------------------------------- true risk (log-odds of default over the term)
COEF = {
    "fico_per_10pts_above_700": -0.13, "dti_per_pt": 0.035, "revol_util_per_10pts": 0.08, "inquiry_each": 0.12,
    "delinq_2y_each": 0.28, "log_income_vs_66k": -0.40, "term_60": 0.45, "emp_lt_1y": 0.20, "not_verified": 0.10,
    "home": {"RENT": 0.0, "MORTGAGE": -0.18, "OWN": -0.06},
    "purpose": {"debt_consolidation": 0.0, "credit_card": -0.15, "home_improvement": -0.12, "major_purchase": 0.0,
                "car": -0.22, "medical": 0.30, "small_business": 0.60, "other": 0.15},
    "channel": {"Online": 0.0, "Branch": -0.10, "Broker": 0.25},
}
z = (COEF["fico_per_10pts_above_700"] * (fico - 700) / 10 + COEF["dti_per_pt"] * (dti - 20)
     + COEF["revol_util_per_10pts"] * (revol - 48) / 10 + COEF["inquiry_each"] * inq + COEF["delinq_2y_each"] * delinq
     + COEF["log_income_vs_66k"] * np.log(income / 66_000) + COEF["term_60"] * (term == 60)
     + COEF["emp_lt_1y"] * (emp == "< 1 year") + COEF["not_verified"] * (verif == "Not Verified")
     + np.vectorize(COEF["home"].get)(home) + np.vectorize(COEF["purpose"].get)(purpose)
     + np.vectorize(COEF["channel"].get)(channel))
# calibrate the intercept: average lifetime default probability 13% at average unemployment
lo, hi = -6.0, 2.0
for _ in range(60):
    mid = (lo + hi) / 2
    lo, hi = (mid, hi) if (1 / (1 + np.exp(-(z + mid)))).mean() < 0.13 else (lo, mid)
INTERCEPT = mid
z = z + INTERCEPT
p_life = 1 / (1 + np.exp(-z))

# lender's own grade: sees most of the risk, with noise; grade mix A 15 / B 28 / C 27 / D 16 / E 9 / F 4 / G 1
lender_score = z + rng.normal(0, 0.45, N)
cuts = np.quantile(lender_score, np.cumsum([.15, .28, .27, .16, .09, .04]))
gi = np.searchsorted(cuts, lender_score)
grade = np.array(list("ABCDEFG"))[gi]
sub = np.empty(N, int)
for g in range(7):
    m = gi == g
    sub[m] = np.minimum(5, 1 + (pd.Series(lender_score[m]).rank(pct=True).to_numpy() * 5).astype(int))
sub_grade = np.char.add(grade, sub.astype(str))
rate_base = np.array([7.0, 10.5, 14.0, 17.8, 21.5, 25.0, 28.5])
rate_trend = np.interp(k, [0, 17, 29, 47], [0.0, 0.0, 2.5, 2.0])
int_rate = np.round(rate_base[gi] + 0.7 * (sub - 1) + rate_trend + rng.normal(0, .15, N), 2)
r_m = int_rate / 1200
installment = np.round(amount * r_m / (1 - (1 + r_m) ** -term), 2)

# ---------------------------------------------------------------- monthly performance
mob_all = np.arange(1, 61)


def age_weights(T):
    w_ = mob_all[:T] ** 1.6 * np.exp(-mob_all[:T] / 7.5)
    return np.concatenate([w_ / w_.sum(), np.zeros(60 - T)])


AGE = {36: age_weights(36), 60: age_weights(60)}
cum_h = -np.log(1 - p_life)
prepay_h = np.where(np.isin(grade, ["A", "B"]), 0.017, 0.011)
tmp_h = np.clip(0.012 * np.exp(0.55 * (z - z.mean())), 0.002, 0.06)


def bal_after(n, ix):
    n = np.minimum(n, term[ix])
    g = (1 + r_m[ix]) ** n
    return np.maximum(amount[ix] * g - installment[ix] * (g - 1) / r_m[ix], 0)


state = np.zeros(N, int)        # 0 current 1 30dpd 2 60dpd 3 90dpd 4 120dpd | 5 charged off 6 paid 7 prepaid
bad = np.zeros(N, bool)
tmp_long = np.zeros(N, bool)
n_paid = np.zeros(N, int)
default_mob = np.full(N, -1)
co_mob = np.full(N, -1)
co_amount = np.zeros(N)
rows = []
LABEL = np.array(["Current", "30 DPD", "60 DPD", "90+ DPD", "90+ DPD", "Charged Off", "Fully Paid", "Prepaid"])
DPD = np.array([0, 30, 60, 90, 120, 0, 0, 0])

for mob in range(1, 70):
    c = k + mob
    active = (state < 5) & (c <= OBS_END)
    if not active.any():
        break
    idx = np.where(active)[0]
    st = state[idx]
    u = rng.random((4, len(idx)))
    paid = np.zeros(len(idx))
    prev_bal = bal_after(n_paid[idx], idx)
    m_idx = min(mob, 60) - 1
    haz = cum_h[idx] * np.where(term[idx] == 36, AGE[36][m_idx], AGE[60][m_idx]) * macro_mult[np.minimum(c[idx], OBS_END)]
    new = st.copy()

    cur = st == 0
    pre = cur & (u[0] < prepay_h[idx]) & (mob < term[idx])
    dflt = cur & ~pre & (u[1] < 1 - np.exp(-haz))
    tmp = cur & ~pre & ~dflt & (u[2] < tmp_h[idx]) & (mob < term[idx])
    pay = cur & ~pre & ~dflt & ~tmp
    new[pre] = 7; paid[pre] = prev_bal[pre] * (1 + r_m[idx][pre])
    new[dflt | tmp] = 1
    bad[idx[dflt]] = True
    tmp_long[idx[tmp]] = u[3][tmp] > 0.65
    paid[pay] = np.minimum(installment[idx][pay], prev_bal[pay] * (1 + r_m[idx][pay]))
    n_paid[idx[pay]] += 1

    d30, d60 = st == 1, st == 2
    b = bad[idx]
    cure30 = d30 & ~b & ~tmp_long[idx]
    roll30 = d30 & ~cure30
    cure60 = d60 & ~b
    new[cure30] = 0; n_paid[idx[cure30]] += 2; paid[cure30] = 2 * installment[idx][cure30]
    new[roll30] = 2
    new[cure60] = 0; n_paid[idx[cure60]] += 3; paid[cure60] = 3 * installment[idx][cure60]
    to90 = d60 & b
    new[to90] = 3; default_mob[idx[to90]] = mob
    new[st == 3] = 4
    co = st == 4
    new[co] = 5; co_mob[idx[co]] = mob; co_amount[idx[co]] = prev_bal[co]

    state[idx] = new
    full = (state[idx] == 0) & (n_paid[idx] >= term[idx])
    state[idx[full]] = 6
    bal = np.where(np.isin(state[idx], [5, 7]), 0, bal_after(n_paid[idx], idx))
    rows.append(pd.DataFrame({"i": idx, "c": c[idx], "mob": mob, "s": state[idx], "paid": np.round(paid, 2),
                              "bal": np.round(bal, 2)}))

lm = pd.concat(rows, ignore_index=True)
loan_ids = np.array([f"NB-{i:06d}" for i in range(1, N + 1)])
loan_months = pd.DataFrame({
    "loan_id": loan_ids[lm.i], "month": cal[lm.c].astype(str), "months_on_book": lm.mob,
    "status": LABEL[lm.s], "days_past_due": DPD[lm.s], "scheduled_payment": installment[lm.i],
    "amount_paid": lm.paid, "balance_end": lm.bal,
    "charge_off_amount": np.where(lm.s == 5, np.round(co_amount[lm.i], 2), 0.0)})

# ---------------------------------------------------------------- recoveries (12 months after charge-off)
REC_MEAN = {"RENT": 0.22, "MORTGAGE": 0.32, "OWN": 0.27}
co_idx = np.where(co_mob > 0)[0]
rec_rows = []
rw = np.array([.22, .16, .13, .1, .09, .07, .06, .05, .04, .03, .03, .02])
for i in co_idx:
    mu = REC_MEAN[home[i]]
    rate = rng.beta(mu * 6, (1 - mu) * 6)
    total = co_amount[i] * rate
    for j in range(12):
        c = k[i] + co_mob[i] + 1 + j
        if c > OBS_END:
            break
        amt = round(total * rw[j] * rng.uniform(.6, 1.4), 2)
        if amt > 0:
            rec_rows.append((loan_ids[i], str(cal[c]), amt))
recoveries = pd.DataFrame(rec_rows, columns=["loan_id", "month", "recovery_amount"])

# ---------------------------------------------------------------- loan extract (+ mess)
issue_date = pd.to_datetime([f"{START + int(m)}-{d:02d}" for m, d in zip(k, issue_day)])
loans = pd.DataFrame({
    "loan_id": loan_ids, "issue_d": issue_date.strftime("%Y-%m-%d"), "channel": channel, "region": region,
    "loan_amnt": amount, "term": np.char.add(np.char.add(" ", term.astype(str)), " months"),
    "int_rate": [f"{r:.2f}%" for r in int_rate], "installment": installment, "grade": grade, "sub_grade": sub_grade,
    "emp_length": emp, "home_ownership": home, "annual_inc": income.astype(object), "verification_status": verif,
    "purpose": purpose, "dti": dti.astype(object), "fico_score": fico.astype(object),
    "revol_util": [f"{r:.1f}%" for r in revol], "inq_last_6mths": inq, "delinq_2yrs": delinq})

inj = {}
broker = np.where(channel == "Broker")[0]                       # broker system exports dates as "Apr-2023"
loans.loc[broker, "issue_d"] = issue_date[broker].strftime("%b-%Y")
inj["issue_date_month_text_format (broker feed, day lost)"] = len(broker)
pool = rng.permutation(np.setdiff1d(np.arange(N), broker))   # each problem goes into its own rows
pos = 0


def pick(n):
    global pos
    sel = pool[pos:pos + n]
    pos += n
    return sel


s = pick(1500); loans.loc[s, "annual_inc"] = [f"${v:,.2f}" for v in income[s]]; inj["income_as_text_with_symbols"] = len(s)
s = pick(300); loans.loc[s, "annual_inc"] = np.nan; inj["income_missing"] = len(s)
s = pick(1750); loans.loc[s, "emp_length"] = "n/a"; inj["emp_length_na"] = len(s)
s = pick(2500)
loans.loc[s, "purpose"] = [rng.choice([p.replace("_", " ").title(), " " + p, p.upper()]) for p in purpose[s]]
inj["purpose_case_or_spacing"] = len(s)
TYPO = {"North": "Nrth", "South": "Sout", "East": "Esat", "West": "Wets", "Central": "Cental", "Coastal": "Coastl"}
s = pick(150); loans.loc[s, "region"] = [TYPO[r] for r in region[s]]; inj["region_typo"] = len(s)
s = pick(60); loans.loc[s, "dti"] = rng.choice([-1.0, 999.0], len(s)); inj["dti_invalid"] = len(s)
s = pick(200); loans.loc[s, "fico_score"] = np.nan; inj["fico_missing"] = len(s)
s = pick(25); loans.loc[s, "home_ownership"] = rng.choice(["ANY", "NONE"], len(s)); inj["home_ownership_other"] = len(s)

test = loans.sample(12, random_state=1).copy()
test["loan_id"] = [f"TEST-{i:04d}" for i in range(1, 13)]
test["loan_amnt"] = 1.0
inj["test_loans"] = len(test)
dups = loans.sample(400, random_state=2)
inj["duplicate_rows"] = len(dups)
loans_out = pd.concat([loans, test, dups]).sample(frac=1, random_state=3).reset_index(drop=True)

lm_dup = loan_months.sample(300, random_state=4)
inj["loan_months_duplicate_rows"] = len(lm_dup)
loan_months_out = pd.concat([loan_months, lm_dup]).sort_values(["loan_id", "months_on_book"], kind="stable")

macro = pd.DataFrame({"month": cal.astype(str), "unemployment_rate": unemp, "base_rate": np.round(base_rate, 2)})

loans_out.to_csv(RAW / "loans.csv", index=False)
loan_months_out.to_csv(RAW / "loan_months.csv", index=False)
recoveries.to_csv(RAW / "recoveries.csv", index=False)
macro.to_csv(RAW / "macro.csv", index=False)

# ---------------------------------------------------------------- truth
ever_def = default_mob > 0
online_post = (channel == "Online") & (k >= POLICY_IDX)
truth = {
    "seed": 5, "loans": N, "issue_period": "2021-01..2024-12", "observed_until": "2025-12",
    "true_model": {"intercept": round(INTERCEPT, 4), "coefficients": COEF,
                   "meaning": "log-odds of default over the full term, at 4.5% unemployment; spread over months on "
                              "book by an age curve peaking around month 12"},
    "average_true_lifetime_pd": round(float(p_life.mean()), 4),
    "macro": {"hazard_multiplier": "exp(0.30 * (unemployment - 4.5))", "unemployment_peak": float(unemp.max()),
              "peak_month": str(cal[int(unemp.argmax())])},
    "policy_change": {"month": "2023-07", "rule": "Online channel max DTI raised from 40% to 50%",
                      "loans_with_dti_40_50": int(((dti > 40) & online_post).sum())},
    "lender_grade_noise_sd": 0.45,
    "recovery_rate_mean_by_home_ownership": REC_MEAN,
    "default_definition": "90+ days past due (third missed payment); charged off at 120 DPD",
    "observed": {"defaults_90dpd": int(ever_def.sum()), "charged_off": int((co_mob > 0).sum()),
                 "default_rate_observed": round(float(ever_def.mean()), 4),
                 "loan_months_rows": len(loan_months), "recovery_rows": len(recoveries)},
    "injected": inj,
    "expected_clean_loans": N,
}
(RAW / "true_effects.json").write_text(json.dumps(truth, indent=2))
print(json.dumps({k_: truth[k_] for k_ in ("average_true_lifetime_pd", "observed", "policy_change", "injected")}, indent=1))
print(f"loans.csv {len(loans_out):,} rows | loan_months.csv {len(loan_months_out):,} | recoveries {len(recoveries):,}")
