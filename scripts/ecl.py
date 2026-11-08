"""
IFRS 9 expected credit loss (ECL) for the open loan book, at 31 Dec 2025 and (for a backtest) at 31 Dec 2024.

Everything is estimated only from data available at the reporting date (no look-ahead):
  * Staging   Stage 3 = 90+ days past due (credit-impaired). Stage 2 = 30-89 DPD, or cured from arrears in the
              last 6 months (probation, a significant increase in credit risk). Stage 1 = everything else.
  * PD        monthly default hazard by months on book, scaled by score band (from the scorecard) and by a
              forward-looking unemployment multiplier estimated from history. 12-month PD for Stage 1, lifetime
              PD (to maturity) for Stage 2, 100% for Stage 3. Loans already 30/60 DPD use observed roll rates.
  * LGD       1 - recovery rate on charge-offs with a complete 12-month workout, by home ownership.
  * EAD       outstanding balance at the reporting date.
  * Scenarios unemployment base (current level), upside (-0.7 pt), downside (+1.5 pt), weighted 60/20/20.
Simplifications (stated in the README): no discounting, no prepayment in the PD term structure, EAD = current balance.

Backtest: the 12-month PDs computed at Dec 2024 are compared with the defaults that actually happened in 2025.
Outputs (data_clean/): ecl_loans.csv, ecl_summary.csv, ecl_backtest.csv, ecl_parameters.json
Run: python scripts/ecl.py
"""
from pathlib import Path
import json
import sqlite3
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
CLEAN = ROOT / "data_clean"
SCEN = {"Upside": (-0.7, 0.2), "Base": (0.0, 0.6), "Downside": (1.5, 0.2)}

with sqlite3.connect(CLEAN / "northbridge.db") as con:
    LM = pd.read_sql("SELECT loan_id, month, months_on_book AS mob, status, days_past_due AS dpd, balance_end, "
                     "charge_off_amount FROM loan_months ORDER BY loan_id, months_on_book", con)
    loans = pd.read_sql("SELECT loan_id, grade, term_months, home_ownership, issue_month FROM loans", con)
    scores = pd.read_sql("SELECT loan_id, score_band FROM loan_scores", con)
    rec = pd.read_sql("SELECT loan_id, month, recovery_amount FROM recoveries", con)
    macro = pd.read_sql("SELECT month, unemployment_rate FROM macro", con).set_index("month").unemployment_rate

LM = LM.merge(scores, on="loan_id").merge(loans[["loan_id", "term_months", "home_ownership", "grade"]], on="loan_id")
# probation flag: the row is Current but the loan was 30+ DPD in one of the previous 6 months
delinq_mob = LM.mob.where(LM.dpd >= 30)
LM["last_delinq_mob"] = delinq_mob.groupby(LM.loan_id).ffill()
LM["probation"] = (LM.status == "Current") & (LM.mob - LM.last_delinq_mob <= 6)
LM["prev_status"] = LM.groupby("loan_id").status.shift().fillna("Current")
LM["prev_probation"] = LM.groupby("loan_id").probation.shift().fillna(False).astype(bool)
# A default "starts" when a current loan misses a payment (30 DPD) and goes on to reach 90+ DPD two months later.
# Hazard = chance that a loan that was Current at the start of the month starts a default that month.
grp = LM.groupby("loan_id").status
LM["next_status"] = grp.shift(-1)
LM["event"] = (LM.status == "30 DPD") & (grp.shift(-2) == "90+ DPD")
LM["at_risk"] = LM.prev_status == "Current"
MAX_MOB = 72


def estimate(asof):
    """Fit PD, LGD and roll-rate parameters on history known at month `asof` (no look-ahead)."""
    p = pd.Period(asof, "M")
    H = LM[(LM.month <= str(p - 2)) & LM.at_risk]           # outcome (90+ two months later) known by asof
    g = H.groupby("mob").event.agg(["sum", "size"])
    h = (g["sum"] / g["size"]).reindex(range(1, MAX_MOB + 1))
    tail = h.dropna().iloc[-6:].mean()                       # extrapolate beyond the oldest MOB observed
    h = h.fillna(tail).clip(lower=1e-5)
    exp_row = H.mob.map(h)
    rel = (H.groupby("score_band").event.sum() / exp_row.groupby(H.score_band).sum()).to_dict()
    exp_rows = exp_row * H.score_band.map(rel)
    prob = H.prev_probation
    prepay = float((H.status == "Prepaid").mean())          # monthly early-repayment rate (loan leaves the book)
    probation_mult = float(H[prob].event.sum() / exp_rows[prob].sum())
    # macro: actual / expected default starts per calendar month vs unemployment (log-linear, weighted)
    mon = pd.DataFrame({"a": H.event, "e": exp_rows, "m": H.month}).groupby("m").sum()
    mon = mon[mon.e > 20]
    u = macro.reindex(mon.index).to_numpy()
    y = np.log(np.maximum(mon.a, 0.5) / mon.e).to_numpy()
    w = mon.e.to_numpy()
    X = np.column_stack([np.ones_like(u), u])
    beta = np.linalg.solve(X.T @ (X * w[:, None]), X.T @ (w * y))
    u_bar = float(np.average(u, weights=w))
    # roll rates for loans already in arrears (next month known by asof)
    seq = LM[LM.month <= str(p - 1)]
    r30 = float((seq.next_status[seq.status == "30 DPD"] == "60 DPD").mean())
    r60 = float((seq.next_status[seq.status == "60 DPD"] == "90+ DPD").mean())
    # LGD: charge-offs with a complete 12-month workout before asof
    co = LM[(LM.status == "Charged Off") & (LM.month <= str(p - 12))][["loan_id", "charge_off_amount", "home_ownership"]]
    rr = rec[rec.month <= asof].groupby("loan_id").recovery_amount.sum()
    co = co.assign(rec=co.loan_id.map(rr).fillna(0))
    lgd = (1 - co.groupby("home_ownership").rec.sum() / co.groupby("home_ownership").charge_off_amount.sum()).to_dict()
    lgd["ALL"] = float(1 - co.rec.sum() / co.charge_off_amount.sum())
    return {"h": h, "rel": rel, "prepay": prepay, "probation_mult": probation_mult, "beta": float(beta[1]), "u_bar": u_bar,
            "roll_30_60": r30, "roll_60_90": r60, "lgd": lgd, "charge_offs_for_lgd": len(co)}


def pd_horizon(par, mob, band, months, mult):
    """Probability of default over `months` months after `mob`, for a performing loan."""
    h = par["h"].to_numpy()
    out = np.zeros(len(mob))
    for i, (m, b, n, k) in enumerate(zip(mob, band, months, mult)):
        if n <= 0:
            continue
        hh = np.clip(h[m: m + n] * par["rel"][b] * k, 0, 0.999)
        # default in month j requires surviving (no default, no early repayment) months before j
        alive = np.concatenate([[1.0], np.cumprod((1 - hh) * (1 - par["prepay"]))[:-1]])
        out[i] = float(np.sum(hh * alive))
    return out


def ecl_at(asof, u_override=None):
    par = estimate(asof)
    B = LM[(LM.month == asof) & LM.status.isin(["Current", "30 DPD", "60 DPD", "90+ DPD"])].copy()
    B["stage"] = np.select([B.status == "90+ DPD", B.status.isin(["30 DPD", "60 DPD"]) | B.probation], [3, 2], 1)
    B["ead"] = B.balance_end
    B["lgd"] = B.home_ownership.map(par["lgd"]).fillna(par["lgd"]["ALL"])
    remaining = (B.term_months - B.mob).clip(lower=1).to_numpy()
    u_now = float(macro[asof]) if u_override is None else u_override
    for s, (shift, _) in SCEN.items():
        mult_s = np.exp(par["beta"] * (u_now + shift - par["u_bar"]))
        k = np.where(B.probation, par["probation_mult"], 1.0) * mult_s
        # a default within 12 months must start (first missed payment) within 10 months
        pd12 = pd_horizon(par, B.mob.to_numpy(), B.score_band.to_numpy(), np.minimum(10, remaining), k)
        pdl = pd_horizon(par, B.mob.to_numpy(), B.score_band.to_numpy(), remaining, k)
        arrears = np.where(B.status == "30 DPD", par["roll_30_60"] * par["roll_60_90"],
                           np.where(B.status == "60 DPD", par["roll_60_90"], np.nan))
        pd12 = np.where(np.isnan(arrears), pd12, arrears)
        pdl = np.where(np.isnan(arrears), pdl, np.maximum(arrears, pdl))
        B[f"pd12_{s}"] = np.where(B.stage == 3, 1.0, pd12)
        B[f"pd_life_{s}"] = np.where(B.stage == 3, 1.0, pdl)
        B[f"ecl_{s}"] = np.where(B.stage == 1, B[f"pd12_{s}"], B[f"pd_life_{s}"]) * B.lgd * B.ead
    B["pd12"] = sum(B[f"pd12_{s}"] * w for s, (_, w) in SCEN.items())
    B["pd_life"] = sum(B[f"pd_life_{s}"] * w for s, (_, w) in SCEN.items())
    B["ecl"] = sum(B[f"ecl_{s}"] * w for s, (_, w) in SCEN.items())
    B["asof"] = asof
    return B, par, u_now


E25, P25, U25 = ecl_at("2025-12")
E24, P24, U24 = ecl_at("2024-12")
U25_ACTUAL = float(macro[[m for m in macro.index if m.startswith("2025")]].mean())
E24x, _, _ = ecl_at("2024-12", u_override=U25_ACTUAL)     # ex-post: Dec-2024 model with 2025's actual unemployment

# ---------------------------------------------------------------- backtest: Dec-2024 PDs vs 2025 defaults
with sqlite3.connect(CLEAN / "northbridge.db") as con:
    out = pd.read_sql("SELECT loan_id, default_month FROM loan_outcomes", con).set_index("loan_id").default_month
bt = E24[E24.stage < 3].copy()
bt["defaulted_2025"] = bt.loan_id.map(out).between("2025-01", "2025-12").astype(int)
bt["pd12_actual_u"] = E24x.set_index("loan_id").loc[bt.loan_id, "pd12_Base"].to_numpy()
back = (bt.groupby("score_band").agg(loans=("loan_id", "size"), predicted_base=("pd12_Base", "sum"),
                                     predicted_weighted=("pd12", "sum"), predicted_actual_unemployment=("pd12_actual_u", "sum"),
                                     actual=("defaulted_2025", "sum")).reset_index())
back["score_band"] = back.score_band.astype(str)
back.loc[len(back)] = ["All", len(bt), bt.pd12_Base.sum(), bt.pd12.sum(), bt.pd12_actual_u.sum(), bt.defaulted_2025.sum()]
pc = ["predicted_base", "predicted_weighted", "predicted_actual_unemployment"]
back[pc] = back[pc].astype(float).round(1)
stage_bt = bt.groupby("stage").agg(loans=("loan_id", "size"), predicted=("pd12", "sum"),
                                   predicted_actual_unemployment=("pd12_actual_u", "sum"), actual=("defaulted_2025", "sum")).round(1)

# ---------------------------------------------------------------- summaries
def summary(E):
    s = E.groupby(["asof", "stage"]).agg(loans=("loan_id", "size"), ead=("ead", "sum"), ecl=("ecl", "sum"),
                                          ecl_base=("ecl_Base", "sum"), ecl_downside=("ecl_Downside", "sum"),
                                          ecl_upside=("ecl_Upside", "sum")).reset_index()
    tot = s.drop(columns=["asof", "stage"]).sum().astype(object); tot["asof"] = E["asof"].iloc[0]; tot["stage"] = "Total"
    s = pd.concat([s.astype({"stage": str}), tot.to_frame().T], ignore_index=True)
    s["coverage_pct"] = (100 * s.ecl / s.ead).astype(float).round(2)
    return s


S = pd.concat([summary(E24), summary(E25)], ignore_index=True)
cols = ["asof", "loan_id", "grade", "score_band", "status", "stage", "mob", "ead", "lgd", "pd12", "pd_life", "ecl",
        "ecl_Base", "ecl_Upside", "ecl_Downside"]
pd.concat([E24[cols], E25[cols]]).round(5).to_csv(CLEAN / "ecl_loans.csv", index=False)
S.round(2).to_csv(CLEAN / "ecl_summary.csv", index=False)
back.to_csv(CLEAN / "ecl_backtest.csv", index=False)


def params(par, u):
    return {"monthly_prepayment_rate": round(par["prepay"], 4), "probation_multiplier": round(par["probation_mult"], 3), "unemployment_beta": round(par["beta"], 3),
            "unemployment_history_avg": round(par["u_bar"], 2), "unemployment_at_date": u,
            "roll_30_to_60": round(par["roll_30_60"], 3), "roll_60_to_90": round(par["roll_60_90"], 3),
            "lgd": {k: round(v, 3) for k, v in par["lgd"].items()}, "charge_offs_for_lgd": par["charge_offs_for_lgd"],
            "score_band_relative_risk": {int(k): round(v, 2) for k, v in sorted(par["rel"].items())}}


grade_cov = E25.groupby("grade").apply(lambda d: 100 * d.ecl.sum() / d.ead.sum(), include_groups=False).round(2).to_dict()
res = {"scenarios": {k: {"unemployment_shift_pts": v[0], "weight": v[1]} for k, v in SCEN.items()},
       "2025-12": params(P25, U25), "2024-12": params(P24, U24),
       "coverage_by_grade_2025_pct": grade_cov,
       "backtest_2025": {"loans": int(len(bt)), "predicted_weighted": round(float(bt.pd12.sum()), 1),
                         "predicted_base": round(float(bt.pd12_Base.sum()), 1),
                         "predicted_with_actual_2025_unemployment": round(float(bt.pd12_actual_u.sum()), 1),
                         "unemployment_2025_actual_avg": round(U25_ACTUAL, 2), "actual": int(bt.defaulted_2025.sum()),
                         "by_stage": stage_bt.reset_index().to_dict("records")}}
(CLEAN / "ecl_parameters.json").write_text(json.dumps(res, indent=2, default=float))
print(S.round(2).to_string(index=False))
print(back.to_string(index=False))
print(json.dumps(res, indent=1, default=float))
