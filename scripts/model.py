"""
Application scorecard: probability that a new loan defaults (90+ DPD) within 24 months.

* Development sample: loans issued 2021-2022. Out-of-time test: loans issued 2023 (all have 24 months of history).
* Scorecard: numeric fields binned, Weight of Evidence (WoE) per bin, logistic regression on the WoE values,
  converted to points (600 points = 50:1 good:bad odds, 20 points to double the odds).
* Benchmarks: the lender's own grade (A-G) and a gradient-boosting challenger.
* Truth check: a plain logistic regression on the raw fields, compared with the generator's true coefficients.

Outputs (data_clean/): scorecard.csv, loan_scores.csv, score_bands.csv, model_metrics.json, truth_check.csv
Run: python scripts/model.py
"""
from pathlib import Path
import json
import sqlite3
import numpy as np
import pandas as pd
from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import HistGradientBoostingClassifier
from sklearn.metrics import roc_auc_score, roc_curve

ROOT = Path(__file__).resolve().parents[1]
CLEAN = ROOT / "data_clean"
truth = json.loads((ROOT / "data_raw" / "true_effects.json").read_text())

with sqlite3.connect(CLEAN / "northbridge.db") as con:
    D = pd.read_sql("SELECT l.*, o.default_24m, o.default_12m FROM loans l JOIN loan_outcomes o USING (loan_id)", con)

D["log_income"] = np.log(D.annual_inc)
D["term_60"] = (D.term_months == 60).astype(int)
D["emp_band"] = np.select([D.emp_years.isna(), D.emp_years < 1, D.emp_years < 5], ["unknown", "<1y", "1-4y"], "5y+")
dev = D[D.vintage_year.isin([2021, 2022])].copy()
oot = D[D.vintage_year == 2023].copy()
y_dev, y_oot = dev.default_24m.astype(int), oot.default_24m.astype(int)

NUMERIC = {"fico_score": 6, "dti": 6, "revol_util": 5, "log_income": 5, "inq_last_6mths": None, "delinq_2yrs": None}
CATEG = ["term_60", "emp_band", "home_ownership", "verification_status", "purpose", "channel"]


def bins_for(col, q):
    if q is None:                                              # counts: 0, 1, 2, 3+
        return [-np.inf, 0.5, 1.5, 2.5, np.inf]
    edges = np.unique(np.quantile(dev[col].dropna(), np.linspace(0, 1, q + 1))[1:-1].round(2))
    return [-np.inf, *edges, np.inf]


def label(df, col):
    if col in NUMERIC:
        b = pd.cut(df[col], BINS[col], right=False)
        return b.astype(str).where(df[col].notna(), "missing")
    return df[col].astype(str)


BINS = {c: bins_for(c, q) for c, q in NUMERIC.items()}
FEATURES = list(NUMERIC) + CATEG
goods, bads = (y_dev == 0).sum(), (y_dev == 1).sum()
WOE, rows = {}, []
for f in FEATURES:
    lab = label(dev, f)
    t = pd.crosstab(lab, y_dev).reindex(columns=[0, 1], fill_value=0)
    woe = np.log(((t[0] + 0.5) / goods) / ((t[1] + 0.5) / bads))
    iv = (((t[0] / goods) - (t[1] / bads)) * woe).sum()
    WOE[f] = woe.to_dict()
    for b in t.index:
        rows.append({"feature": f, "bin": b, "loans": int(t.loc[b].sum()), "bad_rate_pct": round(100 * t.loc[b, 1] / t.loc[b].sum(), 2),
                     "woe": round(woe[b], 4), "iv_feature": round(iv, 4)})


def woe_matrix(df):
    return pd.DataFrame({f: label(df, f).map(WOE[f]).fillna(0.0) for f in FEATURES})


X_dev, X_oot = woe_matrix(dev), woe_matrix(oot)
lr = LogisticRegression(C=1.0, max_iter=2000).fit(X_dev, y_dev)
coef = dict(zip(FEATURES, lr.coef_[0]))

# points: score = offset + factor * ln(good/bad odds);  ln(odds) = -(b0 + sum b_i * woe_i)
PDO, BASE, BASE_ODDS = 20, 600, 50
factor = PDO / np.log(2)
offset = BASE - factor * np.log(BASE_ODDS)
n = len(FEATURES)
card = pd.DataFrame(rows)
card["points"] = (-(card.feature.map(coef) * card.woe + lr.intercept_[0] / n) * factor + offset / n).round(0).astype(int)


def score(df):
    pts = pd.DataFrame({f: label(df, f).map(card[card.feature == f].set_index("bin").points).fillna(
        round(offset / n - lr.intercept_[0] / n * factor)) for f in FEATURES})
    return pts.sum(axis=1)


def ks(y, p):
    fpr, tpr, _ = roc_curve(y, p)
    return float(np.max(tpr - fpr))


def metrics(y, p):
    auc = roc_auc_score(y, p)
    return {"auc": round(auc, 4), "gini": round(2 * auc - 1, 4), "ks": round(ks(y, p), 4)}


p_dev, p_oot = lr.predict_proba(X_dev)[:, 1], lr.predict_proba(X_oot)[:, 1]
grade_rank = {g: i for i, g in enumerate("ABCDEFG")}
gb_feats = ["fico_score", "dti", "revol_util", "log_income", "inq_last_6mths", "delinq_2yrs", "term_60", "emp_years"]
cat_dummies = lambda df: pd.get_dummies(df[["home_ownership", "verification_status", "purpose", "channel"]]).astype(int)
Xg_dev = pd.concat([dev[gb_feats], cat_dummies(dev)], axis=1)
Xg_oot = pd.concat([oot[gb_feats], cat_dummies(oot)], axis=1).reindex(columns=Xg_dev.columns, fill_value=0)
gb = HistGradientBoostingClassifier(max_depth=3, learning_rate=0.05, max_iter=300, random_state=0).fit(Xg_dev, y_dev)

M = {
    "target": "default (90+ DPD) within 24 months of issue",
    "development": {"loans": len(dev), "bad_rate": round(float(y_dev.mean()), 4), "issued": "2021-2022"},
    "out_of_time": {"loans": len(oot), "bad_rate": round(float(y_oot.mean()), 4), "issued": "2023"},
    "scorecard": {"dev": metrics(y_dev, p_dev), "oot": metrics(y_oot, p_oot)},
    "lender_grade": {"dev": metrics(y_dev, dev.grade.map(grade_rank)), "oot": metrics(y_oot, oot.grade.map(grade_rank))},
    "gradient_boosting": {"dev": metrics(y_dev, gb.predict_proba(Xg_dev)[:, 1]), "oot": metrics(y_oot, gb.predict_proba(Xg_oot)[:, 1])},
    "information_value": {f: round(float(card[card.feature == f].iv_feature.iloc[0]), 3) for f in FEATURES},
    "points_scaling": {"base_points": BASE, "base_odds": BASE_ODDS, "pdo": PDO},
}

# ---------------------------------------------------------------- score every loan, bands and calibration
D["score"] = score(D)
D["pd_24m"] = lr.predict_proba(woe_matrix(D))[:, 1]
band_edges = np.quantile(D.loc[D.vintage_year.isin([2021, 2022]), "score"], np.linspace(0, 1, 11))[1:-1]
D["score_band"] = 10 - np.searchsorted(band_edges, D.score, side="right")         # 1 = best ... 10 = worst
lo_hi = D.groupby("score_band").score.agg(["min", "max"])
D["score_range"] = D.score_band.map(lambda b: f"{int(lo_hi.loc[b, 'min'])}-{int(lo_hi.loc[b, 'max'])}")
ev = D[D.vintage_year == 2023]
bands = (ev.groupby(["score_band", "score_range"])
         .agg(loans=("loan_id", "size"), predicted_pd_pct=("pd_24m", "mean"), actual_default_pct=("default_24m", "mean"))
         .reset_index())
bands[["predicted_pd_pct", "actual_default_pct"]] = (bands[["predicted_pd_pct", "actual_default_pct"]] * 100).round(2)
M["calibration_oot"] = {"predicted_pct": round(100 * float(ev.pd_24m.mean()), 2), "actual_pct": round(100 * float(ev.default_24m.mean()), 2)}
top, bottom = bands.iloc[0], bands.iloc[-1]
M["band_spread_oot"] = round(float(bottom.actual_default_pct / top.actual_default_pct), 1)

# within-grade separation: does the score split loans the grade treats as equal?
ev_c = ev[ev.grade == "C"]
med = ev_c.score.median()
M["grade_c_split"] = {"low_score_default_pct": round(100 * ev_c[ev_c.score < med].default_24m.mean(), 2),
                      "high_score_default_pct": round(100 * ev_c[ev_c.score >= med].default_24m.mean(), 2)}

# ---------------------------------------------------------------- truth check: raw-field logistic regression
R = D[D.vintage_year.isin([2021, 2022, 2023])].dropna(subset=["fico_score", "dti", "annual_inc"]).copy()
Xr = pd.DataFrame({
    "fico_per_10pts_above_700": (R.fico_score - 700) / 10, "dti_per_pt": R.dti, "revol_util_per_10pts": R.revol_util / 10,
    "inquiry_each": R.inq_last_6mths, "delinq_2y_each": R.delinq_2yrs, "log_income_vs_66k": np.log(R.annual_inc / 66_000),
    "term_60": R.term_60, "emp_lt_1y": (R.emp_years < 1).astype(int),
    "not_verified": (R.verification_status == "Not Verified").astype(int),
    "home_MORTGAGE": (R.home_ownership == "MORTGAGE").astype(int), "purpose_small_business": (R.purpose == "small_business").astype(int),
    "purpose_medical": (R.purpose == "medical").astype(int), "purpose_car": (R.purpose == "car").astype(int),
    "channel_Broker": (R.channel == "Broker").astype(int), "channel_Branch": (R.channel == "Branch").astype(int)})
raw = LogisticRegression(C=1e6, max_iter=5000).fit(Xr, R.default_24m.astype(int))
tc = truth["true_model"]["coefficients"]
true_v = {k: tc[k] for k in tc if not isinstance(tc[k], dict)}
true_v.update({"home_MORTGAGE": tc["home"]["MORTGAGE"], "purpose_small_business": tc["purpose"]["small_business"],
               "purpose_medical": tc["purpose"]["medical"], "purpose_car": tc["purpose"]["car"],
               "channel_Broker": tc["channel"]["Broker"], "channel_Branch": tc["channel"]["Branch"]})
check = pd.DataFrame({"effect": Xr.columns, "true": [true_v[c] for c in Xr.columns], "estimated": raw.coef_[0].round(3)})
check["same_sign"] = np.sign(check.true) == np.sign(check.estimated)
M["truth_check"] = {"effects": len(check), "same_sign": int(check.same_sign.sum()),
                    "correlation": round(float(np.corrcoef(check.true, check.estimated)[0, 1]), 3)}

card.to_csv(CLEAN / "scorecard.csv", index=False)
bands.to_csv(CLEAN / "score_bands.csv", index=False)
check.to_csv(CLEAN / "truth_check.csv", index=False)
D[["loan_id", "score", "score_band", "score_range", "pd_24m"]].round({"pd_24m": 5}).to_csv(CLEAN / "loan_scores.csv", index=False)
(CLEAN / "model_metrics.json").write_text(json.dumps(M, indent=2))
with sqlite3.connect(CLEAN / "northbridge.db") as con:
    D[["loan_id", "score", "score_band", "score_range", "pd_24m"]].to_sql("loan_scores", con, index=False, if_exists="replace")
print(json.dumps({k: M[k] for k in ("scorecard", "lender_grade", "gradient_boosting", "calibration_oot", "band_spread_oot",
                                    "grade_c_split", "truth_check")}, indent=1))
print(bands.to_string(index=False))
print(check.to_string(index=False))
