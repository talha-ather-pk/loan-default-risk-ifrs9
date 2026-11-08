-- Northbridge Lending: credit risk questions, answered in SQL (SQLite) on data_clean/northbridge.db.
-- Default = 90+ days past due. Rates on "24-month" basis use loans issued up to Dec 2023 (all have 24 months of history).
-- Run all: python scripts/run_sql.py  (writes sql/results.md)

-- name: 01_portfolio_summary
-- How big is the book and what has it lost so far?
SELECT COUNT(*)                                                  AS loans,
       ROUND(SUM(l.loan_amnt) / 1e6, 1)                          AS originated_m,
       ROUND(AVG(l.int_rate), 2)                                 AS avg_rate_pct,
       SUM(o.defaulted)                                          AS defaults,
       ROUND(100.0 * AVG(o.defaulted), 2)                        AS default_rate_pct,
       ROUND(SUM(o.charge_off_amount) / 1e6, 2)                  AS charged_off_m,
       ROUND(SUM(o.recoveries) / 1e6, 2)                         AS recovered_m,
       ROUND((SUM(o.charge_off_amount) - SUM(o.recoveries)) / 1e6, 2) AS net_loss_m,
       SUM(o.is_open)                                            AS open_loans,
       ROUND(SUM(CASE WHEN o.is_open = 1 THEN o.balance_now END) / 1e6, 1) AS open_balance_m
FROM loans l JOIN loan_outcomes o USING (loan_id);

-- name: 02_default_by_grade
-- Does the lender's grade rank risk, and is the price right for it?
SELECT l.grade,
       COUNT(*)                                       AS loans,
       ROUND(AVG(l.int_rate), 2)                      AS avg_rate_pct,
       ROUND(100.0 * AVG(o.default_24m), 2)           AS default_24m_pct,
       ROUND(100.0 * SUM(o.charge_off_amount - o.recoveries) / SUM(l.loan_amnt), 2) AS net_loss_pct_of_amount
FROM loans l JOIN loan_outcomes o USING (loan_id)
WHERE l.issue_month <= '2023-12'
GROUP BY l.grade ORDER BY l.grade;

-- name: 03_default_by_purpose
-- Which loan purposes are riskiest?
SELECT l.purpose,
       COUNT(*)                                       AS loans,
       ROUND(100.0 * AVG(o.default_24m), 2)           AS default_24m_pct,
       ROUND(AVG(l.int_rate), 2)                      AS avg_rate_pct
FROM loans l JOIN loan_outcomes o USING (loan_id)
WHERE l.issue_month <= '2023-12'
GROUP BY l.purpose ORDER BY default_24m_pct DESC;

-- name: 04_fico_and_dti_bands
-- How does default change with credit score and debt-to-income?
WITH b AS (
  SELECT o.default_24m,
         CASE WHEN l.fico_score IS NULL THEN 'missing'
              WHEN l.fico_score < 660 THEN '620-659' WHEN l.fico_score < 700 THEN '660-699'
              WHEN l.fico_score < 740 THEN '700-739' WHEN l.fico_score < 780 THEN '740-779' ELSE '780+' END AS fico_band,
         CASE WHEN l.dti IS NULL THEN 'missing' WHEN l.dti < 10 THEN '00-10' WHEN l.dti < 20 THEN '10-20'
              WHEN l.dti < 30 THEN '20-30' WHEN l.dti < 40 THEN '30-40' ELSE '40-50' END AS dti_band
  FROM loans l JOIN loan_outcomes o USING (loan_id) WHERE l.issue_month <= '2023-12')
SELECT 'FICO' AS factor, fico_band AS band, COUNT(*) AS loans, ROUND(100.0 * AVG(default_24m), 2) AS default_24m_pct
FROM b GROUP BY fico_band
UNION ALL
SELECT 'DTI', dti_band, COUNT(*), ROUND(100.0 * AVG(default_24m), 2) FROM b GROUP BY dti_band
ORDER BY factor, band;

-- name: 05_vintage_curves
-- Cumulative default rate by issue year at 6, 12, 18 and 24 months on book (only vintages old enough for each point).
WITH v AS (
  SELECT l.vintage_year, o.default_mob, o.months_available FROM loans l JOIN loan_outcomes o USING (loan_id))
SELECT CAST(vintage_year AS TEXT) AS vintage,
       COUNT(*) AS loans,
       ROUND(100.0 * AVG(CASE WHEN default_mob <= 6  THEN 1.0 ELSE 0 END), 2) AS mob_6_pct,
       ROUND(100.0 * AVG(CASE WHEN default_mob <= 12 THEN 1.0 ELSE 0 END), 2) AS mob_12_pct,
       CASE WHEN MIN(months_available) >= 18 THEN ROUND(100.0 * AVG(CASE WHEN default_mob <= 18 THEN 1.0 ELSE 0 END), 2) END AS mob_18_pct,
       CASE WHEN MIN(months_available) >= 24 THEN ROUND(100.0 * AVG(CASE WHEN default_mob <= 24 THEN 1.0 ELSE 0 END), 2) END AS mob_24_pct
FROM v GROUP BY vintage_year ORDER BY vintage_year;

-- name: 06_roll_rates
-- Of the loans in each status this month, where are they next month? (all months, Jan 2021 - Dec 2025)
WITH t AS (
  SELECT a.status AS from_status, b.status AS to_status
  FROM loan_months a JOIN loan_months b
    ON b.loan_id = a.loan_id AND b.months_on_book = a.months_on_book + 1
  WHERE a.status IN ('Current', '30 DPD', '60 DPD'))
SELECT from_status,
       COUNT(*) AS loan_months,
       ROUND(100.0 * AVG(to_status = 'Current'), 2)  AS to_current_pct,
       ROUND(100.0 * AVG(to_status = '30 DPD'), 2)   AS to_30_pct,
       ROUND(100.0 * AVG(to_status = '60 DPD'), 2)   AS to_60_pct,
       ROUND(100.0 * AVG(to_status = '90+ DPD'), 2)  AS to_90_pct,
       ROUND(100.0 * AVG(to_status IN ('Prepaid', 'Fully Paid')), 2) AS to_closed_pct
FROM t GROUP BY from_status
ORDER BY CASE from_status WHEN 'Current' THEN 0 WHEN '30 DPD' THEN 1 ELSE 2 END;

-- name: 07_loss_given_default
-- How much of a charged-off balance comes back? Only charge-offs with a full 12-month workout (charged off by Dec 2024).
SELECT l.home_ownership,
       COUNT(*)                                                   AS charge_offs,
       ROUND(SUM(o.charge_off_amount) / 1e6, 2)                   AS charged_off_m,
       ROUND(100.0 * SUM(o.recoveries) / SUM(o.charge_off_amount), 1)       AS recovery_rate_pct,
       ROUND(100.0 - 100.0 * SUM(o.recoveries) / SUM(o.charge_off_amount), 1) AS lgd_pct
FROM loans l JOIN loan_outcomes o USING (loan_id)
WHERE o.charge_off_month <= '2024-12'
GROUP BY l.home_ownership
UNION ALL
SELECT 'All', COUNT(*), ROUND(SUM(o.charge_off_amount) / 1e6, 2),
       ROUND(100.0 * SUM(o.recoveries) / SUM(o.charge_off_amount), 1),
       ROUND(100.0 - 100.0 * SUM(o.recoveries) / SUM(o.charge_off_amount), 1)
FROM loan_outcomes o WHERE o.charge_off_month <= '2024-12';

-- name: 08_policy_change
-- July 2023: online borrowers may have DTI up to 50% (was 40%). 12-month default rate of online loans, before vs after.
SELECT CASE WHEN l.issue_month < '2023-07' THEN '1 before (Jul 2022 - Jun 2023)' ELSE '2 after (Jul 2023 - Dec 2024)' END AS period,
       CASE WHEN l.dti > 40 THEN 'DTI 40-50 (newly allowed)' ELSE 'DTI up to 40' END AS dti_group,
       COUNT(*) AS loans,
       ROUND(100.0 * AVG(o.default_12m), 2) AS default_12m_pct,
       ROUND(AVG(l.int_rate), 2) AS avg_rate_pct
FROM loans l JOIN loan_outcomes o USING (loan_id)
WHERE l.channel = 'Online' AND l.issue_month BETWEEN '2022-07' AND '2024-12'
GROUP BY period, dti_group ORDER BY period, dti_group;

-- name: 09_defaults_vs_unemployment
-- New defaults per 1,000 open loans each quarter, next to unemployment.
WITH m AS (
  SELECT month, COUNT(*) AS open_loans,
         SUM(CASE WHEN status = '90+ DPD' AND days_past_due = 90 THEN 1 ELSE 0 END) AS new_defaults
  FROM loan_months GROUP BY month)
SELECT substr(m.month, 1, 4) || '-Q' || ((CAST(substr(m.month, 6, 2) AS INT) + 2) / 3) AS quarter,
       ROUND(1000.0 * SUM(new_defaults) / SUM(open_loans), 2) AS defaults_per_1000_loan_months,
       ROUND(AVG(x.unemployment_rate), 1) AS unemployment_pct
FROM m JOIN macro x USING (month)
WHERE m.month >= '2021-07'
GROUP BY quarter ORDER BY quarter;
