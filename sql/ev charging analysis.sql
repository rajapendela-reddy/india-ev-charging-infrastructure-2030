-- ============================================================
-- INDIA EV CHARGING INFRASTRUCTURE
-- SQL BUSINESS ANALYSIS
-- ============================================================


-- ============================================================
-- 1. TOP 10 STATES BY CURRENT EV COUNT
-- Business Question:
-- Which states have the largest current EV markets?
-- ============================================================

SELECT
    state,
    evs
FROM charging_infrastructure
ORDER BY evs DESC
LIMIT 10;


-- ============================================================
-- 2. TOP 10 STATES BY EVs PER CHARGING UNIT
-- Business Question:
-- Which states have the highest EV demand relative
-- to available charging infrastructure?
-- ============================================================

SELECT
    state,
    evs,
    charging_infrastructure,
    ROUND(
        CAST(evs AS FLOAT) / charging_infrastructure,
        2
    ) AS evs_per_charging_unit
FROM charging_infrastructure
WHERE charging_infrastructure > 0
ORDER BY evs_per_charging_unit DESC
LIMIT 10;


-- ============================================================
-- 3. INFRASTRUCTURE PRESSURE CLASSIFICATION
-- Business Question:
-- Which states face High, Moderate or Low
-- infrastructure pressure?
-- ============================================================

SELECT
    state,
    evs,
    charging_infrastructure,

    ROUND(
        CAST(evs AS FLOAT) / charging_infrastructure,
        2
    ) AS evs_per_charging_unit,

    CASE
        WHEN CAST(evs AS FLOAT) / charging_infrastructure >= 2
            THEN 'High Pressure'

        WHEN CAST(evs AS FLOAT) / charging_infrastructure >= 1
            THEN 'Moderate Pressure'

        ELSE 'Low Pressure'
    END AS pressure_category

FROM charging_infrastructure

WHERE charging_infrastructure > 0

ORDER BY evs_per_charging_unit DESC;


-- ============================================================
-- 4. YEARLY EV MARKET SIZE
-- Business Question:
-- How has India's EV market changed over time?
-- ============================================================

SELECT
    year,
    SUM(ev_registrations) AS total_evs

FROM historical_ev_long

GROUP BY year

ORDER BY year;


-- ============================================================
-- 5. YEAR-OVER-YEAR EV GROWTH
-- Business Question:
-- How fast is the EV market growing each year?
-- ============================================================

WITH yearly_ev AS (

    SELECT
        year,
        SUM(ev_registrations) AS total_evs

    FROM historical_ev_long

    GROUP BY year
)

SELECT
    year,
    total_evs,

    LAG(total_evs) OVER (
        ORDER BY year
    ) AS previous_year_evs,

    ROUND(
        (
            total_evs
            - LAG(total_evs) OVER (ORDER BY year)
        )
        * 100.0
        / LAG(total_evs) OVER (ORDER BY year),
        2
    ) AS yoy_growth_pct

FROM yearly_ev

ORDER BY year;


-- ============================================================
-- 6. STATE EV MARKET RANKING
-- Business Question:
-- Which states have the largest EV markets?
-- ============================================================

SELECT
    state,
    SUM(ev_registrations) AS total_evs,

    RANK() OVER (
        ORDER BY SUM(ev_registrations) DESC
    ) AS ev_market_rank

FROM historical_ev_long

WHERE year = 2024
  AND state <> 'Grand Total'

GROUP BY state

ORDER BY ev_market_rank;


-- ============================================================
-- 7. EV DEMAND VS CHARGING INFRASTRUCTURE
-- Business Question:
-- Which states have high EV demand relative to
-- available charging infrastructure?
-- ============================================================

SELECT
    e.state,

    SUM(e.ev_registrations) AS evs_2024,

    c.charging_infrastructure,

    ROUND(
        CAST(SUM(e.ev_registrations) AS FLOAT)
        / c.charging_infrastructure,
        2
    ) AS evs_per_charging_unit

FROM historical_ev_long e

JOIN charging_infrastructure c
    ON e.state = c.state

WHERE e.year = 2024
  AND e.state <> 'Grand Total'

GROUP BY
    e.state,
    c.charging_infrastructure

ORDER BY evs_per_charging_unit DESC;
