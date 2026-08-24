# SQL Analysis

## Database
SQLite database: `EV_Market.db`

## How to Run

1. Download `EV_Market.db`
2. Open it using any SQLite-compatible database tool.
3. Open `ev_charging_analysis.sql`
4. Execute the queries.
5. Compare the results with the example outputs below.

## Tables

- `charging_infrastructure`
- `historical_ev`
- `historical_ev_long`

## Queries Included

1. Top 10 states by EV count
2. EVs per charging unit
3. Infrastructure pressure classification
4. Yearly EV market size
5. YoY EV growth using `LAG()`
6. State ranking using `RANK()`
7. EV demand vs charging infrastructure using `JOIN`
