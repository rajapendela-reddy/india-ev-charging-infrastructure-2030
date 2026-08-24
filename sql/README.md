# SQL Analysis

## Overview

This folder contains the SQL analysis used in the **India EV Charging Infrastructure — 2030 Investment Opportunity** project.

SQL was used as a business analysis and validation layer to analyze EV registrations and charging infrastructure across Indian states.

## Database

The analysis uses **SQLite**.

- `EV_Market.db` — SQLite database containing the project tables
- `ev_charging_analysis.sql` — SQL queries used for the analysis

### Tables

- `charging_infrastructure`
- `historical_ev`
- `historical_ev_long`

## Business Questions

The SQL analysis answers:

1. Which states have the largest EV markets?
2. Which states have the highest EVs per charging infrastructure unit?
3. Which states face high infrastructure pressure?
4. How has India's EV market changed over time?
5. How fast is India's EV market growing year over year (YoY)?
6. How do states rank by EV registrations?
7. Which states have high EV demand relative to charging infrastructure?
8. Which states should receive higher investment priority?

## Key Analyses

### 1. EV Infrastructure Analysis

Compares state-level EV registrations with available charging infrastructure.

### 2. Top States by EVs per Charging Unit

Identifies states where EV demand is high relative to available charging infrastructure.

### 3. Infrastructure Pressure Classification

Classifies states into:

- High Pressure
- Moderate Pressure
- Low Pressure

based on EVs per charging unit.

### 4. Yearly EV Market Size

Aggregates EV registrations by year to understand the historical growth of India's EV market.

### 5. Year-over-Year EV Growth

Uses the SQL `LAG()` window function to compare each year's EV registrations with the previous year and calculate YoY growth.

| Year | YoY Growth |
|------|-----------:|
| 2021 | 160.78% |
| 2022 | 208.20% |
| 2023 | 34.46% |
| 2024 | 85.31% |

### 6. State EV Market Ranking

Uses the `RANK()` window function to rank states based on EV registrations.

### 7. EV Demand vs Charging Infrastructure

Uses a `JOIN` between EV registration data and charging infrastructure data to identify states with high EV demand relative to available infrastructure.

### 8. Investment Priority Classification

Uses `CASE WHEN` logic to classify states based on EV demand and infrastructure pressure.

## SQL Concepts Demonstrated

- `SELECT`
- `WHERE`
- `GROUP BY`
- `ORDER BY`
- Aggregate functions
- `CASE WHEN`
- `JOIN`
- CTEs
- Window functions
  - `LAG()`
  - `RANK()`

## How to Run

The queries can be executed using any SQLite-compatible database tool.

1. Download `EV_Market.db`.
2. Open the database using a SQLite-compatible tool.
3. Open `ev_charging_analysis.sql`.
4. Run the queries against the database.

The database contains the tables required by the SQL queries, so the analysis can be reproduced without setting up a separate database server.
