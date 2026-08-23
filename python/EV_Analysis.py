"""EV market intelligence and charging-infrastructure analytics.

Standalone replacement for the original Google Colab notebook export.
Edit the two paths immediately below, or override them with --input/--output.
"""

from pathlib import Path

# =============================================================================
# USER SETTINGS: EDIT THESE INPUT AND OUTPUT PATHS
# =============================================================================
# Relative-path example (recommended when the Excel file is inside this project):
PROJECT_DIRECTORY = Path.cwd().parent

INPUT_EXCEL_PATH = PROJECT_DIRECTORY / "data" / "raw" / "EV_market_model.xlsx"
OUTPUT_DIRECTORY = PROJECT_DIRECTORY / "data" / "processed"

# Windows absolute-path example:
# INPUT_EXCEL_PATH = Path(r"D:\your_folder\EV_market_model.xlsx")
# OUTPUT_DIRECTORY = Path(r"D:\your_folder\EV_market_results")
# =============================================================================

import argparse
import logging
from collections.abc import Mapping

import matplotlib
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from sklearn.cluster import KMeans
from sklearn.linear_model import LinearRegression
from sklearn.metrics import silhouette_score
from sklearn.preprocessing import StandardScaler


# Input workbook settings
CHARGING_SHEET = "Charging_Stations"
HISTORICAL_SHEET = "EV_Registrations"
CHARGING_SOURCE_FILTER = "BEE / EV Yatra"
HISTORICAL_YEARS = [2020, 2021, 2022, 2023, 2024]

# Output settings
OUTPUT_EXCEL_NAME = "EV_Investor_Dashboard_Data.xlsx"
SAVE_PLOTS_BY_DEFAULT = True
SHOW_PLOTS_BY_DEFAULT = False
INVESTOR_RANKING_COLUMNS = [
    "State",
    "EVs",
    "EV_2030",
    "Infrastructure_Gap_2030",
    "Infrastructure_Pressure_Index",
    "Investor_Opportunity_Score",
]
CHARGING_INFRASTRUCTURE_COLUMN_MAP = {
    "State": "State",
    "EVs": "EVs",
    "Operational Public Charging Infrastructure": (
        "Operational Public Charging Infrastructure"
    ),
    "EVs_per_Infrastructure": "EVs_per_Infrastructure",
    "Infrastructure_per_1000_EVs": "infrastructure for 1000 evs",
    "Source": "Source",
    "data year": "data year",
    "Infrastructure_Pressure_Index": "Infrastructure_Pressure_Index",
    "Pressure_Category": "Pressure_Category",
    "EV_Market_Category": "EV_Market_Category",
    "Opportunity_Segment": "Opportunity_Segmen",
}

# Model assumptions retained from the Colab analysis
INVESTOR_ELIGIBILITY_EVS = 150_000
CURRENT_INFRASTRUCTURE = 26_241
EV_STOCK_2025 = 7_290_226
BASE_EVS_PER_INFRASTRUCTURE = 278
STATE_FORECAST_YEARS = 5
SOM_CAPTURE_RATE = 0.10

BASE_GROWTH = {
    2025: 0.50,
    2026: 0.45,
    2027: 0.40,
    2028: 0.35,
    2029: 0.30,
    2030: 0.25,
}
CONSERVATIVE_GROWTH = {
    2025: 0.40,
    2026: 0.35,
    2027: 0.30,
    2028: 0.25,
    2029: 0.20,
    2030: 0.20,
}
AGGRESSIVE_GROWTH = {
    2025: 0.60,
    2026: 0.55,
    2027: 0.50,
    2028: 0.45,
    2029: 0.40,
    2030: 0.35,
}
COST_SCENARIOS = {
    "Conservative": 1_000_000,
    "Base": 1_500_000,
    "Aggressive": 2_500_000,
}

POPULATION_2024 = {
    "Uttar Pradesh": 237_882_000,
    "Maharashtra": 126_385_000,
    "Bihar": 131_055_000,
    "Chandigarh": 1_243_000,
    "Karnataka": 68_866_000,
    "Assam": 36_159_000,
    "Tamil Nadu": 77_563_000,
    "Rajasthan": 81_800_000,
    "Gujarat": 70_700_000,
    "Delhi": 21_248_000,
    "Madhya Pradesh": 87_000_000,
    "Odisha": 46_960_000,
    "Kerala": 35_500_000,
    "Andhra Pradesh": 53_156_000,
    "Chhattisgarh": 30_100_000,
    "Uttarakhand": 11_700_000,
    "Tripura": 4_350_000,
    "Punjab": 30_700_000,
    "Haryana": 30_700_000,
    "Jharkhand": 41_000_000,
    "Jammu & Kashmir": 7_200_000,
    "Goa": 1_610_000,
    "West Bengal": 99_723_000,
    "Puducherry": 1_600_000,
    "Mizoram": 1_240_000,
    "Andaman & Nicobar Islands": 410_000,
    "Himachal Pradesh": 7_400_000,
    "Meghalaya": 3_400_000,
    "Manipur": 3_300_000,
    "Nagaland": 2_300_000,
    "Arunachal Pradesh": 1_600_000,
    "Lakshadweep": 71_000,
    "Ladakh": 320_000,
    "Telangana": 38_100_000,
    "Sikkim": 690_000,
}

LOGGER = logging.getLogger("ev_market_analytics")


def parse_args() -> argparse.Namespace:
    """Read optional command-line path and plot overrides."""
    parser = argparse.ArgumentParser(
        description="Analyze India's EV market from the supplied Excel workbook."
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=INPUT_EXCEL_PATH,
        help=f"Input Excel workbook (default: {INPUT_EXCEL_PATH})",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=OUTPUT_DIRECTORY,
        help=f"Output directory (default: {OUTPUT_DIRECTORY})",
    )
    parser.add_argument(
        "--no-plots",
        action="store_true",
        help="Do not save PNG charts.",
    )
    parser.add_argument(
        "--show-plots",
        action="store_true",
        default=SHOW_PLOTS_BY_DEFAULT,
        help="Open chart windows as well as saving the PNG files.",
    )
    return parser.parse_args()


def require_columns(data: pd.DataFrame, required: list[object], sheet: str) -> None:
    """Raise a useful error when an expected workbook column is missing."""
    missing = [column for column in required if column not in data.columns]
    if missing:
        raise ValueError(
            f"Sheet '{sheet}' is missing required columns: {missing}. "
            f"Available columns: {list(data.columns)}"
        )


def numeric_series(values: pd.Series) -> pd.Series:
    """Convert comma-formatted Excel values to numbers."""
    cleaned = values.astype("string").str.replace(",", "", regex=False).str.strip()
    return pd.to_numeric(cleaned, errors="coerce")


def minmax_score(values: pd.Series) -> pd.Series:
    """Return a stable zero-to-one score, including for missing/constant data."""
    numeric = pd.to_numeric(values, errors="coerce")
    if numeric.notna().sum() == 0:
        return pd.Series(0.0, index=values.index)
    numeric = numeric.fillna(numeric.median())
    minimum, maximum = numeric.min(), numeric.max()
    if np.isclose(minimum, maximum):
        return pd.Series(1.0, index=values.index)
    return (numeric - minimum) / (maximum - minimum)


def load_input_data(input_path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Load and validate the two required Excel sheets."""
    input_path = input_path.expanduser().resolve()
    if not input_path.is_file():
        raise FileNotFoundError(
            f"Input workbook not found: {input_path}\n"
            "Edit INPUT_EXCEL_PATH at the top of this file or use --input."
        )

    workbook = pd.ExcelFile(input_path)
    missing_sheets = [
        sheet
        for sheet in (CHARGING_SHEET, HISTORICAL_SHEET)
        if sheet not in workbook.sheet_names
    ]
    if missing_sheets:
        raise ValueError(
            f"Missing required sheet(s): {missing_sheets}. "
            f"Available sheets: {workbook.sheet_names}"
        )

    charging = pd.read_excel(input_path, sheet_name=CHARGING_SHEET)
    historical = pd.read_excel(input_path, sheet_name=HISTORICAL_SHEET)
    return charging, historical


def clean_charging_data(charging_raw: pd.DataFrame) -> pd.DataFrame:
    """Clean the charging sheet and calculate pressure/opportunity fields."""
    charging = charging_raw.dropna(how="all").copy()
    required = [
        "State",
        "EVs",
        "Operational Public Charging Infrastructure",
        "Source",
        "data year",
    ]
    require_columns(charging, required, CHARGING_SHEET)

    charging["State"] = charging["State"].astype("string").str.strip()
    charging["Source"] = charging["Source"].astype("string").str.strip()
    charging["EVs"] = numeric_series(charging["EVs"])
    infrastructure_column = "Operational Public Charging Infrastructure"
    charging[infrastructure_column] = numeric_series(charging[infrastructure_column])
    charging["data year"] = pd.to_numeric(charging["data year"], errors="coerce")

    charging = charging[
        charging["Source"].str.casefold() == CHARGING_SOURCE_FILTER.casefold()
    ].copy()
    if charging.empty:
        raise ValueError(
            f"No rows in '{CHARGING_SHEET}' match Source="
            f"'{CHARGING_SOURCE_FILTER}'."
        )

    charging = charging.dropna(
        subset=["State", "EVs", infrastructure_column]
    ).copy()
    charging = charging[(charging["EVs"] > 0) & (charging[infrastructure_column] > 0)]

    if charging["State"].duplicated().any():
        duplicate_count = int(charging["State"].duplicated().sum())
        LOGGER.warning(
            "%s duplicate state row(s) remained after source filtering; "
            "keeping the latest data year for each state.",
            duplicate_count,
        )
        charging = (
            charging.sort_values(["State", "data year"], na_position="first")
            .drop_duplicates("State", keep="last")
            .copy()
        )

    charging["EVs_per_Infrastructure"] = (
        charging["EVs"] / charging[infrastructure_column]
    )
    charging["Infrastructure_per_1000_EVs"] = (
        charging[infrastructure_column] / charging["EVs"] * 1000
    )

    national_average = charging["EVs_per_Infrastructure"].mean()
    charging["Infrastructure_Pressure_Index"] = (
        charging["EVs_per_Infrastructure"] / national_average
    )
    charging["Pressure_Category"] = pd.cut(
        charging["Infrastructure_Pressure_Index"],
        bins=[-np.inf, 0.75, 1.25, np.inf],
        labels=["Low Pressure", "Moderate Pressure", "High Pressure"],
    )

    ev_median = charging["EVs"].median()
    charging["EV_Market_Category"] = np.where(
        charging["EVs"] >= ev_median, "High EV Market", "Low EV Market"
    )
    charging["Opportunity_Segment"] = np.select(
        [
            (charging["EV_Market_Category"] == "High EV Market")
            & (charging["Pressure_Category"] == "High Pressure"),
            (charging["EV_Market_Category"] == "High EV Market")
            & (charging["Pressure_Category"] != "High Pressure"),
            (charging["EV_Market_Category"] == "Low EV Market")
            & (charging["Pressure_Category"] == "High Pressure"),
        ],
        ["Priority Investment", "Market Development", "Infrastructure Expansion"],
        default="Low Priority",
    )
    return charging.sort_values("EVs", ascending=False).reset_index(drop=True)


def build_market_diagnostics(
    charging: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, dict[str, float]]:
    """Build Pareto, outlier, and regression tables."""
    infrastructure_column = "Operational Public Charging Infrastructure"
    pareto = charging[["State", "EVs"]].sort_values("EVs", ascending=False).copy()
    total_evs = pareto["EVs"].sum()
    pareto["EV_Market_Share"] = pareto["EVs"] / total_evs * 100
    pareto["Cumulative_Share"] = pareto["EV_Market_Share"].cumsum()

    q1 = charging["EVs_per_Infrastructure"].quantile(0.25)
    q3 = charging["EVs_per_Infrastructure"].quantile(0.75)
    iqr = q3 - q1
    lower_bound, upper_bound = q1 - 1.5 * iqr, q3 + 1.5 * iqr
    outliers = charging[
        (charging["EVs_per_Infrastructure"] < lower_bound)
        | (charging["EVs_per_Infrastructure"] > upper_bound)
    ].copy()
    outliers["Outlier_Lower_Bound"] = lower_bound
    outliers["Outlier_Upper_Bound"] = upper_bound

    regression = charging[["State", "EVs", infrastructure_column]].dropna().copy()
    if len(regression) < 2:
        raise ValueError("At least two valid charging rows are required for regression.")
    model = LinearRegression().fit(regression[["EVs"]], regression[infrastructure_column])
    regression["Predicted_Infrastructure"] = model.predict(regression[["EVs"]])
    regression["Infrastructure_Gap"] = (
        regression["Predicted_Infrastructure"] - regression[infrastructure_column]
    )
    regression["Infrastructure_Gap_Percent"] = np.where(
        regression["Predicted_Infrastructure"] != 0,
        regression["Infrastructure_Gap"]
        / regression["Predicted_Infrastructure"]
        * 100,
        np.nan,
    )

    metrics = {
        "Pearson correlation": regression["EVs"].corr(
            regression[infrastructure_column], method="pearson"
        ),
        "Spearman correlation": regression["EVs"].corr(
            regression[infrastructure_column], method="spearman"
        ),
        "Regression coefficient": float(model.coef_[0]),
        "Regression intercept": float(model.intercept_),
        "Regression R-squared": float(model.score(
            regression[["EVs"]], regression[infrastructure_column]
        )),
    }
    return pareto, outliers, regression, metrics


def label_clusters(cluster_data: pd.DataFrame) -> dict[int, str]:
    """Assign useful labels from cluster profiles instead of arbitrary KMeans IDs."""
    profile = cluster_data.groupby("Cluster").agg(
        mean_evs=("EVs", "mean"),
        mean_pressure=("Infrastructure_Pressure_Index", "mean"),
    )
    remaining = set(int(value) for value in profile.index)
    labels: dict[int, str] = {}

    if remaining:
        largest = int(profile.loc[list(remaining), "mean_evs"].idxmax())
        labels[largest] = "Large EV Markets"
        remaining.remove(largest)
    if remaining:
        constrained = int(profile.loc[list(remaining), "mean_pressure"].idxmax())
        labels[constrained] = "Infrastructure-Constrained Markets"
        remaining.remove(constrained)
    if remaining:
        lowest = int(profile.loc[list(remaining), "mean_evs"].idxmin())
        labels[lowest] = "Low-Demand Markets"
        remaining.remove(lowest)
    for cluster in remaining:
        labels[cluster] = "Mainstream EV Markets"
    return labels


def build_clusters(charging: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Cluster states and report silhouette scores for feasible K values."""
    features = [
        "EVs",
        "EVs_per_Infrastructure",
        "Infrastructure_per_1000_EVs",
        "Infrastructure_Pressure_Index",
    ]
    cluster_data = charging[["State", *features]].dropna().copy()
    if len(cluster_data) < 2:
        raise ValueError("At least two states are required for clustering.")

    scaled = StandardScaler().fit_transform(cluster_data[features])
    quality_rows: list[dict[str, float | int]] = []
    maximum_k = min(8, len(cluster_data) - 1)
    for k in range(2, maximum_k + 1):
        candidate = KMeans(n_clusters=k, random_state=42, n_init=20)
        candidate_labels = candidate.fit_predict(scaled)
        quality_rows.append(
            {
                "K": k,
                "Inertia": float(candidate.inertia_),
                "Silhouette_Score": float(silhouette_score(scaled, candidate_labels)),
            }
        )

    selected_k = min(4, len(cluster_data))
    model = KMeans(n_clusters=selected_k, random_state=42, n_init=20)
    cluster_data["Cluster"] = model.fit_predict(scaled)
    cluster_data["Cluster_Name"] = cluster_data["Cluster"].map(
        label_clusters(cluster_data)
    )
    cluster_data = cluster_data.sort_values(["Cluster", "EVs"], ascending=[True, False])
    return cluster_data, pd.DataFrame(quality_rows)


def normalize_historical_headers(data: pd.DataFrame) -> pd.DataFrame:
    """Convert Excel year headers such as '2020' to integer 2020."""
    renamed: dict[object, object] = {}
    for column in data.columns:
        text = str(column).strip()
        if text.isdigit() and len(text) == 4:
            renamed[column] = int(text)
    return data.rename(columns=renamed)


def build_growth_analysis(
    historical_raw: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Calculate state growth scores and the national annual history."""
    historical = normalize_historical_headers(historical_raw.copy())
    require_columns(historical, ["State", *HISTORICAL_YEARS], HISTORICAL_SHEET)
    historical["State"] = historical["State"].astype("string").str.strip()
    historical = historical[
        ~historical["State"].str.contains("Grand Total", case=False, na=False)
    ].copy()

    historical_long = historical[["State", *HISTORICAL_YEARS]].melt(
        id_vars="State", var_name="Year", value_name="EVs"
    )
    historical_long["Year"] = pd.to_numeric(historical_long["Year"], errors="coerce")
    historical_long["EVs"] = numeric_series(historical_long["EVs"])
    historical_long = historical_long.dropna(subset=["State", "Year", "EVs"])
    historical_long["Year"] = historical_long["Year"].astype(int)

    if historical_long.duplicated(["State", "Year"]).any():
        duplicate_count = int(historical_long.duplicated(["State", "Year"]).sum())
        LOGGER.warning(
            "%s duplicate State/Year historical row(s) found; keeping the first.",
            duplicate_count,
        )
        historical_long = historical_long.drop_duplicates(
            ["State", "Year"], keep="first"
        )

    historical_long = historical_long.sort_values(["State", "Year"]).copy()
    historical_long["YoY_Growth"] = (
        historical_long.groupby("State")["EVs"].pct_change(fill_method=None) * 100
    )

    endpoints = (
        historical_long[historical_long["Year"].isin([2020, 2024])]
        .pivot(index="State", columns="Year", values="EVs")
        .reset_index()
        .rename(columns={2020: "EV_2020", 2024: "EV_2024"})
    )
    endpoints = endpoints.dropna(subset=["EV_2020", "EV_2024"])
    endpoints = endpoints[endpoints["EV_2020"] > 0].copy()
    endpoints["CAGR_2020_2024"] = (
        (endpoints["EV_2024"] / endpoints["EV_2020"]) ** (1 / 4) - 1
    ) * 100
    endpoints["Absolute_Growth"] = endpoints["EV_2024"] - endpoints["EV_2020"]
    endpoints["Growth_Score"] = endpoints["CAGR_2020_2024"].rank(pct=True)

    recent = (
        historical_long[historical_long["Year"].isin([2023, 2024])]
        .pivot(index="State", columns="Year", values="EVs")
        .reset_index()
        .rename(columns={2023: "EV_2023", 2024: "EV_2024"})
    )
    recent = recent.dropna(subset=["EV_2023", "EV_2024"])
    recent = recent[recent["EV_2023"] > 0].copy()
    recent["Growth_2023_2024"] = (
        (recent["EV_2024"] - recent["EV_2023"]) / recent["EV_2023"] * 100
    )
    recent["Recent_Growth_Score"] = recent["Growth_2023_2024"].rank(pct=True)

    growth = endpoints.merge(
        recent[["State", "EV_2023", "Growth_2023_2024", "Recent_Growth_Score"]],
        on="State",
        how="left",
    )
    growth["Final_Growth_Score"] = (
        0.60 * growth["Growth_Score"] + 0.40 * growth["Recent_Growth_Score"]
    )
    annual = historical_long.groupby("Year", as_index=False)["EVs"].sum()
    annual["YoY_Growth"] = annual["EVs"].pct_change(fill_method=None) * 100
    return historical_long, growth, annual


def build_current_opportunity_ranking(
    charging: pd.DataFrame,
    clusters: pd.DataFrame,
    growth: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Combine market size, pressure, growth, and penetration scores."""
    score = charging[
        ["State", "EVs", "Infrastructure_Pressure_Index", "Opportunity_Segment"]
    ].copy()
    score = score.merge(clusters[["State", "Cluster", "Cluster_Name"]], on="State")
    score["Market_Size_Score"] = minmax_score(score["EVs"])
    score["Pressure_Score"] = minmax_score(score["Infrastructure_Pressure_Index"])
    score["Opportunity_Score"] = (
        0.60 * score["Market_Size_Score"] + 0.40 * score["Pressure_Score"]
    )
    score = score.merge(
        growth[
            ["State", "CAGR_2020_2024", "Growth_2023_2024", "Final_Growth_Score"]
        ],
        on="State",
        how="left",
    )
    growth_fill = score["Final_Growth_Score"].median()
    score["Final_Growth_Score"] = score["Final_Growth_Score"].fillna(
        0.0 if pd.isna(growth_fill) else growth_fill
    )
    score["Opportunity_Score_V2"] = (
        0.50 * score["Market_Size_Score"]
        + 0.30 * score["Pressure_Score"]
        + 0.20 * score["Final_Growth_Score"]
    )

    population = pd.DataFrame(
        POPULATION_2024.items(), columns=["State", "Population_2024"]
    )
    score = score.merge(population, on="State", how="left")
    score["EVs_per_1000_People"] = score["EVs"] / score["Population_2024"] * 1000
    score["Penetration_Score"] = score["EVs_per_1000_People"].rank(pct=True)
    score["Penetration_Score"] = score["Penetration_Score"].fillna(0.5)

    score["Opportunity_Score_V3"] = (
        0.45 * score["Market_Size_Score"]
        + 0.25 * score["Pressure_Score"]
        + 0.20 * score["Final_Growth_Score"]
        + 0.10 * score["Penetration_Score"]
    )
    scenario_weights = {
        "Market_Focused": (0.50, 0.25, 0.15, 0.10),
        "Balanced": (0.45, 0.25, 0.20, 0.10),
        "Growth_Focused": (0.35, 0.25, 0.30, 0.10),
    }
    for scenario, weights in scenario_weights.items():
        market_weight, pressure_weight, growth_weight, penetration_weight = weights
        score[f"Score_{scenario}"] = (
            market_weight * score["Market_Size_Score"]
            + pressure_weight * score["Pressure_Score"]
            + growth_weight * score["Final_Growth_Score"]
            + penetration_weight * score["Penetration_Score"]
        )
        score[f"Rank_{scenario}"] = score[f"Score_{scenario}"].rank(
            ascending=False, method="min"
        )

    score = score.sort_values("Opportunity_Score_V3", ascending=False).reset_index(drop=True)
    score["Opportunity_Rank_V3"] = score.index + 1
    sensitivity = score[
        ["State", "Rank_Market_Focused", "Rank_Balanced", "Rank_Growth_Focused"]
    ].copy()
    rank_columns = [
        "Rank_Market_Focused",
        "Rank_Balanced",
        "Rank_Growth_Focused",
    ]
    sensitivity["Average_Rank"] = sensitivity[rank_columns].mean(axis=1)
    sensitivity["Rank_Range"] = (
        sensitivity[rank_columns].max(axis=1) - sensitivity[rank_columns].min(axis=1)
    )
    return score, sensitivity.sort_values("Rank_Balanced")


def scenario_forecast(start_value: float, growth_rates: Mapping[int, float]) -> pd.Series:
    """Compound an annual starting value using the supplied growth assumptions."""
    values: dict[int, float] = {}
    current = float(start_value)
    for year, growth_rate in growth_rates.items():
        current *= 1 + growth_rate
        values[year] = current
    return pd.Series(values, dtype="float64")


def build_forecasts(
    annual: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Create CAGR, linear, scenario, stock, and infrastructure forecasts."""
    indexed = annual.set_index("Year")["EVs"]
    for year in (2020, 2023, 2024):
        if year not in indexed.index:
            raise ValueError(f"Historical EV data does not contain required year {year}.")

    ev_2020, ev_2024 = float(indexed.loc[2020]), float(indexed.loc[2024])
    if ev_2020 <= 0:
        raise ValueError("The national 2020 EV total must be greater than zero.")
    cagr = (ev_2024 / ev_2020) ** (1 / 4) - 1
    forecast_years = np.arange(2025, 2031)
    forecast_comparison = pd.DataFrame({"Year": forecast_years})
    forecast_comparison["CAGR_Forecast"] = [
        ev_2024 * (1 + cagr) ** (year - 2024) for year in forecast_years
    ]
    linear_model = LinearRegression().fit(annual[["Year"]], annual["EVs"])
    forecast_comparison["Linear_Forecast"] = linear_model.predict(
        forecast_comparison[["Year"]]
    )

    scenario = pd.DataFrame({"Year": forecast_years})
    scenario["Conservative"] = scenario_forecast(
        ev_2024, CONSERVATIVE_GROWTH
    ).values
    scenario["Base"] = scenario_forecast(ev_2024, BASE_GROWTH).values
    scenario["Aggressive"] = scenario_forecast(ev_2024, AGGRESSIVE_GROWTH).values

    stock = scenario.copy()
    for column in ("Conservative", "Base", "Aggressive"):
        stock[f"{column}_Stock"] = EV_STOCK_2025 + stock[column].cumsum()

    final_year = stock.loc[stock["Year"] == 2030].iloc[0]
    required = pd.DataFrame(
        {
            "Scenario": ["Conservative", "Base", "Aggressive"],
            "EVs_2030": [
                final_year["Conservative_Stock"],
                final_year["Base_Stock"],
                final_year["Aggressive_Stock"],
            ],
            "EVs_per_Infrastructure": [350, 278, 200],
        }
    )
    required["Required_Infrastructure"] = (
        required["EVs_2030"] / required["EVs_per_Infrastructure"]
    )
    required["Additional_Infrastructure"] = (
        required["Required_Infrastructure"] - CURRENT_INFRASTRUCTURE
    ).clip(lower=0)
    return forecast_comparison, stock, required


def assign_growth_rate(score: float) -> float:
    """Convert a normalized state growth score into an annual assumption."""
    if score >= 0.65:
        return 0.40
    if score >= 0.50:
        return 0.35
    if score >= 0.40:
        return 0.30
    return 0.25


def build_investor_analysis(
    charging: pd.DataFrame,
    growth: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Forecast state demand, rank investments, and calculate opportunity size."""
    state_2030 = charging[
        [
            "State",
            "EVs",
            "Operational Public Charging Infrastructure",
            "Infrastructure_Pressure_Index",
            "Opportunity_Segment",
        ]
    ].copy()
    state_2030 = state_2030.merge(
        growth[["State", "Growth_2023_2024", "Final_Growth_Score"]],
        on="State",
        how="left",
    )
    growth_fill = state_2030["Final_Growth_Score"].median()
    state_2030["Final_Growth_Score"] = state_2030["Final_Growth_Score"].fillna(
        0.0 if pd.isna(growth_fill) else growth_fill
    )
    state_2030["Investor_Eligible"] = np.where(
        state_2030["EVs"] >= INVESTOR_ELIGIBILITY_EVS,
        "Eligible",
        "Not Eligible",
    )
    investors = state_2030[state_2030["Investor_Eligible"] == "Eligible"].copy()
    if investors.empty:
        raise ValueError(
            "No state meets INVESTOR_ELIGIBILITY_EVS. Adjust the model setting "
            "near the top of the script if appropriate."
        )

    investors["Annual_Growth_Assumption"] = investors["Final_Growth_Score"].apply(
        assign_growth_rate
    )
    investors["EV_2030"] = investors["EVs"] * (
        1 + investors["Annual_Growth_Assumption"]
    ) ** STATE_FORECAST_YEARS
    investors["Required_Infrastructure_2030"] = (
        investors["EV_2030"] / BASE_EVS_PER_INFRASTRUCTURE
    )
    infrastructure_column = "Operational Public Charging Infrastructure"
    investors["Infrastructure_Gap_2030"] = (
        investors["Required_Infrastructure_2030"] - investors[infrastructure_column]
    ).clip(lower=0)
    investors["Infrastructure_Gap_Percent_2030"] = (
        investors["Infrastructure_Gap_2030"]
        / investors["Required_Infrastructure_2030"]
        * 100
    )

    investors["EV2030_Score"] = minmax_score(investors["EV_2030"])
    investors["Gap2030_Score"] = minmax_score(investors["Infrastructure_Gap_2030"])
    investors["Pressure_Score_2"] = minmax_score(
        investors["Infrastructure_Pressure_Index"]
    )
    investors["Investor_Opportunity_Score"] = (
        0.40 * investors["EV2030_Score"]
        + 0.40 * investors["Gap2030_Score"]
        + 0.20 * investors["Pressure_Score_2"]
    )
    investors["Score_Market_Focused"] = (
        0.50 * investors["EV2030_Score"]
        + 0.30 * investors["Gap2030_Score"]
        + 0.20 * investors["Pressure_Score_2"]
    )
    investors["Score_Balanced"] = investors["Investor_Opportunity_Score"]
    investors["Score_Infrastructure_Focused"] = (
        0.30 * investors["EV2030_Score"]
        + 0.50 * investors["Gap2030_Score"]
        + 0.20 * investors["Pressure_Score_2"]
    )
    for score_column, rank_column in (
        ("Score_Market_Focused", "Rank_Market"),
        ("Score_Balanced", "Rank_Balanced"),
        ("Score_Infrastructure_Focused", "Rank_Infrastructure"),
    ):
        investors[rank_column] = investors[score_column].rank(
            ascending=False, method="min"
        )

    investors = investors.sort_values(
        "Investor_Opportunity_Score", ascending=False
    ).reset_index(drop=True)
    investors["Rank"] = investors.index + 1
    investors["Recommendation"] = np.select(
        [
            investors["Investor_Opportunity_Score"] >= 0.50,
            investors["Investor_Opportunity_Score"] >= 0.25,
        ],
        ["High Priority", "Medium-High Priority"],
        default="Selective Investment",
    )
    investors["Investment_Tier"] = np.select(
        [investors["Rank"] <= 3, investors["Rank"] <= 7, investors["Rank"] <= 15],
        [
            "Tier 1 - Strategic Priority",
            "Tier 2 - High Potential",
            "Tier 3 - Secondary Opportunity",
        ],
        default="Tier 4 - Low Priority",
    )

    gap_pareto = investors[["State", "Infrastructure_Gap_2030"]].sort_values(
        "Infrastructure_Gap_2030", ascending=False
    ).copy()
    total_gap = gap_pareto["Infrastructure_Gap_2030"].sum()
    if total_gap <= 0:
        raise ValueError("The calculated 2030 infrastructure opportunity is zero.")
    gap_pareto["Gap_Share"] = gap_pareto["Infrastructure_Gap_2030"] / total_gap * 100
    gap_pareto["Cumulative_Gap_Share"] = gap_pareto["Gap_Share"].cumsum()

    top_five_gap = gap_pareto.head(5)["Infrastructure_Gap_2030"].sum()
    strategic_gap = gap_pareto.head(3)["Infrastructure_Gap_2030"].sum()
    market_sizes = []
    for scenario, cost in COST_SCENARIOS.items():
        market_sizes.append(
            {
                "Scenario": scenario,
                "Cost_per_Unit_INR": cost,
                "TAM_INR_Crore": total_gap * cost / 10_000_000,
                "SAM_INR_Crore": top_five_gap * cost / 10_000_000,
                "Strategic_Top3_INR_Crore": strategic_gap * cost / 10_000_000,
                "SOM_10pct_INR_Crore": (
                    strategic_gap * cost * SOM_CAPTURE_RATE / 10_000_000
                ),
            }
        )
    return state_2030, investors, gap_pareto, pd.DataFrame(market_sizes)


def configure_matplotlib(show_plots: bool) -> None:
    """Select a non-interactive backend before importing pyplot when needed."""
    if not show_plots:
        matplotlib.use("Agg")


def finish_figure(
    figure: object,
    output_path: Path,
    save_plots: bool,
    show_plots: bool,
) -> None:
    """Save, optionally show, and close a Matplotlib figure."""
    import matplotlib.pyplot as plt

    figure.tight_layout()
    if save_plots:
        figure.savefig(output_path, dpi=180, bbox_inches="tight")
    if show_plots:
        plt.show()
    plt.close(figure)


def create_plots(
    output_directory: Path,
    charging: pd.DataFrame,
    pareto: pd.DataFrame,
    regression: pd.DataFrame,
    clusters: pd.DataFrame,
    current_sensitivity: pd.DataFrame,
    annual: pd.DataFrame,
    stock_forecast: pd.DataFrame,
    gap_pareto: pd.DataFrame,
    investors: pd.DataFrame,
    save_plots: bool,
    show_plots: bool,
) -> None:
    """Create the eight most useful charts from the original notebook."""
    if not save_plots and not show_plots:
        return
    configure_matplotlib(show_plots)
    import matplotlib.pyplot as plt

    plot_directory = output_directory / "plots"
    if save_plots:
        plot_directory.mkdir(parents=True, exist_ok=True)

    figure, axis_1 = plt.subplots(figsize=(12, 6))
    axis_1.bar(pareto["State"], pareto["EVs"], color="#2563EB")
    axis_1.set_ylabel("EV registrations")
    axis_1.tick_params(axis="x", rotation=90)
    axis_2 = axis_1.twinx()
    axis_2.plot(pareto["State"], pareto["Cumulative_Share"], color="#DC2626", marker="o")
    axis_2.axhline(80, color="#6B7280", linestyle="--")
    axis_2.set_ylabel("Cumulative market share (%)")
    axis_1.set_title("Pareto Analysis of India's EV Market")
    finish_figure(figure, plot_directory / "01_ev_market_pareto.png", save_plots, show_plots)

    figure, axis = plt.subplots(figsize=(11, 7))
    bubble_size = np.clip(charging["EVs"] / 5000, 25, 500)
    axis.scatter(charging["EVs"], charging["Infrastructure_Pressure_Index"], s=bubble_size, alpha=0.7)
    for row in charging.itertuples():
        axis.annotate(row.State, (row.EVs, row.Infrastructure_Pressure_Index), fontsize=7)
    axis.axvline(charging["EVs"].median(), color="#6B7280", linestyle="--")
    axis.axhline(charging["Infrastructure_Pressure_Index"].median(), color="#6B7280", linestyle="--")
    axis.set(xlabel="EV registrations", ylabel="Infrastructure pressure index", title="EV Market Size vs Infrastructure Pressure")
    finish_figure(figure, plot_directory / "02_market_vs_pressure.png", save_plots, show_plots)

    figure, axis = plt.subplots(figsize=(10, 6))
    infrastructure_column = "Operational Public Charging Infrastructure"
    ordered_regression = regression.sort_values("EVs")
    axis.scatter(regression["EVs"], regression[infrastructure_column], alpha=0.7, label="Actual")
    axis.plot(ordered_regression["EVs"], ordered_regression["Predicted_Infrastructure"], color="#DC2626", label="Regression")
    axis.set(xlabel="EV registrations", ylabel="Public charging infrastructure", title="EVs vs Charging Infrastructure")
    axis.legend()
    finish_figure(figure, plot_directory / "03_infrastructure_regression.png", save_plots, show_plots)

    figure, axis = plt.subplots(figsize=(11, 7))
    for cluster_name, subset in clusters.groupby("Cluster_Name"):
        axis.scatter(subset["EVs"], subset["Infrastructure_Pressure_Index"], alpha=0.75, label=cluster_name)
        for row in subset.itertuples():
            axis.annotate(row.State, (row.EVs, row.Infrastructure_Pressure_Index), fontsize=7)
    axis.set(xlabel="EV registrations", ylabel="Infrastructure pressure index", title="EV State Clusters")
    axis.legend(fontsize=8)
    finish_figure(figure, plot_directory / "04_state_clusters.png", save_plots, show_plots)

    top_states = current_sensitivity.head(10)
    figure, axis = plt.subplots(figsize=(12, 7))
    for column, label in (
        ("Rank_Market_Focused", "Market focused"),
        ("Rank_Balanced", "Balanced"),
        ("Rank_Growth_Focused", "Growth focused"),
    ):
        axis.plot(top_states["State"], top_states[column], marker="o", label=label)
    axis.invert_yaxis()
    axis.tick_params(axis="x", rotation=45)
    axis.set(xlabel="State", ylabel="Rank", title="Opportunity Ranking Across Scenarios")
    axis.legend()
    finish_figure(figure, plot_directory / "05_scenario_rank_sensitivity.png", save_plots, show_plots)

    figure, axis = plt.subplots(figsize=(11, 6))
    axis.plot(annual["Year"], annual["EVs"], marker="o", label="Historical registrations")
    for column in ("Conservative_Stock", "Base_Stock", "Aggressive_Stock"):
        axis.plot(stock_forecast["Year"], stock_forecast[column], marker="o", label=column.replace("_", " "))
    axis.set(xlabel="Year", ylabel="EVs", title="India EV Market: Historical and 2030 Scenarios")
    axis.grid(alpha=0.3)
    axis.legend()
    finish_figure(figure, plot_directory / "06_ev_forecast.png", save_plots, show_plots)

    top_gap = gap_pareto.head(10)
    figure, axis_1 = plt.subplots(figsize=(12, 7))
    axis_1.bar(top_gap["State"], top_gap["Infrastructure_Gap_2030"], color="#0F766E")
    axis_1.tick_params(axis="x", rotation=45)
    axis_1.set_ylabel("2030 infrastructure gap")
    axis_2 = axis_1.twinx()
    axis_2.plot(top_gap["State"], top_gap["Cumulative_Gap_Share"], color="#DC2626", marker="o")
    axis_2.axhline(80, color="#6B7280", linestyle="--")
    axis_2.set_ylabel("Cumulative opportunity share (%)")
    axis_1.set_title("Concentration of 2030 Charging Opportunity")
    finish_figure(figure, plot_directory / "07_2030_gap_pareto.png", save_plots, show_plots)

    top_investors = investors.head(10).sort_values("Investor_Opportunity_Score")
    figure, axis = plt.subplots(figsize=(10, 6))
    axis.barh(top_investors["State"], top_investors["Investor_Opportunity_Score"], color="#7C3AED")
    axis.set(xlabel="Investor opportunity score", ylabel="State", title="Top EV Charging Investment Opportunities")
    finish_figure(figure, plot_directory / "08_investor_ranking.png", save_plots, show_plots)


def build_summary(
    charging: pd.DataFrame,
    market_metrics: dict[str, float],
    stock_forecast: pd.DataFrame,
    investors: pd.DataFrame,
    gap_pareto: pd.DataFrame,
) -> pd.DataFrame:
    """Create a compact executive model summary."""
    infrastructure_column = "Operational Public Charging Infrastructure"
    total_evs = charging["EVs"].sum()
    total_infrastructure = charging[infrastructure_column].sum()
    final_forecast = stock_forecast.loc[stock_forecast["Year"] == 2030].iloc[0]
    total_gap = gap_pareto["Infrastructure_Gap_2030"].sum()
    top_three_share = gap_pareto.head(3)["Infrastructure_Gap_2030"].sum() / total_gap * 100
    rows: list[tuple[str, object]] = [
        ("States analyzed", len(charging)),
        ("Current EVs", round(total_evs)),
        ("Current public charging infrastructure", round(total_infrastructure)),
        ("Current EVs per infrastructure", round(total_evs / total_infrastructure, 2)),
        ("Mean state pressure index", round(charging["Infrastructure_Pressure_Index"].mean(), 3)),
        ("Pearson correlation", round(market_metrics["Pearson correlation"], 4)),
        ("Spearman correlation", round(market_metrics["Spearman correlation"], 4)),
        ("Regression R-squared", round(market_metrics["Regression R-squared"], 4)),
        ("2030 conservative EV stock", round(final_forecast["Conservative_Stock"])),
        ("2030 base EV stock", round(final_forecast["Base_Stock"])),
        ("2030 aggressive EV stock", round(final_forecast["Aggressive_Stock"])),
        ("Total eligible-state infrastructure gap", round(total_gap)),
        ("Top 3 share of eligible-state gap (%)", round(top_three_share, 2)),
        ("Top ranked state", investors.iloc[0]["State"]),
    ]
    return pd.DataFrame(rows, columns=["Metric", "Value"])


def format_excel_workbook(output_path: Path) -> None:
    """Apply simple filters, frozen headers, and readable column widths."""
    workbook = load_workbook(output_path)
    for worksheet in workbook.worksheets:
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        for column_cells in worksheet.columns:
            width = min(
                max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
                + 2,
                45,
            )
            worksheet.column_dimensions[column_cells[0].column_letter].width = max(width, 10)
    workbook.save(output_path)


def export_results(output_path: Path, tables: Mapping[str, pd.DataFrame]) -> None:
    """Write all final tables once; this replaces duplicate notebook exports."""
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, table in tables.items():
            table.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    format_excel_workbook(output_path)


def run_analysis(
    input_path: Path,
    output_directory: Path,
    save_plots: bool = SAVE_PLOTS_BY_DEFAULT,
    show_plots: bool = SHOW_PLOTS_BY_DEFAULT,
) -> Path:
    """Run the complete analysis and return the generated Excel path."""
    input_path = input_path.expanduser().resolve()
    output_directory = output_directory.expanduser().resolve()
    output_directory.mkdir(parents=True, exist_ok=True)

    LOGGER.info("Input workbook: %s", input_path)
    LOGGER.info("Output directory: %s", output_directory)
    charging_raw, historical_raw = load_input_data(input_path)
    charging = clean_charging_data(charging_raw)
    pareto, outliers, regression, market_metrics = build_market_diagnostics(charging)
    clusters, cluster_quality = build_clusters(charging)
    historical_long, growth, annual = build_growth_analysis(historical_raw)
    current_ranking, current_sensitivity = build_current_opportunity_ranking(
        charging, clusters, growth
    )
    forecast_comparison, stock_forecast, required_infrastructure = build_forecasts(annual)
    state_2030, investors, gap_pareto, market_size = build_investor_analysis(
        charging, growth
    )
    create_plots(
        output_directory,
        charging,
        pareto,
        regression,
        clusters,
        current_sensitivity,
        annual,
        stock_forecast,
        gap_pareto,
        investors,
        save_plots,
        show_plots,
    )

    output_path = output_directory / OUTPUT_EXCEL_NAME
    export_results(
        output_path,
        {
            "Historical_EV": normalize_historical_headers(historical_raw.copy()),
            "Charging_Infrastructure": charging[
                list(CHARGING_INFRASTRUCTURE_COLUMN_MAP)
            ].rename(columns=CHARGING_INFRASTRUCTURE_COLUMN_MAP),
            "State_2030": state_2030,
            "Investor_Ranking": investors[INVESTOR_RANKING_COLUMNS].copy(),
            "EV_Forecast": stock_forecast,
        },
    )

    LOGGER.info("Analysis complete: %s", output_path)
    LOGGER.info("Top investment opportunities:\n%s", investors.head(10)[
        ["Rank", "State", "Investor_Opportunity_Score", "Infrastructure_Gap_2030", "Recommendation"]
    ].to_string(index=False))
    return output_path


def main() -> None:
    """Command-line entry point."""
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    args = parse_args()
    try:
        output_path = run_analysis(
            args.input,
            args.output,
            save_plots=not args.no_plots,
            show_plots=args.show_plots,
        )
    except (FileNotFoundError, ValueError, KeyError) as error:
        LOGGER.error("%s", error)
        raise SystemExit(1) from error
    print(f"\nFinished. Main output file:\n{output_path}")


if __name__ == "__main__":
    main()
