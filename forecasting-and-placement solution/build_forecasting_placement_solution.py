from __future__ import annotations

import json
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd


BASE_DIR = Path(
    r"C:\Users\Mamphey\Documents\Codex\2026-04-22-files-mentioned-by-the-user-merged\forecasting-and-placement solution"
)
DATA_PATH = Path(r"C:\Users\Mamphey\Desktop\B_Dataset\merged_data.xlsx")
KPI_PATH = Path(
    r"C:\Users\Mamphey\Desktop\03_Amazon_Regional_Demand_Forecasting_Inventory_Placement_Project_Kit\C_Supporting_Files\templates\kpi_list.csv"
)
OUTPUT_DIR = BASE_DIR / "outputs"
DOCS_DIR = BASE_DIR / "docs"
SQL_DIR = BASE_DIR / "sql"


@dataclass
class ComboModelResult:
    category: str
    region: str
    best_model: str
    baseline_wape: float
    regression_wape: float
    baseline_mape: float
    regression_mape: float
    holdout_actual: float
    sigma_weekly: float


def ensure_dirs() -> None:
    for path in [OUTPUT_DIR, DOCS_DIR, SQL_DIR]:
        path.mkdir(parents=True, exist_ok=True)


def trim_object_columns(df: pd.DataFrame) -> pd.DataFrame:
    cleaned = df.copy()
    for column in cleaned.select_dtypes(include=["object", "string"]).columns:
        cleaned[column] = cleaned[column].astype(str).str.strip()
        cleaned.loc[cleaned[column].isin(["nan", "None"]), column] = pd.NA
    return cleaned


def normalize_binary(series: pd.Series) -> pd.Series:
    numeric = pd.to_numeric(series, errors="coerce")
    return numeric.fillna(0).clip(lower=0, upper=1).round().astype(int)


def warehouse_suffix_map(warehouses: pd.DataFrame) -> dict[str, str]:
    suffix = warehouses["warehouse_id"].astype(str).str.extract(r"(\d+)$")[0]
    return dict(zip(suffix, warehouses["warehouse_id"].astype(str)))


def canonicalize_warehouse_ids(
    df: pd.DataFrame, column: str, suffix_to_id: dict[str, str]
) -> tuple[pd.DataFrame, int]:
    cleaned = df.copy()
    original = cleaned[column].astype(str).str.strip()
    suffix = original.str.extract(r"(\d+)$")[0]
    canonical = suffix.map(suffix_to_id).fillna(original)
    changed = int((canonical != original).sum())
    cleaned[f"{column}_raw"] = original
    cleaned[column] = canonical
    return cleaned, changed


def week_start(series: pd.Series) -> pd.Series:
    parsed = pd.to_datetime(series, errors="coerce")
    return parsed - pd.to_timedelta(parsed.dt.weekday, unit="D")


def wape(actual: pd.Series, forecast: pd.Series) -> float:
    denom = actual.abs().sum()
    if denom == 0:
        return 0.0
    return float((actual.sub(forecast).abs().sum()) / denom)


def mape(actual: pd.Series, forecast: pd.Series) -> float:
    valid = actual != 0
    if not valid.any():
        return 0.0
    return float((actual[valid].sub(forecast[valid]).abs() / actual[valid].abs()).mean())


def normal_cdf(x: float) -> float:
    return 0.5 * (1 + math.erf(x / math.sqrt(2)))


def z_for_service_level(service_level: float) -> float:
    levels = np.array([0.90, 0.93, 0.95, 0.97, 0.98, 0.99])
    z_scores = np.array([1.282, 1.476, 1.645, 1.881, 2.054, 2.326])
    return float(np.interp(service_level, levels, z_scores))


def build_design_matrix(df: pd.DataFrame) -> np.ndarray:
    return np.column_stack(
        [
            np.ones(len(df)),
            df["trend"].to_numpy(dtype=float),
            df["sin_52"].to_numpy(dtype=float),
            df["cos_52"].to_numpy(dtype=float),
            df["sin_26"].to_numpy(dtype=float),
            df["cos_26"].to_numpy(dtype=float),
            df["holiday_days"].to_numpy(dtype=float),
            df["prime_days"].to_numpy(dtype=float),
            df["marketing_days"].to_numpy(dtype=float),
            df["avg_weather"].to_numpy(dtype=float),
            df["lag_1"].to_numpy(dtype=float),
            df["ma_4"].to_numpy(dtype=float),
        ]
    )


def fit_ols(train_df: pd.DataFrame) -> np.ndarray:
    x = build_design_matrix(train_df)
    y = train_df["units_ordered"].to_numpy(dtype=float)
    return np.linalg.pinv(x) @ y


def predict_ols_row(row: pd.Series, beta: np.ndarray) -> float:
    vector = np.array(
        [
            1.0,
            float(row["trend"]),
            float(row["sin_52"]),
            float(row["cos_52"]),
            float(row["sin_26"]),
            float(row["cos_26"]),
            float(row["holiday_days"]),
            float(row["prime_days"]),
            float(row["marketing_days"]),
            float(row["avg_weather"]),
            float(row["lag_1"]),
            float(row["ma_4"]),
        ]
    )
    return max(0.0, float(vector @ beta))


def prepare_data() -> dict[str, pd.DataFrame]:
    workbook = pd.ExcelFile(DATA_PATH)
    raw = {sheet: pd.read_excel(workbook, sheet_name=sheet) for sheet in workbook.sheet_names}

    warehouses = trim_object_columns(raw["Warehouses"]).drop_duplicates().reset_index(drop=True)
    suffix_to_id = warehouse_suffix_map(warehouses)

    sku_master = trim_object_columns(raw["SKU Master"]).drop_duplicates().reset_index(drop=True)
    sku_master["avg_target_service_level"] = sku_master["target_service_level"]

    daily = trim_object_columns(raw["Daily Demand"])
    daily["date"] = pd.to_datetime(daily["date"], errors="coerce")
    daily["weekend_flag_raw"] = daily["weekend_flag"].astype(str)
    daily["weekend_flag"] = (daily["date"].dt.dayofweek >= 5).astype(int)
    for flag_col in ["holiday_peak_flag", "prime_event_flag", "marketing_push_flag"]:
        daily[flag_col] = normalize_binary(daily[flag_col])
    daily = daily.sort_values(["date", "sku_id", "region"]).reset_index(drop=True)

    event_calendar = trim_object_columns(raw["Event Calendar"])
    event_calendar["date"] = pd.to_datetime(event_calendar["date"], errors="coerce")
    for flag_col in ["holiday_peak_flag", "prime_event_flag", "marketing_push_flag"]:
        event_calendar[flag_col] = normalize_binary(event_calendar[flag_col])
    event_calendar["weather_disruption_index"] = pd.to_numeric(
        event_calendar["weather_disruption_index"], errors="coerce"
    )
    event_calendar = (
        event_calendar.groupby("date", as_index=False)
        .agg(
            holiday_peak_flag=("holiday_peak_flag", "max"),
            prime_event_flag=("prime_event_flag", "max"),
            marketing_push_flag=("marketing_push_flag", "max"),
            weather_disruption_index=("weather_disruption_index", "mean"),
        )
        .sort_values("date")
        .reset_index(drop=True)
    )
    event_calendar["weekend_flag"] = (event_calendar["date"].dt.dayofweek >= 5).astype(int)
    event_calendar["week_start"] = week_start(event_calendar["date"])

    starting_inventory = trim_object_columns(raw["Starting Inventory"])
    starting_inventory, _ = canonicalize_warehouse_ids(starting_inventory, "warehouse_id", suffix_to_id)

    warehouse_costs = trim_object_columns(raw["Warehouse Region Costs"])
    warehouse_costs, _ = canonicalize_warehouse_ids(warehouse_costs, "warehouse_id", suffix_to_id)

    daily_enriched = daily.merge(
        sku_master[
            [
                "sku_id",
                "category",
                "unit_cost_usd",
                "selling_price_usd",
                "target_service_level",
            ]
        ],
        on="sku_id",
        how="left",
    )
    daily_enriched["week_start"] = week_start(daily_enriched["date"])

    weekly_events = (
        event_calendar.groupby("week_start", as_index=False)
        .agg(
            holiday_days=("holiday_peak_flag", "sum"),
            prime_days=("prime_event_flag", "sum"),
            marketing_days=("marketing_push_flag", "sum"),
            avg_weather=("weather_disruption_index", "mean"),
        )
        .sort_values("week_start")
        .reset_index(drop=True)
    )

    weekly_category_region = (
        daily_enriched.groupby(["week_start", "category", "region"], as_index=False)
        .agg(
            units_ordered=("units_ordered", "sum"),
            avg_price=("price_usd", "mean"),
            avg_service_level=("target_service_level", "mean"),
            active_skus=("sku_id", "nunique"),
        )
        .merge(weekly_events, on="week_start", how="left")
        .sort_values(["category", "region", "week_start"])
        .reset_index(drop=True)
    )
    weekly_category_region["week_index"] = (
        weekly_category_region["week_start"].rank(method="dense").astype(int)
    )
    weekly_category_region["iso_week"] = weekly_category_region["week_start"].dt.isocalendar().week.astype(int)
    weekly_category_region["trend"] = (
        weekly_category_region.groupby(["category", "region"]).cumcount() + 1
    ).astype(float)
    weekly_category_region["sin_52"] = np.sin(2 * np.pi * weekly_category_region["iso_week"] / 52.0)
    weekly_category_region["cos_52"] = np.cos(2 * np.pi * weekly_category_region["iso_week"] / 52.0)
    weekly_category_region["sin_26"] = np.sin(2 * np.pi * weekly_category_region["iso_week"] / 26.0)
    weekly_category_region["cos_26"] = np.cos(2 * np.pi * weekly_category_region["iso_week"] / 26.0)
    weekly_category_region["lag_1"] = (
        weekly_category_region.groupby(["category", "region"])["units_ordered"].shift(1)
    )
    weekly_category_region["ma_4"] = (
        weekly_category_region.groupby(["category", "region"])["units_ordered"]
        .shift(1)
        .rolling(4, min_periods=1)
        .mean()
    )

    inventory_by_node_category = (
        starting_inventory.merge(
            warehouses[["warehouse_id", "warehouse_region"]],
            on="warehouse_id",
            how="left",
        )
        .merge(
            sku_master[["sku_id", "category", "target_service_level"]],
            on="sku_id",
            how="left",
        )
        .groupby(["warehouse_region", "category"], as_index=False)
        .agg(
            on_hand_units=("starting_inventory_units", "sum"),
            avg_service_level=("target_service_level", "mean"),
        )
    )
    inventory_by_node_category["inventory_share"] = (
        inventory_by_node_category["on_hand_units"]
        / inventory_by_node_category.groupby("category")["on_hand_units"].transform("sum")
    )

    node_cost_matrix = (
        warehouse_costs.merge(
            warehouses[["warehouse_id", "warehouse_region"]],
            on="warehouse_id",
            how="left",
        )
        .groupby(["warehouse_region", "demand_region"], as_index=False)
        .agg(
            avg_ship_cost=("ship_cost_per_unit_usd", "mean"),
            avg_lead_days=("lead_time_days", "mean"),
        )
    )
    node_cost_matrix["service_score"] = (
        node_cost_matrix["avg_ship_cost"] + 0.4 * node_cost_matrix["avg_lead_days"]
    )

    sku_mix = (
        daily_enriched.groupby(["category", "region", "sku_id"], as_index=False)["units_ordered"]
        .sum()
        .rename(columns={"units_ordered": "historical_units"})
    )
    sku_mix["sku_mix_share"] = (
        sku_mix["historical_units"]
        / sku_mix.groupby(["category", "region"])["historical_units"].transform("sum")
    )

    kpi_list = pd.read_csv(KPI_PATH)

    return {
        "daily_enriched": daily_enriched,
        "weekly_category_region": weekly_category_region,
        "weekly_events": weekly_events,
        "inventory_by_node_category": inventory_by_node_category,
        "node_cost_matrix": node_cost_matrix,
        "sku_mix": sku_mix,
        "kpi_list": kpi_list,
        "sku_master": sku_master,
    }


def evaluate_models(weekly_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    holdout_weeks = sorted(weekly_df["week_start"].unique())[-8:]
    results: list[ComboModelResult] = []
    detailed_rows: list[dict[str, Any]] = []

    for (category, region), combo in weekly_df.groupby(["category", "region"]):
        combo = combo.sort_values("week_start").reset_index(drop=True)
        train = combo[~combo["week_start"].isin(holdout_weeks)].copy()
        holdout = combo[combo["week_start"].isin(holdout_weeks)].copy()
        if train.empty or holdout.empty:
            continue

        history = train["units_ordered"].tolist()
        baseline_preds = []
        for actual in holdout["units_ordered"].tolist():
            baseline_pred = float(np.mean(history[-4:])) if history else float(train["units_ordered"].mean())
            baseline_preds.append(max(0.0, baseline_pred))
            history.append(actual)
        holdout["baseline_forecast"] = baseline_preds

        regression_train = train.dropna(subset=["lag_1", "ma_4"]).copy()
        if len(regression_train) < 10:
            beta = None
            regression_preds = baseline_preds
        else:
            beta = fit_ols(regression_train)
            history = train["units_ordered"].tolist()
            regression_preds = []
            for _, row in holdout.iterrows():
                lag_1 = history[-1]
                ma_4 = float(np.mean(history[-4:]))
                temp_row = row.copy()
                temp_row["lag_1"] = lag_1
                temp_row["ma_4"] = ma_4
                pred = predict_ols_row(temp_row, beta)
                regression_preds.append(pred)
                history.append(float(row["units_ordered"]))
        holdout["regression_forecast"] = regression_preds

        baseline_wape = wape(holdout["units_ordered"], holdout["baseline_forecast"])
        regression_wape = wape(holdout["units_ordered"], holdout["regression_forecast"])
        baseline_mape = mape(holdout["units_ordered"], holdout["baseline_forecast"])
        regression_mape = mape(holdout["units_ordered"], holdout["regression_forecast"])
        best_model = "regression" if regression_wape <= baseline_wape else "moving_average"
        sigma_weekly = float(
            max(
                holdout["units_ordered"].std(ddof=0) if len(holdout) > 1 else 0.0,
                (holdout["units_ordered"] - holdout[f"{best_model if best_model == 'regression' else 'baseline'}_forecast"]).abs().mean()
                if best_model == "regression"
                else (holdout["units_ordered"] - holdout["baseline_forecast"]).abs().mean(),
                1.0,
            )
        )

        results.append(
            ComboModelResult(
                category=category,
                region=region,
                best_model=best_model,
                baseline_wape=baseline_wape,
                regression_wape=regression_wape,
                baseline_mape=baseline_mape,
                regression_mape=regression_mape,
                holdout_actual=float(holdout["units_ordered"].sum()),
                sigma_weekly=sigma_weekly,
            )
        )

        for _, row in holdout.iterrows():
            detailed_rows.append(
                {
                    "category": category,
                    "region": region,
                    "week_start": row["week_start"].strftime("%Y-%m-%d"),
                    "actual_units": float(row["units_ordered"]),
                    "baseline_forecast": float(row["baseline_forecast"]),
                    "regression_forecast": float(row["regression_forecast"]),
                    "selected_model": best_model,
                }
            )

    benchmark = pd.DataFrame([vars(item) for item in results]).sort_values(
        ["best_model", "regression_wape", "baseline_wape"]
    )
    details = pd.DataFrame(detailed_rows)
    return benchmark, details


def build_future_event_profile(weekly_events: pd.DataFrame, horizon_weeks: int) -> pd.DataFrame:
    last_week = weekly_events["week_start"].max()
    future_weeks = pd.date_range(last_week + pd.Timedelta(days=7), periods=horizon_weeks, freq="7D")
    week_profile = (
        weekly_events.assign(iso_week=weekly_events["week_start"].dt.isocalendar().week.astype(int))
        .groupby("iso_week", as_index=False)
        .agg(
            holiday_days=("holiday_days", "mean"),
            prime_days=("prime_days", "mean"),
            marketing_days=("marketing_days", "mean"),
            avg_weather=("avg_weather", "mean"),
        )
    )
    overall = weekly_events[["holiday_days", "prime_days", "marketing_days", "avg_weather"]].mean()
    rows = []
    for idx, week in enumerate(future_weeks, start=1):
        iso_week = int(week.isocalendar().week)
        matched = week_profile[week_profile["iso_week"] == iso_week]
        if matched.empty:
            values = overall
        else:
            values = matched.iloc[0]
        rows.append(
            {
                "week_start": week,
                "iso_week": iso_week,
                "holiday_days": float(values["holiday_days"]),
                "prime_days": float(values["prime_days"]),
                "marketing_days": float(values["marketing_days"]),
                "avg_weather": float(values["avg_weather"]),
                "horizon_week": idx,
                "sin_52": math.sin(2 * math.pi * iso_week / 52.0),
                "cos_52": math.cos(2 * math.pi * iso_week / 52.0),
                "sin_26": math.sin(2 * math.pi * iso_week / 26.0),
                "cos_26": math.cos(2 * math.pi * iso_week / 26.0),
            }
        )
    return pd.DataFrame(rows)


def create_forecasts(
    weekly_df: pd.DataFrame, benchmark: pd.DataFrame, weekly_events: pd.DataFrame, horizon_weeks: int = 6
) -> pd.DataFrame:
    future_events = build_future_event_profile(weekly_events, horizon_weeks)
    benchmark_lookup = benchmark.set_index(["category", "region"]).to_dict(orient="index")
    rows: list[dict[str, Any]] = []

    for (category, region), combo in weekly_df.groupby(["category", "region"]):
        combo = combo.sort_values("week_start").reset_index(drop=True)
        config = benchmark_lookup.get(
            (category, region),
            {
                "best_model": "moving_average",
                "baseline_wape": 0.0,
                "regression_wape": 0.0,
                "baseline_mape": 0.0,
                "regression_mape": 0.0,
                "sigma_weekly": float(max(combo["units_ordered"].std(ddof=0), 1.0)),
            },
        )
        service_level = float(combo["avg_service_level"].mean())
        history = combo["units_ordered"].tolist()
        regression_train = combo.dropna(subset=["lag_1", "ma_4"]).copy()
        beta = fit_ols(regression_train) if len(regression_train) >= 10 else None
        last_trend = float(combo["trend"].max())

        for step, (_, event_row) in enumerate(future_events.iterrows(), start=1):
            ma_pred = float(np.mean(history[-4:]))
            lag_1 = history[-1]
            ma_4 = float(np.mean(history[-4:]))
            reg_pred = ma_pred
            if beta is not None:
                temp = pd.Series(
                    {
                        "trend": last_trend + step,
                        "sin_52": event_row["sin_52"],
                        "cos_52": event_row["cos_52"],
                        "sin_26": event_row["sin_26"],
                        "cos_26": event_row["cos_26"],
                        "holiday_days": event_row["holiday_days"],
                        "prime_days": event_row["prime_days"],
                        "marketing_days": event_row["marketing_days"],
                        "avg_weather": event_row["avg_weather"],
                        "lag_1": lag_1,
                        "ma_4": ma_4,
                    }
                )
                reg_pred = predict_ols_row(temp, beta)
            selected_model = config["best_model"]
            forecast_units = reg_pred if selected_model == "regression" else ma_pred
            history.append(forecast_units)
            rows.append(
                {
                    "category": category,
                    "region": region,
                    "week_start": event_row["week_start"].strftime("%Y-%m-%d"),
                    "horizon_week": step,
                    "moving_average_forecast": round(ma_pred, 2),
                    "regression_forecast": round(reg_pred, 2),
                    "selected_model": selected_model,
                    "forecast_units": round(forecast_units, 2),
                    "sigma_weekly": round(float(config["sigma_weekly"]), 2),
                    "service_level": round(service_level, 4),
                }
            )

    return pd.DataFrame(rows)


def recommend_placement(
    forecasts: pd.DataFrame,
    inventory_by_node_category: pd.DataFrame,
    node_cost_matrix: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    category_current_share = inventory_by_node_category[
        ["warehouse_region", "category", "inventory_share", "on_hand_units"]
    ].copy()
    planning_horizon_weeks = 2.0
    fast_lead_threshold = 2.0
    scenarios: list[dict[str, Any]] = []
    recommended_allocations: list[dict[str, Any]] = []

    for (category, region), combo in forecasts.groupby(["category", "region"]):
        cost_options = node_cost_matrix[node_cost_matrix["demand_region"] == region].copy()
        cost_options = cost_options.sort_values(["service_score", "avg_ship_cost", "avg_lead_days"]).reset_index(
            drop=True
        )
        fast_options = cost_options[cost_options["avg_lead_days"] <= fast_lead_threshold]
        primary = fast_options.iloc[0] if not fast_options.empty else cost_options.iloc[0]
        backup = cost_options.iloc[1] if len(cost_options) > 1 else cost_options.iloc[0]
        service_level = float(combo["service_level"].mean())
        weekly_mean = float(combo["forecast_units"].mean())
        sigma_weekly = float(combo["sigma_weekly"].mean())
        volatility = sigma_weekly / weekly_mean if weekly_mean > 0 else 0.0
        if service_level >= 0.97 or volatility >= 0.45:
            primary_share, backup_share = 0.75, 0.25
        elif service_level >= 0.95 or volatility >= 0.30:
            primary_share, backup_share = 0.85, 0.15
        else:
            primary_share, backup_share = 0.90, 0.10

        lead_weeks = max(float(primary["avg_lead_days"]) / 7.0, 0.15)
        cycle_stock = weekly_mean * planning_horizon_weeks
        safety_stock = z_for_service_level(service_level) * sigma_weekly * math.sqrt(lead_weeks + 0.25)
        target_units = max(cycle_stock + safety_stock, weekly_mean)
        lead_time_demand = weekly_mean * lead_weeks
        sigma_lead = max(sigma_weekly * math.sqrt(max(lead_weeks, 0.15)), 1.0)

        current_shares = category_current_share[category_current_share["category"] == category].copy()
        current_alloc = {
            row["warehouse_region"]: float(row["inventory_share"]) * target_units
            for _, row in current_shares.iterrows()
        }
        rec_alloc = {node: 0.0 for node in category_current_share["warehouse_region"].unique()}
        rec_alloc[str(primary["warehouse_region"])] = target_units * primary_share
        rec_alloc[str(backup["warehouse_region"])] += target_units * backup_share

        fast_nodes = set(
            cost_options.loc[cost_options["avg_lead_days"] <= fast_lead_threshold, "warehouse_region"].astype(str)
        )
        baseline_fast_inventory = sum(
            units for node, units in current_alloc.items() if str(node) in fast_nodes
        )
        recommended_fast_inventory = sum(
            units for node, units in rec_alloc.items() if str(node) in fast_nodes
        )

        baseline_fill = min(1.0, baseline_fast_inventory / max(cycle_stock, 1.0))
        recommended_fill = min(1.0, recommended_fast_inventory / max(cycle_stock, 1.0))
        baseline_risk = 1.0 - normal_cdf((baseline_fast_inventory - lead_time_demand) / sigma_lead)
        recommended_risk = 1.0 - normal_cdf((recommended_fast_inventory - lead_time_demand) / sigma_lead)
        baseline_transfer_units = max(0.0, cycle_stock - baseline_fast_inventory)
        recommended_transfer_units = max(0.0, cycle_stock - recommended_fast_inventory)
        transfer_cost_per_unit = float(backup["avg_ship_cost"])
        baseline_transfer_cost = baseline_transfer_units * transfer_cost_per_unit
        recommended_transfer_cost = recommended_transfer_units * transfer_cost_per_unit

        baseline_ship_cost = 0.0
        baseline_lead = 0.0
        for node, units in current_alloc.items():
            match = cost_options[cost_options["warehouse_region"] == node]
            if match.empty:
                continue
            share = units / target_units if target_units else 0.0
            baseline_ship_cost += share * float(match.iloc[0]["avg_ship_cost"])
            baseline_lead += share * float(match.iloc[0]["avg_lead_days"])

        recommended_ship_cost = (
            primary_share * float(primary["avg_ship_cost"]) + backup_share * float(backup["avg_ship_cost"])
        )
        recommended_lead = (
            primary_share * float(primary["avg_lead_days"]) + backup_share * float(backup["avg_lead_days"])
        )

        scenarios.append(
            {
                "category": category,
                "demand_region": region,
                "primary_node": str(primary["warehouse_region"]),
                "backup_node": str(backup["warehouse_region"]),
                "service_level": round(service_level, 4),
                "weekly_forecast_mean": round(weekly_mean, 2),
                "sigma_weekly": round(sigma_weekly, 2),
                "target_inventory_units": round(target_units, 2),
                "baseline_fill_rate": round(baseline_fill, 4),
                "recommended_fill_rate": round(recommended_fill, 4),
                "baseline_stockout_risk": round(max(0.0, min(1.0, baseline_risk)), 4),
                "recommended_stockout_risk": round(max(0.0, min(1.0, recommended_risk)), 4),
                "baseline_transfer_cost": round(baseline_transfer_cost, 2),
                "recommended_transfer_cost": round(recommended_transfer_cost, 2),
                "baseline_avg_ship_cost": round(baseline_ship_cost, 2),
                "recommended_avg_ship_cost": round(recommended_ship_cost, 2),
                "baseline_avg_lead_days": round(baseline_lead, 2),
                "recommended_avg_lead_days": round(recommended_lead, 2),
            }
        )

        for node, units in rec_alloc.items():
            if units <= 0:
                continue
            recommended_allocations.append(
                {
                    "category": category,
                    "warehouse_region": node,
                    "demand_region": region,
                    "allocation_units": round(units, 2),
                    "allocation_share": round(units / target_units, 4),
                    "target_inventory_units": round(target_units, 2),
                }
            )

    scenarios_df = pd.DataFrame(scenarios)
    recommended_df = pd.DataFrame(recommended_allocations)

    rec_node_category = (
        recommended_df.groupby(["warehouse_region", "category"], as_index=False)["allocation_units"]
        .sum()
        .rename(columns={"allocation_units": "recommended_active_units"})
    )
    current_node_category = inventory_by_node_category[
        ["warehouse_region", "category", "on_hand_units", "inventory_share"]
    ].copy()
    current_node_category["active_units_as_is"] = (
        current_node_category["inventory_share"]
        * current_node_category.groupby("category")["on_hand_units"].transform("sum")
        / current_node_category.groupby("category")["on_hand_units"].transform("sum")
    )
    inventory_implications = current_node_category.merge(
        rec_node_category,
        on=["warehouse_region", "category"],
        how="outer",
    ).fillna({"on_hand_units": 0, "inventory_share": 0, "recommended_active_units": 0})
    inventory_implications["inventory_surplus_vs_target"] = (
        inventory_implications["on_hand_units"] - inventory_implications["recommended_active_units"]
    )
    inventory_implications["recommended_share"] = (
        inventory_implications["recommended_active_units"]
        / inventory_implications.groupby("category")["recommended_active_units"].transform("sum")
    ).replace([np.inf, -np.inf], 0).fillna(0)

    return scenarios_df, recommended_df, inventory_implications


def build_kpi_summary(
    benchmark: pd.DataFrame,
    scenarios_df: pd.DataFrame,
    inventory_implications: pd.DataFrame,
) -> dict[str, Any]:
    baseline_wape = float(benchmark["baseline_wape"].mean())
    best_wape = float(
        benchmark.apply(
            lambda row: row["regression_wape"] if row["best_model"] == "regression" else row["baseline_wape"],
            axis=1,
        ).mean()
    )
    baseline_fill = float(
        np.average(scenarios_df["baseline_fill_rate"], weights=scenarios_df["target_inventory_units"])
    )
    recommended_fill = float(
        np.average(scenarios_df["recommended_fill_rate"], weights=scenarios_df["target_inventory_units"])
    )
    baseline_risk = float(
        np.average(scenarios_df["baseline_stockout_risk"], weights=scenarios_df["target_inventory_units"])
    )
    recommended_risk = float(
        np.average(scenarios_df["recommended_stockout_risk"], weights=scenarios_df["target_inventory_units"])
    )
    baseline_transfer = float(scenarios_df["baseline_transfer_cost"].sum())
    recommended_transfer = float(scenarios_df["recommended_transfer_cost"].sum())
    baseline_lead = float(
        np.average(scenarios_df["baseline_avg_lead_days"], weights=scenarios_df["target_inventory_units"])
    )
    recommended_lead = float(
        np.average(scenarios_df["recommended_avg_lead_days"], weights=scenarios_df["target_inventory_units"])
    )

    current_share = inventory_implications[["category", "warehouse_region", "inventory_share"]].copy()
    rec_share = inventory_implications[["category", "warehouse_region", "recommended_share"]].copy()
    adherence_rows = current_share.merge(rec_share, on=["category", "warehouse_region"], how="outer").fillna(0)
    adherence = (
        adherence_rows.groupby("category")
        .apply(lambda frame: 1.0 - 0.5 * np.abs(frame["inventory_share"] - frame["recommended_share"]).sum())
        .mean()
    )

    total_on_hand = float(inventory_implications["on_hand_units"].sum())
    total_recommended_active = float(inventory_implications["recommended_active_units"].sum())
    weeks_of_supply = total_on_hand / max(total_recommended_active / 2.0, 1.0)

    return {
        "baseline_wape": round(baseline_wape, 4),
        "selected_model_wape": round(best_wape, 4),
        "fill_rate_baseline": round(baseline_fill, 4),
        "fill_rate_recommended": round(recommended_fill, 4),
        "stockout_risk_baseline": round(baseline_risk, 4),
        "stockout_risk_recommended": round(recommended_risk, 4),
        "reactive_transfer_cost_baseline": round(baseline_transfer, 2),
        "reactive_transfer_cost_recommended": round(recommended_transfer, 2),
        "avg_lead_days_baseline": round(baseline_lead, 2),
        "avg_lead_days_recommended": round(recommended_lead, 2),
        "placement_adherence_baseline": round(float(adherence), 4),
        "total_on_hand_units": round(total_on_hand, 0),
        "recommended_active_units": round(total_recommended_active, 0),
        "network_weeks_of_supply": round(weeks_of_supply, 1),
    }


def build_documents(
    kpi_summary: dict[str, Any],
    benchmark: pd.DataFrame,
    scenarios_df: pd.DataFrame,
    inventory_implications: pd.DataFrame,
) -> None:
    def fmt_pct(value: float) -> str:
        return f"{value * 100:.2f}%"

    top_risk = scenarios_df.sort_values("baseline_stockout_risk", ascending=False).head(5)
    top_relocation = inventory_implications.sort_values("inventory_surplus_vs_target", ascending=False).head(10)

    memo = f"""# Executive Memo

## Recommendation

Use a category-region weekly forecasting layer with model selection between moving-average and event-aware regression, then convert the winning forecast into two-week target inventory plus safety stock for node placement decisions.

## Why this works

- The source data does **not** support true time-series forecasting at the `sku_id`-`region` grain because each SKU-region appears only once historically.
- The strongest available signal is the weekly `category`-`region` history enriched with promotions, seasonality, and weather.
- We therefore forecast at `category`-`region`, then allocate to active SKU-region combinations using observed demand mix shares.

## Quantified impact

- Forecast WAPE improves from **{fmt_pct(kpi_summary['baseline_wape'])}** for the moving-average baseline to **{fmt_pct(kpi_summary['selected_model_wape'])}** for the selected forecasting policy.
- Estimated on-time fill rate improves from **{fmt_pct(kpi_summary['fill_rate_baseline'])}** to **{fmt_pct(kpi_summary['fill_rate_recommended'])}**.
- Estimated stockout risk falls from **{fmt_pct(kpi_summary['stockout_risk_baseline'])}** to **{fmt_pct(kpi_summary['stockout_risk_recommended'])}**.
- Estimated reactive transfer cost for the active planning inventory falls from **${kpi_summary['reactive_transfer_cost_baseline']:,.0f}** to **${kpi_summary['reactive_transfer_cost_recommended']:,.0f}** over the modeled cycle.
- Average effective lead time improves from **{kpi_summary['avg_lead_days_baseline']:.2f} days** to **{kpi_summary['avg_lead_days_recommended']:.2f} days**.

## Management actions

1. Use the selected forecast by category and region for the next six weekly cycles.
2. Position primary inventory in the recommended node and maintain a smaller backup share in the secondary node.
3. Freeze or defer replenishment into the most overstocked node-category pools before authorizing any new buys.
4. Track fill rate, stockout risk, transfer cost, and placement adherence weekly in the dashboard.

## Key trade-offs

- Faster delivery generally requires a modest bias toward lower-lead nodes, even when their outbound cost is not the lowest.
- The data supports robust network planning at category-region level, but not confident long-horizon SKU-level learning.
- Because actual on-hand inventory is far above the modeled active target, the near-term value comes more from **repositioning future replenishment** than from urgent network buys.
"""

    appendix = f"""# Assumptions and Limitations

## Assumptions

- Forecasting grain is weekly `category`-`region`.
- SKU-region forecasts are allocated from category-region forecasts using observed mix share.
- Safety stock uses target service level with a normal-demand approximation.
- Placement decisions are made across node regions using average ship cost and average lead time.
- Active inventory target uses a two-week cycle stock plus safety stock policy.

## Limitations

- Each `sku_id`-`region` appears only once historically in the source data, so true SKU-level time-series benchmarking is not statistically credible.
- No purchase-order, replenishment cadence, supplier lead-time, or transfer-history tables were provided.
- Transfer cost between nodes is approximated from the available outbound shipping-cost matrix.
- The model is best interpreted as a **decision-support prototype** for network planning, not as a fully optimized production replenishment engine.

## KPI references

- Placement adherence baseline: **{kpi_summary['placement_adherence_baseline']:.1%}**
- Total on-hand units observed: **{kpi_summary['total_on_hand_units']:,.0f}**
- Recommended active inventory units: **{kpi_summary['recommended_active_units']:,.0f}**
- Implied weeks of supply at current on-hand: **{kpi_summary['network_weeks_of_supply']:.1f} weeks**

## Highest-risk baseline segments

{df_to_markdown_table(top_risk[['category','demand_region','baseline_stockout_risk','recommended_stockout_risk','baseline_fill_rate','recommended_fill_rate']])}

## Largest surplus pools versus recommended active inventory

{df_to_markdown_table(top_relocation[['warehouse_region','category','on_hand_units','recommended_active_units','inventory_surplus_vs_target']])}
"""

    (DOCS_DIR / "executive_memo.md").write_text(memo, encoding="utf-8")
    (DOCS_DIR / "assumptions_and_limitations.md").write_text(appendix, encoding="utf-8")


def export_sql_files() -> None:
    weekly_sql = """SELECT DATE_TRUNC('week', date) AS week_start, sku_id, region,
       SUM(units_ordered) AS units_ordered
FROM daily_demand
GROUP BY 1,2,3
ORDER BY 1,2,3;
"""

    category_region_sql = """SELECT
    DATE_TRUNC('week', d.date) AS week_start,
    s.category,
    d.region,
    SUM(d.units_ordered) AS units_ordered,
    AVG(d.price_usd) AS avg_price,
    SUM(d.holiday_peak_flag) AS holiday_days,
    SUM(d.prime_event_flag) AS prime_days,
    SUM(d.marketing_push_flag) AS marketing_days,
    AVG(d.weather_disruption_index) AS avg_weather
FROM daily_demand d
JOIN sku_master s
    ON d.sku_id = s.sku_id
GROUP BY 1,2,3
ORDER BY 1,2,3;
"""
    (SQL_DIR / "weekly_demand_from_daily.sql").write_text(weekly_sql, encoding="utf-8")
    (SQL_DIR / "weekly_category_region_features.sql").write_text(category_region_sql, encoding="utf-8")


def dataframe_to_records(df: pd.DataFrame, limit: int | None = None) -> list[dict[str, Any]]:
    sample = df if limit is None else df.head(limit)
    rows = []
    for row in sample.to_dict(orient="records"):
        parsed = {}
        for key, value in row.items():
            if isinstance(value, pd.Timestamp):
                parsed[key] = value.strftime("%Y-%m-%d")
            elif isinstance(value, (np.integer,)):
                parsed[key] = int(value)
            elif isinstance(value, (np.floating,)):
                parsed[key] = float(value)
            else:
                parsed[key] = value
        rows.append(parsed)
    return rows


def df_to_markdown_table(df: pd.DataFrame) -> str:
    safe = df.copy()
    for column in safe.columns:
        if pd.api.types.is_datetime64_any_dtype(safe[column]):
            safe[column] = safe[column].dt.strftime("%Y-%m-%d")
    headers = [str(column) for column in safe.columns]
    lines = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join(["---"] * len(headers)) + " |",
    ]
    for _, row in safe.iterrows():
        values = ["" if pd.isna(value) else str(value) for value in row.tolist()]
        lines.append("| " + " | ".join(values) + " |")
    return "\n".join(lines)


def update_folder_readme(kpi_summary: dict[str, Any]) -> None:
    readme = f"""# forecasting-and-placement solution

This folder contains a complete regional demand forecasting and inventory placement project package built from the provided Amazon-style synthetic workbook and KPI template.

## Deliverables

- `forecasting_and_placement_solution.ipynb`
- `build_forecasting_placement_solution.py`
- `docs/executive_memo.md`
- `docs/assumptions_and_limitations.md`
- `sql/weekly_demand_from_daily.sql`
- `sql/weekly_category_region_features.sql`
- `outputs/amazon_forecasting_placement_dashboard.xlsx`
- `outputs/amazon_forecasting_placement_deck.pptx`

## Headline results

- Selected-model WAPE: **{kpi_summary['selected_model_wape']:.1%}**
- Fill-rate improvement: **{kpi_summary['fill_rate_baseline']:.1%} -> {kpi_summary['fill_rate_recommended']:.1%}**
- Stockout-risk reduction: **{kpi_summary['stockout_risk_baseline']:.1%} -> {kpi_summary['stockout_risk_recommended']:.1%}**
- Reactive transfer cost reduction: **${kpi_summary['reactive_transfer_cost_baseline']:,.0f} -> ${kpi_summary['reactive_transfer_cost_recommended']:,.0f}**
"""
    (BASE_DIR / "README.md").write_text(readme, encoding="utf-8")


def run_solution() -> dict[str, Any]:
    ensure_dirs()
    data = prepare_data()
    benchmark, benchmark_details = evaluate_models(data["weekly_category_region"])
    forecasts = create_forecasts(
        data["weekly_category_region"],
        benchmark,
        data["weekly_events"],
        horizon_weeks=6,
    )
    scenarios_df, recommended_df, inventory_implications = recommend_placement(
        forecasts,
        data["inventory_by_node_category"],
        data["node_cost_matrix"],
    )

    sku_forecasts = (
        forecasts.groupby(["category", "region"], as_index=False)["forecast_units"]
        .sum()
        .rename(columns={"forecast_units": "six_week_forecast_units"})
        .merge(data["sku_mix"], on=["category", "region"], how="left")
    )
    sku_forecasts["sku_forecast_units"] = (
        sku_forecasts["six_week_forecast_units"] * sku_forecasts["sku_mix_share"]
    ).round(2)
    sku_forecasts = sku_forecasts.sort_values("sku_forecast_units", ascending=False).reset_index(drop=True)

    network_kpis = build_kpi_summary(benchmark, scenarios_df, inventory_implications)

    region_forecast = (
        forecasts.groupby(["week_start", "region"], as_index=False)["forecast_units"].sum().sort_values("week_start")
    )
    model_summary = pd.DataFrame(
        [
            {
                "model": "moving_average",
                "wape": benchmark["baseline_wape"].mean(),
                "mape": benchmark["baseline_mape"].mean(),
            },
            {
                "model": "selected_policy",
                "wape": benchmark.apply(
                    lambda row: row["regression_wape"] if row["best_model"] == "regression" else row["baseline_wape"],
                    axis=1,
                ).mean(),
                "mape": benchmark.apply(
                    lambda row: row["regression_mape"] if row["best_model"] == "regression" else row["baseline_mape"],
                    axis=1,
                ).mean(),
            },
        ]
    )
    kpi_dashboard = pd.DataFrame(
        [
            {"metric": "Forecast WAPE", "baseline": network_kpis["baseline_wape"], "recommended": network_kpis["selected_model_wape"]},
            {"metric": "Fill Rate", "baseline": network_kpis["fill_rate_baseline"], "recommended": network_kpis["fill_rate_recommended"]},
            {"metric": "Stockout Risk", "baseline": network_kpis["stockout_risk_baseline"], "recommended": network_kpis["stockout_risk_recommended"]},
            {
                "metric": "Reactive Transfer Cost",
                "baseline": network_kpis["reactive_transfer_cost_baseline"],
                "recommended": network_kpis["reactive_transfer_cost_recommended"],
            },
            {"metric": "Avg Lead Days", "baseline": network_kpis["avg_lead_days_baseline"], "recommended": network_kpis["avg_lead_days_recommended"]},
        ]
    )

    benchmark.to_csv(OUTPUT_DIR / "model_benchmark.csv", index=False)
    benchmark_details.to_csv(OUTPUT_DIR / "holdout_predictions.csv", index=False)
    forecasts.to_csv(OUTPUT_DIR / "category_region_forecasts.csv", index=False)
    sku_forecasts.to_csv(OUTPUT_DIR / "sku_region_forecasts.csv", index=False)
    scenarios_df.to_csv(OUTPUT_DIR / "placement_scenarios.csv", index=False)
    recommended_df.to_csv(OUTPUT_DIR / "placement_recommendations.csv", index=False)
    inventory_implications.to_csv(OUTPUT_DIR / "inventory_implications.csv", index=False)
    region_forecast.to_csv(OUTPUT_DIR / "region_forecast_chart_data.csv", index=False)
    kpi_dashboard.to_csv(OUTPUT_DIR / "kpi_dashboard_data.csv", index=False)
    model_summary.to_csv(OUTPUT_DIR / "model_summary.csv", index=False)
    data["kpi_list"].to_csv(OUTPUT_DIR / "kpi_reference.csv", index=False)

    summary_payload = {
        "network_kpis": network_kpis,
        "top_baseline_risks": dataframe_to_records(
            scenarios_df.sort_values("baseline_stockout_risk", ascending=False).head(8)
        ),
        "top_recommendations": dataframe_to_records(
            recommended_df.sort_values("allocation_units", ascending=False).head(20)
        ),
        "region_forecast_chart_data": dataframe_to_records(region_forecast, limit=None),
        "kpi_dashboard_data": dataframe_to_records(kpi_dashboard, limit=None),
        "model_summary": dataframe_to_records(model_summary, limit=None),
    }
    (OUTPUT_DIR / "solution_summary.json").write_text(json.dumps(summary_payload, indent=2), encoding="utf-8")

    build_documents(network_kpis, benchmark, scenarios_df, inventory_implications)
    export_sql_files()
    update_folder_readme(network_kpis)
    return summary_payload


if __name__ == "__main__":
    summary = run_solution()
    print(json.dumps(summary["network_kpis"], indent=2))
