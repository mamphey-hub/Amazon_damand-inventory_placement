from __future__ import annotations

import argparse
import json
import shutil
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd


DEFAULT_INPUT = Path(r"C:\Users\Mamphey\Desktop\B_Dataset\merged_data.xlsx")
DEFAULT_OUTPUT_DIR = Path(r"C:\Users\Mamphey\Documents\Codex\2026-04-22-files-mentioned-by-the-user-merged\analysis_outputs")


def trim_object_columns(df: pd.DataFrame) -> pd.DataFrame:
    cleaned = df.copy()
    for column in cleaned.select_dtypes(include=["object", "string"]).columns:
        cleaned[column] = cleaned[column].astype(str).str.strip()
        cleaned.loc[cleaned[column].isin(["nan", "None"]), column] = pd.NA
    return cleaned


def parse_date_column(df: pd.DataFrame, column: str) -> pd.DataFrame:
    cleaned = df.copy()
    cleaned[column] = pd.to_datetime(cleaned[column], errors="coerce")
    return cleaned


def normalize_binary(series: pd.Series) -> pd.Series:
    numeric = pd.to_numeric(series, errors="coerce")
    mapped = numeric.fillna(0).clip(lower=0, upper=1).round().astype("Int64")
    return mapped.astype(int)


def warehouse_suffix_map(warehouses: pd.DataFrame) -> dict[str, str]:
    suffix = warehouses["warehouse_id"].astype(str).str.extract(r"(\d+)$")[0]
    return dict(zip(suffix, warehouses["warehouse_id"].astype(str)))


def canonicalize_warehouse_ids(df: pd.DataFrame, column: str, suffix_to_id: dict[str, str]) -> tuple[pd.DataFrame, int]:
    cleaned = df.copy()
    raw_series = cleaned[column].astype(str).str.strip()
    suffix = raw_series.str.extract(r"(\d+)$")[0]
    canonical = suffix.map(suffix_to_id)
    changed_mask = canonical.notna() & canonical.ne(raw_series)
    insert_at = cleaned.columns.get_loc(column) + 1
    cleaned.insert(insert_at, f"{column}_raw", raw_series)
    cleaned[column] = canonical.fillna(raw_series)
    return cleaned, int(changed_mask.sum())


def serialise_value(value: Any) -> Any:
    if pd.isna(value):
        return None
    if isinstance(value, (pd.Timestamp, np.datetime64)):
        return pd.Timestamp(value).strftime("%Y-%m-%d")
    if isinstance(value, (np.integer,)):
        return int(value)
    if isinstance(value, (np.floating,)):
        return float(value)
    return value


def dataframe_to_records(df: pd.DataFrame, limit: int = 10) -> list[dict[str, Any]]:
    if df.empty:
        return []
    return [
        {column: serialise_value(value) for column, value in row.items()}
        for row in df.head(limit).to_dict(orient="records")
    ]


def summarize_columns(df: pd.DataFrame) -> list[dict[str, Any]]:
    profiles: list[dict[str, Any]] = []
    for column in df.columns:
        series = df[column]
        profile: dict[str, Any] = {
            "column": column,
            "dtype": str(series.dtype),
            "non_null_count": int(series.notna().sum()),
            "null_count": int(series.isna().sum()),
            "null_pct": round(float(series.isna().mean() * 100), 2),
            "unique_count": int(series.nunique(dropna=True)),
        }
        if pd.api.types.is_numeric_dtype(series):
            profile.update(
                {
                    "min": serialise_value(series.min()),
                    "p25": serialise_value(series.quantile(0.25)),
                    "median": serialise_value(series.median()),
                    "mean": serialise_value(series.mean()),
                    "p75": serialise_value(series.quantile(0.75)),
                    "max": serialise_value(series.max()),
                    "std": serialise_value(series.std()),
                }
            )
        elif pd.api.types.is_datetime64_any_dtype(series):
            profile.update(
                {
                    "min": serialise_value(series.min()),
                    "max": serialise_value(series.max()),
                }
            )
        else:
            top_values = (
                series.astype(str)
                .value_counts(dropna=False)
                .head(5)
                .rename_axis("value")
                .reset_index(name="count")
            )
            profile["top_values"] = dataframe_to_records(top_values, limit=5)
        profiles.append(profile)
    return profiles


def make_sheet_profile(df: pd.DataFrame) -> dict[str, Any]:
    return {
        "rows": int(len(df)),
        "columns": int(df.shape[1]),
        "duplicate_rows": int(df.duplicated().sum()),
        "total_missing_cells": int(df.isna().sum().sum()),
        "columns_profile": summarize_columns(df),
        "sample_rows": dataframe_to_records(df, limit=5),
    }


def df_to_html_table(df: pd.DataFrame, index: bool = False) -> str:
    safe = df.copy()
    for column in safe.columns:
        if pd.api.types.is_datetime64_any_dtype(safe[column]):
            safe[column] = safe[column].dt.strftime("%Y-%m-%d")
    return safe.to_html(index=index, border=0, classes="table table-sm")


def df_to_markdown_table(df: pd.DataFrame) -> str:
    safe = df.copy()
    for column in safe.columns:
        if pd.api.types.is_datetime64_any_dtype(safe[column]):
            safe[column] = safe[column].dt.strftime("%Y-%m-%d")

    headers = [str(column) for column in safe.columns]
    header_row = "| " + " | ".join(headers) + " |"
    separator_row = "| " + " | ".join(["---"] * len(headers)) + " |"
    body_rows = []
    for _, row in safe.iterrows():
        values = ["" if pd.isna(value) else str(value) for value in row.tolist()]
        body_rows.append("| " + " | ".join(values) + " |")
    return "\n".join([header_row, separator_row, *body_rows])


def build_html_report(
    quality_summary: pd.DataFrame,
    overview: dict[str, Any],
    region_summary: pd.DataFrame,
    category_summary: pd.DataFrame,
    monthly_summary: pd.DataFrame,
    flag_summary: pd.DataFrame,
    profile_summary: pd.DataFrame,
    sheet_profiles: dict[str, dict[str, Any]],
    download_links: list[dict[str, str]],
) -> str:
    download_items = "".join(
        [
            (
                f'<li><a href="{item["href"]}" download>{item["label"]}</a>'
                f' <span class="muted">({item["href"]})</span></li>'
            )
            for item in download_links
        ]
    )
    sections: list[str] = []
    sections.append(
        f"""
        <h1>Merged Data Cleaning, EDA, and Profiling</h1>
        <p>This report summarizes the cleaning actions applied to the workbook and the main exploratory findings from the enriched daily demand dataset.</p>
        <section class="downloads">
          <h2>Downloads</h2>
          <p>Use these links to download the cleaned workbook, reports, source-ready CSV exports, or the full ZIP bundle.</p>
          <ul>
            {download_items}
          </ul>
        </section>
        <h2>Key Metrics</h2>
        <table class="table">
            <tr><th>Metric</th><th>Value</th></tr>
            <tr><td>Daily demand rows</td><td>{overview['daily_rows']:,}</td></tr>
            <tr><td>Unique SKUs</td><td>{overview['unique_skus']:,}</td></tr>
            <tr><td>Unique regions</td><td>{overview['unique_regions']}</td></tr>
            <tr><td>Date range</td><td>{overview['date_range']}</td></tr>
            <tr><td>Total units ordered</td><td>{overview['total_units']:,}</td></tr>
            <tr><td>Total revenue (USD)</td><td>{overview['total_revenue']:,.2f}</td></tr>
            <tr><td>Average unit price (USD)</td><td>{overview['avg_price']:,.2f}</td></tr>
            <tr><td>Weather / units correlation</td><td>{overview['weather_correlation']:.3f}</td></tr>
        </table>
        <h2>Cleaning Summary</h2>
        {df_to_html_table(quality_summary)}
        <h2>Workbook Profile Summary</h2>
        {df_to_html_table(profile_summary)}
        <h2>Demand by Region</h2>
        {df_to_html_table(region_summary)}
        <h2>Demand by Category</h2>
        {df_to_html_table(category_summary)}
        <h2>Monthly Trend</h2>
        {df_to_html_table(monthly_summary)}
        <h2>Event and Calendar Effect</h2>
        {df_to_html_table(flag_summary)}
        """
    )

    for sheet_name, profile in sheet_profiles.items():
        column_df = pd.DataFrame(profile["columns_profile"])
        sample_df = pd.DataFrame(profile["sample_rows"])
        sections.append(
            f"""
            <details>
                <summary><strong>{sheet_name}</strong> profile</summary>
                <p>Rows: {profile['rows']:,} | Columns: {profile['columns']} | Duplicate rows: {profile['duplicate_rows']} | Missing cells: {profile['total_missing_cells']}</p>
                <h3>Columns</h3>
                {df_to_html_table(column_df)}
                <h3>Sample Rows</h3>
                {df_to_html_table(sample_df)}
            </details>
            """
        )

    return f"""
    <html>
      <head>
        <meta charset="utf-8" />
        <title>Merged Data Report</title>
        <style>
          body {{
            font-family: Segoe UI, Arial, sans-serif;
            margin: 24px;
            color: #1f2937;
            background: #f8fafc;
          }}
          h1, h2, h3 {{
            color: #0f172a;
          }}
          .table {{
            border-collapse: collapse;
            width: 100%;
            margin-bottom: 24px;
            background: white;
          }}
          .table th, .table td {{
            border: 1px solid #cbd5e1;
            padding: 8px 10px;
            text-align: left;
            vertical-align: top;
          }}
          .table th {{
            background: #e2e8f0;
          }}
          details {{
            margin: 18px 0;
            padding: 12px;
            background: white;
            border: 1px solid #cbd5e1;
          }}
          summary {{
            cursor: pointer;
            font-size: 16px;
          }}
          .downloads {{
            margin: 18px 0 24px;
            padding: 16px;
            background: white;
            border: 1px solid #cbd5e1;
          }}
          .downloads ul {{
            margin: 0;
            padding-left: 20px;
          }}
          .downloads li {{
            margin: 8px 0;
          }}
          .muted {{
            color: #475569;
            font-size: 13px;
          }}
        </style>
      </head>
      <body>
        {''.join(sections)}
      </body>
    </html>
    """


def main() -> None:
    parser = argparse.ArgumentParser(description="Clean, profile, and summarize merged_data.xlsx")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    args = parser.parse_args()

    input_path = args.input
    output_dir = args.output_dir
    cleaned_dir = output_dir / "cleaned_data"
    output_dir.mkdir(parents=True, exist_ok=True)
    cleaned_dir.mkdir(parents=True, exist_ok=True)

    workbook = pd.ExcelFile(input_path)
    raw_sheets = {sheet_name: pd.read_excel(workbook, sheet_name=sheet_name) for sheet_name in workbook.sheet_names}

    cleaned_sheets: dict[str, pd.DataFrame] = {}
    cleaning_actions: list[dict[str, Any]] = []

    warehouses = trim_object_columns(raw_sheets["Warehouses"])
    warehouses = warehouses.drop_duplicates().reset_index(drop=True)
    cleaned_sheets["Warehouses"] = warehouses
    suffix_to_id = warehouse_suffix_map(warehouses)

    sku_master = trim_object_columns(raw_sheets["SKU Master"]).drop_duplicates().reset_index(drop=True)
    cleaned_sheets["SKU Master"] = sku_master

    daily = trim_object_columns(raw_sheets["Daily Demand"])
    daily = parse_date_column(daily, "date")
    daily["weekend_flag_raw"] = daily["weekend_flag"].astype(str)
    daily["weekend_flag"] = (daily["date"].dt.dayofweek >= 5).astype(int)
    for flag_col in ["holiday_peak_flag", "prime_event_flag", "marketing_push_flag"]:
        daily[flag_col] = normalize_binary(daily[flag_col])
    daily = daily.drop_duplicates().sort_values(["date", "sku_id", "region"]).reset_index(drop=True)
    cleaned_sheets["Daily Demand"] = daily
    cleaning_actions.append(
        {
            "sheet": "Daily Demand",
            "issue_fixed": "Parsed dates and replaced corrupted weekend_flag values using the calendar date",
            "before_rows": int(len(raw_sheets["Daily Demand"])),
            "after_rows": int(len(daily)),
        }
    )

    event_calendar = trim_object_columns(raw_sheets["Event Calendar"])
    event_calendar = parse_date_column(event_calendar, "date")
    event_calendar["weekend_flag_raw"] = event_calendar["weekend_flag"].astype(str)
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
            source_rows_collapsed=("date", "size"),
        )
        .sort_values("date")
        .reset_index(drop=True)
    )
    event_calendar["weekend_flag"] = (event_calendar["date"].dt.dayofweek >= 5).astype(int)
    cleaned_sheets["Event Calendar"] = event_calendar
    cleaning_actions.append(
        {
            "sheet": "Event Calendar",
            "issue_fixed": "Collapsed duplicate dates to one calendar row per day, inferred weekend_flag, and averaged weather_disruption_index",
            "before_rows": int(len(raw_sheets["Event Calendar"])),
            "after_rows": int(len(event_calendar)),
        }
    )

    starting_inventory = trim_object_columns(raw_sheets["Starting Inventory"])
    starting_inventory, inventory_id_changes = canonicalize_warehouse_ids(
        starting_inventory, "warehouse_id", suffix_to_id
    )
    starting_inventory = starting_inventory.drop_duplicates().reset_index(drop=True)
    cleaned_sheets["Starting Inventory"] = starting_inventory
    cleaning_actions.append(
        {
            "sheet": "Starting Inventory",
            "issue_fixed": "Canonicalized warehouse_id values to the IDs defined in the Warehouses sheet",
            "before_rows": int(len(raw_sheets["Starting Inventory"])),
            "after_rows": int(len(starting_inventory)),
            "id_updates": inventory_id_changes,
        }
    )

    warehouse_region_costs = trim_object_columns(raw_sheets["Warehouse Region Costs"])
    warehouse_region_costs, cost_id_changes = canonicalize_warehouse_ids(
        warehouse_region_costs, "warehouse_id", suffix_to_id
    )
    warehouse_region_costs = warehouse_region_costs.drop_duplicates().reset_index(drop=True)
    cleaned_sheets["Warehouse Region Costs"] = warehouse_region_costs
    cleaning_actions.append(
        {
            "sheet": "Warehouse Region Costs",
            "issue_fixed": "Canonicalized warehouse_id values to the IDs defined in the Warehouses sheet",
            "before_rows": int(len(raw_sheets["Warehouse Region Costs"])),
            "after_rows": int(len(warehouse_region_costs)),
            "id_updates": cost_id_changes,
        }
    )

    weekly = trim_object_columns(raw_sheets["Weekly Demand"])
    weekly = parse_date_column(weekly, "week")
    weekly["week_raw"] = weekly["week"]
    weekly["week"] = weekly["week"] - pd.to_timedelta(weekly["week"].dt.weekday, unit="D")
    weekly = weekly.drop_duplicates().sort_values(["week", "sku_id", "region"]).reset_index(drop=True)
    cleaned_sheets["Weekly Demand"] = weekly
    cleaning_actions.append(
        {
            "sheet": "Weekly Demand",
            "issue_fixed": "Normalized week values to the Monday week start",
            "before_rows": int(len(raw_sheets["Weekly Demand"])),
            "after_rows": int(len(weekly)),
        }
    )

    source_index = raw_sheets["Index"].copy()
    cleaned_index = source_index[["Sheet", "Source File"]].copy()
    cleaned_index["Rows"] = cleaned_index["Sheet"].map({sheet: len(df) for sheet, df in cleaned_sheets.items()})
    cleaned_index["Columns"] = cleaned_index["Sheet"].map({sheet: df.shape[1] for sheet, df in cleaned_sheets.items()})
    note_lookup = {row["sheet"]: row["issue_fixed"] for row in cleaning_actions}
    cleaned_index["Cleaning Notes"] = cleaned_index["Sheet"].map(note_lookup).fillna("Trimmed text and preserved schema")
    cleaned_sheets["Index"] = cleaned_index

    daily_enriched = (
        cleaned_sheets["Daily Demand"]
        .merge(cleaned_sheets["SKU Master"], on="sku_id", how="left")
        .merge(
            cleaned_sheets["Event Calendar"][
                [
                    "date",
                    "holiday_peak_flag",
                    "prime_event_flag",
                    "marketing_push_flag",
                    "weather_disruption_index",
                    "weekend_flag",
                ]
            ],
            on="date",
            how="left",
            suffixes=("", "_calendar"),
        )
    )
    daily_enriched["revenue_usd"] = daily_enriched["units_ordered"] * daily_enriched["price_usd"]
    daily_enriched["gross_margin_usd"] = daily_enriched["units_ordered"] * (
        daily_enriched["price_usd"] - daily_enriched["unit_cost_usd"]
    )
    daily_enriched["month"] = daily_enriched["date"].dt.to_period("M").astype(str)

    daily_date_range = f"{daily_enriched['date'].min():%Y-%m-%d} to {daily_enriched['date'].max():%Y-%m-%d}"
    overview = {
        "daily_rows": int(len(daily_enriched)),
        "unique_skus": int(daily_enriched["sku_id"].nunique()),
        "unique_regions": int(daily_enriched["region"].nunique()),
        "date_range": daily_date_range,
        "total_units": int(daily_enriched["units_ordered"].sum()),
        "total_revenue": float(daily_enriched["revenue_usd"].sum()),
        "avg_price": float(daily_enriched["price_usd"].mean()),
        "weather_correlation": float(
            daily_enriched[["units_ordered", "weather_disruption_index"]].corr(numeric_only=True).iloc[0, 1]
        ),
    }

    region_summary = (
        daily_enriched.groupby("region", as_index=False)
        .agg(
            total_units=("units_ordered", "sum"),
            avg_units=("units_ordered", "mean"),
            total_revenue=("revenue_usd", "sum"),
            avg_price=("price_usd", "mean"),
        )
        .sort_values("total_units", ascending=False)
        .reset_index(drop=True)
    )

    category_summary = (
        daily_enriched.groupby("category", as_index=False)
        .agg(
            total_units=("units_ordered", "sum"),
            total_revenue=("revenue_usd", "sum"),
            gross_margin_usd=("gross_margin_usd", "sum"),
            avg_service_level=("target_service_level", "mean"),
        )
        .sort_values("total_units", ascending=False)
        .reset_index(drop=True)
    )

    monthly_summary = (
        daily_enriched.groupby("month", as_index=False)
        .agg(
            total_units=("units_ordered", "sum"),
            total_revenue=("revenue_usd", "sum"),
            avg_weather_disruption=("weather_disruption_index", "mean"),
        )
        .sort_values("month")
        .reset_index(drop=True)
    )

    flag_frames = []
    for flag_column in ["holiday_peak_flag", "prime_event_flag", "marketing_push_flag", "weekend_flag"]:
        grouped = (
            daily_enriched.groupby(flag_column, as_index=False)
            .agg(
                avg_units=("units_ordered", "mean"),
                median_units=("units_ordered", "median"),
                avg_revenue=("revenue_usd", "mean"),
                row_count=("sku_id", "size"),
            )
            .rename(columns={flag_column: "flag_value"})
        )
        grouped.insert(0, "flag_name", flag_column)
        flag_frames.append(grouped)
    flag_summary = pd.concat(flag_frames, ignore_index=True)

    warehouse_match_before_inventory = raw_sheets["Starting Inventory"]["warehouse_id"].isin(
        cleaned_sheets["Warehouses"]["warehouse_id"]
    )
    warehouse_match_after_inventory = cleaned_sheets["Starting Inventory"]["warehouse_id"].isin(
        cleaned_sheets["Warehouses"]["warehouse_id"]
    )
    warehouse_match_before_costs = raw_sheets["Warehouse Region Costs"]["warehouse_id"].isin(
        cleaned_sheets["Warehouses"]["warehouse_id"]
    )
    warehouse_match_after_costs = cleaned_sheets["Warehouse Region Costs"]["warehouse_id"].isin(
        cleaned_sheets["Warehouses"]["warehouse_id"]
    )

    quality_summary = pd.DataFrame(
        [
            {
                "check": "Daily Demand weekend_flag distinct values",
                "before": int(raw_sheets["Daily Demand"]["weekend_flag"].astype(str).nunique()),
                "after": int(cleaned_sheets["Daily Demand"]["weekend_flag"].nunique()),
                "status": "fixed",
            },
            {
                "check": "Event Calendar rows",
                "before": int(len(raw_sheets["Event Calendar"])),
                "after": int(len(cleaned_sheets["Event Calendar"])),
                "status": "deduplicated",
            },
            {
                "check": "Starting Inventory warehouse match rate",
                "before": round(float(warehouse_match_before_inventory.mean()), 4),
                "after": round(float(warehouse_match_after_inventory.mean()), 4),
                "status": "fixed",
            },
            {
                "check": "Warehouse Region Costs warehouse match rate",
                "before": round(float(warehouse_match_before_costs.mean()), 4),
                "after": round(float(warehouse_match_after_costs.mean()), 4),
                "status": "fixed",
            },
            {
                "check": "Weekly Demand distinct week values",
                "before": int(pd.to_datetime(raw_sheets["Weekly Demand"]["week"], errors="coerce").nunique()),
                "after": int(cleaned_sheets["Weekly Demand"]["week"].nunique()),
                "status": "normalized",
            },
        ]
    )

    sheet_profiles = {sheet_name: make_sheet_profile(df) for sheet_name, df in cleaned_sheets.items()}
    profile_summary = pd.DataFrame(
        [
            {
                "sheet": sheet_name,
                "rows": profile["rows"],
                "columns": profile["columns"],
                "duplicate_rows": profile["duplicate_rows"],
                "missing_cells": profile["total_missing_cells"],
            }
            for sheet_name, profile in sheet_profiles.items()
        ]
    ).sort_values("sheet")

    cleaning_log = pd.DataFrame(cleaning_actions)

    for sheet_name, df in cleaned_sheets.items():
        export_df = df.copy()
        for column in export_df.columns:
            if pd.api.types.is_datetime64_any_dtype(export_df[column]):
                export_df[column] = export_df[column].dt.strftime("%Y-%m-%d")
        export_df.to_csv(cleaned_dir / f"{sheet_name.lower().replace(' ', '_')}.csv", index=False)

    daily_enriched_export = daily_enriched.copy()
    for column in daily_enriched_export.columns:
        if pd.api.types.is_datetime64_any_dtype(daily_enriched_export[column]):
            daily_enriched_export[column] = daily_enriched_export[column].dt.strftime("%Y-%m-%d")
    daily_enriched_export.to_csv(cleaned_dir / "daily_demand_enriched.csv", index=False)

    with pd.ExcelWriter(output_dir / "merged_data_cleaned.xlsx", engine="openpyxl") as writer:
        cleaned_sheets["Index"].to_excel(writer, sheet_name="Index", index=False)
        cleaned_sheets["Daily Demand"].to_excel(writer, sheet_name="Daily Demand", index=False)
        cleaned_sheets["Event Calendar"].to_excel(writer, sheet_name="Event Calendar", index=False)
        cleaned_sheets["SKU Master"].to_excel(writer, sheet_name="SKU Master", index=False)
        cleaned_sheets["Starting Inventory"].to_excel(writer, sheet_name="Starting Inventory", index=False)
        cleaned_sheets["Warehouse Region Costs"].to_excel(writer, sheet_name="Warehouse Region Costs", index=False)
        cleaned_sheets["Warehouses"].to_excel(writer, sheet_name="Warehouses", index=False)
        cleaned_sheets["Weekly Demand"].to_excel(writer, sheet_name="Weekly Demand", index=False)

    (output_dir / "cleaning_log.csv").write_text(cleaning_log.to_csv(index=False), encoding="utf-8")
    (output_dir / "quality_summary.csv").write_text(quality_summary.to_csv(index=False), encoding="utf-8")
    (output_dir / "region_summary.csv").write_text(region_summary.to_csv(index=False), encoding="utf-8")
    (output_dir / "category_summary.csv").write_text(category_summary.to_csv(index=False), encoding="utf-8")
    (output_dir / "monthly_summary.csv").write_text(monthly_summary.to_csv(index=False), encoding="utf-8")
    (output_dir / "flag_summary.csv").write_text(flag_summary.to_csv(index=False), encoding="utf-8")

    profile_json = {
        "overview": overview,
        "quality_summary": dataframe_to_records(quality_summary, limit=len(quality_summary)),
        "sheet_profiles": sheet_profiles,
    }
    (output_dir / "data_profile.json").write_text(json.dumps(profile_json, indent=2), encoding="utf-8")

    markdown_lines = [
        "# Merged Data Cleaning, EDA, and Profiling",
        "",
        "## Key Findings",
        f"- Daily demand covers {overview['date_range']} with {overview['daily_rows']:,} rows across {overview['unique_skus']:,} SKUs.",
        f"- Total units ordered: {overview['total_units']:,}.",
        f"- Total revenue: ${overview['total_revenue']:,.2f}.",
        f"- Average selling price in Daily Demand: ${overview['avg_price']:,.2f}.",
        f"- Weather disruption correlation with units ordered: {overview['weather_correlation']:.3f}.",
        "",
        "## Cleaning Actions",
    ]
    for action in cleaning_actions:
        update_text = f" ({action['id_updates']} IDs updated)" if "id_updates" in action else ""
        markdown_lines.append(
            f"- {action['sheet']}: {action['issue_fixed']} [rows: {action['before_rows']} -> {action['after_rows']}]"
            f"{update_text}."
        )

    markdown_lines.extend(
        [
            "",
            "## Top Regions by Units",
            df_to_markdown_table(region_summary),
            "",
            "## Top Categories by Units",
            df_to_markdown_table(category_summary),
            "",
            "## Monthly Trend",
            df_to_markdown_table(monthly_summary),
            "",
            "## Event Effects",
            df_to_markdown_table(flag_summary),
            "",
        ]
    )
    (output_dir / "eda_summary.md").write_text("\n".join(markdown_lines), encoding="utf-8")

    bundle_path = output_dir.parent / "analysis_outputs_bundle.zip"
    if bundle_path.exists():
        bundle_path.unlink()
    shutil.make_archive(str(bundle_path.with_suffix("")), "zip", root_dir=output_dir)

    download_links = [
        {"label": "Full analysis ZIP bundle", "href": "../analysis_outputs_bundle.zip"},
        {"label": "Cleaned workbook", "href": "merged_data_cleaned.xlsx"},
        {"label": "HTML profile report", "href": "data_profile.html"},
        {"label": "JSON data profile", "href": "data_profile.json"},
        {"label": "EDA markdown summary", "href": "eda_summary.md"},
        {"label": "Cleaning log CSV", "href": "cleaning_log.csv"},
        {"label": "Quality summary CSV", "href": "quality_summary.csv"},
        {"label": "Region summary CSV", "href": "region_summary.csv"},
        {"label": "Category summary CSV", "href": "category_summary.csv"},
        {"label": "Monthly summary CSV", "href": "monthly_summary.csv"},
        {"label": "Flag summary CSV", "href": "flag_summary.csv"},
    ]
    cleaned_csv_dir = output_dir / "cleaned_data"
    for csv_path in sorted(cleaned_csv_dir.glob("*.csv")):
        download_links.append(
            {"label": f"Cleaned CSV: {csv_path.name}", "href": f"cleaned_data/{csv_path.name}"}
        )

    html_report = build_html_report(
        quality_summary=quality_summary,
        overview=overview,
        region_summary=region_summary,
        category_summary=category_summary,
        monthly_summary=monthly_summary,
        flag_summary=flag_summary,
        profile_summary=profile_summary,
        sheet_profiles=sheet_profiles,
        download_links=download_links,
    )
    (output_dir / "data_profile.html").write_text(html_report, encoding="utf-8")

    print(json.dumps({"output_dir": str(output_dir), "overview": overview}, indent=2))


if __name__ == "__main__":
    main()
