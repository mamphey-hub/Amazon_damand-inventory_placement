import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const BASE_DIR =
  "C:/Users/Mamphey/Documents/Codex/2026-04-22-files-mentioned-by-the-user-merged/forecasting-and-placement solution";
const OUTPUT_DIR = path.join(BASE_DIR, "outputs");

function parseCsv(text) {
  const rows = [];
  let row = [];
  let value = "";
  let inQuotes = false;
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    const next = text[i + 1];
    if (ch === '"') {
      if (inQuotes && next === '"') {
        value += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === "," && !inQuotes) {
      row.push(value);
      value = "";
    } else if ((ch === "\n" || ch === "\r") && !inQuotes) {
      if (ch === "\r" && next === "\n") i += 1;
      row.push(value);
      if (row.length > 1 || row[0] !== "") rows.push(row);
      row = [];
      value = "";
    } else {
      value += ch;
    }
  }
  if (value.length > 0 || row.length > 0) {
    row.push(value);
    rows.push(row);
  }
  const [header, ...body] = rows;
  return body.map((line) =>
    Object.fromEntries(
      header.map((key, index) => {
        const raw = line[index] ?? "";
        const num = Number(raw);
        if (raw !== "" && !Number.isNaN(num)) return [key, num];
        return [key, raw];
      }),
    ),
  );
}

function percent(value) {
  return Math.round(value * 1000) / 10;
}

function formatHeader(range) {
  range.format = {
    fill: "#0F172A",
    font: { bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
  };
}

function formatSubheader(range) {
  range.format = {
    fill: "#E2E8F0",
    font: { bold: true, color: "#0F172A" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
  };
}

function setTitle(sheet, title, subtitle) {
  sheet.getRange("A1:H1").merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A1").format = {
    fill: "#0B1F33",
    font: { bold: true, color: "#FFFFFF", size: 18 },
    horizontalAlignment: "left",
    verticalAlignment: "center",
  };
  sheet.getRange("A2:H2").merge();
  sheet.getRange("A2").values = [[subtitle]];
  sheet.getRange("A2").format = {
    fill: "#DCE8F5",
    font: { color: "#1E293B", italic: true },
    horizontalAlignment: "left",
    verticalAlignment: "center",
  };
}

async function main() {
  const summary = JSON.parse(await fs.readFile(path.join(OUTPUT_DIR, "solution_summary.json"), "utf8"));
  const benchmark = parseCsv(await fs.readFile(path.join(OUTPUT_DIR, "model_benchmark.csv"), "utf8"));
  const forecasts = parseCsv(await fs.readFile(path.join(OUTPUT_DIR, "category_region_forecasts.csv"), "utf8"));
  const scenarios = parseCsv(await fs.readFile(path.join(OUTPUT_DIR, "placement_scenarios.csv"), "utf8"));
  const inventory = parseCsv(await fs.readFile(path.join(OUTPUT_DIR, "inventory_implications.csv"), "utf8"));
  const kpis = parseCsv(await fs.readFile(path.join(OUTPUT_DIR, "kpi_reference.csv"), "utf8"));

  const workbook = Workbook.create();
  const dashboard = workbook.worksheets.add("Dashboard");
  const benchmarkSheet = workbook.worksheets.add("Benchmark");
  const forecastSheet = workbook.worksheets.add("Forecasts");
  const placementSheet = workbook.worksheets.add("Placement");
  const inventorySheet = workbook.worksheets.add("Inventory");
  const kpiSheet = workbook.worksheets.add("KPI Glossary");
  const assumptionSheet = workbook.worksheets.add("Assumptions");

  dashboard.showGridLines = false;
  setTitle(
    dashboard,
    "Regional Forecasting and Placement Dashboard",
    "Category-region forecast benchmark, network service KPIs, and recommended node placements.",
  );

  dashboard.getRange("A4:B8").values = [
    ["Metric", "Value"],
    ["Selected-model WAPE", summary.network_kpis.selected_model_wape],
    ["On-time fill rate", summary.network_kpis.fill_rate_recommended],
    ["Stockout risk", summary.network_kpis.stockout_risk_recommended],
    ["Avg lead days", summary.network_kpis.avg_lead_days_recommended],
  ];
  formatHeader(dashboard.getRange("A4:B4"));
  dashboard.getRange("A5:A8").format = { font: { bold: true } };
  dashboard.getRange("B5:B7").format.numberFormat = "0.0%";
  dashboard.getRange("B8").format.numberFormat = "0.00";

  dashboard.getRange("D4:E8").values = [
    ["Metric", "Value"],
    ["Baseline fill rate", summary.network_kpis.fill_rate_baseline],
    ["Baseline WAPE", summary.network_kpis.baseline_wape],
    ["Transfer cost baseline", summary.network_kpis.reactive_transfer_cost_baseline],
    ["Placement adherence", summary.network_kpis.placement_adherence_baseline],
  ];
  formatHeader(dashboard.getRange("D4:E4"));
  dashboard.getRange("D5:D8").format = { font: { bold: true } };
  dashboard.getRange("E5:E6").format.numberFormat = "0.0%";
  dashboard.getRange("E7").format.numberFormat = "$#,##0";
  dashboard.getRange("E8").format.numberFormat = "0.0%";

  dashboard.getRange("A11:F17").values = [
    ["Category", "Demand Region", "Primary Node", "Backup Node", "Fill Rate", "Lead Days"],
    ...scenarios
      .sort((a, b) => b.recommended_fill_rate - a.recommended_fill_rate)
      .slice(0, 6)
      .map((row) => [
        row.category,
        row.demand_region,
        row.primary_node,
        row.backup_node,
        row.recommended_fill_rate,
        row.recommended_avg_lead_days,
      ]),
  ];
  formatHeader(dashboard.getRange("A11:F11"));
  dashboard.getRange("E12:E17").format.numberFormat = "0.0%";
  dashboard.getRange("F12:F17").format.numberFormat = "0.00";

  dashboard.getRange("J2:N8").values = [
    ["Week", "East", "North", "South", "West"],
    ...(() => {
      const grouped = new Map();
      for (const row of summary.region_forecast_chart_data) {
        const existing = grouped.get(row.week_start) ?? { Week: row.week_start };
        existing[row.region] = row.forecast_units;
        grouped.set(row.week_start, existing);
      }
      return [...grouped.values()].map((row) => [
        row.Week,
        row.East ?? 0,
        row.North ?? 0,
        row.South ?? 0,
        row.West ?? 0,
      ]);
    })(),
  ];
  const forecastChart = dashboard.charts.add("line", dashboard.getRange("J2:N8"));
  forecastChart.setPosition("H2", "P18");
  forecastChart.title = "Six-Week Regional Demand Forecast";
  forecastChart.hasLegend = true;
  forecastChart.xAxis = { axisType: "textAxis" };
  forecastChart.yAxis = { numberFormatCode: "0" };

  dashboard.getRange("J20:L25").values = [
    ["KPI", "Baseline", "Recommended"],
    ...summary.kpi_dashboard_data.map((row) => [row.metric, row.baseline, row.recommended]),
  ];
  const kpiChart = dashboard.charts.add("bar", dashboard.getRange("J20:L25"));
  kpiChart.setPosition("H20", "P36");
  kpiChart.title = "Baseline vs Recommended KPI Comparison";
  kpiChart.hasLegend = true;
  kpiChart.barOptions.direction = "column";
  kpiChart.barOptions.grouping = "clustered";
  kpiChart.xAxis = { axisType: "textAxis" };

  setTitle(
    benchmarkSheet,
    "Model Benchmark",
    "Weekly category-region holdout comparison between moving average and event-aware regression.",
  );
  const benchmarkMatrix = [
    [
      "Category",
      "Region",
      "Best Model",
      "Baseline WAPE",
      "Regression WAPE",
      "Baseline MAPE",
      "Regression MAPE",
      "Holdout Actual",
    ],
    ...benchmark.map((row) => [
      row.category,
      row.region,
      row.best_model,
      row.baseline_wape,
      row.regression_wape,
      row.baseline_mape,
      row.regression_mape,
      row.holdout_actual,
    ]),
  ];
  benchmarkSheet.getRange(`A4:H${benchmarkMatrix.length + 3}`).values = benchmarkMatrix;
  formatHeader(benchmarkSheet.getRange("A4:H4"));
  benchmarkSheet.getRange(`D5:G${benchmarkMatrix.length + 3}`).format.numberFormat = "0.0%";
  benchmarkSheet.getRange("A4:H28").format.wrapText = true;
  benchmarkSheet.freezePanes.freezeRows(4);
  benchmarkSheet.tables.add(`A4:H${benchmarkMatrix.length + 3}`, true, "BenchmarkTable");

  setTitle(
    forecastSheet,
    "Forecast Output",
    "Six-week forward demand projections at category-region level with selected model tags.",
  );
  const forecastMatrix = [
    ["Category", "Region", "Week Start", "Horizon Week", "Selected Model", "Forecast Units", "Sigma Weekly", "Service Level"],
    ...forecasts.map((row) => [
      row.category,
      row.region,
      row.week_start,
      row.horizon_week,
      row.selected_model,
      row.forecast_units,
      row.sigma_weekly,
      row.service_level,
    ]),
  ];
  forecastSheet.getRange(`A4:H${forecastMatrix.length + 3}`).values = forecastMatrix;
  formatHeader(forecastSheet.getRange("A4:H4"));
  forecastSheet.getRange(`H5:H${forecastMatrix.length + 3}`).format.numberFormat = "0.0%";
  forecastSheet.freezePanes.freezeRows(4);
  forecastSheet.tables.add(`A4:H${forecastMatrix.length + 3}`, true, "ForecastTable");

  setTitle(
    placementSheet,
    "Placement Scenarios",
    "Primary and backup node recommendations with fill-rate, stockout, cost, and lead-time implications.",
  );
  const placementMatrix = [
    [
      "Category",
      "Demand Region",
      "Primary Node",
      "Backup Node",
      "Target Inventory Units",
      "Baseline Fill Rate",
      "Recommended Fill Rate",
      "Baseline Stockout Risk",
      "Recommended Stockout Risk",
      "Baseline Transfer Cost",
      "Recommended Transfer Cost",
      "Baseline Lead Days",
      "Recommended Lead Days",
    ],
    ...scenarios.map((row) => [
      row.category,
      row.demand_region,
      row.primary_node,
      row.backup_node,
      row.target_inventory_units,
      row.baseline_fill_rate,
      row.recommended_fill_rate,
      row.baseline_stockout_risk,
      row.recommended_stockout_risk,
      row.baseline_transfer_cost,
      row.recommended_transfer_cost,
      row.baseline_avg_lead_days,
      row.recommended_avg_lead_days,
    ]),
  ];
  placementSheet.getRange(`A4:M${placementMatrix.length + 3}`).values = placementMatrix;
  formatHeader(placementSheet.getRange("A4:M4"));
  placementSheet.getRange(`F5:I${placementMatrix.length + 3}`).format.numberFormat = "0.0%";
  placementSheet.getRange(`J5:K${placementMatrix.length + 3}`).format.numberFormat = "$#,##0";
  placementSheet.getRange(`L5:M${placementMatrix.length + 3}`).format.numberFormat = "0.00";
  placementSheet.getRange(`F5:G${placementMatrix.length + 3}`).conditionalFormats.addColorScale({
    colors: ["#FCA5A5", "#FDE68A", "#86EFAC"],
  });
  placementSheet.freezePanes.freezeRows(4);
  placementSheet.tables.add(`A4:M${placementMatrix.length + 3}`, true, "PlacementTable");

  setTitle(
    inventorySheet,
    "Inventory Implications",
    "Current node inventory versus recommended active inventory footprint.",
  );
  const inventoryRows = inventory
    .sort((a, b) => b.inventory_surplus_vs_target - a.inventory_surplus_vs_target)
    .slice(0, 30);
  const inventoryMatrix = [
    [
      "Warehouse Region",
      "Category",
      "On Hand Units",
      "Current Share",
      "Recommended Active Units",
      "Recommended Share",
      "Surplus vs Target",
    ],
    ...inventoryRows.map((row) => [
      row.warehouse_region,
      row.category,
      row.on_hand_units,
      row.inventory_share,
      row.recommended_active_units,
      row.recommended_share,
      row.inventory_surplus_vs_target,
    ]),
  ];
  inventorySheet.getRange(`A4:G${inventoryMatrix.length + 3}`).values = inventoryMatrix;
  formatHeader(inventorySheet.getRange("A4:G4"));
  inventorySheet.getRange(`D5:D${inventoryMatrix.length + 3}`).format.numberFormat = "0.0%";
  inventorySheet.getRange(`F5:F${inventoryMatrix.length + 3}`).format.numberFormat = "0.0%";
  inventorySheet.getRange(`G5:G${inventoryMatrix.length + 3}`).conditionalFormats.addDataBar({
    color: "#2563EB",
    gradient: true,
  });
  inventorySheet.freezePanes.freezeRows(4);
  inventorySheet.tables.add(`A4:G${inventoryMatrix.length + 3}`, true, "InventoryTable");

  setTitle(kpiSheet, "KPI Glossary", "Reference metrics requested in the project template.");
  const kpiMatrix = [
    ["KPI Name", "Formula or Logic", "Business Use"],
    ...kpis.map((row) => [row.kpi_name, row.formula_or_logic, row.business_use]),
  ];
  kpiSheet.getRange(`A4:C${kpiMatrix.length + 3}`).values = kpiMatrix;
  formatHeader(kpiSheet.getRange("A4:C4"));
  kpiSheet.getRange(`A4:C${kpiMatrix.length + 3}`).format.wrapText = true;
  kpiSheet.tables.add(`A4:C${kpiMatrix.length + 3}`, true, "KpiGlossaryTable");

  setTitle(
    assumptionSheet,
    "Assumptions and Use Notes",
    "Business rules used to convert forecasts into inventory and placement decisions.",
  );
  assumptionSheet.getRange("A4:B11").values = [
    ["Topic", "Decision Rule"],
    ["Forecast grain", "Weekly category-region, then allocated to SKU-region using observed demand mix share."],
    ["Benchmarking", "Moving average versus event-aware regression; choose the lower-WAPE method by category-region."],
    ["Safety stock", "Two-week cycle stock plus safety stock using target service level and normal-demand approximation."],
    ["Placement logic", "Primary node chosen on service-weighted score; backup node provides resilience buffer."],
    ["Transfer metric", "Reactive transfer cost estimates expedited cross-node units needed to protect on-time fill."],
    ["Inventory implication", "Use current on-hand only as directional placement signal because source inventory is heavily overstocked."],
    ["Limitations", "No true SKU-region time-series history, no supplier lead times, and no transfer-history table."],
  ];
  formatHeader(assumptionSheet.getRange("A4:B4"));
  assumptionSheet.getRange("A5:A11").format = { font: { bold: true } };
  assumptionSheet.getRange("A4:B11").format.wrapText = true;
  assumptionSheet.getRange("A4:B11").format.columnWidth = 34;

  const inspection = await workbook.inspect({
    kind: "table",
    range: "Dashboard!A1:F17",
    include: "values",
    tableMaxRows: 20,
    tableMaxCols: 8,
  });
  console.log(inspection.ndjson);
  const errorScan = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 50 },
    summary: "final formula error scan",
  });
  console.log(errorScan.ndjson);
  await workbook.render({ sheetName: "Dashboard", range: "A1:P36", scale: 1.2 });

  await fs.mkdir(OUTPUT_DIR, { recursive: true });
  const xlsx = await SpreadsheetFile.exportXlsx(workbook);
  await xlsx.save(path.join(OUTPUT_DIR, "amazon_forecasting_placement_dashboard.xlsx"));
}

await main();
