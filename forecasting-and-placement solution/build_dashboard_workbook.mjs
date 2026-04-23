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

function sumBy(items, keyFn, valueFn) {
  const map = new Map();
  for (const item of items) {
    const key = keyFn(item);
    map.set(key, (map.get(key) ?? 0) + valueFn(item));
  }
  return map;
}

async function main() {
  const summary = JSON.parse(await fs.readFile(path.join(OUTPUT_DIR, "solution_summary.json"), "utf8"));
  const benchmark = parseCsv(await fs.readFile(path.join(OUTPUT_DIR, "model_benchmark.csv"), "utf8"));
  const forecasts = parseCsv(await fs.readFile(path.join(OUTPUT_DIR, "category_region_forecasts.csv"), "utf8"));
  const scenarios = parseCsv(await fs.readFile(path.join(OUTPUT_DIR, "placement_scenarios.csv"), "utf8"));
  const inventory = parseCsv(await fs.readFile(path.join(OUTPUT_DIR, "inventory_implications.csv"), "utf8"));
  const kpis = parseCsv(await fs.readFile(path.join(OUTPUT_DIR, "kpi_reference.csv"), "utf8"));

  const fillImprovementByRegion = [...sumBy(
    scenarios,
    (row) => row.demand_region,
    (row) => row.recommended_fill_rate - row.baseline_fill_rate,
  )]
    .map(([region, fill_gain]) => ({ region, fill_gain }))
    .sort((a, b) => b.fill_gain - a.fill_gain);

  const transferByCategory = [...sumBy(
    scenarios,
    (row) => row.category,
    (row) => row.baseline_transfer_cost - row.recommended_transfer_cost,
  )]
    .map(([category, transfer_savings]) => ({ category, transfer_savings }))
    .sort((a, b) => b.transfer_savings - a.transfer_savings);

  const modelMix = [...sumBy(
    benchmark,
    (row) => row.best_model,
    () => 1,
  )]
    .map(([model, combos]) => ({ model, combos }))
    .sort((a, b) => b.combos - a.combos);

  const benchmarkByCategory = [...sumBy(
    benchmark,
    (row) => row.category,
    (row) => row.best_model === "regression" ? row.regression_wape : row.baseline_wape,
  )]
    .map(([category, totalWape]) => {
      const count = benchmark.filter((row) => row.category === category).length;
      return { category, selected_wape: totalWape / Math.max(count, 1) };
    })
    .sort((a, b) => b.selected_wape - a.selected_wape);

  const leadByRegion = scenarios
    .map((row) => ({
      demand_region: row.demand_region,
      baseline_avg_lead_days: row.baseline_avg_lead_days,
      recommended_avg_lead_days: row.recommended_avg_lead_days,
    }))
    .sort((a, b) => b.baseline_avg_lead_days - a.baseline_avg_lead_days);

  const riskSegments = scenarios
    .map((row) => ({
      label: `${row.category}-${row.demand_region}`,
      baseline_stockout_risk: row.baseline_stockout_risk,
      recommended_stockout_risk: row.recommended_stockout_risk,
    }))
    .sort((a, b) => b.baseline_stockout_risk - a.baseline_stockout_risk)
    .slice(0, 8);

  const surplusByNode = [...sumBy(
    inventory,
    (row) => row.warehouse_region,
    (row) => row.inventory_surplus_vs_target,
  )]
    .map(([warehouse_region, surplus_units]) => ({ warehouse_region, surplus_units }))
    .sort((a, b) => b.surplus_units - a.surplus_units);

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

  dashboard.getRange("R2:S6").values = [
    ["Region", "Fill Gain"],
    ...fillImprovementByRegion.map((row) => [row.region, row.fill_gain]),
  ];
  const fillChart = dashboard.charts.add("bar", dashboard.getRange("R2:S6"));
  fillChart.setPosition("Q2", "X18");
  fillChart.title = "Fill-Rate Improvement by Region";
  fillChart.hasLegend = false;
  fillChart.barOptions.direction = "column";
  fillChart.xAxis = { axisType: "textAxis" };
  fillChart.yAxis = { numberFormatCode: "0.0%" };

  dashboard.getRange("R20:S26").values = [
    ["Category", "Transfer Savings"],
    ...transferByCategory.map((row) => [row.category, row.transfer_savings]),
  ];
  const transferChart = dashboard.charts.add("bar", dashboard.getRange("R20:S26"));
  transferChart.setPosition("Q20", "X36");
  transferChart.title = "Transfer-Cost Savings by Category";
  transferChart.hasLegend = false;
  transferChart.barOptions.direction = "column";
  transferChart.xAxis = { axisType: "textAxis" };
  transferChart.yAxis = { numberFormatCode: "$#,##0" };

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
  benchmarkSheet.getRange("J4:K6").values = [
    ["Model", "Winning Combos"],
    ...modelMix.map((row) => [row.model, row.combos]),
  ];
  formatSubheader(benchmarkSheet.getRange("J4:K4"));
  const modelMixChart = benchmarkSheet.charts.add("doughnut", benchmarkSheet.getRange("J4:K6"));
  modelMixChart.setPosition("J8", "P22");
  modelMixChart.title = "Winning Model Mix";
  modelMixChart.hasLegend = true;
  benchmarkSheet.getRange("J24:K30").values = [
    ["Category", "Selected WAPE"],
    ...benchmarkByCategory.map((row) => [row.category, row.selected_wape]),
  ];
  formatSubheader(benchmarkSheet.getRange("J24:K24"));
  const benchmarkCategoryChart = benchmarkSheet.charts.add("bar", benchmarkSheet.getRange("J24:K30"));
  benchmarkCategoryChart.setPosition("J32", "P48");
  benchmarkCategoryChart.title = "Selected WAPE by Category";
  benchmarkCategoryChart.hasLegend = false;
  benchmarkCategoryChart.barOptions.direction = "column";
  benchmarkCategoryChart.xAxis = { axisType: "textAxis" };
  benchmarkCategoryChart.yAxis = { numberFormatCode: "0.0%" };

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
  forecastSheet.getRange("J4:N10").values = [
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
  formatSubheader(forecastSheet.getRange("J4:N4"));
  const regionForecastChart = forecastSheet.charts.add("line", forecastSheet.getRange("J4:N10"));
  regionForecastChart.setPosition("J12", "S30");
  regionForecastChart.title = "Regional Forecast Trajectory";
  regionForecastChart.hasLegend = true;
  regionForecastChart.xAxis = { axisType: "textAxis" };
  regionForecastChart.yAxis = { numberFormatCode: "0" };

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
  placementSheet.getRange(`O4:Q${leadByRegion.length + 4}`).values = [
    ["Demand Region", "Baseline Lead Days", "Recommended Lead Days"],
    ...leadByRegion.map((row) => [
      row.demand_region,
      row.baseline_avg_lead_days,
      row.recommended_avg_lead_days,
    ]),
  ];
  formatSubheader(placementSheet.getRange("O4:Q4"));
  const leadChart = placementSheet.charts.add("bar", placementSheet.getRange(`O4:Q${leadByRegion.length + 4}`));
  leadChart.setPosition("O8", "W24");
  leadChart.title = "Lead-Time Comparison by Region";
  leadChart.hasLegend = true;
  leadChart.barOptions.direction = "column";
  leadChart.barOptions.grouping = "clustered";
  leadChart.xAxis = { axisType: "textAxis" };
  leadChart.yAxis = { numberFormatCode: "0.00" };
  placementSheet.getRange(`O28:Q${riskSegments.length + 28}`).values = [
    ["Segment", "Baseline Risk", "Recommended Risk"],
    ...riskSegments.map((row) => [
      row.label,
      row.baseline_stockout_risk,
      row.recommended_stockout_risk,
    ]),
  ];
  formatSubheader(placementSheet.getRange("O28:Q28"));
  const riskChart = placementSheet.charts.add("bar", placementSheet.getRange(`O28:Q${riskSegments.length + 28}`));
  riskChart.setPosition("O32", "W48");
  riskChart.title = "Top Baseline Stockout-Risk Segments";
  riskChart.hasLegend = true;
  riskChart.barOptions.direction = "column";
  riskChart.barOptions.grouping = "clustered";
  riskChart.xAxis = { axisType: "textAxis" };
  riskChart.yAxis = { numberFormatCode: "0.00%" };

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
  inventorySheet.getRange(`J4:K${surplusByNode.length + 4}`).values = [
    ["Warehouse Region", "Surplus Units"],
    ...surplusByNode.map((row) => [row.warehouse_region, row.surplus_units]),
  ];
  formatSubheader(inventorySheet.getRange("J4:K4"));
  const surplusChart = inventorySheet.charts.add("bar", inventorySheet.getRange(`J4:K${surplusByNode.length + 4}`));
  surplusChart.setPosition("J8", "Q24");
  surplusChart.title = "Inventory Surplus by Node";
  surplusChart.hasLegend = false;
  surplusChart.barOptions.direction = "column";
  surplusChart.xAxis = { axisType: "textAxis" };
  surplusChart.yAxis = { numberFormatCode: "#,##0" };

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
    range: "Dashboard!A1:X36",
    include: "values",
    tableMaxRows: 20,
    tableMaxCols: 24,
  });
  console.log(inspection.ndjson);
  const errorScan = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 50 },
    summary: "final formula error scan",
  });
  console.log(errorScan.ndjson);
  await workbook.render({ sheetName: "Dashboard", range: "A1:X36", scale: 1.1 });
  await workbook.render({ sheetName: "Placement", range: "A1:W48", scale: 1.0 });
  await workbook.render({ sheetName: "Inventory", range: "A1:Q24", scale: 1.0 });

  await fs.mkdir(OUTPUT_DIR, { recursive: true });
  const xlsx = await SpreadsheetFile.exportXlsx(workbook);
  await xlsx.save(path.join(OUTPUT_DIR, "amazon_forecasting_placement_dashboard.xlsx"));
}

await main();
