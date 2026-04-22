import fs from "node:fs/promises";
import path from "node:path";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const BASE_DIR =
  "C:/Users/Mamphey/Documents/Codex/2026-04-22-files-mentioned-by-the-user-merged/forecasting-and-placement solution";
const OUTPUT_DIR = path.join(BASE_DIR, "outputs");

const W = 1280;
const H = 720;
const COLORS = {
  navy: "#0B1F33",
  slate: "#334155",
  soft: "#F8FAFC",
  blue: "#2563EB",
  green: "#16A34A",
  amber: "#D97706",
  red: "#DC2626",
  border: "#CBD5E1",
  paleBlue: "#DBEAFE",
  paleGreen: "#DCFCE7",
  paleAmber: "#FEF3C7",
};
const FONT = {
  title: "Poppins",
  body: "Lato",
};

function addPanel(slide, left, top, width, height, fill = "#FFFFFF") {
  return slide.shapes.add({
    geometry: "roundRect",
    position: { left, top, width, height },
    fill,
    line: { style: "solid", fill: COLORS.border, width: 1 },
  });
}

function addText(slide, text, left, top, width, height, options = {}) {
  const shape = slide.shapes.add({
    geometry: "rect",
    position: { left, top, width, height },
    fill: "#FFFFFF00",
    line: { style: "solid", fill: "#FFFFFF00", width: 0 },
  });
  shape.text = text;
  shape.text.typeface = options.typeface ?? FONT.body;
  shape.text.fontSize = options.fontSize ?? 22;
  shape.text.color = options.color ?? COLORS.slate;
  shape.text.bold = options.bold ?? false;
  shape.text.alignment = options.alignment ?? "left";
  shape.text.verticalAlignment = options.verticalAlignment ?? "top";
  shape.text.insets = { left: 10, right: 10, top: 8, bottom: 8 };
  return shape;
}

function addTitleBlock(slide, kicker, title, subtitle) {
  addText(slide, kicker, 56, 36, 320, 28, {
    fontSize: 16,
    color: COLORS.blue,
    bold: true,
    typeface: FONT.body,
  });
  addText(slide, title, 56, 62, 760, 92, {
    fontSize: 32,
    color: COLORS.navy,
    bold: true,
    typeface: FONT.title,
  });
  addText(slide, subtitle, 56, 148, 900, 56, {
    fontSize: 18,
    color: COLORS.slate,
    typeface: FONT.body,
  });
}

function addKpiCard(slide, left, top, label, value, fill) {
  addPanel(slide, left, top, 250, 110, fill);
  addText(slide, label, left + 12, top + 8, 226, 34, {
    fontSize: 16,
    color: COLORS.slate,
    bold: true,
  });
  addText(slide, value, left + 12, top + 40, 226, 52, {
    fontSize: 28,
    color: COLORS.navy,
    bold: true,
    typeface: FONT.title,
  });
}

async function main() {
  const summary = JSON.parse(await fs.readFile(path.join(OUTPUT_DIR, "solution_summary.json"), "utf8"));
  const memo = await fs.readFile(path.join(BASE_DIR, "docs", "executive_memo.md"), "utf8");

  const presentation = Presentation.create({ slideSize: { width: W, height: H } });
  presentation.theme.colorScheme = {
    name: "Ops",
    themeColors: {
      bg1: COLORS.soft,
      tx1: COLORS.navy,
      tx2: COLORS.slate,
      accent1: COLORS.blue,
      accent2: COLORS.green,
      accent3: COLORS.amber,
      accent4: COLORS.red,
      accent5: "#1D4ED8",
      accent6: "#0F766E",
    },
  };

  const slide1 = presentation.slides.add();
  slide1.background.fill = COLORS.soft;
  addTitleBlock(
    slide1,
    "AMAZON NETWORK PLANNING CASE",
    "Regional demand forecasting and inventory placement decision package",
    "Weekly category-region forecasting with a placement layer designed to raise fill rate, reduce avoidable transfers, and improve delivery speed.",
  );
  addKpiCard(slide1, 56, 250, "Selected-model WAPE", `${(summary.network_kpis.selected_model_wape * 100).toFixed(1)}%`, COLORS.paleBlue);
  addKpiCard(slide1, 326, 250, "On-time fill rate", `${(summary.network_kpis.fill_rate_recommended * 100).toFixed(1)}%`, COLORS.paleGreen);
  addKpiCard(slide1, 596, 250, "Avg lead days", `${summary.network_kpis.avg_lead_days_recommended.toFixed(2)}`, COLORS.paleAmber);
  addKpiCard(slide1, 866, 250, "Transfer cost", `$${summary.network_kpis.reactive_transfer_cost_recommended.toFixed(0)}`, "#FEE2E2");
  addPanel(slide1, 56, 400, 1168, 240, "#FFFFFF");
  addText(
    slide1,
    "Decision statement",
    72,
    418,
    220,
    28,
    { fontSize: 18, bold: true, color: COLORS.blue },
  );
  addText(
    slide1,
    "Forecast at the category-region level, allocate to active SKU-region mixes, and pre-position inventory into low-lead nodes with a backup buffer rather than planning off the current overstocked network footprint.",
    72,
    452,
    1100,
    90,
    { fontSize: 22, color: COLORS.navy, bold: true, typeface: FONT.title },
  );
  addText(
    slide1,
    `Current on-hand is roughly ${summary.network_kpis.network_weeks_of_supply.toFixed(0)} weeks of modeled active supply, so the immediate operational lever is future replenishment and placement discipline rather than new buys.`,
    72,
    548,
    1080,
    60,
    { fontSize: 18, color: COLORS.slate },
  );

  const slide2 = presentation.slides.add();
  slide2.background.fill = COLORS.soft;
  addTitleBlock(
    slide2,
    "DATA REALITY",
    "Modeling approach chosen to fit the grain of the source data",
    "The input supports a strong regional planning model, but not a statistically credible SKU-region time series for direct forecasting.",
  );
  addPanel(slide2, 56, 220, 540, 420, "#FFFFFF");
  addPanel(slide2, 628, 220, 596, 420, "#FFFFFF");
  addText(slide2, "What the data supports", 72, 238, 260, 28, { fontSize: 18, bold: true, color: COLORS.blue });
  addText(
    slide2,
    "- Weekly demand history by category and region\n- Promotional and calendar signals\n- Warehouse-region cost and lead-time trade-offs\n- Current node inventory mix by category",
    72,
    274,
    470,
    180,
    { fontSize: 20 },
  );
  addText(slide2, "Critical limitation", 644, 238, 260, 28, { fontSize: 18, bold: true, color: COLORS.red });
  addText(
    slide2,
    "Each SKU-region appears only once historically in the demand file. That means a true per-SKU forecasting benchmark would overstate confidence and understate risk.",
    644,
    276,
    520,
    120,
    { fontSize: 22, color: COLORS.navy, bold: true, typeface: FONT.title },
  );
  addText(
    slide2,
    "Chosen solution:\n1. Forecast weekly category-region demand.\n2. Benchmark moving average vs event-aware regression.\n3. Allocate winning forecast to active SKU-region mix.\n4. Convert forecasts into target inventory and node placement rules.",
    644,
    418,
    520,
    170,
    { fontSize: 19 },
  );

  const slide3 = presentation.slides.add();
  slide3.background.fill = COLORS.soft;
  addTitleBlock(
    slide3,
    "FORECAST BENCHMARK",
    "Event-aware regression outperforms the moving-average baseline",
    "The selected policy improves forecast accuracy enough to justify using it as the planning signal for downstream placement decisions.",
  );
  addPanel(slide3, 56, 220, 720, 410, "#FFFFFF");
  addPanel(slide3, 810, 220, 414, 410, "#FFFFFF");
  const benchmarkChart = slide3.charts.add("bar");
  benchmarkChart.position = { left: 82, top: 256, width: 660, height: 330 };
  benchmarkChart.title = "Model benchmark";
  benchmarkChart.categories = ["WAPE", "MAPE"];
  benchmarkChart.hasLegend = true;
  benchmarkChart.barOptions.direction = "column";
  benchmarkChart.barOptions.grouping = "clustered";
  const baselineSeries = benchmarkChart.series.add("Moving average");
  baselineSeries.values = [
    summary.network_kpis.baseline_wape,
    summary.model_summary.find((row) => row.model === "moving_average")?.mape ?? summary.network_kpis.baseline_wape,
  ];
  baselineSeries.fill = COLORS.red;
  const selectedSeries = benchmarkChart.series.add("Selected policy");
  selectedSeries.values = [
    summary.network_kpis.selected_model_wape,
    summary.model_summary.find((row) => row.model === "selected_policy")?.mape ?? summary.network_kpis.selected_model_wape,
  ];
  selectedSeries.fill = COLORS.blue;
  benchmarkChart.yAxis = { numberFormatCode: "0.0%" };
  benchmarkChart.xAxis = { axisType: "textAxis" };
  addText(slide3, "Why it wins", 836, 252, 200, 28, { fontSize: 18, bold: true, color: COLORS.blue });
  addText(
    slide3,
    "- Captures promotional lift\n- Uses seasonal week patterns\n- Preserves a transparent benchmark\n- Allows segment-level model selection",
    836,
    292,
    330,
    160,
    { fontSize: 20 },
  );
  addText(
    slide3,
    `Accuracy moved from ${(summary.network_kpis.baseline_wape * 100).toFixed(1)}% WAPE to ${(summary.network_kpis.selected_model_wape * 100).toFixed(1)}% WAPE, which is modest but directionally meaningful for inventory positioning.`,
    836,
    470,
    330,
    110,
    { fontSize: 18, color: COLORS.slate },
  );

  const slide4 = presentation.slides.add();
  slide4.background.fill = COLORS.soft;
  addTitleBlock(
    slide4,
    "PLACEMENT LOGIC",
    "Place most inventory in the best service node and keep a smaller backup buffer",
    "The network policy favors low-lead nodes first, then uses a secondary buffer node to protect service where volatility or service targets are high.",
  );
  const nodeLabels = [
    ["North", 110, 300],
    ["West", 290, 470],
    ["Central", 520, 360],
    ["East", 760, 250],
    ["South", 980, 470],
  ];
  for (const [label, left, top] of nodeLabels) {
    addPanel(slide4, left, top, 140, 70, "#FFFFFF");
    addText(slide4, label, left + 10, top + 14, 120, 30, {
      fontSize: 20,
      bold: true,
      alignment: "center",
      color: COLORS.navy,
    });
  }
  addText(slide4, "Illustrative node map", 100, 248, 220, 28, {
    fontSize: 18,
    bold: true,
    color: COLORS.blue,
  });
  addPanel(slide4, 900, 220, 290, 300, "#FFFFFF");
  addText(slide4, "Placement rule", 920, 240, 180, 28, { fontSize: 18, bold: true, color: COLORS.blue });
  addText(
    slide4,
    "Primary node share\n- 75% to 90%\n\nBackup node share\n- 10% to 25%\n\nTrigger a larger backup buffer when service target is high or weekly volatility is elevated.",
    920,
    278,
    240,
    190,
    { fontSize: 19 },
  );
  addText(
    slide4,
    "The result is a cleaner active footprint that serves demand faster than the current evenly spread node mix.",
    900,
    540,
    280,
    60,
    { fontSize: 18, color: COLORS.slate },
  );

  const slide5 = presentation.slides.add();
  slide5.background.fill = COLORS.soft;
  addTitleBlock(
    slide5,
    "KPI IMPACT",
    "Network service improves sharply while avoidable transfer cost falls",
    "The policy is designed to improve speed and fill performance without relying on more total inventory.",
  );
  addPanel(slide5, 56, 220, 730, 420, "#FFFFFF");
  addPanel(slide5, 820, 220, 404, 420, "#FFFFFF");
  const impactChart = slide5.charts.add("bar");
  impactChart.position = { left: 86, top: 258, width: 670, height: 320 };
  impactChart.title = "Baseline vs recommended";
  impactChart.categories = ["Fill rate", "Transfer cost", "Lead days"];
  impactChart.hasLegend = true;
  impactChart.barOptions.direction = "column";
  impactChart.barOptions.grouping = "clustered";
  const impactBaseline = impactChart.series.add("Baseline");
  impactBaseline.values = [
    summary.network_kpis.fill_rate_baseline,
    summary.network_kpis.reactive_transfer_cost_baseline,
    summary.network_kpis.avg_lead_days_baseline,
  ];
  impactBaseline.fill = COLORS.red;
  const impactRecommended = impactChart.series.add("Recommended");
  impactRecommended.values = [
    summary.network_kpis.fill_rate_recommended,
    summary.network_kpis.reactive_transfer_cost_recommended,
    summary.network_kpis.avg_lead_days_recommended,
  ];
  impactRecommended.fill = COLORS.green;
  impactChart.xAxis = { axisType: "textAxis" };
  addText(slide5, "Headline improvements", 846, 250, 240, 28, { fontSize: 18, bold: true, color: COLORS.blue });
  addText(
    slide5,
    `- Fill rate: ${(summary.network_kpis.fill_rate_baseline * 100).toFixed(1)}% -> ${(summary.network_kpis.fill_rate_recommended * 100).toFixed(1)}%\n- Transfer cost: $${summary.network_kpis.reactive_transfer_cost_baseline.toFixed(0)} -> $${summary.network_kpis.reactive_transfer_cost_recommended.toFixed(0)}\n- Lead days: ${summary.network_kpis.avg_lead_days_baseline.toFixed(2)} -> ${summary.network_kpis.avg_lead_days_recommended.toFixed(2)}\n- Placement adherence baseline: ${(summary.network_kpis.placement_adherence_baseline * 100).toFixed(1)}%`,
    846,
    292,
    320,
    220,
    { fontSize: 19 },
  );

  const slide6 = presentation.slides.add();
  slide6.background.fill = COLORS.soft;
  addTitleBlock(
    slide6,
    "ACTION PLAN",
    "Use the model as a planning layer and govern to a small set of weekly KPIs",
    "The business case is strongest when replenishment decisions are disciplined and the data limitation is acknowledged up front.",
  );
  addPanel(slide6, 56, 220, 560, 400, "#FFFFFF");
  addPanel(slide6, 648, 220, 576, 400, "#FFFFFF");
  addText(slide6, "Actions for the sponsor", 76, 242, 220, 28, { fontSize: 18, bold: true, color: COLORS.blue });
  addText(
    slide6,
    "1. Use the selected weekly forecast as the planning baseline.\n2. Position inventory in the recommended primary and backup nodes.\n3. Pause replenishment into the most overstocked node-category pools.\n4. Review fill rate, risk, transfer cost, and adherence weekly.\n5. Upgrade the data model before scaling to true SKU-level ML.",
    76,
    280,
    500,
    250,
    { fontSize: 20 },
  );
  addText(slide6, "Trade-offs and limitations", 668, 242, 240, 28, { fontSize: 18, bold: true, color: COLORS.red });
  addText(
    slide6,
    "The source data does not contain repeat SKU-region history, supplier lead times, or node-to-node transfer history. This solution is therefore best used as a transparent planning prototype and prioritization tool rather than a production optimizer.",
    668,
    282,
    520,
    170,
    { fontSize: 20 },
  );
  addText(
    slide6,
    memo.split("## Key trade-offs")[0].slice(0, 420),
    668,
    470,
    520,
    120,
    { fontSize: 16, color: COLORS.slate },
  );

  await presentation.export({ slide: slide1, format: "png", scale: 1 });
  await presentation.export({ slide: slide5, format: "png", scale: 1 });
  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(path.join(OUTPUT_DIR, "amazon_forecasting_placement_deck.pptx"));
}

await main();
