# Executive Memo

## Recommendation

Use a category-region weekly forecasting layer with model selection between moving-average and event-aware regression, then convert the winning forecast into two-week target inventory plus safety stock for node placement decisions.

## Why this works

- The source data does **not** support true time-series forecasting at the `sku_id`-`region` grain because each SKU-region appears only once historically.
- The strongest available signal is the weekly `category`-`region` history enriched with promotions, seasonality, and weather.
- We therefore forecast at `category`-`region`, then allocate to active SKU-region combinations using observed demand mix shares.

## Quantified impact

- Forecast WAPE improves from **57.92%** for the moving-average baseline to **54.64%** for the selected forecasting policy.
- Estimated on-time fill rate improves from **80.45%** to **97.80%**.
- Estimated stockout risk falls from **0.03%** to **0.00%**.
- Estimated reactive transfer cost for the active planning inventory falls from **$484** to **$50** over the modeled cycle.
- Average effective lead time improves from **2.33 days** to **1.70 days**.

## Management actions

1. Use the selected forecast by category and region for the next six weekly cycles.
2. Position primary inventory in the recommended node and maintain a smaller backup share in the secondary node.
3. Freeze or defer replenishment into the most overstocked node-category pools before authorizing any new buys.
4. Track fill rate, stockout risk, transfer cost, and placement adherence weekly in the dashboard.

## Key trade-offs

- Faster delivery generally requires a modest bias toward lower-lead nodes, even when their outbound cost is not the lowest.
- The data supports robust network planning at category-region level, but not confident long-horizon SKU-level learning.
- Because actual on-hand inventory is far above the modeled active target, the near-term value comes more from **repositioning future replenishment** than from urgent network buys.
