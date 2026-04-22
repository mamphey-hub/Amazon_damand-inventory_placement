# Assumptions and Limitations

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

- Placement adherence baseline: **66.8%**
- Total on-hand units observed: **8,926,517**
- Recommended active inventory units: **3,835**
- Implied weeks of supply at current on-hand: **4655.4 weeks**

## Highest-risk baseline segments

| category | demand_region | baseline_stockout_risk | recommended_stockout_risk | baseline_fill_rate | recommended_fill_rate |
| --- | --- | --- | --- | --- | --- |
| Electronics | North | 0.0028 | 0.0 | 0.9643 | 1.0 |
| Toys | North | 0.0018 | 0.0 | 1.0 | 1.0 |
| Toys | West | 0.001 | 0.0 | 1.0 | 1.0 |
| Home | North | 0.0002 | 0.0 | 0.8688 | 1.0 |
| Toys | South | 0.0001 | 0.0 | 1.0 | 1.0 |

## Largest surplus pools versus recommended active inventory

| warehouse_region | category | on_hand_units | recommended_active_units | inventory_surplus_vs_target |
| --- | --- | --- | --- | --- |
| East | Electronics | 667674 | 44.84 | 667629.16 |
| Central | Electronics | 659498 | 226.20999999999998 | 659271.79 |
| North | Electronics | 635836 | 203.32 | 635632.68 |
| South | Electronics | 544860 | 0.0 | 544860.0 |
| West | Electronics | 489397 | 137.31 | 489259.69 |
| North | Pet | 422813 | 216.87 | 422596.13 |
| Central | Pet | 388245 | 264.46999999999997 | 387980.53 |
| Central | Toys | 366111 | 212.06 | 365898.94 |
| North | Toys | 362129 | 210.61 | 361918.39 |
| East | Toys | 353652 | 41.75 | 353610.25 |
