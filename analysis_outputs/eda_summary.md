# Merged Data Cleaning, EDA, and Profiling

## Key Findings
- Daily demand covers 2024-07-01 to 2026-01-16 with 5,000 rows across 5,000 SKUs.
- Total units ordered: 219,212.
- Total revenue: $27,316,800.17.
- Average selling price in Daily Demand: $125.13.
- Weather disruption correlation with units ordered: 0.006.

## Cleaning Actions
- Daily Demand: Parsed dates and replaced corrupted weekend_flag values using the calendar date [rows: 5000 -> 5000].
- Event Calendar: Collapsed duplicate dates to one calendar row per day, inferred weekend_flag, and averaged weather_disruption_index [rows: 5000 -> 562].
- Starting Inventory: Canonicalized warehouse_id values to the IDs defined in the Warehouses sheet [rows: 5000 -> 5000] (4000 IDs updated).
- Warehouse Region Costs: Canonicalized warehouse_id values to the IDs defined in the Warehouses sheet [rows: 5000 -> 5000] (4000 IDs updated).
- Weekly Demand: Normalized week values to the Monday week start [rows: 5000 -> 5000].

## Top Regions by Units
| region | total_units | avg_units | total_revenue | avg_price |
| --- | --- | --- | --- | --- |
| East | 60570 | 47.099533437013996 | 7657391.44 | 127.02965785381028 |
| North | 55369 | 45.53371710526316 | 6797173.9 | 123.37329769736841 |
| West | 53854 | 43.15224358974359 | 6742769.95 | 125.70678685897437 |
| South | 49419 | 39.5352 | 6119464.88 | 124.326824 |

## Top Categories by Units
| category | total_units | total_revenue | gross_margin_usd | avg_service_level |
| --- | --- | --- | --- | --- |
| Electronics | 73464 | 9167292.11 | 3718365.4299999997 | 0.9628254349130174 |
| Toys | 36791 | 4592616.95 | 2671817.27 | 0.971 |
| Kitchen | 36482 | 4520969.82 | 3194648.45 | 0.9453333333333334 |
| Pet | 36412 | 4463116.83 | 1717749.09 | 0.9826610576923076 |
| Home | 24095 | 3043708.74 | 580222.8600000001 | 0.9645189189189189 |
| Beauty | 11968 | 1529095.72 | 788874.92 | 0.9310000000000002 |

## Monthly Trend
| month | total_units | total_revenue | avg_weather_disruption |
| --- | --- | --- | --- |
| 2024-07 | 1483 | 214612.89 | 0.048125 |
| 2024-08 | 4195 | 521351.95 | 0.1258762886597938 |
| 2024-09 | 6940 | 853407.45 | 0.13515337423312884 |
| 2024-10 | 9301 | 1116755.57 | 0.2698689956331878 |
| 2024-11 | 10947 | 1292688.81 | 0.14364312267657992 |
| 2024-12 | 13613 | 1710441.56 | 0.25379411764705884 |
| 2025-01 | 17069 | 2085026.16 | 0.16178403755868545 |
| 2025-02 | 14521 | 1762602.94 | 0.12056338028169013 |
| 2025-03 | 15961 | 2049349.18 | 0.11624338624338623 |
| 2025-04 | 19635 | 2384801.52 | 0.15788888888888888 |
| 2025-05 | 19716 | 2488426.51 | 0.14417249417249417 |
| 2025-06 | 19100 | 2422975.46 | 0.20900249376558605 |
| 2025-07 | 19129 | 2398432.77 | 0.13665835411471322 |
| 2025-08 | 14372 | 1788799.58 | 0.17166666666666666 |
| 2025-09 | 12787 | 1617986.82 | 0.09729241877256317 |
| 2025-10 | 9693 | 1223270.73 | 0.07311004784688994 |
| 2025-11 | 6625 | 859564.85 | 0.013835616438356173 |
| 2025-12 | 3253 | 427442.08 | 0.14355263157894738 |
| 2026-01 | 872 | 98863.34 | 0.05454545454545454 |

## Event Effects
| flag_name | flag_value | avg_units | median_units | avg_revenue | row_count |
| --- | --- | --- | --- | --- | --- |
| holiday_peak_flag | 0 | 43.90303319799379 | 42.0 | 5453.354430379747 | 4187 |
| holiday_peak_flag | 1 | 43.53013530135301 | 42.0 | 5514.88950799508 | 813 |
| prime_event_flag | 0 | 43.54225352112676 | 42.0 | 5435.978021566902 | 4544 |
| prime_event_flag | 1 | 46.833333333333336 | 45.5 | 5736.219385964912 | 456 |
| marketing_push_flag | 0 | 43.31198824270418 | 42.0 | 5388.678509342851 | 4763 |
| marketing_push_flag | 1 | 54.50210970464135 | 54.0 | 6964.2381012658225 | 237 |
| weekend_flag | 0 | 43.8670183231538 | 42.0 | 5502.239589117157 | 3602 |
| weekend_flag | 1 | 43.77896995708154 | 42.5 | 5363.185386266095 | 1398 |
