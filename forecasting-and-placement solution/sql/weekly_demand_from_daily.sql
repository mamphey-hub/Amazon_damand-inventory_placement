SELECT DATE_TRUNC('week', date) AS week_start, sku_id, region,
       SUM(units_ordered) AS units_ordered
FROM daily_demand
GROUP BY 1,2,3
ORDER BY 1,2,3;
