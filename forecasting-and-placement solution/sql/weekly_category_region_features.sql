SELECT
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
