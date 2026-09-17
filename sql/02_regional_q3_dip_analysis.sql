-- Query 02: Isolate Quarterly Seasonality & Quantify Q3 Contraction
WITH QuarterlyRevenue AS (
    SELECT 
        CASE 
            WHEN EXTRACT(MONTH FROM CAST(OrderDate AS DATE)) BETWEEN 1 AND 3 THEN 'Q1'
            WHEN EXTRACT(MONTH FROM CAST(OrderDate AS DATE)) BETWEEN 4 AND 6 THEN 'Q2'
            WHEN EXTRACT(MONTH FROM CAST(OrderDate AS DATE)) BETWEEN 7 AND 9 THEN 'Q3'
            ELSE 'Q4'
        END AS SalesQuarter,
        SUM(CalculatedSalesUSD) AS GrossQuarterRevenue
    FROM raw_orders
    GROUP BY 1
)
SELECT 
    SalesQuarter,
    GrossQuarterRevenue,
    ROUND((GrossQuarterRevenue - LAG(GrossQuarterRevenue) OVER (ORDER BY SalesQuarter)) / LAG(GrossQuarterRevenue) OVER (ORDER BY SalesQuarter) * 100, 2) AS QoQ_Growth_Pct
FROM QuarterlyRevenue;
