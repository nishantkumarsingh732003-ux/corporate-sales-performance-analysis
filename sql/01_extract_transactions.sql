-- Query 01: Extract cleaned transaction base for Excel ingestion
SELECT 
    t.OrderID, 
    t.OrderDate, 
    t.Region, 
    t.ProductCategory, 
    t.Quantity, 
    c.UnitPriceUSD, 
    (t.Quantity * c.UnitPriceUSD) AS GrossRevenueUSD, 
    t.CalculatedSalesUSD 
FROM raw_orders t 
INNER JOIN product_catalog c 
    ON t.ProductCategory = c.ProductCategory 
WHERE t.CalculatedSalesUSD > 0 
ORDER BY t.OrderDate ASC;