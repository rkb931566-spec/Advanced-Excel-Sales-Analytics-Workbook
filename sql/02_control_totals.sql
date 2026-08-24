-- ============================================================================
-- Control totals. Every number the workbook computes with formulas is checked
-- against this query. If Excel and BigQuery disagree, the workbook is wrong.
--
-- Verified 2026-08-03 - results are reproduced in docs/verified_numbers.md
-- ============================================================================

WITH sales AS (
  SELECT
    EXTRACT(YEAR FROM oi.created_at) AS yr,
    oi.order_id,
    oi.sale_price                    AS revenue,
    oi.sale_price - p.cost           AS profit
  FROM `bigquery-public-data.thelook_ecommerce.order_items` oi
  JOIN `bigquery-public-data.thelook_ecommerce.products` p
    ON p.id = oi.product_id
  WHERE oi.status NOT IN ('Cancelled', 'Returned')
    AND DATE(oi.created_at) BETWEEN '2024-01-01' AND '2025-12-31'
)

SELECT
  yr,
  COUNT(*)                                        AS line_items,
  COUNT(DISTINCT order_id)                        AS orders,
  ROUND(SUM(revenue), 2)                          AS revenue,
  ROUND(SUM(profit), 2)                           AS profit,
  ROUND(SUM(profit) / SUM(revenue), 4)            AS gross_margin,
  ROUND(SUM(revenue) / COUNT(DISTINCT order_id), 2) AS avg_order_value
FROM sales
GROUP BY yr

UNION ALL

SELECT
  NULL                                            AS yr,
  COUNT(*),
  COUNT(DISTINCT order_id),
  ROUND(SUM(revenue), 2),
  ROUND(SUM(profit), 2),
  ROUND(SUM(profit) / SUM(revenue), 4),
  ROUND(SUM(revenue) / COUNT(DISTINCT order_id), 2)
FROM sales

ORDER BY yr;

-- Note on the order count: summing the per-year order counts gives 39,775,
-- while counting distinct orders across the whole window gives 39,743. The
-- 32-order gap is orders whose line items straddle 31 Dec / 1 Jan, counted once
-- in each year. The workbook reports the window-level figure, 39,743.
