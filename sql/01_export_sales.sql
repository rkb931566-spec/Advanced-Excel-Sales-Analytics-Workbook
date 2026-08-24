-- ============================================================================
-- Export the analysis grain for the Excel workbook.
--
-- Source : bigquery-public-data.thelook_ecommerce (public, free to query, never
--          expires - chosen so the whole project stays reproducible from a
--          BigQuery sandbox with no billing account attached)
-- Grain  : one row per order line item
-- Window : 2024-01-01 .. 2025-12-31 (two full years, so YoY is meaningful)
-- Output : 57,542 rows -> data/thelook_sales_2024_2025.csv
--
-- Revenue recognition: Cancelled and Returned items are excluded. They exist in
-- the source and would otherwise inflate revenue by counting sales that never
-- completed.
-- ============================================================================

SELECT
  oi.order_id,
  DATE(oi.created_at)                    AS order_date,
  oi.user_id,

  -- customer attributes, joined from the users dimension
  u.country,
  u.gender                               AS customer_gender,
  u.age                                  AS customer_age,
  u.traffic_source,

  -- product attributes, joined from the products dimension
  p.department,
  p.category,
  p.brand,
  SUBSTR(p.name, 1, 60)                  AS product_name,

  -- money. cost comes from the product record, so profit is per line item.
  ROUND(oi.sale_price, 2)                AS revenue,
  ROUND(p.cost, 2)                       AS cost,
  ROUND(oi.sale_price - p.cost, 2)       AS profit

FROM `bigquery-public-data.thelook_ecommerce.order_items` oi
JOIN `bigquery-public-data.thelook_ecommerce.products` p
  ON p.id = oi.product_id
JOIN `bigquery-public-data.thelook_ecommerce.users` u
  ON u.id = oi.user_id
WHERE oi.status NOT IN ('Cancelled', 'Returned')
  AND DATE(oi.created_at) BETWEEN '2024-01-01' AND '2025-12-31'
ORDER BY oi.order_id;   -- sorted so the workbook can flag the first line of
                        -- each order with a single comparison to the row above
