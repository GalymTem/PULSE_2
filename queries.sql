-- name: pie_revenue_by_category
SELECT category,
       SUM(revenue) AS revenue
FROM (
  -- Example: adapt to your schema. Replace with your real tables/joins.
  -- For the Chinook-style demo below, you can map "category" to Genre.
  SELECT g."Name" AS category,
         (il."UnitPrice" * il."Quantity") AS revenue
  FROM "InvoiceLine" il
  JOIN "Track" t  ON t."TrackId" = il."TrackId"
  JOIN "Genre" g  ON g."GenreId" = t."GenreId"
) x
GROUP BY category
ORDER BY revenue DESC
LIMIT 10
;
----------------------------------------------------------------

-- name: bar_top_sellers_by_revenue
-- If you don't have "seller" in your schema, pick another dimension (e.g., artist).
SELECT ar."Name" AS seller_id,
       SUM(il."UnitPrice" * il."Quantity") AS revenue
FROM "Artist" ar
JOIN "Album"  al ON al."ArtistId" = ar."ArtistId"
JOIN "Track"  t  ON t."AlbumId"   = al."AlbumId"
JOIN "InvoiceLine" il ON il."TrackId" = t."TrackId"
GROUP BY ar."Name"
ORDER BY revenue DESC
LIMIT 10
;
----------------------------------------------------------------

-- name: barh_avg_review_by_category
-- Demo placeholder: average review isn’t in Chinook. Replace with your field or compute a proxy.
-- For demonstration, we fake "avg_score" and "n_reviews" by grouping durations/prices.
WITH cat AS (
  SELECT g."Name" AS category,
         AVG(t."UnitPrice")::numeric(10,2) AS avg_score,   -- <- replace with real review avg
         COUNT(*)::int AS n_reviews                         -- <- replace with real reviews count
  FROM "Track" t
  JOIN "Genre" g ON g."GenreId" = t."GenreId"
  GROUP BY g."Name"
)
SELECT * FROM cat WHERE n_reviews >= 50
ORDER BY avg_score DESC, n_reviews DESC
LIMIT 20
;
----------------------------------------------------------------

-- name: line_daily_revenue_2010_2014
WITH days AS (
  SELECT d::date AS day
  FROM generate_series('2010-01-01'::date, '2014-12-31'::date, '1 day') AS g(d)
),
rev AS (
  SELECT DATE(i."InvoiceDate") AS day,
         SUM(il."UnitPrice" * il."Quantity") AS revenue
  FROM "InvoiceLine" il
  JOIN "Invoice" i ON i."InvoiceId" = il."InvoiceId"
  WHERE i."InvoiceDate" >= '2010-01-01'::date
    AND i."InvoiceDate" <  '2015-01-01'::date
  GROUP BY 1
)
SELECT d.day, COALESCE(r.revenue, 0) AS revenue
FROM days d
LEFT JOIN rev r USING (day)
ORDER BY d.day;

----------------------------------------------------------------

-- name: hist_order_value
-- Per-invoice totals → distribution
WITH per_invoice AS (
  SELECT il."InvoiceId",
         SUM(il."UnitPrice" * il."Quantity") AS order_value
  FROM "InvoiceLine" il
  GROUP BY il."InvoiceId"
)
SELECT order_value FROM per_invoice
;
----------------------------------------------------------------

-- name: duration_by_genre_minutes
SELECT
  (t."Milliseconds"::float/60000.0) AS duration_min,
  g."Name"                           AS genre
FROM "InvoiceLine" il
JOIN "Track" t ON t."TrackId" = il."TrackId"
JOIN "Genre" g ON g."GenreId" = t."GenreId"
WHERE t."Milliseconds" IS NOT NULL;

----------------------------------------------------------------

-- name: timeslider_monthly_revenue_by_country
SELECT TO_CHAR(date_trunc('month', i."InvoiceDate"), 'YYYY-MM') AS month,
       COALESCE(c."Country", 'Unknown') AS country,
       SUM(il."UnitPrice" * il."Quantity") AS revenue
FROM "InvoiceLine" il
JOIN "Invoice"  i ON i."InvoiceId" = il."InvoiceId"
JOIN "Customer" c ON c."CustomerId" = i."CustomerId"
GROUP BY 1, 2
ORDER BY 1, 2
;


----------------------------------------------------------------

-- name: heatmap_genre_country_revenue
SELECT
  g."Name"                  AS genre,
  i."BillingCountry"        AS country,
  SUM(il."Quantity" * il."UnitPrice") AS value   -- revenue
FROM "Invoice" i
JOIN "InvoiceLine" il ON il."InvoiceId" = i."InvoiceId"
JOIN "Track" t        ON t."TrackId"    = il."TrackId"
JOIN "Genre" g        ON g."GenreId"    = t."GenreId"
GROUP BY 1, 2
ORDER BY value DESC;

----------------------------------------------------------------
-- name: Sunburst with Geography
SELECT
  i."BillingCountry" AS country,
  g."Name"          AS genre,
  a."Name"          AS artist,
  SUM(il."Quantity" * il."UnitPrice") AS value
FROM "Invoice" i
JOIN "InvoiceLine" il ON il."InvoiceId" = i."InvoiceId"
JOIN "Track" t        ON t."TrackId"    = il."TrackId"
JOIN "Album" al       ON al."AlbumId"   = t."AlbumId"
JOIN "Artist" a       ON a."ArtistId"   = al."ArtistId"
JOIN "Genre" g        ON g."GenreId"    = t."GenreId"
GROUP BY 1,2,3
ORDER BY value DESC;

----------------------------------------------------------------
-- name: treemap_genre_artist_revenue
SELECT
  g."Name"                          AS genre,
  a."Name"                          AS artist,
  SUM(il."Quantity" * il."UnitPrice") AS revenue,     -- area
  AVG(il."UnitPrice")                 AS avg_price    -- color
FROM "InvoiceLine" il
JOIN "Track"  t  ON t."TrackId"  = il."TrackId"
JOIN "Album"  al ON al."AlbumId" = t."AlbumId"
JOIN "Artist" a  ON a."ArtistId" = al."ArtistId"
JOIN "Genre"  g  ON g."GenreId"  = t."GenreId"
GROUP BY 1,2
ORDER BY revenue DESC;

--------------------------------------------------------------
-- name: wordcloud_top_tracks
SELECT
  t."Name" AS track,
  COUNT(il."InvoiceLineId") AS sales_count
FROM "InvoiceLine" il
JOIN "Track" t ON t."TrackId" = il."TrackId"
GROUP BY t."Name"
ORDER BY sales_count DESC;
