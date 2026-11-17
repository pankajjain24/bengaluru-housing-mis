-- Bengaluru housing market: MIS queries (DuckDB). Run: python src/run_sql.py
-- Only rows that pass the validation rules (is_valid = 1) feed the KPIs.

CREATE OR REPLACE VIEW listings AS SELECT * FROM read_csv_auto('data/cleaned_listings.csv');

-- Q1. Headline KPIs
SELECT
    COUNT(*)                                   AS valid_listings,
    COUNT(DISTINCT location)                   AS localities,
    ROUND(AVG(price_lakh), 1)                  AS avg_price_lakh,
    ROUND(MEDIAN(price_lakh), 1)               AS median_price_lakh,
    ROUND(AVG(price_per_sqft))                 AS avg_price_per_sqft,
    ROUND(MEDIAN(price_per_sqft))              AS median_price_per_sqft
FROM listings WHERE is_valid = 1;

-- Q2. Data completeness report: how much of each field is filled (all cleaned rows)
SELECT 'bath'      AS field, ROUND(100.0 * COUNT(bath)      / COUNT(*), 1) AS pct_filled FROM listings
UNION ALL SELECT 'balcony', ROUND(100.0 * COUNT(balcony)    / COUNT(*), 1) FROM listings
UNION ALL SELECT 'society', ROUND(100.0 * SUM(has_society)  / COUNT(*), 1) FROM listings;

-- Q3. By area type: where does the market sit?
SELECT area_type,
       COUNT(*)                                AS listings,
       ROUND(100.0 * COUNT(*) / SUM(COUNT(*)) OVER (), 1) AS share_pct,
       ROUND(AVG(price_per_sqft))              AS avg_price_per_sqft
FROM listings WHERE is_valid = 1
GROUP BY area_type ORDER BY listings DESC;

-- Q4. By bedroom count: size and price per sqft
SELECT bhk, COUNT(*) AS listings, ROUND(AVG(total_sqft)) AS avg_sqft,
       ROUND(AVG(price_lakh), 1) AS avg_price_lakh, ROUND(AVG(price_per_sqft)) AS avg_price_per_sqft
FROM listings WHERE is_valid = 1 AND bhk BETWEEN 1 AND 5
GROUP BY bhk ORDER BY bhk;

-- Q5. Locality ranking (min 30 valid listings), with premium vs the city average
WITH city AS (SELECT AVG(price_per_sqft) AS city_avg FROM listings WHERE is_valid = 1)
SELECT l.location, COUNT(*) AS listings, ROUND(AVG(l.price_per_sqft)) AS avg_price_per_sqft,
       ROUND(100.0 * (AVG(l.price_per_sqft) / c.city_avg - 1), 1) AS premium_vs_city_pct
FROM listings l CROSS JOIN city c
WHERE l.is_valid = 1
GROUP BY l.location, c.city_avg HAVING COUNT(*) >= 30
ORDER BY avg_price_per_sqft DESC;

-- Q6. Ready-to-move vs upcoming
SELECT availability_group, COUNT(*) AS listings, ROUND(AVG(price_per_sqft)) AS avg_price_per_sqft
FROM listings WHERE is_valid = 1 GROUP BY availability_group;

-- Q7. Price bands (lakh): how supply is spread
SELECT CASE WHEN price_lakh < 50 THEN '1. Under 50L' WHEN price_lakh < 100 THEN '2. 50L-1Cr'
            WHEN price_lakh < 200 THEN '3. 1Cr-2Cr' ELSE '4. Over 2Cr' END AS price_band,
       COUNT(*) AS listings, ROUND(100.0 * COUNT(*) / SUM(COUNT(*)) OVER (), 1) AS share_pct
FROM listings WHERE is_valid = 1 GROUP BY price_band ORDER BY price_band;
