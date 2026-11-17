# Bengaluru Housing Market: MIS Report and Data Quality Audit

A reporting project on **13,320 real Bengaluru property listings**. It shows the work an MIS / data analyst does day to day: audit messy data, document every cleaning decision, build a formula-driven Excel MIS workbook, and cross-check it against SQL.

**Stack:** Python (pandas) · SQL (DuckDB) · Excel (formulas, charts)

## What I did
1. **Audited the raw data** (`src/clean.py`). Removed 529 exact duplicates, standardised area types, converted 46 listings given in other units (sq. metres, sq. yards, acres and so on) to sqft, averaged 192 area ranges such as "2100 - 2850", and parsed BHK from text ("2 BHK", "4 Bedroom").
2. **Validated the records.** Flagged listings under 300 sqft per bedroom, bathroom counts that don't fit the BHK, and price-per-sqft outliers within each BHK group. 1,391 rows (10.9%) were excluded from KPIs. Nothing was deleted silently: every rule logs its count.
3. **Built the Excel MIS workbook** (`output/Bengaluru_Housing_MIS.xlsx`). Every KPI is a live formula (`COUNTIFS`, `AVERAGEIFS`) over the Clean Data sheet: summary KPIs, price bands, area-type, BHK and locality reports, and a Data Quality sheet with reconciliation checks.
4. **Cross-checked with SQL** (`sql/analysis.sql`). The workbook and the DuckDB queries give the same numbers, for example 11,383 valid listings and an average of ₹5,800 per sqft.

## Findings (valid listings, asking prices)
- **Average price ₹101 lakh, median ₹68 lakh.** The mean is pulled up by a few expensive listings, so the median is the better "typical" figure.
- **Plot listings cost the most per sqft** (about ₹9,073) against about ₹5,429 for super built-up homes.
- **Price per sqft rises with size**: about ₹4,879 for 2 BHK, ₹9,177 for 4 BHK.
- **Supply is concentrated in the middle:** 45.8% of valid listings are ₹50 lakh to ₹1 crore, and 26.4% are under ₹50 lakh.
- **78% of listings are ready to move.**
- **Locality premiums vary widely.** Among localities with 30+ valid listings, Rajaji Nagar averages about 116% above the city average and Koramangala about 57% above (small samples, so treat as indicative).
- **Data completeness:** society is filled for only 58% of listings, which limits any building-level analysis.

## Run it
```bash
pip install -r requirements.txt
python src/clean.py          # cleaning + data quality log
python src/run_sql.py        # SQL results
python src/build_excel.py    # builds output/Bengaluru_Housing_MIS.xlsx
```
The Excel file ships with calculated values. If you rebuild it, open it in Excel once so the formulas calculate.

## Limits
- Prices are **asking prices**, not completed sales.
- There is no listing date (possession dates carry no year), so this is a snapshot, not a trend.
- There is no listing ID, so identical rows were treated as duplicates. Two genuinely identical flats would be merged.
- Price-per-sqft outliers were defined statistically (1.5 x IQR within each BHK), not by checking each listing.

