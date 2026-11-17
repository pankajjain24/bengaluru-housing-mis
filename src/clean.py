"""
Bengaluru housing listings: cleaning and data-quality audit.

Raw file: 13,320 real listings (public Kaggle dataset, mirrored on GitHub).
Every rule below logs how many rows it touched, so the cleaning is auditable.

Run:  python src/clean.py
Writes: data/cleaned_listings.csv, output/quality_log.json
"""
import json, re
from pathlib import Path
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "data" / "Bengaluru_House_Data.csv"
OUT = ROOT / "output"; OUT.mkdir(exist_ok=True)

# Convert non-sqft units to square feet (standard conversion factors).
UNIT_TO_SQFT = {"Sq. Meter": 10.7639, "Sq. Yards": 9.0, "Acres": 43560.0,
                "Cents": 435.6, "Guntha": 1089.0, "Perch": 272.25, "Grounds": 2400.0}

def parse_sqft(v):
    """Return (sqft, note). note in {'plain','range_mid','converted','unparseable'}."""
    s = str(v).strip()
    try:
        return float(s), "plain"
    except ValueError:
        pass
    if " - " in s:
        a, b = s.split(" - ")
        try:
            return (float(a) + float(b)) / 2, "range_mid"
        except ValueError:
            return np.nan, "unparseable"
    m = re.match(r"^([\d\.]+)\s*(.+)$", s)
    if m and m.group(2).strip() in UNIT_TO_SQFT:
        return float(m.group(1)) * UNIT_TO_SQFT[m.group(2).strip()], "converted"
    return np.nan, "unparseable"

def main():
    log = {}
    df = pd.read_csv(RAW)
    log["raw_rows"] = int(len(df))

    # 1. Exact duplicates (no listing ID exists, so identical rows are treated as repeats)
    n = len(df); df = df.drop_duplicates().copy()
    log["exact_duplicates_removed"] = int(n - len(df))

    # 2. Text standardisation
    df["area_type"] = df["area_type"].str.replace(r"\s+", " ", regex=True).str.strip()
    df["location"] = df["location"].str.strip()
    log["missing_location_removed"] = int(df["location"].isna().sum())
    df = df.dropna(subset=["location"])

    # 3. Society: 40%+ missing, so keep rows but flag
    log["missing_society_flagged"] = int(df["society"].isna().sum())
    df["has_society"] = df["society"].notna().astype(int)
    df = df.drop(columns=["society"])

    # 4. BHK from the 'size' text ("2 BHK", "4 Bedroom", "1 RK")
    log["missing_size_removed"] = int(df["size"].isna().sum())
    df = df.dropna(subset=["size"]).copy()
    df["bhk"] = df["size"].str.extract(r"^(\d+)")[0].astype(int)
    df = df.drop(columns=["size"])

    # 5. Total area: ranges, other units, unparseable
    parsed = df["total_sqft"].apply(parse_sqft)
    df["total_sqft_num"] = [p[0] for p in parsed]
    notes = pd.Series([p[1] for p in parsed], index=df.index)
    log["sqft_range_averaged"] = int((notes == "range_mid").sum())
    log["sqft_unit_converted"] = int((notes == "converted").sum())
    log["sqft_unparseable_removed"] = int((notes == "unparseable").sum())
    df = df[notes != "unparseable"].drop(columns=["total_sqft"]).rename(columns={"total_sqft_num": "total_sqft"})

    # 6. Missing counts kept as blanks (not guessed), but logged
    log["missing_bath_kept_blank"] = int(df["bath"].isna().sum())
    log["missing_balcony_kept_blank"] = int(df["balcony"].isna().sum())

    # 7. Availability: ready vs dated (possession dates carry no year, so grouped)
    df["availability_group"] = np.where(df["availability"].str.strip().eq("Ready To Move"),
                                        "Ready to move", "Upcoming (dated)")

    # 8. Derived measure: price is in lakhs of rupees
    df["price_lakh"] = df["price"]
    df["price_per_sqft"] = (df["price_lakh"] * 100000 / df["total_sqft"]).round(0)
    df = df.drop(columns=["price", "availability"])

    # 9. Validation rules: flag, keep the row, exclude flagged rows from KPIs
    df["flag_small_area"] = (df["total_sqft"] / df["bhk"] < 300).astype(int)       # under 300 sqft per bedroom
    df["flag_bath_mismatch"] = (df["bath"] > df["bhk"] + 2).astype(int)            # more than 2 extra bathrooms
    q1 = df.groupby("bhk")["price_per_sqft"].transform(lambda s: s.quantile(0.25))
    q3 = df.groupby("bhk")["price_per_sqft"].transform(lambda s: s.quantile(0.75))
    iqr = q3 - q1
    df["flag_price_outlier"] = ((df["price_per_sqft"] < q1 - 1.5 * iqr) | (df["price_per_sqft"] > q3 + 1.5 * iqr)).astype(int)
    df["is_valid"] = (1 - df[["flag_small_area", "flag_bath_mismatch", "flag_price_outlier"]].max(axis=1)).astype(int)
    log["flag_small_area"] = int(df["flag_small_area"].sum())
    log["flag_bath_mismatch"] = int(df["flag_bath_mismatch"].sum())
    log["flag_price_outlier"] = int(df["flag_price_outlier"].sum())
    log["rows_after_cleaning"] = int(len(df))
    log["valid_rows_used_for_kpis"] = int(df["is_valid"].sum())
    log["rows_flagged_excluded"] = int(len(df) - df["is_valid"].sum())

    df.insert(0, "listing_id", range(1, len(df) + 1))
    df.to_csv(ROOT / "data" / "cleaned_listings.csv", index=False)
    (OUT / "quality_log.json").write_text(json.dumps(log, indent=2))
    for k, v in log.items(): print(f"  {k:<32}{v:>8,}")

if __name__ == "__main__":
    main()
