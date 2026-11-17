"""Builds the formula-driven Excel MIS workbook from the cleaned data.
Every KPI is a live formula over the 'Clean Data' sheet, so it recalculates if the data changes.
Run: python src/build_excel.py   (then recalc with LibreOffice if you want cached values)
"""
import json
from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
df = pd.read_csv(ROOT / "data" / "cleaned_listings.csv")
log = json.loads((ROOT / "output" / "quality_log.json").read_text())
N = len(df); LAST = N + 1

F = "Arial"
H = Font(name=F, bold=True, color="FFFFFF", size=10); HF = PatternFill("solid", fgColor="1F4E79")
T = Font(name=F, bold=True, size=14, color="1F4E79"); B = Font(name=F, size=10); BB = Font(name=F, size=10, bold=True)
INP = Font(name=F, size=10, color="0000FF"); NOTE = Font(name=F, size=9, italic=True, color="666666")
thin = Side(style="thin", color="D0D0D0"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin)

wb = Workbook()
ws_s = wb.active; ws_s.title = "MIS Summary"
ws_b = wb.create_sheet("BHK Report"); ws_l = wb.create_sheet("Locality Report")
ws_q = wb.create_sheet("Data Quality"); ws_d = wb.create_sheet("Clean Data")

# ---------- Clean Data ----------
cols = ["listing_id", "area_type", "location", "bhk", "bath", "balcony", "has_society", "total_sqft",
        "availability_group", "price_lakh", "price_per_sqft", "flag_small_area", "flag_bath_mismatch",
        "flag_price_outlier", "is_valid"]
for j, c in enumerate(cols, 1):
    cell = ws_d.cell(1, j, c); cell.font = H; cell.fill = HF
for i, row in enumerate(df[cols].itertuples(index=False), 2):
    for j, v in enumerate(row, 1):
        ws_d.cell(i, j, None if pd.isna(v) else (v.item() if hasattr(v, "item") else v)).font = B
for j, w in enumerate([10, 20, 28, 6, 6, 8, 11, 11, 18, 11, 13, 14, 16, 16, 9], 1):
    ws_d.column_dimensions[get_column_letter(j)].width = w
ws_d.freeze_panes = "A2"; ws_d.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{LAST}"
C = {c: f"'Clean Data'!${get_column_letter(j)}$2:${get_column_letter(j)}${LAST}" for j, c in enumerate(cols, 1)}
VALID = C["is_valid"]

def hdr(ws, r, labels, c0=1):
    for k, t in enumerate(labels):
        x = ws.cell(r, c0 + k, t); x.font = H; x.fill = HF; x.alignment = Alignment(horizontal="center", wrap_text=True); x.border = BOX
def put(ws, ref, v, font=B, fmt=None):
    c = ws[ref]; c.value = v; c.font = font; c.border = BOX
    if fmt: c.number_format = fmt
    return c

# ---------- MIS Summary ----------
ws_s["A1"] = "Bengaluru Housing Market: MIS Summary"; ws_s["A1"].font = T
ws_s["A2"] = "Source: public Bengaluru house listings dataset (13,320 raw rows). Price in INR lakh. KPIs use rows that pass the validation rules (is_valid = 1)."; ws_s["A2"].font = NOTE
hdr(ws_s, 4, ["KPI", "Value"])
kp = [("Cleaned listings", f"=COUNTA({C['listing_id']})", "#,##0"),
      ("Valid listings (used for KPIs)", f"=SUM({VALID})", "#,##0"),
      ("Flagged and excluded", "=B5-B6", "#,##0"),
      ("% excluded", "=IF(B5=0,0,B7/B5)", "0.0%"),
      ("Avg price (INR lakh)", f"=AVERAGEIFS({C['price_lakh']},{VALID},1)", "#,##0.0"),
      ("Avg price per sqft (INR)", f"=AVERAGEIFS({C['price_per_sqft']},{VALID},1)", "#,##0"),
      ("Ready-to-move share", f'=COUNTIFS({C["availability_group"]},"Ready to move",{VALID},1)/B6', "0.0%")]
for k, (lab, f, fmt) in enumerate(kp, 5):
    put(ws_s, f"A{k}", lab, BB); put(ws_s, f"B{k}", f, B, fmt)

ws_s["A13"] = "Supply by price band (INR lakh)"; ws_s["A13"].font = BB
hdr(ws_s, 14, ["Band", "From (>=)", "To (<)", "Listings", "Share"])
bands = [("Under 50L", 0, 50), ("50L to 1Cr", 50, 100), ("1Cr to 2Cr", 100, 200), ("Over 2Cr", 200, 100000)]
for k, (nm, lo, hi) in enumerate(bands, 15):
    put(ws_s, f"A{k}", nm); put(ws_s, f"B{k}", lo, INP); put(ws_s, f"C{k}", hi, INP)
    put(ws_s, f"D{k}", f'=COUNTIFS({C["price_lakh"]},">="&B{k},{C["price_lakh"]},"<"&C{k},{VALID},1)', B, "#,##0")
    put(ws_s, f"E{k}", f"=IF($B$6=0,0,D{k}/$B$6)", B, "0.0%")
put(ws_s, "A19", "Total", BB); put(ws_s, "D19", "=SUM(D15:D18)", BB, "#,##0"); put(ws_s, "E19", "=SUM(E15:E18)", BB, "0.0%")
ws_s["A20"] = "Blue cells are inputs you can change (band limits)."; ws_s["A20"].font = NOTE

ws_s["A22"] = "By area type"; ws_s["A22"].font = BB
hdr(ws_s, 23, ["Area type", "Listings", "Share", "Avg price/sqft (INR)"])
for k, at in enumerate(["Super built-up Area", "Built-up Area", "Plot Area", "Carpet Area"], 24):
    put(ws_s, f"A{k}", at, INP)
    put(ws_s, f"B{k}", f"=COUNTIFS({C['area_type']},A{k},{VALID},1)", B, "#,##0")
    put(ws_s, f"C{k}", f"=IF($B$6=0,0,B{k}/$B$6)", B, "0.0%")
    put(ws_s, f"D{k}", f"=IFERROR(AVERAGEIFS({C['price_per_sqft']},{C['area_type']},A{k},{VALID},1),0)", B, "#,##0")
for col, w in zip("ABCDE", [34, 16, 14, 20, 12]): ws_s.column_dimensions[col].width = w
ch = BarChart(); ch.type = "col"; ch.title = "Listings by price band"; ch.legend = None
ch.add_data(Reference(ws_s, min_col=4, min_row=14, max_row=18), titles_from_data=True)
ch.set_categories(Reference(ws_s, min_col=1, min_row=15, max_row=18)); ch.height = 7; ch.width = 13
ws_s.add_chart(ch, "G4")

# ---------- BHK Report ----------
ws_b["A1"] = "Report by bedroom count (BHK)"; ws_b["A1"].font = T
hdr(ws_b, 3, ["BHK", "Listings", "Avg sqft", "Avg price (lakh)", "Avg price/sqft (INR)"])
for k, b in enumerate([1, 2, 3, 4, 5], 4):
    put(ws_b, f"A{k}", b, INP)
    put(ws_b, f"B{k}", f"=COUNTIFS({C['bhk']},A{k},{VALID},1)", B, "#,##0")
    put(ws_b, f"C{k}", f"=IFERROR(AVERAGEIFS({C['total_sqft']},{C['bhk']},A{k},{VALID},1),0)", B, "#,##0")
    put(ws_b, f"D{k}", f"=IFERROR(AVERAGEIFS({C['price_lakh']},{C['bhk']},A{k},{VALID},1),0)", B, "#,##0.0")
    put(ws_b, f"E{k}", f"=IFERROR(AVERAGEIFS({C['price_per_sqft']},{C['bhk']},A{k},{VALID},1),0)", B, "#,##0")
for col, w in zip("ABCDE", [8, 12, 12, 16, 20]): ws_b.column_dimensions[col].width = w
ch = BarChart(); ch.type = "col"; ch.title = "Avg price per sqft by BHK"; ch.legend = None
ch.add_data(Reference(ws_b, min_col=5, min_row=3, max_row=8), titles_from_data=True)
ch.set_categories(Reference(ws_b, min_col=1, min_row=4, max_row=8)); ch.height = 7; ch.width = 13
ws_b.add_chart(ch, "G3")

# ---------- Locality Report ----------
valid = df[df.is_valid == 1]
g = valid.groupby("location").agg(n=("listing_id", "count"), p=("price_per_sqft", "mean"))
top = g[g.n >= 30].sort_values("p", ascending=False).head(15).index.tolist()
ws_l["A1"] = "Top 15 localities by avg price per sqft (min 30 valid listings)"; ws_l["A1"].font = T
ws_l["A2"] = "Locality names were selected from the data in Python (blue = input); every number is a formula."; ws_l["A2"].font = NOTE
hdr(ws_l, 4, ["Locality", "Listings", "Avg price/sqft (INR)", "Premium vs city avg"])
for k, loc in enumerate(top, 5):
    put(ws_l, f"A{k}", loc, INP)
    put(ws_l, f"B{k}", f"=COUNTIFS({C['location']},A{k},{VALID},1)", B, "#,##0")
    put(ws_l, f"C{k}", f"=AVERAGEIFS({C['price_per_sqft']},{C['location']},A{k},{VALID},1)", B, "#,##0")
    put(ws_l, f"D{k}", f"=C{k}/'MIS Summary'!$B$10-1", B, "0.0%")
for col, w in zip("ABCD", [28, 12, 22, 20]): ws_l.column_dimensions[col].width = w
ch = BarChart(); ch.type = "bar"; ch.title = "Avg price per sqft (INR)"; ch.legend = None
ch.add_data(Reference(ws_l, min_col=3, min_row=4, max_row=19), titles_from_data=True)
ch.set_categories(Reference(ws_l, min_col=1, min_row=5, max_row=19)); ch.height = 10; ch.width = 14
ws_l.add_chart(ch, "F4")

# ---------- Data Quality ----------
ws_q["A1"] = "Data quality log"; ws_q["A1"].font = T
ws_q["A2"] = "Counts come from src/clean.py (quality_log.json). Check cells below reconcile the workbook to the log."; ws_q["A2"].font = NOTE
hdr(ws_q, 4, ["Step", "Rows", "Treatment"])
rows = [("Raw rows", log["raw_rows"], "Starting point"),
        ("Exact duplicates removed", log["exact_duplicates_removed"], "No listing ID exists, so identical rows are treated as repeats"),
        ("Missing location removed", log["missing_location_removed"], "Cannot be placed on a locality"),
        ("Missing size removed", log["missing_size_removed"], "BHK cannot be derived"),
        ("Area given as a range, averaged", log["sqft_range_averaged"], "Midpoint of the range used"),
        ("Area in other units, converted", log["sqft_unit_converted"], "Sq. Meter, Sq. Yards, Acres, etc. converted to sqft"),
        ("Society missing, flagged", log["missing_society_flagged"], "Kept; has_society = 0"),
        ("Bathrooms missing, left blank", log["missing_bath_kept_blank"], "Not guessed"),
        ("Balcony missing, left blank", log["missing_balcony_kept_blank"], "Not guessed"),
        ("Flag: under 300 sqft per bedroom", log["flag_small_area"], "Unrealistic area; excluded from KPIs"),
        ("Flag: bathrooms exceed BHK by 2+", log["flag_bath_mismatch"], "Likely entry error; excluded from KPIs"),
        ("Flag: price/sqft outlier within BHK group", log["flag_price_outlier"], "Outside 1.5 x IQR; excluded from KPIs"),
        ("Rows after cleaning", log["rows_after_cleaning"], "Rows in the Clean Data sheet"),
        ("Valid rows used for KPIs", log["valid_rows_used_for_kpis"], "is_valid = 1")]
for k, (a, b, c) in enumerate(rows, 5):
    put(ws_q, f"A{k}", a); put(ws_q, f"B{k}", b, INP, "#,##0"); put(ws_q, f"C{k}", c)
ws_q["A21"] = "Reconciliation checks"; ws_q["A21"].font = BB
put(ws_q, "A22", "Clean Data rows = log"); put(ws_q, "B22", "='MIS Summary'!B5", B, "#,##0"); put(ws_q, "C22", '=IF(B22=B17,"OK","CHECK")', BB)
put(ws_q, "A23", "Valid rows = log"); put(ws_q, "B23", "='MIS Summary'!B6", B, "#,##0"); put(ws_q, "C23", '=IF(B23=B18,"OK","CHECK")', BB)
for col, w in zip("ABC", [44, 12, 62]): ws_q.column_dimensions[col].width = w

wb.save(ROOT / "output" / "Bengaluru_Housing_MIS.xlsx"); print("saved", N, "rows")
