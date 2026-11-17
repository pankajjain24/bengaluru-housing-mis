"""Runs sql/analysis.sql with DuckDB and prints each result. Run: python src/run_sql.py"""
import os
from pathlib import Path
import duckdb, pandas as pd
ROOT = Path(__file__).resolve().parent.parent; os.chdir(ROOT)
pd.set_option("display.width", 130); pd.set_option("display.max_rows", 40)
con = duckdb.connect()
text = (ROOT / "sql" / "analysis.sql").read_text()
for stmt in [s.strip() for s in text.split(";") if s.strip()]:
    lines = [l for l in stmt.splitlines()]
    title = next((l.lstrip("- ").strip() for l in lines if l.strip().startswith("-- Q")), "")
    sql = "\n".join(l for l in lines if not l.strip().startswith("--")).strip()
    if not sql: continue
    res = con.execute(sql)
    if sql.upper().startswith("CREATE"): continue
    print("\n" + "=" * 90 + "\n" + title + "\n" + "=" * 90)
    print(res.fetch_df().head(15).to_string(index=False))
