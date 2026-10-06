# /// script
# requires-python = ">=3.10"
# dependencies = ["pandas"]
# ///
"""把 data/monthly.csv 整理成網頁可直接載入的 docs/data.js（window.BI_DATA）。

用法：uv run etl/build_data.py
"""
import json
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
df = pd.read_csv(ROOT / "data" / "monthly.csv", encoding="utf-8-sig")
mp = pd.read_csv(ROOT / "data" / "region_mapping.csv", encoding="utf-8-sig")

months = df[["year_roc", "year_ad", "month"]].drop_duplicates().sort_values(["year_roc", "month"]).values.tolist()
mi = {(r, m): i for i, (r, _, m) in enumerate(months)}
regions = mp.region.tolist()
ri = {r: i for i, r in enumerate(regions)}

rows = [
    [mi[(r.year_roc, r.month)], ri[r.region], r.births_male, r.births_female, r.deaths_male, r.deaths_female,
     r.marriages, r.divorces, r.pop_total, r.pop_male, r.pop_female]
    for r in df.itertuples()
]
data = {
    "months": months,  # [民國年, 西元年, 月]
    "regions": [{"name": r.region, "area": r.area, "type": r.type} for r in mp.itertuples()],
    "cols": ["m", "r", "births_male", "births_female", "deaths_male", "deaths_female", "marriages", "divorces", "pop_total", "pop_male", "pop_female"],
    "rows": rows,
}
out = ROOT / "docs" / "data.js"
out.parent.mkdir(exist_ok=True)
out.write_text("window.BI_DATA = " + json.dumps(data, ensure_ascii=False, separators=(",", ":")) + ";\n", encoding="utf-8")

chk = df[df.year_roc == 114].births_total.sum()
print(f"{len(rows)} 列、{len(months)} 個月、{len(regions)} 個縣市；民國114年出生數合計 {chk}")
print(f"{out}  {out.stat().st_size / 1024:.1f} KB")
