# -*- coding: utf-8 -*-
"""
R1 青少年社区活动资源：第一批官方数据下载器
下载：
1) USAspending.gov：21st CCLC（Assistance Listing 84.287）联邦资助记录，FY2024-FY2026
2) U.S. Census：2024 Vintage 各州县级单岁年龄人口 Excel 文件（后续提取14-18岁）

运行：
    py -m pip install requests pandas openpyxl
    py R1_download_data.py

重要：
USAspending记录是联邦资助 award，不等于独立项目/活动地点数量。
Census人口文件是人口分母原始数据，需按MSA县界汇总。
"""

from pathlib import Path
import time
import requests
import pandas as pd

OUT = Path("R1_downloads")
OUT.mkdir(exist_ok=True)

# ---------- A. USAspending：21st CCLC awards ----------
API = "https://api.usaspending.gov/api/v2/search/spending_by_award/"

payload_base = {
    "filters": {
        "program_numbers": ["84.287"],
        "award_type_codes": ["02", "03", "04", "05"],
        "time_period": [{
            "start_date": "2023-10-01",
            "end_date": "2026-09-30"
        }]
    },
    "fields": [
        "Award ID",
        "Recipient Name",
        "Recipient UEI",
        "Awarding Agency",
        "Funding Agency",
        "Start Date",
        "End Date",
        "Award Amount",
        "Description",
        "CFDA Number",
        "Assistance Listings",
        "Place of Performance City Code",
        "Place of Performance State Code",
        "Place of Performance Zip5",
        "Primary Place of Performance",
        "Recipient Location"
    ],
    "limit": 100,
    "page": 1,
    "order": "desc",
    "sort": "Start Date",
    "subawards": False
}

def download_21cclc():
    rows = []
    page = 1
    while True:
        payload = dict(payload_base)
        payload["page"] = page
        print(f"USAspending: downloading page {page} ...")
        r = requests.post(API, json=payload, timeout=90)
        r.raise_for_status()
        data = r.json()

        batch = data.get("results", [])
        rows.extend(batch)
        meta = data.get("page_metadata", {})
        print(f"  received {len(batch)} rows; total={len(rows)}")

        if not batch or not meta.get("hasNext", False):
            break
        page += 1
        time.sleep(0.25)

    if not rows:
        print("WARNING: API returned no awards. Check API response / filter.")
        return

    df = pd.json_normalize(rows)
    csv_path = OUT / "USAspending_21CCLC_84-287_FY2024-2026.csv"
    xlsx_path = OUT / "USAspending_21CCLC_84-287_FY2024-2026.xlsx"
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    df.to_excel(xlsx_path, index=False)
    print(f"Saved: {csv_path}")
    print(f"Saved: {xlsx_path}")
    print("NOTE: Awardee/place-of-performance records are not a census of all local program sites.")

# ---------- B. Census：2024 Vintage county single-year age files ----------
# File names follow Census published state single-year age/sex table pattern.
STATES = {
    "AL": ("01", "Alabama"),
    "AZ": ("04", "Arizona"),
    "CA": ("06", "California"),
    "DE": ("10", "Delaware"),
    "FL": ("12", "Florida"),
    "KS": ("20", "Kansas"),
    "MA": ("25", "Massachusetts"),
    "MD": ("24", "Maryland"),
    "MO": ("29", "Missouri"),
    "NH": ("33", "New Hampshire"),
    "NJ": ("34", "New Jersey"),
    "NY": ("36", "New York"),
    "OH": ("39", "Ohio"),
    "PA": ("42", "Pennsylvania"),
    "TX": ("48", "Texas"),
}

CENSUS_BASE = (
    "https://www2.census.gov/programs-surveys/popest/tables/"
    "2020-2024/state/asrh/"
)

def download_census():
    session = requests.Session()
    for abbr, (fips, name) in STATES.items():
        filename = f"sc-est2024-syasex-{fips}.xlsx"
        url = CENSUS_BASE + filename
        dest = OUT / f"Census_2024_SYASEX_{abbr}_{fips}.xlsx"
        if dest.exists() and dest.stat().st_size > 1000:
            print(f"Census exists, skip: {dest.name}")
            continue
        print(f"Census: {name} ({abbr})")
        try:
            r = session.get(url, timeout=90)
            r.raise_for_status()
            content_type = r.headers.get("Content-Type", "")
            if len(r.content) < 1000 or b"<html" in r.content[:500].lower():
                print(f"  FAILED: response is not an Excel file: {url}")
                continue
            dest.write_bytes(r.content)
            print(f"  Saved: {dest}")
        except Exception as e:
            print(f"  FAILED: {url}\n  {e}")

if __name__ == "__main__":
    print("=== Step 1: 21st CCLC award data ===")
    download_21cclc()
    print("\n=== Step 2: Census single-year age files ===")
    download_census()
    print("\nFinished. Files are in:", OUT.resolve())
