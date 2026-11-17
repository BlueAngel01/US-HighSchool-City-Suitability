# -*- coding: utf-8 -*-
"""
R1 Cultural Facilities - 2023 CBP completion version
Completes Cleveland by aggregating its official county-level CBP records,
and derives Cleveland's 2023 total population from the same five counties.

Main indicators:
  512131 = Motion Picture Theaters (except Drive-Ins)
  711110 = Theater Companies and Dinner Theaters
  712110 = Museums

Supplementary only:
  512132 = Drive-In Motion Picture Theaters
  711310 = Promoters of Performing Arts, Sports, and Similar Events with Facilities

Install:
  py -m pip install pandas openpyxl requests

Run:
  py R1_Cultural_Facilities_CBP_10MSA_补全版.py

Files downloaded automatically:
  cbp23msa.zip
  cbp23co.zip
  cbsa-est2023-alldata.csv
  co-est2023-alldata.csv

Output:
  R1_Cultural_Facilities_2023_10MSA_补全版.xlsx

Important:
- Uses total MSA population, not youth population.
- CBP counts employer establishments, not buildings/screens/seats.
- Cleveland MSA (2022-2023 CBP geography): Cuyahoga, Geauga, Lake,
  Lorain, Medina counties. The program aggregates county CBP establishments
  for Cleveland rather than treating its missing MSA row as zero.
- If any Cleveland county value is suppressed/unpublished for an industry,
  the Cleveland aggregate is marked incomplete and is NOT fabricated.
"""

from pathlib import Path
import zipfile
import requests
import pandas as pd

YEAR = 2023
BASE = "https://www2.census.gov/programs-surveys"
CBP_MSA_URL = f"{BASE}/cbp/datasets/2023/cbp23msa.zip"
CBP_CO_URL = f"{BASE}/cbp/datasets/2023/cbp23co.zip"
POP_MSA_URL = f"{BASE}/popest/datasets/2020-2023/metro/totals/cbsa-est2023-alldata.csv"
POP_CO_URL = f"{BASE}/popest/datasets/2020-2023/counties/totals/co-est2023-alldata.csv"

FILES = {
    "cbp_msa": Path("cbp23msa.zip"),
    "cbp_co": Path("cbp23co.zip"),
    "pop_msa": Path("cbsa-est2023-alldata.csv"),
    "pop_co": Path("co-est2023-alldata.csv"),
}
OUT = Path("R1_Cultural_Facilities_2023_10MSA_补全版.xlsx")

MSAS = {
    "New York": "35620",
    "Philadelphia": "37980",
    "Boston": "14460",
    "Cleveland": "17460",
    "Kansas City": "28140",
    "Jacksonville": "27260",
    "Birmingham": "13820",
    "Houston": "26420",
    "Phoenix": "38060",
    "San Diego": "41740",
}

# Official CBP 2022-2023 MSA county geography for Cleveland-Elyria, OH (17460).
# County FIPS = state FIPS 39 + county FIPS.
CLEVELAND_COUNTIES = {
    "39035": "Cuyahoga County, OH",
    "39055": "Geauga County, OH",
    "39085": "Lake County, OH",
    "39093": "Lorain County, OH",
    "39103": "Medina County, OH",
}

INDUSTRIES = {
    "512131": ("电影院（普通电影院）", "Motion Picture Theaters (except Drive-Ins)"),
    "512132": ("汽车影院（补充）", "Drive-In Motion Picture Theaters"),
    "711110": ("剧院公司及晚餐剧场", "Theater Companies and Dinner Theaters"),
    "711310": ("有场地的表演艺术/体育/类似活动主办机构（补充宽口径）",
               "Promoters of Performing Arts, Sports, and Similar Events with Facilities"),
    "712110": ("博物馆", "Museums"),
}
MAIN_CODES = ["512131", "711110", "712110"]


def download(url, path):
    if path.exists() and path.stat().st_size > 1000:
        print(f"Using cached: {path}")
        return
    print(f"Downloading: {url}")
    r = requests.get(url, timeout=240, headers={"User-Agent": "Mozilla/5.0"})
    r.raise_for_status()
    path.write_bytes(r.content)
    print(f"Saved: {path} ({len(r.content):,} bytes)")


def read_zip_table(zip_path, preferred):
    with zipfile.ZipFile(zip_path) as z:
        names = [n for n in z.namelist()
                 if n.lower().endswith((".txt", ".csv")) and not n.endswith("/")]
        if not names:
            raise RuntimeError(f"No TXT/CSV found in {zip_path}: {z.namelist()}")
        name = next((n for n in names if preferred.lower() in Path(n).name.lower()), names[0])
        print("Reading:", name)
        with z.open(name) as f:
            sample = f.read(8192).decode("utf-8-sig", errors="replace")
            f.seek(0)
            first = sample.splitlines()[0] if sample.splitlines() else ""
            sep = max([",", "\t", "|"], key=lambda x: first.count(x))
            if first.count(sep) == 0:
                raise RuntimeError(f"Could not detect delimiter in {name}: {first[:250]}")
            d = pd.read_csv(f, sep=sep, dtype=str, low_memory=False, encoding="utf-8-sig")
    d.columns = [str(c).strip().upper() for c in d.columns]
    return d


def read_population_csv(path):
    encodings = ["utf-8-sig", "cp1252", "latin1"]
    last = None
    for enc in encodings:
        try:
            d = pd.read_csv(path, dtype=str, low_memory=False, encoding=enc)
            d.columns = [str(c).strip().upper() for c in d.columns]
            return d
        except UnicodeDecodeError as e:
            last = e
    raise last


def find_col(df, candidates, required=True):
    for c in candidates:
        if c.upper() in df.columns:
            return c.upper()
    if required:
        raise RuntimeError(f"Missing columns {candidates}; available={list(df.columns)}")
    return None


def normalize_code(s, width=5):
    return s.astype(str).str.replace(r"\.0$", "", regex=True).str.strip().str.zfill(width)


def load_cbp_msa():
    download(CBP_MSA_URL, FILES["cbp_msa"])
    d = read_zip_table(FILES["cbp_msa"], "msa")
    geo = find_col(d, ["MSA", "CBSA"])
    naics = find_col(d, ["NAICS", "NAICS2017"])
    est = find_col(d, ["EST", "ESTAB"])
    d["GEO"] = normalize_code(d[geo])
    d["IND"] = d[naics].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    d["EST_NUM"] = pd.to_numeric(d[est], errors="coerce")
    d["EST_RAW"] = d[est].astype(str).str.strip()
    return d


def load_cbp_county():
    download(CBP_CO_URL, FILES["cbp_co"])
    d = read_zip_table(FILES["cbp_co"], "co")
    state = find_col(d, ["FIPSTATE", "STATE"])
    county = find_col(d, ["FIPSCTY", "COUNTY"])
    naics = find_col(d, ["NAICS", "NAICS2017"])
    est = find_col(d, ["EST", "ESTAB"])
    d["COUNTY_FIPS"] = normalize_code(d[state], 2) + normalize_code(d[county], 3)
    d["IND"] = d[naics].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    d["EST_NUM"] = pd.to_numeric(d[est], errors="coerce")
    d["EST_RAW"] = d[est].astype(str).str.strip()
    return d


def load_population_msa():
    download(POP_MSA_URL, FILES["pop_msa"])
    d = read_population_csv(FILES["pop_msa"])
    cbsa = find_col(d, ["CBSA"])
    pop = find_col(d, ["POPESTIMATE2023"])
    d["GEO"] = normalize_code(d[cbsa])
    d["POP2023"] = pd.to_numeric(d[pop], errors="coerce")
    return d


def load_population_county():
    download(POP_CO_URL, FILES["pop_co"])
    d = read_population_csv(FILES["pop_co"])
    state = find_col(d, ["STATE"])
    county = find_col(d, ["COUNTY"])
    pop = find_col(d, ["POPESTIMATE2023"])
    d["COUNTY_FIPS"] = normalize_code(d[state], 2) + normalize_code(d[county], 3)
    d["POP2023"] = pd.to_numeric(d[pop], errors="coerce")
    return d


def get_cleveland_county_sum(county_cbp, county_pop, code):
    rows = []
    for naics, (cn, en) in INDUSTRIES.items():
        sub = county_cbp[
            county_cbp["COUNTY_FIPS"].isin(CLEVELAND_COUNTIES)
            & (county_cbp["IND"] == naics)
        ].copy()
        found = {r["COUNTY_FIPS"]: r for _, r in sub.iterrows()}
        vals = []
        status_parts = []
        for fips, county_name in CLEVELAND_COUNTIES.items():
            if fips not in found:
                vals.append(None)
                status_parts.append(f"{county_name}: no row")
            else:
                raw = found[fips]["EST_RAW"]
                val = found[fips]["EST_NUM"]
                if pd.isna(val):
                    vals.append(None)
                    status_parts.append(f"{county_name}: suppressed/not numeric ({raw})")
                else:
                    vals.append(float(val))
                    status_parts.append(f"{county_name}: {int(val)}")
        if all(v is not None for v in vals):
            total = sum(vals)
            status = "county_sum_complete"
        else:
            total = pd.NA
            status = "county_sum_incomplete; not treated as zero; " + "; ".join(status_parts)
        rows.append({
            "MSA": "Cleveland", "CBSA": code, "NAICS": naics,
            "文化设施类别": cn, "CBP经营场所数量": total,
            "Cleveland县级汇总状态": status,
            "县级明细": " | ".join(status_parts)
        })
    cp = county_pop[county_pop["COUNTY_FIPS"].isin(CLEVELAND_COUNTIES)].copy()
    pop_found = {r["COUNTY_FIPS"]: r["POP2023"] for _, r in cp.iterrows()}
    pop_vals = [pop_found.get(fips) for fips in CLEVELAND_COUNTIES]
    if len(cp) == len(CLEVELAND_COUNTIES) and all(pd.notna(v) for v in pop_vals):
        population = sum(pop_vals)
        pop_status = "sum_of_5_county_2023_population"
    else:
        population = pd.NA
        pop_status = "county population incomplete"
    return rows, population, pop_status


def main():
    cbp_msa = load_cbp_msa()
    cbp_county = load_cbp_county()
    pop_msa = load_population_msa()
    pop_county = load_population_county()

    # Main MSA published records.
    cbp_msa = cbp_msa[
        cbp_msa["GEO"].isin(MSAS.values())
        & cbp_msa["IND"].isin(INDUSTRIES)
    ].copy()
    pop_map = {
        r["GEO"]: r["POP2023"]
        for _, r in pop_msa.drop_duplicates("GEO").iterrows()
    }

    # Cleveland fallback: sum county records and county total population.
    cleveland_rows, cleveland_population, cleveland_pop_status = get_cleveland_county_sum(
        cbp_county, pop_county, "17460"
    )
    pop_map["17460"] = cleveland_population

    rows = []
    for msa, cbsa in MSAS.items():
        population = pop_map.get(cbsa, pd.NA)
        for naics, (cn, en) in INDUSTRIES.items():
            if msa == "Cleveland":
                item = next(x for x in cleveland_rows if x["NAICS"] == naics)
                count = item["CBP经营场所数量"]
                status = item["Cleveland县级汇总状态"]
                detail = item["县级明细"]
                source_geo = "2023 CBP county sum: 5 counties"
            else:
                hit = cbp_msa[(cbp_msa["GEO"] == cbsa) & (cbp_msa["IND"] == naics)]
                if len(hit):
                    count = hit.iloc[0]["EST_NUM"]
                    raw = hit.iloc[0]["EST_RAW"]
                    status = "reported" if pd.notna(count) else f"suppressed_or_not_numeric ({raw})"
                else:
                    count = pd.NA
                    status = "no MSA row; not zero"
                detail = ""
                source_geo = "2023 CBP MSA file"
            rate = (
                float(count) / float(population) * 10000
                if pd.notna(count) and pd.notna(population) and float(population) > 0
                else pd.NA
            )
            rows.append({
                "MSA": msa, "CBSA": cbsa, "2023总人口": population,
                "NAICS": naics, "文化设施类别": cn,
                "CBP经营场所数量": count,
                "每万人口经营场所数": rate,
                "数据状态": status, "数据地理来源": source_geo,
                "补充说明": detail if msa == "Cleveland" else ""
            })

    long = pd.DataFrame(rows)
    main_long = long[long["NAICS"].isin(MAIN_CODES)].copy()
    counts = main_long.pivot(index=["MSA", "CBSA", "2023总人口"],
                             columns="文化设施类别", values="CBP经营场所数量").reset_index()
    rates = main_long.pivot(index=["MSA", "CBSA", "2023总人口"],
                            columns="文化设施类别", values="每万人口经营场所数").reset_index()

    # Cleveland county audit sheet for all five counties and selected industries.
    cleveland_audit = cbp_county[
        cbp_county["COUNTY_FIPS"].isin(CLEVELAND_COUNTIES)
        & cbp_county["IND"].isin(INDUSTRIES)
    ].copy()
    cleveland_audit["County name"] = cleveland_audit["COUNTY_FIPS"].map(CLEVELAND_COUNTIES)
    cleveland_audit = cleveland_audit[
        ["COUNTY_FIPS", "County name", "IND", "EST_RAW", "EST_NUM"]
    ].sort_values(["IND", "COUNTY_FIPS"])

    pop_audit = pop_county[pop_county["COUNTY_FIPS"].isin(CLEVELAND_COUNTIES)].copy()
    pop_audit["County name"] = pop_audit["COUNTY_FIPS"].map(CLEVELAND_COUNTIES)
    pop_audit = pop_audit[["COUNTY_FIPS", "County name", "POP2023"]]

    method = pd.DataFrame([
        ["统计年份", "2023"],
        ["人口分母", "MSA全部常住人口；Cleveland由其5个县2023人口相加；不使用青少年人口。"],
        ["标准化公式", "每万人口经营场所数 = 经营场所数量 / MSA总人口 × 10,000。"],
        ["CBP统计单位", "有雇员的经营场所（establishments），不是建筑物、银幕数、座位数或全部公共设施。"],
        ["主指标", "512131普通电影院；711110剧院公司及晚餐剧场；712110博物馆。"],
        ["补充指标", "512132汽车影院；711310有场地的表演艺术/体育/类似活动主办机构，分别列示，不混入主指标。"],
        ["Cleveland处理", "原MSA文件缺少17460记录；改用CBP县级文件，对Cuyahoga、Geauga、Lake、Lorain、Medina五县按NAICS求和。"],
        ["Cleveland人口", "同五县2023年7月1日总人口相加，保持地理边界一致。"],
        ["缺失处理", "县级任一组成县行业值缺失/抑制时，不计算完整县级合计，不以0替代。其他MSA的未发布值也保留缺失。"],
        ["官方CBP数据页", "https://www.census.gov/data/datasets/2023/econ/cbp/2023-cbp.html"],
        ["CBP MSA地理参考", "https://www.census.gov/programs-surveys/cbp/technical-documentation/reference/metro-area-geography-reference.html"],
        ["Cleveland MSA边界", "Cuyahoga、Geauga、Lake、Lorain、Medina counties（CBSA 17460）。"]
    ], columns=["项目", "统计口径与来源"])

    with pd.ExcelWriter(OUT, engine="openpyxl") as writer:
        long.to_excel(writer, sheet_name="10MSA完整长表", index=False)
        counts.to_excel(writer, sheet_name="三项主指标数量", index=False)
        rates.to_excel(writer, sheet_name="三项主指标每万人口率", index=False)
        cleveland_audit.to_excel(writer, sheet_name="Cleveland县级CBP核查", index=False)
        pop_audit.to_excel(writer, sheet_name="Cleveland县级人口核查", index=False)
        method.to_excel(writer, sheet_name="论文口径与来源", index=False)

        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="34495E")
                cell.alignment = Alignment(wrap_text=True, vertical="center")
            for col in ws.columns:
                letter = get_column_letter(col[0].column)
                width = min(max(max(len(str(c.value or "")) for c in list(col)[:100]) + 2, 12), 60)
                ws.column_dimensions[letter].width = width

    print("\n=== Main indicators, all-population standardized ===")
    print(main_long[["MSA", "2023总人口", "NAICS", "CBP经营场所数量",
                     "每万人口经营场所数", "数据状态"]].to_string(index=False))
    print("\nCleveland population status:", cleveland_pop_status)
    print("\nSaved:", OUT.resolve())
    print("If Cleveland still has missing values, inspect the county audit sheet; do not fill suppressed cells with zero.")


if __name__ == "__main__":
    main()
