# -*- coding: utf-8 -*-
"""
2023 Census County Business Patterns (CBP): cultural facilities in 10 MSAs.

Downloads the official MSA-level CBP CSV ZIP (no Census API key required),
filters the specified 2017 NAICS industries, and exports an auditable Excel.

Install:
    py -m pip install pandas openpyxl requests

Run:
    py R1_Cultural_Facilities_CBP_10MSA.py

Output:
    CBP_2023_10MSA_Cultural_Facilities.xlsx
    cbp23msa.zip (cached source file)

Important:
- CBP counts business establishments with paid employees, not every physical
  building, public institution, nonprofit venue, or individual screen.
- NAICS 711110 = Theater Companies and Dinner Theaters.
- NAICS 711310 = Promoters of Performing Arts, Sports, and Similar Events
  with Facilities. This is broader than theaters alone; reported separately.
- NAICS 712110 = Museums.
- 512131 and 512132 are movie theaters and drive-in theaters.
- Suppressed values are not converted to zero.
"""

from pathlib import Path
import zipfile
import requests
import pandas as pd

YEAR = 2023
ZIP_URL = "https://www2.census.gov/programs-surveys/cbp/datasets/2023/cbp23msa.zip"
ZIP_FILE = Path("cbp23msa.zip")
OUT_FILE = Path("CBP_2023_10MSA_Cultural_Facilities.xlsx")

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

INDUSTRIES = {
    "512131": ("电影院（不含汽车影院）", "Motion Picture Theaters (except Drive-Ins)"),
    "512132": ("汽车影院", "Drive-In Motion Picture Theaters"),
    "711110": ("剧院公司及晚餐剧场", "Theater Companies and Dinner Theaters"),
    "711310": ("有场地的表演艺术活动主办机构（较宽口径）",
               "Promoters of Performing Arts, Sports, and Similar Events with Facilities"),
    "712110": ("博物馆", "Museums"),
}

def download_source():
    if ZIP_FILE.exists() and ZIP_FILE.stat().st_size > 100_000:
        print(f"Using cached source: {ZIP_FILE.resolve()}")
        return
    print("Downloading official Census CBP MSA ZIP...")
    r = requests.get(ZIP_URL, timeout=180, headers={"User-Agent":"Mozilla/5.0"})
    r.raise_for_status()
    ZIP_FILE.write_bytes(r.content)
    if not zipfile.is_zipfile(ZIP_FILE):
        ZIP_FILE.unlink(missing_ok=True)
        raise RuntimeError("Downloaded file is not a ZIP. Check the official Census URL.")
    print(f"Downloaded: {ZIP_FILE.resolve()} ({ZIP_FILE.stat().st_size:,} bytes)")

def read_msa_csv():
    download_source()
    with zipfile.ZipFile(ZIP_FILE) as z:
        # Census distributes this file as cbp23msa.txt, not .csv.
        data_files = [
            n for n in z.namelist()
            if n.lower().endswith((".csv", ".txt"))
            and not n.endswith("/")
        ]
        if not data_files:
            raise RuntimeError("No CSV/TXT data file found in CBP MSA ZIP: " + repr(z.namelist()))

        # Prefer the MSA file; the official 2023 archive normally contains cbp23msa.txt.
        target = next((n for n in data_files if "msa" in Path(n).name.lower()), data_files[0])
        print("Reading:", target)
        with z.open(target) as f:
            # CBP .txt is a delimited text file. Detect comma, tab, or pipe
            # from the header instead of assuming the extension indicates format.
            sample = f.read(8192).decode("utf-8-sig", errors="replace")
            f.seek(0)
            first_line = sample.splitlines()[0] if sample.splitlines() else ""
            candidates = [",", "\\t", "|"]
            delimiter = max(candidates, key=lambda d: first_line.count(d))
            if first_line.count(delimiter) == 0:
                raise RuntimeError(
                    "Could not detect delimiter in " + target +
                    ". First line: " + first_line[:300]
                )
            df = pd.read_csv(
                f, sep=delimiter, dtype=str, low_memory=False,
                encoding="utf-8-sig"
            )
    df.columns = [c.strip().upper() for c in df.columns]
    print("Columns detected:", ", ".join(df.columns[:20]))
    return df

def find_column(df, names, required=True):
    for name in names:
        if name.upper() in df.columns:
            return name.upper()
    if required:
        raise RuntimeError(f"Missing columns {names}; actual columns: {list(df.columns)}")
    return None

def main():
    df = read_msa_csv()

    cbsa_col = find_column(df, ["MSA", "CBSA"])
    naics_col = find_column(df, ["NAICS2017", "NAICS"])
    estab_col = find_column(df, ["EST", "ESTAB"])
    name_col = find_column(df, ["NAME"], required=False)
    label_col = find_column(df, ["NAICS2017_LABEL", "NAICS_LABEL"], required=False)
    flag_col = find_column(df, ["EST_F", "ESTAB_F"], required=False)

    df["CBSA_CODE"] = df[cbsa_col].astype(str).str.replace(r"\.0$", "", regex=True).str.zfill(5)
    df["NAICS_CODE"] = df[naics_col].astype(str).str.replace(r"\.0$", "", regex=True)
    df = df[df["CBSA_CODE"].isin(MSAS.values()) & df["NAICS_CODE"].isin(INDUSTRIES)].copy()

    # CBP MSA file uses MSA as the 5-digit metropolitan area code and EST as establishment count.
    # Retain the record-level flags; never coerce suppressed values to zero.
    df["ESTAB_NUMERIC"] = pd.to_numeric(df[estab_col], errors="coerce")
    df["ESTAB_STATUS"] = "reported"
    if flag_col:
        df.loc[df[flag_col].fillna("").ne(""), "ESTAB_STATUS"] = (
            "flagged: " + df.loc[df[flag_col].fillna("").ne(""), flag_col].astype(str)
        )
    df.loc[df[estab_col].isna() | df[estab_col].astype(str).str.strip().isin(["", "N", "S"]), "ESTAB_STATUS"] = "suppressed_or_not_reported"

    reverse = {v:k for k,v in MSAS.items()}
    df["MSA"] = df["CBSA_CODE"].map(reverse)
    df["NAICS_NAME_EN"] = df["NAICS_CODE"].map(lambda x: INDUSTRIES[x][1])
    df["指标类别（中文）"] = df["NAICS_CODE"].map(lambda x: INDUSTRIES[x][0])
    if name_col:
        df["Census geography name"] = df[name_col]
    if label_col:
        df["Census NAICS label"] = df[label_col]

    detail_cols = ["MSA","CBSA_CODE","NAICS_CODE","指标类别（中文）","NAICS_NAME_EN",
                   estab_col,"ESTAB_NUMERIC","ESTAB_STATUS"]
    if flag_col:
        detail_cols.append(flag_col)
    if name_col:
        detail_cols.append(name_col)
    if label_col:
        detail_cols.append(label_col)
    detail = df[detail_cols].sort_values(["MSA","NAICS_CODE"])

    # Wide output: each category separate; blank means no publishable value,
    # not zero. Cinema total adds regular + drive-in only if both are numeric.
    wide = []
    for msa, cbsa in MSAS.items():
        rec = {"MSA":msa,"CBSA":cbsa}
        sub = df[df["CBSA_CODE"] == cbsa]
        for code, (cn, en) in INDUSTRIES.items():
            row = sub[sub["NAICS_CODE"] == code]
            rec[f"{code}_{cn}"] = row.iloc[0]["ESTAB_NUMERIC"] if len(row) else pd.NA
            rec[f"{code}_flag/status"] = row.iloc[0]["ESTAB_STATUS"] if len(row) else "record_not_found"
        c1 = rec.get("512131_电影院（不含汽车影院）")
        c2 = rec.get("512132_汽车影院")
        rec["电影院合计（两项均有数值才计算）"] = (
            c1 + c2 if pd.notna(c1) and pd.notna(c2) else pd.NA
        )
        t1 = rec.get("711110_剧院公司及晚餐剧场")
        t2 = rec.get("711310_有场地的表演艺术活动主办机构（较宽口径）")
        rec["剧院相关机构合计（两项均有数值；非纯物理剧场数）"] = (
            t1 + t2 if pd.notna(t1) and pd.notna(t2) else pd.NA
        )
        wide.append(rec)
    summary = pd.DataFrame(wide)

    methodology = pd.DataFrame([
        ["数据集","U.S. Census Bureau, 2023 County Business Patterns (CBP), MSA file."],
        ["下载方式","官方CBP MSA ZIP；本程序直接下载，无需Census API key。"],
        ["统计单位","Establishments（有雇员的经营场所/机构），不是银幕数、座位数或每一栋建筑数。"],
        ["覆盖范围","CBP覆盖有雇员的企业；无雇员机构、部分公共部门直接运营设施可能不在统计范围内。"],
        ["电影院","NAICS 512131（普通电影院）与512132（汽车影院）分列；两项都有数值时才合计。"],
        ["剧院","NAICS 711110为剧院公司及晚餐剧场；711310为有场地的表演艺术/体育/类似活动主办机构，范围较宽，必须分列，不能把合计解释为纯剧场建筑数。"],
        ["博物馆","NAICS 712110 Museums。"],
        ["缺失与抑制","空值、抑制值或缺少记录均不当作0；保留状态字段。"],
        ["年份","2023。所有类别采用同一年份，确保横向可比。"],
        ["官方数据页", "https://www.census.gov/data/datasets/2023/econ/cbp/2023-cbp.html"],
        ["官方API变量说明", "https://api.census.gov/data/2023/cbp/variables.html"],
    ], columns=["项目","说明"])

    with pd.ExcelWriter(OUT_FILE, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="10MSA汇总", index=False)
        detail.to_excel(writer, sheet_name="NAICS原始记录", index=False)
        methodology.to_excel(writer, sheet_name="数据口径与限制", index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            from openpyxl.styles import Font, PatternFill, Alignment
            from openpyxl.utils import get_column_letter
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="34495E")
                cell.alignment = Alignment(wrap_text=True, vertical="center")
            for col in ws.columns:
                letter = get_column_letter(col[0].column)
                width = min(max(max((len(str(c.value or "")) for c in list(col)[:100]), default=10)+2, 12), 48)
                ws.column_dimensions[letter].width = width

    print("\n=== 2023 CBP Cultural Facilities: 10 MSA ===")
    print(summary[["MSA","CBSA","电影院合计（两项均有数值才计算）",
                   "711110_剧院公司及晚餐剧场",
                   "711310_有场地的表演艺术活动主办机构（较宽口径）",
                   "712110_博物馆"]].to_string(index=False))
    print("\nSaved:", OUT_FILE.resolve())
    print("Note: blank/suppressed values are not zero.")

if __name__ == "__main__":
    main()
