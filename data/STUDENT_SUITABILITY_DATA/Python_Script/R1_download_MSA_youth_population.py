# -*- coding: utf-8 -*-
"""
R1 - 直接下载10个MSA的14-18岁人口分母（Census 2024）
并将用户上传的州级SC-EST2024-SYASEX文件替换为正确的MSA级数据。

安装：
    py -m pip install requests pandas openpyxl

运行：
    py R1_download_MSA_youth_population.py

输出：
    R1_MSA_population_2024/
      cbsa-est2024-syasex.csv
      R1_10MSA_age14_18_2024.xlsx
      R1_10MSA_age14_18_2024.csv

说明：
Census CBSA-EST2024-SYASEX是MSA/CBSA层级的单岁年龄数据。
YEAR=6为2024年7月1日估计；AGE=14..18。
"""

from pathlib import Path
import requests
import pandas as pd

OUT = Path("R1_MSA_population_2024")
OUT.mkdir(exist_ok=True)

URL = (
    "https://www2.census.gov/programs-surveys/popest/"
    "datasets/2020-2024/metro/asrh/cbsa-est2024-syasex.csv"
)

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

def main():
    dest = OUT / "cbsa-est2024-syasex.csv"
    print("Downloading Census CBSA single-year age dataset...")
    r = requests.get(URL, timeout=120)
    r.raise_for_status()
    dest.write_bytes(r.content)
    print("Saved:", dest.resolve())

    df = pd.read_csv(dest, dtype={"CBSA": str, "YEAR": str, "AGE": str})
    df.columns = [c.strip().upper() for c in df.columns]

    required = {"CBSA", "NAME", "YEAR", "AGE"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(
            f"Unexpected Census CSV layout. Missing columns: {missing}. "
            f"Actual columns: {df.columns.tolist()}"
        )

    # Determine total-population column robustly.
    pop_col = next(
        (c for c in ["TOT_POP", "POPESTIMATE", "TOTAL_POPULATION"] if c in df.columns),
        None
    )
    if pop_col is None:
        raise ValueError(f"Cannot identify total population column: {df.columns.tolist()}")

    # Census YEAR=6 is July 1, 2024; single-year AGE 14 through 18 inclusive.
    sub = df[
        df["CBSA"].isin(MSAS.values())
        & (df["YEAR"].astype(str).str.replace(".0", "", regex=False) == "6")
        & (pd.to_numeric(df["AGE"], errors="coerce").between(14, 18))
    ].copy()

    sub[pop_col] = pd.to_numeric(sub[pop_col], errors="coerce")
    sub["AGE"] = pd.to_numeric(sub["AGE"], errors="coerce")

    detail = (
        sub.groupby(["CBSA", "NAME", "AGE"], as_index=False)[pop_col]
        .sum()
        .sort_values(["CBSA", "AGE"])
    )
    summary = (
        detail.groupby(["CBSA", "NAME"], as_index=False)[pop_col]
        .sum()
        .rename(columns={pop_col: "POP_AGE_14_18_2024"})
    )
    reverse = {v: k for k, v in MSAS.items()}
    summary.insert(0, "MSA_REQUESTED", summary["CBSA"].map(reverse))
    summary["YEAR"] = 2024
    summary["SOURCE"] = URL

    # Make sure all ten requested MSA codes were present.
    found = set(summary["CBSA"])
    missing_codes = set(MSAS.values()) - found
    if missing_codes:
        print("WARNING: MSA codes not found:", sorted(missing_codes))

    out_xlsx = OUT / "R1_10MSA_age14_18_2024.xlsx"
    out_csv = OUT / "R1_10MSA_age14_18_2024.csv"
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="MSA汇总", index=False)
        detail.to_excel(writer, sheet_name="年龄明细", index=False)
        pd.DataFrame(
            [{"MSA": k, "CBSA": v} for k, v in MSAS.items()]
        ).to_excel(writer, sheet_name="CBSA代码", index=False)
    summary.to_csv(out_csv, index=False, encoding="utf-8-sig")

    print("\\n10 MSA youth population (ages 14-18):")
    print(summary.to_string(index=False))
    print("\\nSaved:", out_xlsx.resolve())
    print("Saved:", out_csv.resolve())

if __name__ == "__main__":
    main()
