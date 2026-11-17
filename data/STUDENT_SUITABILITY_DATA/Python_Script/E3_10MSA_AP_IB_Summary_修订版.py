# -*- coding: utf-8 -*-
"""
E3 AP/IB 十个MSA统计程序（修订版）
输入：E3_CRDC_Eligibility_Audit.csv
输出：E3_10MSA_AP_IB_Summary_修订版.xlsx

固定规则：
- Yes、No 是有效回答；-9、-5、空值及其他值排除，不推定为No。
- 不再计算 Yes/(Yes+No) 覆盖率（当前数据No为0，会产生虚假的100%）。
- 另行计算“确认Yes学校占全部高中比例”= Yes学校数 / V5高中学校总数。
- AP/IB参与学生人数只汇总Yes学校中已报告的M+F。
- 参与学生比例 = 已报告参与人数 / MSA全部高中学生数。
- 课程数量只汇总AP=Yes学校的有效数值。
"""

from pathlib import Path
import pandas as pd
import numpy as np

BASE = Path.cwd()
INPUT = BASE / "E3_CRDC_Eligibility_Audit.csv"
OUTPUT = BASE / "E3_10MSA_AP_IB_Summary_修订版.xlsx"

if not INPUT.exists():
    raise FileNotFoundError(
        "找不到 E3_CRDC_Eligibility_Audit.csv。"
        "请将程序放在该CSV所在文件夹运行。"
    )

df = pd.read_csv(INPUT, dtype=str, encoding="utf-8-sig", low_memory=False)

required = [
    "MSA_NAME", "CBSA", "NCESSCH", "SCH_NAME", "HS_STUDENT_COUNT",
    "SCH_APENR_IND", "SCH_APCOURSES", "TOT_APENR_M", "TOT_APENR_F",
    "SCH_IBENR_IND", "TOT_IBENR_M", "TOT_IBENR_F"
]
missing = [c for c in required if c not in df.columns]
if missing:
    raise ValueError("审计文件缺少必要字段：" + ", ".join(missing))

for c in ["SCH_APENR_IND", "SCH_IBENR_IND"]:
    df[c] = df[c].astype("string").str.strip()

for c in [
    "HS_STUDENT_COUNT", "SCH_APCOURSES",
    "TOT_APENR_M", "TOT_APENR_F",
    "TOT_IBENR_M", "TOT_IBENR_F"
]:
    df[c + "_NUM"] = pd.to_numeric(df[c], errors="coerce")
    df.loc[df[c + "_NUM"] < 0, c + "_NUM"] = np.nan

for prefix, indicator in [("AP", "SCH_APENR_IND"), ("IB", "SCH_IBENR_IND")]:
    df[prefix + "_VALID"] = df[indicator].isin(["Yes", "No"])
    df[prefix + "_YES"] = df[indicator].eq("Yes")
    df[prefix + "_NO"] = df[indicator].eq("No")
    df[prefix + "_EXCLUDED"] = ~df[prefix + "_VALID"]

for prefix in ["AP", "IB"]:
    m = df[f"TOT_{prefix}ENR_M_NUM"]
    f = df[f"TOT_{prefix}ENR_F_NUM"]
    df[f"{prefix}_MF_PARTICIPANTS"] = np.where(
        df[f"{prefix}_YES"] & m.notna() & f.notna(),
        m + f,
        np.nan
    )

rows = []
for msa, g in df.groupby("MSA_NAME", dropna=False, sort=True):
    school_n = len(g)
    total_students = g["HS_STUDENT_COUNT_NUM"].sum(min_count=1)

    row = {
        "MSA_NAME": msa,
        "CBSA": g["CBSA"].dropna().iloc[0] if g["CBSA"].notna().any() else "",
        "高中学校数": school_n,
        "高中学生总人数": total_students
    }

    for prefix in ["AP", "IB"]:
        yes_n = int(g[f"{prefix}_YES"].sum())
        no_n = int(g[f"{prefix}_NO"].sum())
        excluded_n = int(g[f"{prefix}_EXCLUDED"].sum())

        row[f"{prefix}确认Yes学校数"] = yes_n
        row[f"{prefix}确认Yes学校占全部高中比例"] = (
            yes_n / school_n if school_n else np.nan
        )
        row[f"{prefix}有效回答Yes学校数"] = yes_n
        row[f"{prefix}有效回答No学校数"] = no_n
        row[f"{prefix}排除记录数"] = excluded_n

        participant = g.loc[g[f"{prefix}_YES"], f"{prefix}_MF_PARTICIPANTS"].dropna()
        participant_total = participant.sum() if len(participant) else np.nan
        row[f"{prefix}参与学生人数_M加F"] = participant_total
        row[f"{prefix}参与人数可计算学校数"] = int(len(participant))
        row[f"{prefix}参与学生人数占MSA高中生比例"] = (
            participant_total / total_students
            if pd.notna(participant_total) and pd.notna(total_students) and total_students > 0
            else np.nan
        )

        if prefix == "AP":
            courses = g.loc[g["AP_YES"], "SCH_APCOURSES_NUM"].dropna()
            row["AP课程数合计"] = courses.sum() if len(courses) else np.nan
            row["AP课程数均值_仅Yes学校"] = courses.mean() if len(courses) else np.nan
            row["AP课程数有效学校数"] = int(len(courses))

    rows.append(row)

summary = pd.DataFrame(rows)

detail_cols = [
    "MSA_NAME", "CBSA", "NCESSCH", "SCH_NAME", "HS_STUDENT_COUNT",
    "SCH_APENR_IND", "AP_VALID", "AP_YES", "AP_NO", "AP_EXCLUDED",
    "SCH_APCOURSES", "TOT_APENR_M", "TOT_APENR_F", "AP_MF_PARTICIPANTS",
    "SCH_IBENR_IND", "IB_VALID", "IB_YES", "IB_NO", "IB_EXCLUDED",
    "TOT_IBENR_M", "TOT_IBENR_F", "IB_MF_PARTICIPANTS"
]
detail = df[[c for c in detail_cols if c in df.columns]].copy()
detail = detail.rename(columns={
    "AP_VALID": "AP是否有效Yes/No回答",
    "AP_YES": "AP是否Yes",
    "AP_NO": "AP是否No",
    "AP_EXCLUDED": "AP是否排除",
    "IB_VALID": "IB是否有效Yes/No回答",
    "IB_YES": "IB是否Yes",
    "IB_NO": "IB是否No",
    "IB_EXCLUDED": "IB是否排除",
    "AP_MF_PARTICIPANTS": "AP参与人数_M+F",
    "IB_MF_PARTICIPANTS": "IB参与人数_M+F"
})

rules = pd.DataFrame([
    ["Yes", "纳入对应项目的确认参与学校数。"],
    ["No", "作为明确未参与记录保留；当前数据中No数量为0。"],
    ["-9、-5、空值及其他编码", "排除，不推定为No，不进入Yes/No有效回答统计。"],
    ["确认Yes学校占全部高中比例", "Yes学校数 ÷ V5对应MSA高中学校总数；只表示已确认开设比例，不称为有效回答覆盖率。"],
    ["AP课程数", "只统计AP=Yes学校中有效的SCH_APCOURSES数值。"],
    ["AP/IB参与学生人数", "只汇总对应项目Yes学校中已报告的男生与女生人数（M+F）；负数哨兵值不作为0。"],
    ["AP/IB参与学生人数占MSA高中生比例", "已报告参与学生人数 ÷ V5该MSA全部高中学生总人数；不推断被排除学校的参与情况。"],
    ["排除记录数", "保留非Yes/No记录数量供审计，不进入Yes/No比例分子或分母。"]
], columns=["项目", "修订后的统计口径"])

with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
    summary.to_excel(writer, sheet_name="10MSA汇总_修订", index=False)
    detail.to_excel(writer, sheet_name="学校明细", index=False)
    rules.to_excel(writer, sheet_name="统计口径", index=False)
    for ws in writer.book.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cells in ws.columns:
            letter = cells[0].column_letter
            width = max((len(str(c.value)) if c.value is not None else 0) for c in cells[:300])
            ws.column_dimensions[letter].width = min(max(width + 2, 12), 38)

print("Completed.")
print("Input rows:", len(df))
print("MSA count:", df["MSA_NAME"].nunique(dropna=True))
print("Output:", OUTPUT)
print("Removed misleading Yes/(Yes+No) coverage rate.")
print("Excluded -9, -5, blank and other non-Yes/No values from Yes/No statistics.")
