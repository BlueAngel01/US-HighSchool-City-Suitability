# -*- coding: utf-8 -*-
"""
E3 / AP-IB 汇总程序
输入：当前目录下的 E3_CRDC_Eligibility_Audit.csv
输出：E3_10MSA_AP_IB_Summary.xlsx

统计规则：
1. AP、IB 分开统计。
2. Yes / No 为有效回答；-9、-5、空值及其他编码不进入有效回答分母。
3. AP/IB覆盖率 = Yes / (Yes + No)。
4. 同时报告有效回答学校数、有效回答率、未纳入统计数。
5. AP课程数量仅对 AP=Yes 的学校汇总。
6. 学生参与人数汇总 M+F；X 或被抑制人数不擅自补零。
"""

from pathlib import Path
import pandas as pd
import numpy as np

BASE = Path.cwd()
INPUT = BASE / "E3_CRDC_Eligibility_Audit.csv"
OUTPUT = BASE / "E3_10MSA_AP_IB_Summary.xlsx"

if not INPUT.exists():
    raise FileNotFoundError(
        "找不到 E3_CRDC_Eligibility_Audit.csv。"
        "请将本程序放在该文件所在目录，或在该目录运行程序。"
    )

df = pd.read_csv(INPUT, dtype=str, encoding="utf-8-sig", low_memory=False)

required = [
    "MSA_NAME", "CBSA", "NCESSCH", "SCH_NAME", "HS_STUDENT_COUNT",
    "SCH_APENR_IND", "SCH_APCOURSES", "TOT_APENR_M", "TOT_APENR_F",
    "SCH_IBENR_IND", "TOT_IBENR_M", "TOT_IBENR_F"
]
missing = [c for c in required if c not in df.columns]
if missing:
    raise ValueError("审计文件缺少必要字段: " + ", ".join(missing))

# 清理文本编码
for c in ["SCH_APENR_IND", "SCH_IBENR_IND"]:
    df[c] = df[c].astype("string").str.strip()

# 学校层面的有效性标记：只把明确 Yes/No 纳入指标分母
for prefix, col in [("AP", "SCH_APENR_IND"), ("IB", "SCH_IBENR_IND")]:
    df[f"{prefix}_VALID"] = df[col].isin(["Yes", "No"])
    df[f"{prefix}_YES"] = df[col].eq("Yes")
    df[f"{prefix}_NO"] = df[col].eq("No")
    df[f"{prefix}_EXCLUDED"] = ~df[f"{prefix}_VALID"]

# 将数值字段转为数字；负数哨兵值转为缺失，不作为0
numeric_cols = [
    "HS_STUDENT_COUNT", "SCH_APCOURSES",
    "TOT_APENR_M", "TOT_APENR_F",
    "TOT_IBENR_M", "TOT_IBENR_F"
]
for c in numeric_cols:
    df[c + "_NUM"] = pd.to_numeric(df[c], errors="coerce")
    df.loc[df[c + "_NUM"] < 0, c + "_NUM"] = np.nan

# 学生参与人数：只对明确 Yes 的学校统计；M/F 两项均有数值时才计算合计
for prefix in ["AP", "IB"]:
    m = df[f"TOT_{prefix}ENR_M_NUM"]
    f = df[f"TOT_{prefix}ENR_F_NUM"]
    df[f"{prefix}_MF_PARTICIPANTS"] = np.where(
        df[f"{prefix}_YES"] & m.notna() & f.notna(),
        m + f,
        np.nan
    )

# 每个 MSA 汇总
summary_rows = []
for msa, g in df.groupby("MSA_NAME", dropna=False, sort=True):
    row = {
        "MSA_NAME": msa,
        "CBSA": g["CBSA"].dropna().iloc[0] if g["CBSA"].notna().any() else "",
        "V5高中学校数": len(g),
        "V5高中学生数": pd.to_numeric(g["HS_STUDENT_COUNT"], errors="coerce").clip(lower=0).sum()
    }

    for prefix in ["AP", "IB"]:
        valid_n = int(g[f"{prefix}_VALID"].sum())
        yes_n = int(g[f"{prefix}_YES"].sum())
        no_n = int(g[f"{prefix}_NO"].sum())
        excluded_n = int(g[f"{prefix}_EXCLUDED"].sum())

        row[f"{prefix}_Yes学校数"] = yes_n
        row[f"{prefix}_No学校数"] = no_n
        row[f"{prefix}_有效回答学校数"] = valid_n
        row[f"{prefix}_排除学校数(-9/-5/空值等)"] = excluded_n
        row[f"{prefix}_有效回答率"] = valid_n / len(g) if len(g) else np.nan
        # 若没有明确 No，仍按用户固定公式计算；分母为0则留空
        row[f"{prefix}_覆盖率(Yes/(Yes+No))"] = yes_n / valid_n if valid_n else np.nan

        # 课程数仅对 AP=Yes 学校统计
        if prefix == "AP":
            course_values = g.loc[g["AP_YES"], "SCH_APCOURSES_NUM"].dropna()
            row["AP课程数合计(仅Yes学校)"] = course_values.sum() if len(course_values) else np.nan
            row["AP课程数均值(仅Yes学校)"] = course_values.mean() if len(course_values) else np.nan
            row["AP课程数有效学校数"] = int(course_values.size)

        # 学生参与人数：仅汇总 M+F 均有报告的 Yes 学校
        pcol = f"{prefix}_MF_PARTICIPANTS"
        known = g[pcol].dropna()
        row[f"{prefix}_参与学生人数(M+F，已报告)"] = known.sum() if len(known) else np.nan
        row[f"{prefix}_参与人数可计算学校数"] = int(known.size)

        # 分母只取该项目 Yes/No 有效学校的高中生人数
        valid_enrollment = pd.to_numeric(
            g.loc[g[f"{prefix}_VALID"], "HS_STUDENT_COUNT"],
            errors="coerce"
        )
        valid_enrollment = valid_enrollment[valid_enrollment >= 0]
        denom = valid_enrollment.sum()
        row[f"{prefix}_有效学校高中生人数"] = denom
        row[f"{prefix}_参与率(M+F已报告人数/有效学校高中生人数)"] = (
            known.sum() / denom if len(known) and denom > 0 else np.nan
        )

    summary_rows.append(row)

summary = pd.DataFrame(summary_rows)

# 全部学校明细：保留原始编码，另附有效性和计算字段，方便审计
detail_cols = [
    "MSA_NAME", "CBSA", "NCESSCH", "SCH_NAME", "HS_STUDENT_COUNT",
    "SCH_APENR_IND", "AP_VALID", "AP_YES", "AP_NO", "AP_EXCLUDED",
    "SCH_APCOURSES", "TOT_APENR_M", "TOT_APENR_F",
    "AP_MF_PARTICIPANTS",
    "SCH_IBENR_IND", "IB_VALID", "IB_YES", "IB_NO", "IB_EXCLUDED",
    "TOT_IBENR_M", "TOT_IBENR_F", "IB_MF_PARTICIPANTS"
]
detail = df[[c for c in detail_cols if c in df.columns]].copy()

# 编码说明
rules = pd.DataFrame([
    ["Yes", "纳入有效回答；计入Yes学校数及覆盖率分子"],
    ["No", "纳入有效回答；计入No学校数及覆盖率分母"],
    ["-9", "排除，不进入Yes/No分母"],
    ["-5", "排除，不进入Yes/No分母"],
    ["空值/未匹配/其他编码", "排除，不进入Yes/No分母；保留在排除数量中"],
    ["AP课程数量", "只对AP=Yes学校汇总；负数编码不当作0"],
    ["AP/IB学生人数", "只统计Yes学校中M、F均有非负数值的记录；X不擅自补零"]
], columns=["原始值/字段", "本次固定处理规则"])

with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
    summary.to_excel(writer, sheet_name="10MSA汇总", index=False)
    detail.to_excel(writer, sheet_name="学校明细", index=False)
    rules.to_excel(writer, sheet_name="统计口径", index=False)

    # 简单设置列宽与冻结首行
    for ws in writer.book.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col_cells in ws.columns:
            letter = col_cells[0].column_letter
            max_len = max(
                (len(str(cell.value)) if cell.value is not None else 0)
                for cell in col_cells[:300]
            )
            ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 34)

print("Completed.")
print("Input rows:", len(df))
print("MSA count:", df["MSA_NAME"].nunique(dropna=True))
print("Output:", OUTPUT)
print("Note: -9, -5, blank and other non-Yes/No values are excluded from the Yes/No denominator.")
