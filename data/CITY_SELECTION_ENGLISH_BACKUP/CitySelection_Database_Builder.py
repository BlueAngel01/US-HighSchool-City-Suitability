# -*- coding: utf-8 -*-
"""
美国九大地区 2024 MSA 五指标与产业结构数据库合并程序
输出：
  1) 美国九大地区_2024MSA综合数据库.xlsx
  2) 美国九大地区_2024MSA综合数据库.sqlite

用法：
  py ".\美国九大地区_2024MSA数据库合并.py"

将本程序与九个地区的“五指标_完整.xlsx”放在同一文件夹。
西南中部工作簿若损坏/无法读取，程序会自动使用该地区的三个CSV文件回退：
  西南中部_2024五指标_25MSA.csv
  西南中部_2024产业结构_13类_325行.csv
  西南中部_2024GDP人均GDP.csv
"""

from pathlib import Path
import sqlite3
import math
import sys
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from zipfile import BadZipFile

BASE = Path(__file__).resolve().parent
OUT_XLSX = BASE / "美国九大地区_2024MSA综合数据库.xlsx"
OUT_SQLITE = BASE / "美国九大地区_2024MSA综合数据库.sqlite"

# 固定顺序：美国人口普查局九大 Census Divisions
REGIONS = [
    ("1", "新英格兰 New England", "新英格兰_11MSA五指标_完整.xlsx", 11),
    ("2", "中大西洋 Middle Atlantic", "中大西洋_18MSA五指标_完整.xlsx", 18),
    ("3", "东北中部 East North Central", "东北中部_24MSA五指标_完整.xlsx", 24),
    ("4", "西北中部 West North Central", "西北中部_13MSA五指标_完整.xlsx", 13),
    ("5", "南大西洋 South Atlantic", "南大西洋_42MSA五指标_完整.xlsx", 42),
    ("6", "东南中部 East South Central", "东南中部_16MSA五指标_完整.xlsx", 16),
    ("7", "西南中部 West South Central", "西南中部_25MSA五指标_完整.xlsx", 25),
    ("8", "山地 Mountain", "山地地区_15MSA五指标_完整.xlsx", 15),
    ("9", "太平洋 Pacific", "太平洋地区_31MSA五指标_完整.xlsx", 31),
]

INDUSTRY_NAMES = [
    "农业、林业、渔业与采矿", "建筑业", "制造业", "批发贸易",
    "零售贸易", "运输、仓储与公用事业", "信息业", "金融与保险",
    "房地产与租赁", "专业、科学、管理及行政服务",
    "教育、医疗与社会服务", "艺术、娱乐、住宿与餐饮",
    "其他服务及公共管理",
]

def read_region_xlsx(path):
    """读取地区完整工作簿，按既定模板的两张工作表读取。"""
    wb = load_workbook(path, data_only=True, read_only=True)
    if "五指标总表" not in wb.sheetnames or "产业结构明细13类" not in wb.sheetnames:
        raise ValueError(f"{path.name} 缺少指定工作表")
    ws1, ws2 = wb["五指标总表"], wb["产业结构明细13类"]
    five_rows = list(ws1.values)
    ind_rows = list(ws2.values)
    if not five_rows or not ind_rows:
        raise ValueError(f"{path.name} 工作表为空")
    five = pd.DataFrame(five_rows[1:], columns=five_rows[0])
    industry = pd.DataFrame(ind_rows[1:], columns=ind_rows[0])
    wb.close()
    return normalize_workbook_tables(five, industry)

def normalize_workbook_tables(five, industry):
    """统一不同版本工作簿的列名为标准数据库字段。"""
    # 按模板列位置映射，兼容太平洋地区早期版本的表头文字差异。
    if five.shape[1] < 8 or industry.shape[1] < 10:
        raise ValueError("工作表列数不足，无法按统一模板读取")
    five = five.iloc[:, :8].copy()
    five.columns = [
        "原始清单序号", "MSA", "CBSA", "2024人口",
        "2024人口密度_人每平方英里", "2024人均GDP_美元",
        "产业结构摘要", "2020至2024_CAGR_百分比"
    ]
    industry = industry.iloc[:, :10].copy()
    industry.columns = [
        "原始清单序号", "MSA", "CBSA", "行业",
        "男性就业", "女性就业", "就业合计",
        "就业占比_百分比", "ACS男性变量", "ACS女性变量"
    ]
    return five, industry

def read_west_south_central_csv():
    """西南中部工作簿存在损坏时，用已生成的原始CSV及GDP补充表重建。"""
    f_path = BASE / "西南中部_2024五指标_25MSA.csv"
    i_path = BASE / "西南中部_2024产业结构_13类_325行.csv"
    g_path = BASE / "西南中部_2024GDP人均GDP.csv"
    for p in (f_path, i_path, g_path):
        if not p.exists():
            raise FileNotFoundError(
                f"西南中部工作簿不可读取，且缺少回退文件：{p.name}"
            )
    f = pd.read_csv(f_path, encoding="utf-8-sig")
    i = pd.read_csv(i_path, encoding="utf-8-sig")
    g = pd.read_csv(g_path, encoding="utf-8-sig")
    f = f.merge(g[["CBSA", "GDP_PC_2024"]], on="CBSA", how="left", validate="one_to_one")
    # 若主五指标CSV中的GDP为空，则以GDP补充CSV为准。
    if "2024人均GDP_美元" in f.columns:
        f["2024人均GDP_美元"] = f["2024人均GDP_美元"].fillna(f["GDP_PC_2024"])
    else:
        f["2024人均GDP_美元"] = f["GDP_PC_2024"]
    five = f[[
        "原始清单序号", "MSA", "CBSA", "2024人口",
        "2024人口密度_人每平方英里", "2024人均GDP_美元",
        "2020至2024_CAGR_百分比"
    ]].copy()
    five["产业结构摘要"] = ""
    industry = i.rename(columns={
        "原始清单序号": "原始清单序号",
        "产业类别": "行业",
        "就业合计": "就业合计",
        "就业占比_百分比": "就业占比_百分比",
    })[[
        "原始清单序号", "MSA", "CBSA", "行业",
        "男性就业", "女性就业", "就业合计",
        "就业占比_百分比", "ACS男性变量", "ACS女性变量"
    ]].copy()
    return five, industry

def as_number(series):
    return pd.to_numeric(series, errors="coerce")

def main():
    all_five, all_industry, audit = [], [], []
    for region_no, region_name, filename, expected_n in REGIONS:
        path = BASE / filename
        try:
            if region_no == "7":
                # 即使西南中部xlsx存在，也优先用CSV回退源以规避已知损坏的xlsx。
                try:
                    five, industry = read_region_xlsx(path)
                    audit_source = filename
                except Exception as exc:
                    print(f"提示：西南中部工作簿读取失败（{exc}），改用CSV重建。")
                    five, industry = read_west_south_central_csv()
                    audit_source = "西南中部三份CSV回退"
            else:
                if not path.exists():
                    raise FileNotFoundError(f"缺少文件：{filename}")
                five, industry = read_region_xlsx(path)
                audit_source = filename

            five = five.copy()
            industry = industry.copy()
            five.insert(0, "地区编号", region_no)
            five.insert(1, "地区", region_name)
            industry.insert(0, "地区编号", region_no)
            industry.insert(1, "地区", region_name)

            if len(five) != expected_n:
                raise ValueError(f"{region_name}五指标城市数应为{expected_n}，实际{len(five)}")
            if len(industry) != expected_n * 13:
                raise ValueError(f"{region_name}产业记录应为{expected_n*13}，实际{len(industry)}")

            all_five.append(five)
            all_industry.append(industry)
            audit.append({
                "地区编号": region_no, "地区": region_name,
                "预期城市数": expected_n, "实际城市数": len(five),
                "预期产业记录": expected_n*13, "实际产业记录": len(industry),
                "读取来源": audit_source, "状态": "通过"
            })
            print(f"OK {region_name}: {len(five)}城，{len(industry)}条产业记录")
        except Exception as exc:
            audit.append({
                "地区编号": region_no, "地区": region_name,
                "预期城市数": expected_n, "实际城市数": 0,
                "预期产业记录": expected_n*13, "实际产业记录": 0,
                "读取来源": filename, "状态": f"失败：{exc}"
            })
            print(f"ERROR {region_name}: {exc}")

    if len(all_five) != 9:
        print("\n未能完整读取九个地区。请检查上方ERROR信息；未生成正式数据库。")
        sys.exit(2)

    five = pd.concat(all_five, ignore_index=True)
    industry = pd.concat(all_industry, ignore_index=True)
    audit_df = pd.DataFrame(audit)

    # Numeric normalization
    numeric_five = ["原始清单序号", "CBSA", "2024人口",
                    "2024人口密度_人每平方英里", "2024人均GDP_美元",
                    "2020至2024_CAGR_百分比"]
    for col in numeric_five:
        five[col] = as_number(five[col])
    numeric_ind = ["原始清单序号", "CBSA", "男性就业", "女性就业",
                   "就业合计", "就业占比_百分比"]
    for col in numeric_ind:
        industry[col] = as_number(industry[col])

    # Hard validation: no duplicate CBSA, no missing core data, 13 sectors per MSA.
    if five["CBSA"].isna().any() or five["CBSA"].duplicated().any():
        raise ValueError("CBSA为空或重复，数据库拒绝生成")
    if five[["MSA", "2024人口", "2024人口密度_人每平方英里",
             "2024人均GDP_美元", "2020至2024_CAGR_百分比"]].isna().any().any():
        missing = five[five[["MSA", "2024人口", "2024人口密度_人每平方英里",
                             "2024人均GDP_美元", "2020至2024_CAGR_百分比"]].isna().any(axis=1)]
        raise ValueError("五指标存在空值，未生成正式数据库：\n" +
                         missing[["地区", "MSA", "CBSA"]].to_string(index=False))
    if industry.duplicated(["CBSA", "行业"]).any():
        raise ValueError("同一MSA存在重复行业记录")
    counts = industry.groupby("CBSA").size()
    if not counts.eq(13).all() or set(counts.index) != set(five["CBSA"]):
        raise ValueError("产业结构未做到每个MSA恰好13类，或与五指标城市不匹配")
    if (industry["就业合计"] != industry["男性就业"] + industry["女性就业"]).any():
        raise ValueError("产业男性、女性就业人数与合计不一致")
    totals = industry.groupby("CBSA")["就业合计"].sum()
    # If source sheet has 2024 total employment field (not in final five-indicator table),
    # the industry detail sum is the authoritative employment denominator used for shares.
    shares = industry.groupby("CBSA")["就业占比_百分比"].sum()
    bad_shares = shares[(shares - 100).abs() > 0.02]
    if len(bad_shares):
        raise ValueError("产业占比合计不为100%（容差0.02）：\n" + bad_shares.to_string())

    # Industry wide matrices: columns are the 13 industry classes.
    industry["行业"] = industry["行业"].astype(str).str.replace(r"^\d+\.\s*", "", regex=True)
    categories = list(dict.fromkeys(industry["行业"].tolist()))
    if len(categories) != 13:
        raise ValueError(f"检测到的行业类别不是13类，而是{len(categories)}类：{categories}")
    key_cols = ["地区编号", "地区", "原始清单序号", "MSA", "CBSA"]
    share_matrix = industry.pivot(index="CBSA", columns="行业", values="就业占比_百分比").reset_index()
    emp_matrix = industry.pivot(index="CBSA", columns="行业", values="就业合计").reset_index()
    meta = five[key_cols + ["2024人口"]].drop_duplicates("CBSA")
    share_matrix = meta.merge(share_matrix, on="CBSA", validate="one_to_one")
    emp_matrix = meta.merge(emp_matrix, on="CBSA", validate="one_to_one")

    # Region summary.
    region_summary = five.groupby(["地区编号", "地区"], sort=False).agg({
        "CBSA": "nunique",
        "2024人口": "sum",
        "2024人口密度_人每平方英里": "mean",
        "2024人均GDP_美元": "mean",
        "2020至2024_CAGR_百分比": "mean"
    }).reset_index().rename(columns={
        "CBSA": "MSA数量",
        "2024人口": "2024总人口",
        "2024人口密度_人每平方英里": "平均人口密度",
        "2024人均GDP_美元": "平均人均GDP",
        "2020至2024_CAGR_百分比": "平均人口CAGR"
    })

    # Add source/method notes and basic checks.
    qa = pd.DataFrame([
        ["地区数", 9, five["地区编号"].nunique(), "PASS" if five["地区编号"].nunique()==9 else "FAIL"],
        ["MSA总数", 195, len(five), "PASS" if len(five)==195 else "FAIL"],
        ["产业明细记录", 2535, len(industry), "PASS" if len(industry)==2535 else "FAIL"],
        ["唯一CBSA数", 195, five["CBSA"].nunique(), "PASS" if five["CBSA"].nunique()==195 else "FAIL"],
        ["每城13类行业", 13, int(counts.min()), "PASS" if counts.eq(13).all() else "FAIL"],
        ["产业占比合计", 100, round(float(shares.min()), 6), "PASS" if (shares-100).abs().max()<=0.02 else "FAIL"],
        ["人均GDP空值数", 0, int(five["2024人均GDP_美元"].isna().sum()), "PASS" if five["2024人均GDP_美元"].notna().all() else "FAIL"],
    ], columns=["检查项目", "预期值", "实际值", "结果"])

    # Write Excel database.
    sheets = {
        "全国五指标总表": five,
        "产业结构明细": industry,
        "产业占比矩阵": share_matrix,
        "产业就业人数矩阵": emp_matrix,
        "地区汇总": region_summary,
        "数据质量检查": qa,
        "地区读取日志": audit_df,
    }
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
        for sheet, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet, index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            ws.sheet_view.showGridLines = False
            ws.row_dimensions[1].height = 30
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            for col_cells in ws.columns:
                max_len = min(max(len(str(c.value or "")) for c in col_cells[: min(ws.max_row, 80)]) + 2, 42)
                ws.column_dimensions[get_column_letter(col_cells[0].column)].width = max(12, max_len)
            # readable numeric formats
            for row in ws.iter_rows(min_row=2):
                for c in row:
                    if isinstance(c.value, float):
                        c.number_format = "0.0000"

    # Write SQLite relational database, replacing the prior output atomically at completion.
    if OUT_SQLITE.exists():
        OUT_SQLITE.unlink()
    with sqlite3.connect(OUT_SQLITE) as conn:
        five.to_sql("msa_five_indicators", conn, index=False)
        industry.to_sql("msa_industry_detail", conn, index=False)
        share_matrix.to_sql("msa_industry_share_matrix", conn, index=False)
        emp_matrix.to_sql("msa_industry_employment_matrix", conn, index=False)
        region_summary.to_sql("region_summary", conn, index=False)
        audit_df.to_sql("source_audit", conn, index=False)
        conn.execute("CREATE UNIQUE INDEX idx_msa_cbsa ON msa_five_indicators(CBSA)")
        conn.execute("CREATE INDEX idx_industry_cbsa ON msa_industry_detail(CBSA)")
        conn.execute("CREATE INDEX idx_industry_region ON msa_industry_detail(地区编号)")
        conn.commit()

    print("\n数据库生成完成：")
    print(f"Excel:  {OUT_XLSX}")
    print(f"SQLite: {OUT_SQLITE}")
    print(f"MSA总数：{len(five)}；产业明细：{len(industry)}；地区数：{five['地区编号'].nunique()}")
    print("注意：本程序只合并和校验数据，不重新计算或改变任何地区原有指标。")

if __name__ == "__main__":
    main()
