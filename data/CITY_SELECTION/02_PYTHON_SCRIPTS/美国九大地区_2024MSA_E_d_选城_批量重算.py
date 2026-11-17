# -*- coding: utf-8 -*-
"""
美国九大地区 2024 MSA：产业结构 E -> 五指标距离 d -> 代表城市筛选
全部九区一次性重算；太平洋31个MSA作为一个完整地区，不拆分南北。

输入：
  美国九大地区_2024MSA综合数据库.sqlite

输出：
  美国九大地区_2024MSA_E_d_最终结果汇总.xlsx
  美国九大地区_2024MSA_最终选城汇总.csv

计算口径（九区完全一致，区域内独立计算）：
  1) 13类产业就业份额 = 各行业就业合计 / 该MSA 13类行业就业合计
  2) H(i,j) = sqrt(0.5 * sum_k((sqrt(p_ik)-sqrt(p_jk))^2))
  3) H_RMS = sqrt(mean(H_ij^2, i<j))
  4) E_ij = H_ij / H_RMS * sqrt(2n/(n-1))
  5) ln(人口)、ln(人口密度)、人均GDP、人口CAGR分别在本地区做总体Z-score（ddof=0）
  6) d_ij = sqrt(0.2 * (dZ人口^2+dZ密度^2+dZGDP^2+E_ij^2+dZ增长率^2))
  7) 第一城：平均到本地区所有MSA的d最小（Medoid）
  8) 第二城：固定第一城，最小化平均最近中心距离；改善率>=20%才纳入
  9) 精确并列保留并报告，不擅自裁决。

依赖：
  py -m pip install openpyxl

运行：
  cd "$env:USERPROFILE\\Downloads"
  py ".\\美国九大地区_2024MSA_E_d_选城_批量重算.py"

如数据库不在同一目录：
  py ".\\美国九大地区_2024MSA_E_d_选城_批量重算.py" --db "D:\\data\\美国九大地区_2024MSA综合数据库.sqlite"
"""

from __future__ import annotations

import argparse
import csv
import math
import sqlite3
import sys
from pathlib import Path
from collections import defaultdict

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import ColorScaleRule
    from openpyxl.comments import Comment
except ImportError:
    raise SystemExit("缺少 openpyxl。请先运行：py -m pip install openpyxl")


BASE = Path(__file__).resolve().parent
DEFAULT_DB = BASE / "美国九大地区_2024MSA综合数据库.sqlite"
OUTPUT_XLSX = BASE / "美国九大地区_2024MSA_E_d_最终结果汇总.xlsx"
OUTPUT_CSV = BASE / "美国九大地区_2024MSA_最终选城汇总.csv"

REGIONS = [
    ("1", "新英格兰 New England", "NE", 11),
    ("2", "中大西洋 Middle Atlantic", "MA", 18),
    ("3", "东北中部 East North Central", "ENC", 24),
    ("4", "西北中部 West North Central", "WNC", 13),
    ("5", "南大西洋 South Atlantic", "SA", 42),
    ("6", "东南中部 East South Central", "ESC", 16),
    ("7", "西南中部 West South Central", "WSC", 25),
    ("8", "山地 Mountain", "MTN", 15),
    ("9", "太平洋 Pacific", "PAC", 31),
]

EXPECTED_TOTAL_MSA = 195
EXPECTED_INDUSTRY_ROWS = 195 * 13
WEIGHTS = {
    "人口规模": 0.20,
    "人口密度": 0.20,
    "人均GDP": 0.20,
    "产业结构": 0.20,
    "人口增长率": 0.20,
}
SECOND_CITY_THRESHOLD = 0.20
TIE_TOL = 1e-10

# 数据库列名来自“美国九大地区_2024MSA数据库合并.py”中实际写入SQLite的标准字段。
FIVE_TABLE = "msa_five_indicators"
INDUSTRY_TABLE = "msa_industry_detail"


def as_float(value, field, msa):
    if value is None or str(value).strip() == "":
        raise ValueError(f"{msa} 的字段“{field}”为空")
    try:
        x = float(str(value).replace(",", "").strip())
    except Exception:
        raise ValueError(f"{msa} 的字段“{field}”无法转为数字：{value!r}")
    if not math.isfinite(x):
        raise ValueError(f"{msa} 的字段“{field}”不是有限数值：{value!r}")
    return x


def as_int(value, field, msa):
    x = as_float(value, field, msa)
    if abs(x - round(x)) > 1e-9:
        raise ValueError(f"{msa} 的字段“{field}”应为整数，实际为{x}")
    return int(round(x))


def norm_code(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    s = str(value).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def table_columns(conn, table):
    rows = conn.execute(f'PRAGMA table_info("{table}")').fetchall()
    return {r[1] for r in rows}


def require_columns(conn, table, required):
    try:
        cols = table_columns(conn, table)
    except sqlite3.Error as exc:
        raise ValueError(f"无法读取数据库表 {table}：{exc}")
    missing = sorted(set(required) - cols)
    if missing:
        raise ValueError(
            f"数据库表 {table} 缺少字段：{missing}\n"
            f"实际字段：{sorted(cols)}"
        )


def load_database(db_path: Path):
    if not db_path.exists():
        raise FileNotFoundError(
            f"找不到全国SQLite数据库：{db_path}\n"
            "请把本程序与“美国九大地区_2024MSA综合数据库.sqlite”放在同一文件夹，"
            "或用 --db 指定数据库路径。"
        )

    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    try:
        require_columns(conn, FIVE_TABLE, [
            "地区编号", "地区", "原始清单序号", "MSA", "CBSA", "2024人口",
            "2024人口密度_人每平方英里", "2024人均GDP_美元",
            "2020至2024_CAGR_百分比"
        ])
        require_columns(conn, INDUSTRY_TABLE, [
            "地区编号", "地区", "原始清单序号", "MSA", "CBSA", "行业",
            "男性就业", "女性就业", "就业合计"
        ])

        five_rows = conn.execute(
            f'SELECT * FROM "{FIVE_TABLE}" ORDER BY CAST("地区编号" AS INTEGER), '
            f'CAST("原始清单序号" AS INTEGER)'
        ).fetchall()
        industry_rows = conn.execute(
            f'SELECT rowid AS _rowid, * FROM "{INDUSTRY_TABLE}" ORDER BY rowid'
        ).fetchall()
    finally:
        conn.close()

    if len(five_rows) != EXPECTED_TOTAL_MSA:
        raise ValueError(
            f"全国五指标表应有{EXPECTED_TOTAL_MSA}个MSA，实际{len(five_rows)}个。"
            "请先检查全国数据库，不要在不完整数据上生成正式结果。"
        )
    if len(industry_rows) != EXPECTED_INDUSTRY_ROWS:
        raise ValueError(
            f"全国产业明细应有{EXPECTED_INDUSTRY_ROWS}条，实际{len(industry_rows)}条。"
        )

    five_by_region = defaultdict(list)
    global_cbsa = set()
    for r in five_rows:
        region_id = norm_code(r["地区编号"])
        msa = str(r["MSA"]).strip()
        cbsa = norm_code(r["CBSA"])
        if region_id not in {x[0] for x in REGIONS}:
            raise ValueError(f"{msa} 的地区编号异常：{region_id}")
        if not cbsa:
            raise ValueError(f"{msa} 的CBSA为空")
        if cbsa in global_cbsa:
            raise ValueError(f"全国五指标表出现重复CBSA：{cbsa}")
        global_cbsa.add(cbsa)

        city = {
            "region_id": region_id,
            "region": str(r["地区"]).strip(),
            "no": as_int(r["原始清单序号"], "原始清单序号", msa),
            "name": msa,
            "cbsa": cbsa,
            "population": as_float(r["2024人口"], "2024人口", msa),
            "density": as_float(r["2024人口密度_人每平方英里"], "2024人口密度_人每平方英里", msa),
            "gdp_pc": as_float(r["2024人均GDP_美元"], "2024人均GDP_美元", msa),
            "cagr": as_float(r["2020至2024_CAGR_百分比"], "2020至2024_CAGR_百分比", msa),
        }
        if city["population"] <= 0 or city["density"] <= 0 or city["gdp_pc"] <= 0:
            raise ValueError(f"{msa} 的人口、密度、人均GDP必须大于0")
        five_by_region[region_id].append(city)

    # 保留数据库产业明细的首次出现顺序，确保13类行业顺序一致。
    sector_order = []
    jobs_by_cbsa = defaultdict(dict)
    industry_region_by_cbsa = {}
    industry_msa_by_cbsa = {}
    industry_counts = defaultdict(int)
    for r in industry_rows:
        region_id = norm_code(r["地区编号"])
        cbsa = norm_code(r["CBSA"])
        msa = str(r["MSA"]).strip()
        sector = str(r["行业"]).strip()
        if not sector:
            raise ValueError(f"{msa} 存在空行业名称")
        if sector not in sector_order:
            sector_order.append(sector)
        if sector in jobs_by_cbsa[cbsa]:
            raise ValueError(f"{msa} 的行业“{sector}”重复")
        male = as_int(r["男性就业"], "男性就业", msa)
        female = as_int(r["女性就业"], "女性就业", msa)
        total = as_int(r["就业合计"], "就业合计", msa)
        if min(male, female, total) < 0:
            raise ValueError(f"{msa} {sector} 的就业人数不能为负")
        if male + female != total:
            raise ValueError(f"{msa} {sector} 男性+女性就业人数不等于就业合计")
        jobs_by_cbsa[cbsa][sector] = total
        industry_region_by_cbsa[cbsa] = region_id
        industry_msa_by_cbsa[cbsa] = msa
        industry_counts[cbsa] += 1

    if len(sector_order) != 13:
        raise ValueError(f"全国行业类别应为13类，实际{len(sector_order)}类：{sector_order}")
    if set(jobs_by_cbsa) != global_cbsa:
        missing = sorted(global_cbsa - set(jobs_by_cbsa))
        extra = sorted(set(jobs_by_cbsa) - global_cbsa)
        raise ValueError(f"五指标与产业表CBSA不一致。缺产业记录={missing}；多余产业记录={extra}")

    for cbsa in global_cbsa:
        if industry_counts[cbsa] != 13:
            raise ValueError(
                f"{industry_msa_by_cbsa[cbsa]} 应有13条产业记录，实际{industry_counts[cbsa]}条"
            )
        if set(jobs_by_cbsa[cbsa]) != set(sector_order):
            raise ValueError(f"{industry_msa_by_cbsa[cbsa]} 的13类行业不完整")

    # 逐区核验城市数与产业记录数。
    region_data = {}
    for region_id, region_name, short, expected_n in REGIONS:
        cities = sorted(five_by_region.get(region_id, []), key=lambda x: x["no"])
        if len(cities) != expected_n:
            raise ValueError(
                f"{region_name} 应有{expected_n}城，实际{len(cities)}城。"
                "数据库版本或地区归属可能不一致。"
            )
        for c in cities:
            if c["region"] != region_name:
                raise ValueError(
                    f"{c['name']} 的地区文字为“{c['region']}”，"
                    f"但地区编号{region_id}应为“{region_name}”"
                )
        region_data[region_id] = {
            "region_id": region_id,
            "region": region_name,
            "short": short,
            "expected_n": expected_n,
            "cities": cities,
            "sectors": sector_order,
            "jobs_by_cbsa": jobs_by_cbsa,
        }

    return region_data, five_rows, industry_rows, sector_order


def zscore(values):
    n = len(values)
    mean = sum(values) / n
    variance = sum((x - mean) ** 2 for x in values) / n  # 总体方差，ddof=0
    sd = math.sqrt(variance)
    if sd <= 1e-14:
        raise ValueError("某个指标在该地区内没有变化，标准差为0，不能计算Z-score")
    return [(x - mean) / sd for x in values], mean, sd


def hellinger_matrix(shares):
    n = len(shares)
    H = [[0.0] * n for _ in range(n)]
    for i in range(n):
        for j in range(i + 1, n):
            h = math.sqrt(
                0.5 * sum(
                    (math.sqrt(shares[i][k]) - math.sqrt(shares[j][k])) ** 2
                    for k in range(len(shares[i]))
                )
            )
            H[i][j] = h
            H[j][i] = h
    return H


def matrix_rms_upper(matrix):
    n = len(matrix)
    pair_count = n * (n - 1) // 2
    if pair_count <= 0:
        raise ValueError("城市数量不足，无法计算配对RMS")
    return math.sqrt(
        sum(matrix[i][j] ** 2 for i in range(n) for j in range(i + 1, n))
        / pair_count
    )


def calculate_region(data):
    cities = data["cities"]
    sectors = data["sectors"]
    n = len(cities)

    jobs_matrix = []
    shares = []
    totals = []
    for city in cities:
        sector_jobs = data["jobs_by_cbsa"][city["cbsa"]]
        row = [sector_jobs[s] for s in sectors]
        total = sum(row)
        if total <= 0:
            raise ValueError(f"{city['name']} 的13类产业就业总数为0")
        p = [v / total for v in row]
        if abs(sum(p) - 1.0) > 1e-12:
            raise ValueError(f"{city['name']} 的产业占比合计不等于100%")
        jobs_matrix.append(row)
        shares.append(p)
        totals.append(total)

    H = hellinger_matrix(shares)
    h_rms = matrix_rms_upper(H)
    if h_rms <= 1e-14:
        raise ValueError(f"{data['region']} 的所有城市产业结构完全相同，无法校准E")
    K = math.sqrt(2 * n / (n - 1)) / h_rms
    E = [[H[i][j] * K for j in range(n)] for i in range(n)]

    # 四项连续指标：人口、密度取自然对数；GDP和CAGR保留原值后做Z-score。
    raw_vectors = [
        [math.log(c["population"]) for c in cities],
        [math.log(c["density"]) for c in cities],
        [c["gdp_pc"] for c in cities],
        [c["cagr"] for c in cities],
    ]
    metric_names = ["ln(2024人口)", "ln(2024人口密度)", "2024人均GDP", "2020-2024人口CAGR"]
    z_vectors = []
    stats = []
    for name, vals in zip(metric_names, raw_vectors):
        z, mean, sd = zscore(vals)
        z_vectors.append(z)
        stats.append({"指标": name, "变换后均值": mean, "总体标准差ddof0": sd})

    Z = [[z_vectors[k][i] for k in range(4)] for i in range(n)]

    # 五项等权：每项20%；E作为产业结构的两两距离。
    D = [[0.0] * n for _ in range(n)]
    for i in range(n):
        for j in range(i + 1, n):
            scalar_sq = sum((Z[i][k] - Z[j][k]) ** 2 for k in range(4))
            d = math.sqrt(0.20 * (scalar_sq + E[i][j] ** 2))
            D[i][j] = d
            D[j][i] = d

    # 矩阵性质检查。
    for M, label in ((H, "H"), (E, "E"), (D, "d")):
        for i in range(n):
            if abs(M[i][i]) > 1e-12:
                raise AssertionError(f"{label}矩阵对角线不为0")
            for j in range(n):
                if abs(M[i][j] - M[j][i]) > 1e-12:
                    raise AssertionError(f"{label}矩阵不对称")

    # 单城覆盖距离 / Medoid。
    single_scores = [
        sum(D[i][candidate] for i in range(n)) / n
        for candidate in range(n)
    ]
    best_single = min(single_scores)
    medoid_indices = [
        i for i, score in enumerate(single_scores)
        if abs(score - best_single) <= TIE_TOL
    ]

    # 对每一个Medoid并列候选分别计算第二城市，不擅自打破并列。
    pair_scenarios = []
    pair_rows = []
    for medoid in medoid_indices:
        candidates = []
        for second in range(n):
            if second == medoid:
                continue
            d2 = sum(min(D[i][medoid], D[i][second]) for i in range(n)) / n
            improvement = (single_scores[medoid] - d2) / single_scores[medoid]
            candidates.append({
                "first_idx": medoid,
                "second_idx": second,
                "d1": single_scores[medoid],
                "d2": d2,
                "improvement": improvement,
            })
        best_d2 = min(x["d2"] for x in candidates)
        best_second_rows = [
            x for x in candidates if abs(x["d2"] - best_d2) <= TIE_TOL
        ]
        qualifies = ((single_scores[medoid] - best_d2) / single_scores[medoid]
                     >= SECOND_CITY_THRESHOLD - TIE_TOL)
        for x in candidates:
            x["is_best_tie"] = abs(x["d2"] - best_d2) <= TIE_TOL
            x["qualifies"] = qualifies and x["is_best_tie"]
            pair_rows.append(x)
        pair_scenarios.append({
            "first_idx": medoid,
            "d1": single_scores[medoid],
            "best_d2": best_d2,
            "best_improvement": (single_scores[medoid] - best_d2) / single_scores[medoid],
            "qualifies": qualifies,
            "best_second_indices": [x["second_idx"] for x in best_second_rows],
            "all_candidates": candidates,
        })

    return {
        **data,
        "n": n,
        "pair_count": n * (n - 1) // 2,
        "jobs_matrix": jobs_matrix,
        "job_totals": totals,
        "shares": shares,
        "H": H,
        "h_rms": h_rms,
        "K": K,
        "E": E,
        "Z": Z,
        "z_stats": stats,
        "D": D,
        "single_scores": single_scores,
        "medoid_indices": medoid_indices,
        "pair_scenarios": pair_scenarios,
        "pair_rows": pair_rows,
    }


def style_sheet(ws, widths=None, freeze="A2"):
    ws.freeze_panes = freeze
    ws.sheet_view.showGridLines = False
    ws.auto_filter.ref = ws.dimensions if ws.max_row > 1 else None
    ws.row_dimensions[1].height = 32
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="155E75")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col in range(1, ws.max_column + 1):
        letter = get_column_letter(col)
        if widths and col <= len(widths):
            width = widths[col - 1]
        else:
            width = 16
        ws.column_dimensions[letter].width = width


def add_table_sheet(wb, title, headers, rows, widths=None, number_formats=None):
    ws = wb.create_sheet(title)
    ws.append(headers)
    for row in rows:
        ws.append(row)
    style_sheet(ws, widths)
    if number_formats:
        for col, fmt in number_formats.items():
            for r in range(2, ws.max_row + 1):
                ws.cell(r, col).number_format = fmt
    return ws


def add_matrix_sheet(wb, title, cities, matrix, decimals=6):
    headers = ["原始序号", "MSA", "CBSA"] + [
        f"{c['name']} [{c['cbsa']}]" for c in cities
    ]
    rows = []
    for i, city in enumerate(cities):
        rows.append([city["no"], city["name"], city["cbsa"]] + matrix[i])
    ws = add_table_sheet(
        wb, title, headers, rows,
        widths=[11, 42, 12] + [23] * len(cities)
    )
    last_col = get_column_letter(ws.max_column)
    if ws.max_row >= 2 and ws.max_column >= 4:
        ws.conditional_formatting.add(
            f"D2:{last_col}{ws.max_row}",
            ColorScaleRule(
                start_type="min", start_color="FEE2E2",
                mid_type="percentile", mid_value=50, mid_color="FEF3C7",
                end_type="max", end_color="DCFCE7"
            )
        )
    for row in ws.iter_rows(min_row=2, min_col=4):
        for cell in row:
            cell.number_format = f"0.{('0' * decimals)}"
    return ws


def region_result_text(result):
    first_names = [result["cities"][i]["name"] for i in result["medoid_indices"]]
    scenarios = []
    for s in result["pair_scenarios"]:
        first = result["cities"][s["first_idx"]]["name"]
        if s["qualifies"]:
            second_names = [result["cities"][i]["name"] for i in s["best_second_indices"]]
            scenarios.append(f"{first} → {' / '.join(second_names)}")
        else:
            scenarios.append(f"{first} → 未达到20%")
    return " / ".join(scenarios)


def build_workbook(results, five_rows, industry_rows, sector_order, db_path):
    wb = Workbook()
    ws0 = wb.active
    ws0.title = "九区最终汇总"

    summary_headers = [
        "地区编号", "地区", "MSA数量", "城市配对数", "H_RMS", "产业校准系数K",
        "第一城Medoid", "第一城平均距离D1",
        "第二城结果（每个Medoid方案）", "最佳双城距离D2",
        "最佳改善率", "20%门槛", "最终城市席位数", "备注"
    ]
    ws0.append(summary_headers)

    final_rows = []
    summary_csv_rows = []
    param_rows = []
    quality_rows = []

    for r in results:
        first_names = [r["cities"][i]["name"] for i in r["medoid_indices"]]
        first_display = " / ".join(first_names)
        scenario_texts = []
        d2_values = []
        improvement_values = []
        seat_count_values = []
        for s in r["pair_scenarios"]:
            first_name = r["cities"][s["first_idx"]]["name"]
            d2_values.append(s["best_d2"])
            improvement_values.append(s["best_improvement"])
            if s["qualifies"]:
                second_names = [r["cities"][i]["name"] for i in s["best_second_indices"]]
                scenario_texts.append(f"{first_name} → {' / '.join(second_names)}")
                seat_count_values.append(2)
            else:
                scenario_texts.append(f"{first_name} → 不增加第二城（改善率不足20%）")
                seat_count_values.append(1)

            scenario_no = len(scenario_texts)
            final_rows.append([
                r["region_id"], r["region"], scenario_no, 1, first_name,
                r["cities"][s["first_idx"]]["cbsa"], "第一城Medoid",
                s["d1"], None, None
            ])
            if s["qualifies"]:
                for idx in s["best_second_indices"]:
                    final_rows.append([
                        r["region_id"], r["region"], scenario_no, 2,
                        r["cities"][idx]["name"], r["cities"][idx]["cbsa"],
                        "第二城并列最优候选（同一席位，择一）",
                        s["d1"], s["best_d2"], s["best_improvement"]
                    ])

        best_d2_display = " / ".join(f"{x:.6f}" for x in d2_values)
        best_improve_display = " / ".join(f"{x:.2%}" for x in improvement_values)
        threshold_display = "通过" if all(x >= SECOND_CITY_THRESHOLD - TIE_TOL for x in improvement_values) else (
            "部分方案通过" if any(x >= SECOND_CITY_THRESHOLD - TIE_TOL for x in improvement_values) else "未通过"
        )
        second_display = "；".join(scenario_texts)
        selected_seats = "；".join(str(x) for x in seat_count_values)
        note = "Medoid或第二城存在数学并列时保留并列候选，不强行裁决。" if (
            len(r["medoid_indices"]) > 1 or any(len(s["best_second_indices"]) > 1 for s in r["pair_scenarios"])
        ) else ""
        ws0.append([
            r["region_id"], r["region"], r["n"], r["pair_count"], r["h_rms"], r["K"],
            first_display, min(r["single_scores"][i] for i in r["medoid_indices"]),
            second_display, best_d2_display, best_improve_display,
            threshold_display, selected_seats, note
        ])

        summary_csv_rows.append({
            "地区编号": r["region_id"],
            "地区": r["region"],
            "MSA数量": r["n"],
            "配对数": r["pair_count"],
            "H_RMS": f"{r['h_rms']:.10f}",
            "K": f"{r['K']:.10f}",
            "第一城Medoid": first_display,
            "第一城平均距离D1": f"{min(r['single_scores'][i] for i in r['medoid_indices']):.10f}",
            "第二城方案": second_display,
            "最佳双城距离D2": best_d2_display,
            "改善率": best_improve_display,
            "20%门槛": threshold_display,
        })

        for stat in r["z_stats"]:
            param_rows.append([
                r["region_id"], r["region"], r["n"], r["h_rms"], r["K"],
                stat["指标"], stat["变换后均值"], stat["总体标准差ddof0"]
            ])

        # Region-level QA.
        shares_ok = all(abs(sum(p) - 1.0) <= 1e-12 for p in r["shares"])
        matrix_ok = True
        for M in (r["H"], r["E"], r["D"]):
            for i in range(r["n"]):
                if abs(M[i][i]) > 1e-12:
                    matrix_ok = False
                for j in range(r["n"]):
                    if abs(M[i][j] - M[j][i]) > 1e-12:
                        matrix_ok = False
        quality_rows.append([
            r["region_id"], r["region"], r["n"], r["n"] * 13,
            len(r["sectors"]), "通过" if shares_ok else "失败",
            "通过" if matrix_ok else "失败",
            "通过" if r["h_rms"] > 0 else "失败"
        ])

        code = r["short"]
        cities = r["cities"]
        sectors = r["sectors"]

        # Raw five-indicator table.
        raw_rows = []
        for i, c in enumerate(cities):
            raw_rows.append([
                c["no"], c["name"], c["cbsa"], c["population"], c["density"],
                c["gdp_pc"], c["cagr"], r["job_totals"][i],
                math.log(c["population"]), math.log(c["density"])
            ])
        add_table_sheet(
            wb, f'{r["region_id"]}_{code}_原始指标',
            ["原始序号", "MSA", "CBSA", "2024人口", "人口密度_人/平方英里",
             "2024人均GDP_美元", "2020-2024人口CAGR_百分比",
             "13类产业就业合计", "ln人口", "ln人口密度"],
            raw_rows,
            widths=[11, 42, 12, 17, 22, 20, 23, 22, 16, 18],
            number_formats={4: "#,##0", 5: "0.0000", 6: "$#,##0.00",
                            7: "0.000000", 8: "#,##0", 9: "0.000000", 10: "0.000000"}
        )

        # Industry employment matrix.
        emp_rows = [
            [c["no"], c["name"], c["cbsa"]] + r["jobs_matrix"][i] + [r["job_totals"][i]]
            for i, c in enumerate(cities)
        ]
        add_table_sheet(
            wb, f'{r["region_id"]}_{code}_产业就业',
            ["原始序号", "MSA", "CBSA"] + sectors + ["13类就业合计"],
            emp_rows,
            widths=[11, 42, 12] + [20] * 13 + [20],
            number_formats={j + 4: "#,##0" for j in range(14)}
        )

        # Industry share matrix.
        share_rows = [
            [c["no"], c["name"], c["cbsa"]] + r["shares"][i] + [sum(r["shares"][i])]
            for i, c in enumerate(cities)
        ]
        add_table_sheet(
            wb, f'{r["region_id"]}_{code}_产业占比',
            ["原始序号", "MSA", "CBSA"] + sectors + ["占比合计"],
            share_rows,
            widths=[11, 42, 12] + [20] * 13 + [16],
            number_formats={j + 4: "0.0000%" for j in range(14)}
        )

        add_matrix_sheet(wb, f'{r["region_id"]}_{code}_H矩阵', cities, r["H"])
        add_matrix_sheet(wb, f'{r["region_id"]}_{code}_E矩阵', cities, r["E"])

        # Z-score table.
        z_rows = []
        for i, c in enumerate(cities):
            z_rows.append([
                c["no"], c["name"], c["cbsa"],
                math.log(c["population"]), math.log(c["density"]),
                c["gdp_pc"], c["cagr"],
                r["Z"][i][0], r["Z"][i][1], r["Z"][i][2], r["Z"][i][3],
                r["single_scores"][i]
            ])
        add_table_sheet(
            wb, f'{r["region_id"]}_{code}_Z标准化',
            ["原始序号", "MSA", "CBSA", "ln人口", "ln人口密度",
             "人均GDP原值", "CAGR原值", "Z_ln人口", "Z_ln密度",
             "Z_GDP", "Z_CAGR", "单城平均距离D1"],
            z_rows,
            widths=[11, 42, 12, 16, 18, 20, 18, 16, 16, 16, 16, 22],
            number_formats={j: "0.000000" for j in range(4, 13)}
        )

        add_matrix_sheet(wb, f'{r["region_id"]}_{code}_d矩阵', cities, r["D"])

        # Single-city ranking.
        order = sorted(range(r["n"]), key=lambda i: (r["single_scores"][i], cities[i]["no"]))
        best_single = min(r["single_scores"])
        single_rows = []
        for rank, i in enumerate(order, 1):
            single_rows.append([
                rank, cities[i]["no"], cities[i]["name"], cities[i]["cbsa"],
                r["single_scores"][i],
                "是" if abs(r["single_scores"][i] - best_single) <= TIE_TOL else ""
            ])
        add_table_sheet(
            wb, f'{r["region_id"]}_{code}_单城排名',
            ["排名", "原始序号", "MSA", "CBSA", "单城平均距离D1", "是否Medoid并列最优"],
            single_rows,
            widths=[10, 12, 42, 12, 22, 22],
            number_formats={5: "0.000000"}
        )

        # Second-city coverage ranking, repeated per tied Medoid scenario.
        pair_rank_rows = []
        for scenario_no, s in enumerate(r["pair_scenarios"], 1):
            ordered_pairs = sorted(
                s["all_candidates"],
                key=lambda x: (x["d2"], cities[x["second_idx"]]["no"])
            )
            for rank, x in enumerate(ordered_pairs, 1):
                pair_rank_rows.append([
                    scenario_no,
                    cities[s["first_idx"]]["name"],
                    rank,
                    cities[x["second_idx"]]["name"],
                    cities[x["second_idx"]]["cbsa"],
                    x["d1"], x["d2"], x["improvement"],
                    "是" if x["is_best_tie"] else "",
                    "纳入第二城" if x["qualifies"] else (
                        "最佳候选但未达到20%" if x["is_best_tie"] else ""
                    )
                ])
        add_table_sheet(
            wb, f'{r["region_id"]}_{code}_双城覆盖',
            ["Medoid方案", "固定第一城", "候选排名", "第二城候选", "CBSA",
             "D1", "D2", "改善率", "是否并列最优", "准入判断"],
            pair_rank_rows,
            widths=[12, 42, 12, 42, 12, 16, 16, 16, 18, 24],
            number_formats={6: "0.000000", 7: "0.000000", 8: "0.00%"}
        )

    # Summary formatting.
    style_sheet(ws0, [10, 28, 12, 12, 16, 18, 48, 20, 90, 30, 24, 18, 18, 48])
    for row in range(2, ws0.max_row + 1):
        for col in (5, 6, 8):
            ws0.cell(row, col).number_format = "0.000000"
        ws0.cell(row, 9).alignment = Alignment(vertical="top", wrap_text=True)
        ws0.cell(row, 10).alignment = Alignment(vertical="top", wrap_text=True)
        ws0.cell(row, 11).alignment = Alignment(vertical="top", wrap_text=True)

    # Regional parameters.
    add_table_sheet(
        wb, "九区标准化参数",
        ["地区编号", "地区", "MSA数", "H_RMS", "K", "指标", "变换后均值", "总体标准差ddof0"],
        param_rows,
        widths=[10, 28, 12, 16, 16, 24, 22, 24],
        number_formats={4: "0.0000000000", 5: "0.0000000000",
                        7: "0.000000", 8: "0.000000"}
    )

    # Final selected-city seats, with tied candidates clearly shown as alternatives.
    add_table_sheet(
        wb, "最终入选城市",
        ["地区编号", "地区", "Medoid方案", "席位", "城市", "CBSA",
         "角色", "D1", "D2", "改善率"],
        final_rows,
        widths=[10, 28, 12, 10, 44, 12, 42, 16, 16, 16],
        number_formats={8: "0.000000", 9: "0.000000", 10: "0.00%"}
    )

    # Global data quality.
    global_checks = [
        ["MSA总数", EXPECTED_TOTAL_MSA, len(five_rows), "通过" if len(five_rows) == EXPECTED_TOTAL_MSA else "失败"],
        ["全国产业明细记录", EXPECTED_INDUSTRY_ROWS, len(industry_rows), "通过" if len(industry_rows) == EXPECTED_INDUSTRY_ROWS else "失败"],
        ["全国行业类别数", 13, len(sector_order), "通过" if len(sector_order) == 13 else "失败"],
        ["地区数量", 9, len(results), "通过" if len(results) == 9 else "失败"],
        ["太平洋MSA数量（不拆分）", 31, next(r["n"] for r in results if r["region_id"] == "9"), "通过"],
    ]
    qa_rows = [
        ["全国", "总体", item, f"预期：{expected}；实际：{actual}", status]
        for item, expected, actual, status in global_checks
    ]
    for r in results:
        qa_rows.append([
            r["region_id"], r["region"], "MSA数量", r["n"],
            "通过" if r["n"] == r["expected_n"] else "失败"
        ])
        qa_rows.append([
            r["region_id"], r["region"], "产业记录数量", r["n"] * 13,
            "通过" if r["n"] * 13 == r["expected_n"] * 13 else "失败"
        ])
        qa_rows.append([
            r["region_id"], r["region"], "产业占比合计", "每城100%",
            "通过" if all(abs(sum(p) - 1.0) <= 1e-12 for p in r["shares"]) else "失败"
        ])
        qa_rows.append([
            r["region_id"], r["region"], "H/E/d对称及对角线", "通过",
            "通过"
        ])
    add_table_sheet(
        wb, "数据质量检查",
        ["地区编号", "地区", "检查项目", "实际值", "结果"],
        qa_rows,
        widths=[12, 28, 30, 24, 16]
    )

    # Methodology and definitions.
    method_rows = [
        ["数据源", f"全国SQLite数据库：{db_path.name}"],
        ["统计单位", "美国九大 Census Divisions 内的MSA；Pacific为完整31城，不拆分南北。"],
        ["区域计算原则", "每个地区独立计算产业H、H_RMS、K、E、四项Z-score和d；不使用全国195城的均值、标准差或统一产业校准系数。"],
        ["产业数据", "2024 ACS C24030 的13类行业就业人数；就业占比由13类就业合计重新计算。"],
        ["行业占比", "p_ik = MSA i 的第k类行业就业合计 / 该MSA 13类行业就业合计。"],
        ["Hellinger距离", "H_ij = sqrt(0.5 × Σ_k (sqrt(p_ik)-sqrt(p_jk))²)。"],
        ["配对数", "每个地区 n(n-1)/2，只使用上三角非对角配对计算RMS。"],
        ["H_RMS", "sqrt(Σ_{i<j} H_ij² / [n(n-1)/2])。"],
        ["产业校准", "K = sqrt(2n/(n-1)) / H_RMS；E_ij = K × H_ij。沿用太平洋地区产业结构矩阵程序。"],
        ["人口规模", "ln(2024人口)，在本地区内做Z-score。"],
        ["人口密度", "ln(2024人口密度)，在本地区内做Z-score。"],
        ["人均GDP", "2024人均GDP美元值，在本地区内做Z-score。"],
        ["人口增长率", "2020至2024人口CAGR，在本地区内做Z-score。"],
        ["Z-score", "Z=(x-本地区均值)/本地区总体标准差；ddof=0。"],
        ["五指标权重", "人口规模20%、人口密度20%、人均GDP20%、产业结构20%、人口增长率20%。"],
        ["综合距离d", "d_ij = sqrt(0.2 × [(ΔZ人口)²+(ΔZ密度)²+(ΔZ人均GDP)²+E_ij²+(ΔZ增长率)²])。"],
        ["第一座城市", "Medoid：D1(c)=本地区所有城市到候选城市c的平均d；选择D1最小者。"],
        ["第二座城市", "固定第一城，计算D2(c2)=mean_i min[d(i,c1),d(i,c2)]；选择D2最小者。"],
        ["改善率", "R=(D1-D2)/D1。R≥20%才增加第二城市。"],
        ["城市权重", "覆盖计算对本地区MSA等权，不按人口加权。"],
        ["并列规则", "若Medoid或第二城候选在容差1e-10内并列，全部列出作为并列候选，不私自指定唯一胜者。"],
        ["重要限制", "九区之间的d为各自区域内相对距离，不直接用于九区之间的数值排名。"],
        ["输出", "包含九区原始指标、13类就业/占比矩阵、H、E、Z、d、单城排名、双城覆盖排名、最终入选席位与质量检查。"],
    ]
    add_table_sheet(wb, "计算方法与公式", ["项目", "定义、公式或说明"], method_rows,
                    widths=[26, 120])
    ws_method = wb["计算方法与公式"]
    for row in ws_method.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    # Nationwide raw data audit sheets (all records are directly from the database).
    five_raw_rows = []
    for r in sorted(five_rows, key=lambda x: (int(norm_code(x["地区编号"])), int(x["原始清单序号"]))):
        five_raw_rows.append([
            r["地区编号"], r["地区"], r["原始清单序号"], r["MSA"], r["CBSA"],
            r["2024人口"], r["2024人口密度_人每平方英里"], r["2024人均GDP_美元"],
            r["2020至2024_CAGR_百分比"]
        ])
    add_table_sheet(
        wb, "全国195城五指标原始表",
        ["地区编号", "地区", "原始序号", "MSA", "CBSA", "2024人口",
         "2024人口密度_人/平方英里", "2024人均GDP_美元", "2020-2024_CAGR_百分比"],
        five_raw_rows,
        widths=[10, 28, 12, 48, 12, 18, 24, 22, 22],
        number_formats={6: "#,##0", 7: "0.0000", 8: "$#,##0.00", 9: "0.000000"}
    )

    industry_raw_rows = []
    for r in industry_rows:
        industry_raw_rows.append([
            r["地区编号"], r["地区"], r["原始清单序号"], r["MSA"], r["CBSA"],
            r["行业"], r["男性就业"], r["女性就业"], r["就业合计"],
            r["就业占比_百分比"]
        ])
    add_table_sheet(
        wb, "全国2535条产业明细",
        ["地区编号", "地区", "原始序号", "MSA", "CBSA", "行业",
         "男性就业", "女性就业", "就业合计", "源表就业占比_百分比"],
        industry_raw_rows,
        widths=[10, 28, 12, 48, 12, 36, 16, 16, 16, 22],
        number_formats={7: "#,##0", 8: "#,##0", 9: "#,##0", 10: "0.0000"}
    )

    # Reorder sheets: summaries first, regional detail follows, national raw tables last.
    desired_first = [
        "九区最终汇总", "最终入选城市", "九区标准化参数",
        "数据质量检查", "计算方法与公式"
    ]
    desired_last = ["全国195城五指标原始表", "全国2535条产业明细"]
    order = desired_first + [
        s for s in wb.sheetnames if s not in desired_first + desired_last
    ] + desired_last
    wb._sheets = [wb[s] for s in order]

    wb.properties.title = "美国九大地区2024 MSA产业E与五指标d最终城市筛选"
    wb.properties.subject = "195个MSA；九区独立计算；太平洋31城整体计算"
    wb.properties.creator = "OpenAI"
    wb.save(OUTPUT_XLSX)

    # CSV summary: easy to import into the paper's working files.
    with OUTPUT_CSV.open("w", encoding="utf-8-sig", newline="") as f:
        if summary_csv_rows:
            writer = csv.DictWriter(f, fieldnames=list(summary_csv_rows[0].keys()))
            writer.writeheader()
            writer.writerows(summary_csv_rows)

    return summary_csv_rows


def main():
    parser = argparse.ArgumentParser(
        description="从全国195个MSA SQLite数据库一次性重算九区产业E、五指标d和代表城市。"
    )
    parser.add_argument("--db", type=Path, default=DEFAULT_DB, help="全国SQLite数据库路径")
    args = parser.parse_args()

    try:
        region_data, five_rows, industry_rows, sector_order = load_database(args.db)
        results = []
        print("数据库完整性检查通过，开始九区重新计算……")
        for region_id, region_name, short, expected_n in REGIONS:
            result = calculate_region(region_data[region_id])
            results.append(result)
            first_names = [result["cities"][i]["name"] for i in result["medoid_indices"]]
            print(
                f"OK {region_name}: {result['n']}城，"
                f"配对={result['pair_count']}，"
                f"H_RMS={result['h_rms']:.8f}，K={result['K']:.8f}，"
                f"Medoid={' / '.join(first_names)}"
            )
            for s in result["pair_scenarios"]:
                first = result["cities"][s["first_idx"]]["name"]
                second = " / ".join(result["cities"][i]["name"] for i in s["best_second_indices"])
                print(
                    f"   {first} -> {second if s['qualifies'] else '不增加第二城'}；"
                    f"D1={s['d1']:.6f}，D2={s['best_d2']:.6f}，"
                    f"改善率={s['best_improvement']:.2%}"
                )

        print("\n所有地区计算完成，正在生成Excel和CSV……")
        build_workbook(results, five_rows, industry_rows, sector_order, args.db)
        print("\n完成：")
        print(f"Excel：{OUTPUT_XLSX.resolve()}")
        print(f"CSV：  {OUTPUT_CSV.resolve()}")
        print("地区数：9；MSA总数：195；产业明细：2535；太平洋：31城（不拆分）")
        print("算法：区域内Hellinger→RMS校准E→五指标d→Medoid→第二城20%覆盖改善。")

    except Exception as exc:
        print("\n计算未完成，未生成正式结果。")
        print(f"原因：{exc}")
        sys.exit(2)


if __name__ == "__main__":
    main()
