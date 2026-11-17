# -*- coding: utf-8 -*-
"""
美国九大地区 V3.0：人口总量与 GDP 经济总量双权重独立 P-Median 模型
====================================================================

研究目标
--------
从全国 195 个 MSA 中选出 10 座代表性城市：
1. 九大 Census Divisions 各至少入选 1 座；
2. 总数严格为 10 座；
3. 不设大型/中小型/特大型城市配额；
4. 五项城市特征指标基准权重各 20%；
5. 人口总量权重模型 V3-P 与 GDP 总量权重模型 V3-G 独立运行；
6. 两模型各进行 1,000 次（默认）Dirichlet 权重扰动蒙特卡洛检验。

五指标距离
----------
四个连续变量：ln(2024人口)、ln(2024人口密度)、2024人均GDP、2020-2024人口CAGR。
四个连续变量在各 Census Division 内分别做总体 Z-score（ddof=0）。
产业结构：13 类行业就业份额，计算 Hellinger 距离，再按区域 RMS 校准：
    H_ij = sqrt(0.5 * sum_k((sqrt(s_ik)-sqrt(s_jk))^2))
    H_RMS = sqrt(mean(H_ij^2, i<j))
    K = sqrt(2*n/(n-1)) / H_RMS
    E_ij = K * H_ij
五指标基准距离：
    d_ij = sqrt(0.2 * (dZ_ln人口^2 + dZ_ln密度^2 + dZ_GDPpc^2
                       + E_ij^2 + dZ_CAGR^2))

代表性权重
----------
V3-P：w_i = P_i / sum(P)
V3-G：GDP_total_i = P_i * GDPpc_i；w_i = GDP_total_i / sum(GDP_total)
GDP 总量是用数据库内 2024 人均 GDP × 2024 人口推导，程序会在输出中明确记录。
因此，GDPpc 必须是同一统计年度、同一美元价格口径、覆盖相同 MSA 边界的有效值。

精确优化
--------
总共 10 个名额、9 个地区且每区至少 1 个，因此必然是 8 个地区各选 1 座，
另 1 个地区选 2 座。对每个地区精确枚举所有单城与双城组合，计算加权覆盖损失；
先求各区最佳单城，再将第二个名额分配给“增加第二城后损失下降最多”的地区。
这是当前约束下的精确全局最优解，不是随机聚类或近似算法。

蒙特卡洛
--------
五指标权重扰动：alpha ~ Dirichlet(20,20,20,20,20)，均值仍为 20% 等权。
每次用 d_ij(alpha)=sqrt(sum_k alpha_k * component_ijk) 重新优化。
默认每个模型 1,000 次；两模型使用同一组 alpha 抽样，便于配对比较。

输入
----
美国九大地区_2024MSA综合数据库.sqlite
表：msa_five_indicators、msa_industry_detail
可选表：source_audit

依赖
----
py -m pip install openpyxl

运行（PowerShell）
-----------------
cd "$env:USERPROFILE\\Downloads"
py ".\美国九大地区_V3_人口与GDP双权重PMedian.py"

可选参数
--------
--db "D:\\data\\美国九大地区_2024MSA综合数据库.sqlite"
--mc 1000
--seed 20261116
--skip-mc  （仅调试时跳过蒙特卡洛；正式论文结果不要跳过）

重要
----
1. 原始 SQLite 数据库只读，不会被修改。
2. 数据缺失、重复、地区数量不符或产业类别不完整时，程序停止，不自动填补。
3. 本程序不联网，不调用 BEA API，不会自动补造 GDP 数据。
4. GDP 总量由现有数据库的人均 GDP 与人口推导；如果人均 GDP 来源或口径不可靠，
   应先修正数据库，再运行 V3-G。
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import itertools
import json
import math
import random
import shutil
import sqlite3
import statistics
import sys
import time
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import ColorScaleRule
except ImportError:
    raise SystemExit("缺少 openpyxl。请先运行：py -m pip install openpyxl")


BASE = Path(__file__).resolve().parent
DEFAULT_DB = BASE / "美国九大地区_2024MSA综合数据库.sqlite"
PREFIX = "美国九大地区_V3"
OUT_XLSX = BASE / f"{PREFIX}_双模型完整结果.xlsx"
OUT_SQLITE = BASE / f"{PREFIX}_中间数据.sqlite"
OUT_LOG = BASE / f"{PREFIX}_计算日志.txt"
OUT_SELECTED_P = BASE / f"{PREFIX}_人口加权_最终10城.csv"
OUT_SELECTED_G = BASE / f"{PREFIX}_GDP加权_最终10城.csv"
OUT_ASSIGN = BASE / f"{PREFIX}_195城覆盖分配.csv"
OUT_CANDIDATES = BASE / f"{PREFIX}_全部区域候选组合.csv"
OUT_CROSS = BASE / f"{PREFIX}_双模型交叉评价.csv"
OUT_MC_FREQ = BASE / f"{PREFIX}_蒙特卡洛入选频率.csv"
OUT_MC_RUNS = BASE / f"{PREFIX}_蒙特卡洛逐次结果.csv"

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
REGION_BY_ID = {x[0]: x for x in REGIONS}
EXPECTED_TOTAL_MSA = 195
EXPECTED_INDUSTRY_ROWS = 195 * 13
EXPECTED_SECTORS = 13
TOTAL_SELECTED = 10
TIE_TOL = 1e-10
DIRICHLET_SHAPE = 20.0
BASE_ALPHA = (0.2, 0.2, 0.2, 0.2, 0.2)
FIVE_TABLE = "msa_five_indicators"
INDUSTRY_TABLE = "msa_industry_detail"


# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------

def log(message: str, lines: list[str]) -> None:
    print(message)
    lines.append(str(message))


def norm_code(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    s = str(value).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def as_float(value, field: str, msa: str) -> float:
    if value is None or str(value).strip() == "":
        raise ValueError(f"{msa} 的字段“{field}”为空")
    try:
        x = float(str(value).replace(",", "").replace("$", "").replace("%", "").strip())
    except Exception as exc:
        raise ValueError(f"{msa} 的字段“{field}”无法转为数字：{value!r}") from exc
    if not math.isfinite(x):
        raise ValueError(f"{msa} 的字段“{field}”不是有限数值：{value!r}")
    return x


def as_int(value, field: str, msa: str) -> int:
    x = as_float(value, field, msa)
    if abs(x - round(x)) > 1e-8:
        raise ValueError(f"{msa} 的字段“{field}”应为整数，实际为{x}")
    return int(round(x))


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def table_columns(conn: sqlite3.Connection, table: str) -> set[str]:
    rows = conn.execute(f'PRAGMA table_info("{table}")').fetchall()
    return {row[1] for row in rows}


def require_columns(conn: sqlite3.Connection, table: str, required: list[str]) -> None:
    cols = table_columns(conn, table)
    if not cols:
        raise ValueError(f"数据库缺少数据表：{table}")
    missing = sorted(set(required) - cols)
    if missing:
        raise ValueError(f"数据表 {table} 缺少字段：{missing}；实际字段：{sorted(cols)}")


def percentile(values: list[float], p: float) -> float:
    if not values:
        return float("nan")
    xs = sorted(values)
    pos = (len(xs) - 1) * p
    lo = int(math.floor(pos))
    hi = int(math.ceil(pos))
    if lo == hi:
        return xs[lo]
    return xs[lo] * (hi - pos) + xs[hi] * (pos - lo)


def safe_sheet_name(name: str) -> str:
    for ch in "[]:*?/\\":
        name = name.replace(ch, "_")
    return name[:31]


def csv_write(path: Path, rows: list[dict], fieldnames: list[str] | None = None) -> None:
    if fieldnames is None:
        fieldnames = list(rows[0].keys()) if rows else []
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        if not fieldnames:
            return
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


# ---------------------------------------------------------------------------
# 数据库读取与严格校验
# ---------------------------------------------------------------------------

def load_database(db_path: Path, log_lines: list[str]):
    if not db_path.exists():
        raise FileNotFoundError(
            f"找不到全国 SQLite 数据库：{db_path}\n"
            "请将本程序与“美国九大地区_2024MSA综合数据库.sqlite”放在同一文件夹，"
            "或使用 --db 指定数据库路径。"
        )

    digest = sha256_file(db_path)
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    try:
        integrity = conn.execute("PRAGMA integrity_check").fetchone()[0]
        if str(integrity).lower() != "ok":
            raise ValueError(f"SQLite integrity_check 未通过：{integrity}")

        require_columns(conn, FIVE_TABLE, [
            "地区编号", "地区", "原始清单序号", "MSA", "CBSA", "2024人口",
            "2024人口密度_人每平方英里", "2024人均GDP_美元", "2020至2024_CAGR_百分比"
        ])
        require_columns(conn, INDUSTRY_TABLE, [
            "地区编号", "地区", "原始清单序号", "MSA", "CBSA", "行业",
            "男性就业", "女性就业", "就业合计"
        ])

        five_rows = [dict(r) for r in conn.execute(
            f'SELECT * FROM "{FIVE_TABLE}" ORDER BY CAST("地区编号" AS INTEGER), '
            f'CAST("原始清单序号" AS INTEGER)'
        ).fetchall()]
        industry_rows = [dict(r) for r in conn.execute(
            f'SELECT rowid AS _rowid, * FROM "{INDUSTRY_TABLE}" ORDER BY rowid'
        ).fetchall()]

        source_audit_rows = []
        if "source_audit" in {
            r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
        }:
            source_audit_rows = [dict(r) for r in conn.execute('SELECT * FROM "source_audit"').fetchall()]
    finally:
        conn.close()

    if len(five_rows) != EXPECTED_TOTAL_MSA:
        raise ValueError(f"五指标表应有195个MSA，实际{len(five_rows)}个")
    if len(industry_rows) != EXPECTED_INDUSTRY_ROWS:
        raise ValueError(f"产业明细表应有2535条记录，实际{len(industry_rows)}条")

    cities_by_cbsa = {}
    cities_by_region = defaultdict(list)
    for row in five_rows:
        region_id = norm_code(row["地区编号"])
        msa = str(row["MSA"]).strip()
        cbsa = norm_code(row["CBSA"])
        if region_id not in REGION_BY_ID:
            raise ValueError(f"{msa} 的地区编号不属于九大地区：{region_id}")
        if not msa or not cbsa:
            raise ValueError("五指标表存在空 MSA 名称或 CBSA")
        if cbsa in cities_by_cbsa:
            raise ValueError(f"五指标表 CBSA 重复：{cbsa}")
        region_name = str(row["地区"]).strip()
        expected_region_name = REGION_BY_ID[region_id][1]
        if region_name != expected_region_name:
            raise ValueError(
                f"{msa} 地区名称与地区编号不一致：{region_id} / {region_name}"
            )

        city = {
            "region_id": region_id,
            "region": region_name,
            "no": as_int(row["原始清单序号"], "原始清单序号", msa),
            "name": msa,
            "cbsa": cbsa,
            "population": as_float(row["2024人口"], "2024人口", msa),
            "density": as_float(row["2024人口密度_人每平方英里"], "2024人口密度_人每平方英里", msa),
            "gdp_pc": as_float(row["2024人均GDP_美元"], "2024人均GDP_美元", msa),
            "cagr": as_float(row["2020至2024_CAGR_百分比"], "2020至2024_CAGR_百分比", msa),
            "raw": row,
        }
        if city["population"] <= 0 or city["density"] <= 0 or city["gdp_pc"] <= 0:
            raise ValueError(f"{msa} 的人口、人口密度、人均GDP必须全部大于0")
        city["gdp_total"] = city["population"] * city["gdp_pc"]
        cities_by_cbsa[cbsa] = city
        cities_by_region[region_id].append(city)

    for region_id, region_name, short, expected_n in REGIONS:
        cities = cities_by_region.get(region_id, [])
        if len(cities) != expected_n:
            raise ValueError(f"{region_name} 应有{expected_n}城，实际{len(cities)}城")
        cities.sort(key=lambda c: c["no"])

    sector_order = []
    jobs_by_cbsa = defaultdict(dict)
    industry_count = Counter()
    industry_raw_by_cbsa = defaultdict(list)
    for row in industry_rows:
        cbsa = norm_code(row["CBSA"])
        msa = str(row["MSA"]).strip()
        region_id = norm_code(row["地区编号"])
        sector = str(row["行业"]).strip()
        if cbsa not in cities_by_cbsa:
            raise ValueError(f"产业表出现五指标表之外的 CBSA：{cbsa} ({msa})")
        city = cities_by_cbsa[cbsa]
        if region_id != city["region_id"]:
            raise ValueError(f"{msa} 五指标表与产业表的地区编号不一致")
        if not sector:
            raise ValueError(f"{msa} 存在空行业名称")
        if sector not in sector_order:
            sector_order.append(sector)
        if sector in jobs_by_cbsa[cbsa]:
            raise ValueError(f"{msa} 的行业“{sector}”重复")
        male = as_int(row["男性就业"], "男性就业", msa)
        female = as_int(row["女性就业"], "女性就业", msa)
        total = as_int(row["就业合计"], "就业合计", msa)
        if min(male, female, total) < 0:
            raise ValueError(f"{msa} {sector} 就业人数不能为负")
        if male + female != total:
            raise ValueError(f"{msa} {sector} 男性+女性就业人数不等于就业合计")
        jobs_by_cbsa[cbsa][sector] = total
        industry_count[cbsa] += 1
        industry_raw_by_cbsa[cbsa].append(row)

    if len(sector_order) != EXPECTED_SECTORS:
        raise ValueError(f"全国行业类别应为13类，实际{len(sector_order)}类：{sector_order}")
    if set(jobs_by_cbsa) != set(cities_by_cbsa):
        missing = sorted(set(cities_by_cbsa) - set(jobs_by_cbsa))
        extra = sorted(set(jobs_by_cbsa) - set(cities_by_cbsa))
        raise ValueError(f"五指标表与产业表 CBSA 不一致；缺产业={missing}；多余={extra}")

    for cbsa, city in cities_by_cbsa.items():
        if industry_count[cbsa] != EXPECTED_SECTORS:
            raise ValueError(f"{city['name']} 应有13条产业记录，实际{industry_count[cbsa]}条")
        if set(jobs_by_cbsa[cbsa]) != set(sector_order):
            raise ValueError(f"{city['name']} 的13类产业不完整")
        if sum(jobs_by_cbsa[cbsa].values()) <= 0:
            raise ValueError(f"{city['name']} 的13类产业就业合计为0")

    # 全国归一化权重：两套方案完全独立。
    population_sum = sum(c["population"] for c in cities_by_cbsa.values())
    gdp_sum = sum(c["gdp_total"] for c in cities_by_cbsa.values())
    if population_sum <= 0 or gdp_sum <= 0:
        raise ValueError("全国人口总和或GDP总和无效")
    for city in cities_by_cbsa.values():
        city["w_pop"] = city["population"] / population_sum
        city["w_gdp"] = city["gdp_total"] / gdp_sum

    # 记录核验摘要。
    qa = [
        {"检查项目": "SQLite完整性", "预期": "ok", "实际": "ok", "结果": "通过"},
        {"检查项目": "MSA总数", "预期": 195, "实际": len(five_rows), "结果": "通过"},
        {"检查项目": "产业明细记录数", "预期": 2535, "实际": len(industry_rows), "结果": "通过"},
        {"检查项目": "行业类别数", "预期": 13, "实际": len(sector_order), "结果": "通过"},
        {"检查项目": "地区数量", "预期": 9, "实际": len(cities_by_region), "结果": "通过"},
        {"检查项目": "CBSA唯一性", "预期": "195个唯一值", "实际": len(cities_by_cbsa), "结果": "通过"},
        {"检查项目": "每城13类行业", "预期": "每个MSA均为13类", "实际": "已逐城核验", "结果": "通过"},
        {"检查项目": "人均GDP", "预期": "195城均为有效正数", "实际": "已逐城核验", "结果": "通过"},
        {"检查项目": "GDP总量派生", "预期": "人口×人均GDP", "实际": "195城均已计算", "结果": "通过"},
        {"检查项目": "人口权重归一化", "预期": "权重和=1", "实际": sum(c["w_pop"] for c in cities_by_cbsa.values()), "结果": "通过"},
        {"检查项目": "GDP权重归一化", "预期": "权重和=1", "实际": sum(c["w_gdp"] for c in cities_by_cbsa.values()), "结果": "通过"},
    ]

    region_data = {}
    for region_id, region_name, short, expected_n in REGIONS:
        region_data[region_id] = {
            "region_id": region_id,
            "region": region_name,
            "short": short,
            "cities": cities_by_region[region_id],
            "sectors": sector_order,
            "jobs_by_cbsa": jobs_by_cbsa,
        }

    log(f"数据库校验通过：195城、2535条产业记录、13类行业、9个地区。", log_lines)
    log(f"全国人口总和（候选195城）：{population_sum:,.0f}", log_lines)
    log(f"全国GDP总和（由人口×人均GDP推导）：${gdp_sum:,.0f}", log_lines)
    log(f"输入数据库SHA256：{digest}", log_lines)
    return region_data, five_rows, industry_rows, sector_order, cities_by_cbsa, qa, source_audit_rows, digest


# ---------------------------------------------------------------------------
# 九区五指标距离矩阵
# ---------------------------------------------------------------------------

def zscore(values: list[float], label: str, region_name: str):
    n = len(values)
    mean = sum(values) / n
    sd = math.sqrt(sum((x - mean) ** 2 for x in values) / n)  # ddof=0
    if sd <= 1e-14:
        raise ValueError(f"{region_name} 的 {label} 区域内标准差为0，无法Z-score")
    return [(x - mean) / sd for x in values], mean, sd


def calculate_region(region: dict) -> dict:
    cities = region["cities"]
    sectors = region["sectors"]
    jobs_by_cbsa = region["jobs_by_cbsa"]
    n = len(cities)

    employment = []
    employment_totals = []
    shares = []
    for city in cities:
        row = [jobs_by_cbsa[city["cbsa"]][sector] for sector in sectors]
        total = sum(row)
        if total <= 0:
            raise ValueError(f"{city['name']} 的13类行业就业合计为0")
        employment.append(row)
        employment_totals.append(total)
        shares.append([v / total for v in row])
        if abs(sum(shares[-1]) - 1.0) > 1e-12:
            raise ValueError(f"{city['name']} 的产业占比合计不等于1")

    H = [[0.0] * n for _ in range(n)]
    for i in range(n):
        for j in range(i + 1, n):
            h = math.sqrt(0.5 * sum(
                (math.sqrt(shares[i][k]) - math.sqrt(shares[j][k])) ** 2
                for k in range(EXPECTED_SECTORS)
            ))
            H[i][j] = H[j][i] = h

    pair_count = n * (n - 1) // 2
    h_rms = math.sqrt(
        sum(H[i][j] ** 2 for i in range(n) for j in range(i + 1, n)) / pair_count
    )
    if h_rms <= 1e-14:
        raise ValueError(f"{region['region']} 的产业 H_RMS 为0，无法校准")
    K = math.sqrt(2.0 * n / (n - 1.0)) / h_rms
    E = [[K * H[i][j] for j in range(n)] for i in range(n)]

    transformed = [
        [math.log(c["population"]) for c in cities],
        [math.log(c["density"]) for c in cities],
        [c["gdp_pc"] for c in cities],
        [c["cagr"] for c in cities],
    ]
    labels = ["ln人口", "ln人口密度", "人均GDP", "人口CAGR"]
    z_columns = []
    z_stats = []
    for label, values in zip(labels, transformed):
        z, mean, sd = zscore(values, label, region["region"])
        z_columns.append(z)
        z_stats.append({"指标": label, "变换后均值": mean, "总体标准差_ddof0": sd})
    Z = [[z_columns[k][i] for k in range(4)] for i in range(n)]

    # 五个分量矩阵：前四项为区域Z-score差平方，第五项为产业E距离平方。
    components = [[[0.0] * n for _ in range(n)] for _ in range(5)]
    for i in range(n):
        for j in range(i + 1, n):
            for k in range(4):
                value = (Z[i][k] - Z[j][k]) ** 2
                components[k][i][j] = components[k][j][i] = value
            components[4][i][j] = components[4][j][i] = E[i][j] ** 2

    D = distance_from_components(components, BASE_ALPHA)
    return {
        **region,
        "n": n,
        "pair_count": pair_count,
        "employment": employment,
        "employment_totals": employment_totals,
        "shares": shares,
        "H": H,
        "h_rms": h_rms,
        "K": K,
        "E": E,
        "Z": Z,
        "z_stats": z_stats,
        "components": components,
        "D": D,
    }


def distance_from_components(components: list, alpha: tuple[float, ...]) -> list[list[float]]:
    n = len(components[0])
    D = [[0.0] * n for _ in range(n)]
    for i in range(n):
        for j in range(i + 1, n):
            squared = sum(alpha[k] * components[k][i][j] for k in range(5))
            d = math.sqrt(max(0.0, squared))
            D[i][j] = D[j][i] = d
    return D


# ---------------------------------------------------------------------------
# 精确区域约束 P-Median 求解
# ---------------------------------------------------------------------------

def choose_best(candidates: list[dict]) -> tuple[dict, list[dict]]:
    best_cost = min(x["cost"] for x in candidates)
    ties = [x for x in candidates if abs(x["cost"] - best_cost) <= TIE_TOL]
    chosen = min(ties, key=lambda x: x["cbsa_key"])
    return chosen, ties


def solve_model(results: list[dict], weight_key: str, distance_matrices: dict | None = None,
                collect_candidates: bool = True) -> dict:
    """每区至少1席、总共10席；精确求出唯一额外席位应放在哪一区。"""
    local_results = []
    all_candidate_rows = []

    for region in results:
        rid = region["region_id"]
        cities = region["cities"]
        D = distance_matrices[rid] if distance_matrices is not None else region["D"]
        n = region["n"]
        weights = [c[weight_key] for c in cities]

        one_candidates = []
        for j in range(n):
            cost = sum(weights[i] * D[i][j] for i in range(n))
            cbsa_key = (cities[j]["cbsa"],)
            one_candidates.append({
                "seat_count": 1, "indices": (j,), "cost": cost,
                "cbsa_key": cbsa_key,
                "names": (cities[j]["name"],),
            })

        two_candidates = []
        for j, k in itertools.combinations(range(n), 2):
            cost = 0.0
            for i in range(n):
                nearest = D[i][j] if D[i][j] <= D[i][k] else D[i][k]
                cost += weights[i] * nearest
            cbsa_key = tuple(sorted((cities[j]["cbsa"], cities[k]["cbsa"])))
            two_candidates.append({
                "seat_count": 2, "indices": (j, k), "cost": cost,
                "cbsa_key": cbsa_key,
                "names": tuple(cities[x]["name"] for x in (j, k)),
            })

        best_one, one_ties = choose_best(one_candidates)
        best_two, two_ties = choose_best(two_candidates)
        gain = best_one["cost"] - best_two["cost"]
        if gain < -TIE_TOL:
            raise AssertionError(f"{region['region']} 双城损失高于单城损失，优化计算异常")
        gain = max(0.0, gain)

        local_results.append({
            "region_id": rid,
            "region": region["region"],
            "best_one": best_one,
            "one_ties": one_ties,
            "best_two": best_two,
            "two_ties": two_ties,
            "gain": gain,
            "one_candidates": one_candidates,
            "two_candidates": two_candidates,
        })

        if collect_candidates:
            chosen_one_key = best_one["cbsa_key"]
            chosen_two_key = best_two["cbsa_key"]
            for item in one_candidates + two_candidates:
                is_local_best = (
                    (item["seat_count"] == 1 and abs(item["cost"] - best_one["cost"]) <= TIE_TOL)
                    or (item["seat_count"] == 2 and abs(item["cost"] - best_two["cost"]) <= TIE_TOL)
                )
                all_candidate_rows.append({
                    "模型": "人口加权" if weight_key == "w_pop" else "GDP加权",
                    "地区编号": rid,
                    "地区": region["region"],
                    "席位数": item["seat_count"],
                    "候选城市": "；".join(item["names"]),
                    "候选CBSA": "；".join(item["cbsa_key"]),
                    "加权覆盖损失": item["cost"],
                    "是否地区最优": "是" if is_local_best else "否",
                    "是否最终采用": "否",
                    "单城最优并列数": len(one_ties),
                    "双城最优并列数": len(two_ties),
                })

    # 9 regions each have one seat. The tenth seat goes to the region with
    # the largest reduction from its best one-center to best two-center solution.
    max_gain = max(x["gain"] for x in local_results)
    gain_ties = [x for x in local_results if abs(x["gain"] - max_gain) <= TIE_TOL]
    extra_region = min(gain_ties, key=lambda x: x["region_id"])
    selected_by_region = {}
    region_summary = []
    for item in local_results:
        selected = item["best_two"] if item["region_id"] == extra_region["region_id"] else item["best_one"]
        selected_by_region[item["region_id"]] = list(selected["indices"])
        region_summary.append({
            "模型": "人口加权" if weight_key == "w_pop" else "GDP加权",
            "地区编号": item["region_id"],
            "地区": item["region"],
            "MSA数量": next(r["n"] for r in results if r["region_id"] == item["region_id"]),
            "单城最优损失": item["best_one"]["cost"],
            "双城最优损失": item["best_two"]["cost"],
            "增加第二城的损失下降": item["gain"],
            "单城最优并列数": len(item["one_ties"]),
            "双城最优并列数": len(item["two_ties"]),
            "是否获得第二席": "是" if item["region_id"] == extra_region["region_id"] else "否",
            "最终入选城市": "；".join(
                next(r for r in results if r["region_id"] == item["region_id"])["cities"][i]["name"]
                for i in selected["indices"]
            ),
        })

    selected_rows = []
    for region in results:
        rid = region["region_id"]
        for seat, idx in enumerate(selected_by_region[rid], 1):
            c = region["cities"][idx]
            selected_rows.append({
                "模型": "人口加权" if weight_key == "w_pop" else "GDP加权",
                "地区编号": rid,
                "地区": region["region"],
                "地区席位": seat,
                "MSA": c["name"],
                "CBSA": c["cbsa"],
                "2024人口": c["population"],
                "2024人均GDP_美元": c["gdp_pc"],
                "2024GDP总量_推导美元": c["gdp_total"],
                "人口权重": c["w_pop"],
                "GDP权重": c["w_gdp"],
            })

    best_cost = sum(
        next(x for x in local_results if x["region_id"] == rid)["best_two"]["cost"]
        if rid == extra_region["region_id"]
        else next(x for x in local_results if x["region_id"] == rid)["best_one"]["cost"]
        for rid in selected_by_region
    )

    # Mark final local configurations in candidate export.
    if collect_candidates:
        for row in all_candidate_rows:
            rid = row["地区编号"]
            final_indices = selected_by_region[rid]
            final_cbsas = tuple(sorted(
                next(r for r in results if r["region_id"] == rid)["cities"][i]["cbsa"]
                for i in final_indices
            ))
            if row["候选CBSA"].split("；") == list(final_cbsas):
                row["是否最终采用"] = "是"

    # Verify the total 10-seat / nine-region constraint.
    seat_total = sum(len(v) for v in selected_by_region.values())
    if seat_total != TOTAL_SELECTED or any(len(v) < 1 for v in selected_by_region.values()):
        raise AssertionError(f"最终席位约束失败：总数={seat_total}，地区覆盖={selected_by_region}")

    return {
        "model": "人口加权" if weight_key == "w_pop" else "GDP加权",
        "weight_key": weight_key,
        "selected_by_region": selected_by_region,
        "selected_rows": selected_rows,
        "region_summary": region_summary,
        "candidate_rows": all_candidate_rows,
        "best_cost": best_cost,
        "extra_region": extra_region["region_id"],
        "extra_region_name": extra_region["region"],
        "extra_region_gain": max_gain,
        "extra_region_ties": [x["region_id"] for x in gain_ties],
        "local_results": local_results,
        "seat_total": seat_total,
    }


def build_assignments(results: list[dict], solution: dict, distance_matrices: dict | None = None):
    rows = []
    contributions = 0.0
    weight_key = solution["weight_key"]
    for region in results:
        rid = region["region_id"]
        cities = region["cities"]
        D = distance_matrices[rid] if distance_matrices is not None else region["D"]
        centers = solution["selected_by_region"][rid]
        for i, city in enumerate(cities):
            nearest_distance = min(D[i][j] for j in centers)
            nearest_indices = [j for j in centers if abs(D[i][j] - nearest_distance) <= TIE_TOL]
            nearest_idx = min(nearest_indices, key=lambda j: cities[j]["cbsa"])
            weight = city[weight_key]
            contribution = weight * nearest_distance
            contributions += contribution
            rows.append({
                "模型": solution["model"],
                "地区编号": rid,
                "地区": region["region"],
                "MSA": city["name"],
                "CBSA": city["cbsa"],
                "2024人口": city["population"],
                "2024人均GDP_美元": city["gdp_pc"],
                "2024GDP总量_推导美元": city["gdp_total"],
                "人口权重": city["w_pop"],
                "GDP权重": city["w_gdp"],
                "本模型使用权重": weight,
                "最近代表城市": cities[nearest_idx]["name"],
                "代表城市CBSA": cities[nearest_idx]["cbsa"],
                "并列最近代表城市": "；".join(cities[j]["name"] for j in nearest_indices),
                "最近距离d": nearest_distance,
                "目标函数贡献": contribution,
                "本MSA是否入选": "是" if i in centers else "否",
            })
    if abs(contributions - solution["best_cost"]) > 1e-8:
        raise AssertionError(
            f"{solution['model']} 195城覆盖贡献之和 {contributions:.12f} "
            f"与优化目标值 {solution['best_cost']:.12f} 不一致"
        )
    return rows, contributions


def evaluate_selected(results: list[dict], selected_by_region: dict, weight_key: str,
                      distance_matrices: dict | None = None) -> float:
    total = 0.0
    for region in results:
        rid = region["region_id"]
        D = distance_matrices[rid] if distance_matrices is not None else region["D"]
        centers = selected_by_region[rid]
        for i, city in enumerate(region["cities"]):
            total += city[weight_key] * min(D[i][j] for j in centers)
    return total


# ---------------------------------------------------------------------------
# 蒙特卡洛权重扰动
# ---------------------------------------------------------------------------

def generate_dirichlet_draws(count: int, seed: int) -> list[tuple[float, ...]]:
    rng = random.Random(seed)
    draws = []
    for _ in range(count):
        raw = [rng.gammavariate(DIRICHLET_SHAPE, 1.0) for _ in range(5)]
        total = sum(raw)
        draws.append(tuple(x / total for x in raw))
    return draws


def monte_carlo_model(results: list[dict], weight_key: str, base_solution: dict,
                      alpha_draws: list[tuple[float, ...]], model_label: str):
    frequency = Counter()
    run_rows = []
    jaccards = []
    objectives = []
    baseline_set = {
        row["CBSA"] for row in base_solution["selected_rows"]
    }

    for run_no, alpha in enumerate(alpha_draws, 1):
        matrices = {
            r["region_id"]: distance_from_components(r["components"], alpha)
            for r in results
        }
        solution = solve_model(results, weight_key, distance_matrices=matrices, collect_candidates=False)
        selected_set = {
            c["cbsa"]
            for r in results
            for c in [r["cities"][i] for i in solution["selected_by_region"][r["region_id"]]]
        }
        inter = len(baseline_set & selected_set)
        union = len(baseline_set | selected_set)
        jac = inter / union if union else 1.0
        for cbsa in selected_set:
            frequency[cbsa] += 1
        jaccards.append(jac)
        objectives.append(solution["best_cost"])
        run_rows.append({
            "模型": model_label,
            "模拟序号": run_no,
            "alpha_ln人口": alpha[0],
            "alpha_ln人口密度": alpha[1],
            "alpha_人均GDP": alpha[2],
            "alpha_产业结构E": alpha[3],
            "alpha_人口CAGR": alpha[4],
            "最优目标值": solution["best_cost"],
            "第二席地区": solution["extra_region_name"],
            "第二席地区编号": solution["extra_region"],
            "与基准10城Jaccard": jac,
            "入选CBSA": ";".join(sorted(selected_set)),
        })

    summary = {
        "模型": model_label,
        "模拟次数": len(alpha_draws),
        "基准目标值": base_solution["best_cost"],
        "Jaccard平均值": statistics.mean(jaccards) if jaccards else None,
        "Jaccard最低值": min(jaccards) if jaccards else None,
        "Jaccard第5百分位": percentile(jaccards, 0.05) if jaccards else None,
        "目标值平均数": statistics.mean(objectives) if objectives else None,
        "目标值第5百分位": percentile(objectives, 0.05) if objectives else None,
        "目标值第95百分位": percentile(objectives, 0.95) if objectives else None,
        "第二席地区频数": json.dumps(dict(Counter(r["第二席地区"] for r in run_rows)), ensure_ascii=False),
    }
    return frequency, run_rows, summary


# ---------------------------------------------------------------------------
# 结果表构建
# ---------------------------------------------------------------------------

def build_national_weight_rows(cities_by_cbsa: dict) -> list[dict]:
    rows = []
    for city in sorted(cities_by_cbsa.values(), key=lambda c: (int(c["region_id"]), c["no"])):
        rows.append({
            "地区编号": city["region_id"],
            "地区": city["region"],
            "原始清单序号": city["no"],
            "MSA": city["name"],
            "CBSA": city["cbsa"],
            "2024人口": city["population"],
            "2024人口密度_人每平方英里": city["density"],
            "2024人均GDP_美元": city["gdp_pc"],
            "2020至2024_CAGR_百分比": city["cagr"],
            "2024GDP总量_人口乘人均GDP": city["gdp_total"],
            "人口总量权重": city["w_pop"],
            "GDP经济总量权重": city["w_gdp"],
        })
    return rows


def build_z_rows(results: list[dict]) -> list[dict]:
    rows = []
    for r in results:
        for i, c in enumerate(r["cities"]):
            rows.append({
                "地区编号": r["region_id"], "地区": r["region"],
                "MSA": c["name"], "CBSA": c["cbsa"],
                "ln人口": math.log(c["population"]),
                "ln人口密度": math.log(c["density"]),
                "人均GDP原值": c["gdp_pc"], "CAGR原值": c["cagr"],
                "Z_ln人口": r["Z"][i][0],
                "Z_ln人口密度": r["Z"][i][1],
                "Z_人均GDP": r["Z"][i][2],
                "Z_CAGR": r["Z"][i][3],
            })
    return rows


def build_share_rows(results: list[dict]) -> list[dict]:
    rows = []
    for r in results:
        for i, c in enumerate(r["cities"]):
            row = {"地区编号": r["region_id"], "地区": r["region"], "MSA": c["name"], "CBSA": c["cbsa"]}
            row.update({sector: r["shares"][i][k] for k, sector in enumerate(r["sectors"])})
            row["占比合计"] = sum(r["shares"][i])
            rows.append(row)
    return rows


def build_employment_rows(results: list[dict]) -> list[dict]:
    rows = []
    for r in results:
        for i, c in enumerate(r["cities"]):
            row = {"地区编号": r["region_id"], "地区": r["region"], "MSA": c["name"], "CBSA": c["cbsa"]}
            row.update({sector: r["employment"][i][k] for k, sector in enumerate(r["sectors"])})
            row["13类就业合计"] = r["employment_totals"][i]
            rows.append(row)
    return rows


def build_parameter_rows(results: list[dict]) -> list[dict]:
    rows = []
    for r in results:
        for stat in r["z_stats"]:
            rows.append({
                "地区编号": r["region_id"], "地区": r["region"], "MSA数量": r["n"],
                "配对数": r["pair_count"], "H_RMS": r["h_rms"], "K": r["K"],
                "指标": stat["指标"], "变换后均值": stat["变换后均值"],
                "总体标准差_ddof0": stat["总体标准差_ddof0"],
            })
    return rows


def build_pairwise_rows(results: list[dict]) -> list[dict]:
    rows = []
    for r in results:
        for i in range(r["n"]):
            for j in range(i + 1, r["n"]):
                rows.append({
                    "地区编号": r["region_id"], "地区": r["region"],
                    "CBSA_i": r["cities"][i]["cbsa"], "MSA_i": r["cities"][i]["name"],
                    "CBSA_j": r["cities"][j]["cbsa"], "MSA_j": r["cities"][j]["name"],
                    "Hellinger_H": r["H"][i][j], "校准后产业距离_E": r["E"][i][j],
                    "五指标基准距离_d": r["D"][i][j],
                })
    return rows


def build_cross_evaluation(results: list[dict], pop_solution: dict, gdp_solution: dict) -> list[dict]:
    rows = []
    solutions = [("人口模型入选10城", pop_solution), ("GDP模型入选10城", gdp_solution)]
    for selection_label, selection in solutions:
        for eval_label, weight_key in [("人口权重评价", "w_pop"), ("GDP权重评价", "w_gdp")]:
            loss = evaluate_selected(results, selection["selected_by_region"], weight_key)
            rows.append({
                "城市组合来源": selection_label,
                "评价标准": eval_label,
                "加权覆盖损失": loss,
                "入选10城": "；".join(row["MSA"] for row in selection["selected_rows"]),
                "第二席地区": selection["extra_region_name"],
            })
    return rows


def build_region_model_summary(pop_solution: dict, gdp_solution: dict) -> list[dict]:
    return pop_solution["region_summary"] + gdp_solution["region_summary"]


def build_mc_frequency_rows(results: list[dict], pop_freq: Counter, gdp_freq: Counter,
                            mc_count: int) -> list[dict]:
    rows = []
    for r in results:
        for c in r["cities"]:
            p = pop_freq.get(c["cbsa"], 0)
            g = gdp_freq.get(c["cbsa"], 0)
            rows.append({
                "地区编号": c["region_id"], "地区": c["region"], "MSA": c["name"], "CBSA": c["cbsa"],
                "人口模型入选次数": p, "人口模型入选频率": p / mc_count if mc_count else None,
                "GDP模型入选次数": g, "GDP模型入选频率": g / mc_count if mc_count else None,
            })
    return rows


# ---------------------------------------------------------------------------
# Excel 格式化
# ---------------------------------------------------------------------------

def add_sheet(wb: Workbook, title: str, headers: list[str], rows: list, widths: list[int] | None = None):
    ws = wb.create_sheet(safe_sheet_name(title))
    ws.append(headers)
    for row in rows:
        if isinstance(row, dict):
            ws.append([row.get(h) for h in headers])
        else:
            ws.append(list(row))
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    if ws.max_row >= 1 and ws.max_column >= 1:
        ws.auto_filter.ref = ws.dimensions
    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 32

    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        if widths and col_idx <= len(widths):
            width = widths[col_idx - 1]
        else:
            vals = [str(ws.cell(row, col_idx).value or "") for row in range(1, min(ws.max_row, 100) + 1)]
            width = min(max(max((len(v) for v in vals), default=10) + 2, 12), 38)
        ws.column_dimensions[letter].width = width

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top")
            header = headers[cell.column - 1]
            if isinstance(cell.value, float):
                if "权重" in header or "频率" in header or "占比" in header or "Jaccard" in header or "alpha_" in header:
                    cell.number_format = "0.0000%"
                elif any(x in header for x in ["人口", "就业", "GDP总量", "人均GDP"]):
                    cell.number_format = "#,##0.00"
                else:
                    cell.number_format = "0.000000"
    return ws


def add_matrix_sheet(wb: Workbook, title: str, cities: list[dict], matrix: list[list[float]]):
    headers = ["CBSA", "MSA"] + [f"{c['name']} [{c['cbsa']}]" for c in cities]
    rows = []
    for i, city in enumerate(cities):
        rows.append([city["cbsa"], city["name"]] + matrix[i])
    widths = [12, 42] + [18] * len(cities)
    return add_sheet(wb, title, headers, rows, widths)


def build_workbook(path: Path, results: list[dict], five_rows: list[dict], industry_rows: list[dict],
                   sector_order: list[str], cities_by_cbsa: dict, qa_rows: list[dict],
                   source_audit_rows: list[dict], pop_solution: dict, gdp_solution: dict,
                   pop_assign: list[dict], gdp_assign: list[dict], cross_rows: list[dict],
                   candidate_rows: list[dict], mc_freq_rows: list[dict], mc_run_rows: list[dict],
                   mc_summary_rows: list[dict], mc_count: int, seed: int, db_hash: str):
    wb = Workbook()
    default = wb.active
    wb.remove(default)

    summary_rows = [
        {"模型": "V3-P 人口总量加权", "最终城市数": 10, "地区覆盖数": 9,
         "最优加权覆盖损失": pop_solution["best_cost"], "第二席地区": pop_solution["extra_region_name"],
         "第二席边际损失下降": pop_solution["extra_region_gain"],
         "第二席并列地区": ";".join(pop_solution["extra_region_ties"])},
        {"模型": "V3-G GDP经济总量加权", "最终城市数": 10, "地区覆盖数": 9,
         "最优加权覆盖损失": gdp_solution["best_cost"], "第二席地区": gdp_solution["extra_region_name"],
         "第二席边际损失下降": gdp_solution["extra_region_gain"],
         "第二席并列地区": ";".join(gdp_solution["extra_region_ties"])},
    ]
    add_sheet(wb, "总览", ["模型", "最终城市数", "地区覆盖数", "最优加权覆盖损失", "第二席地区", "第二席边际损失下降", "第二席并列地区"], summary_rows)
    add_sheet(wb, "人口模型_最终10城", list(pop_solution["selected_rows"][0].keys()), pop_solution["selected_rows"])
    add_sheet(wb, "GDP模型_最终10城", list(gdp_solution["selected_rows"][0].keys()), gdp_solution["selected_rows"])
    add_sheet(wb, "双模型交叉评价", list(cross_rows[0].keys()), cross_rows)
    add_sheet(wb, "九区席位与损失", list((pop_solution["region_summary"] + gdp_solution["region_summary"])[0].keys()),
              pop_solution["region_summary"] + gdp_solution["region_summary"])

    national_weights = build_national_weight_rows(cities_by_cbsa)
    add_sheet(wb, "全国195城与双权重", list(national_weights[0].keys()), national_weights)
    add_sheet(wb, "人口模型_覆盖分配", list(pop_assign[0].keys()), pop_assign)
    add_sheet(wb, "GDP模型_覆盖分配", list(gdp_assign[0].keys()), gdp_assign)
    add_sheet(wb, "全部候选组合", list(candidate_rows[0].keys()), candidate_rows)
    add_sheet(wb, "蒙特卡洛_稳定性摘要", list(mc_summary_rows[0].keys()), mc_summary_rows)
    add_sheet(wb, "蒙特卡洛_入选频率", list(mc_freq_rows[0].keys()), mc_freq_rows)
    if mc_run_rows:
        add_sheet(wb, "蒙特卡洛_逐次结果", list(mc_run_rows[0].keys()), mc_run_rows)

    raw_five = []
    for row in five_rows:
        raw_five.append({k: row.get(k) for k in [
            "地区编号", "地区", "原始清单序号", "MSA", "CBSA", "2024人口",
            "2024人口密度_人每平方英里", "2024人均GDP_美元", "2020至2024_CAGR_百分比"
        ]})
    add_sheet(wb, "全国195城原始指标", list(raw_five[0].keys()), raw_five)

    emp_rows = build_employment_rows(results)
    add_sheet(wb, "全国产业就业矩阵", list(emp_rows[0].keys()), emp_rows)
    share_rows = build_share_rows(results)
    add_sheet(wb, "全国产业占比矩阵", list(share_rows[0].keys()), share_rows)
    z_rows = build_z_rows(results)
    add_sheet(wb, "全国Z标准化", list(z_rows[0].keys()), z_rows)
    parameter_rows = build_parameter_rows(results)
    add_sheet(wb, "九区标准化参数", list(parameter_rows[0].keys()), parameter_rows)

    for r in results:
        code = r["short"]
        add_matrix_sheet(wb, f"H_{code}", r["cities"], r["H"])
        add_matrix_sheet(wb, f"E_{code}", r["cities"], r["E"])
        add_matrix_sheet(wb, f"d_{code}", r["cities"], r["D"])

    add_sheet(wb, "数据质量检查", list(qa_rows[0].keys()), qa_rows)
    if source_audit_rows:
        source_headers = list(source_audit_rows[0].keys())
        add_sheet(wb, "原始来源审计", source_headers, source_audit_rows)

    methodology = [
        {"项目": "候选范围", "固定定义": "全国195个MSA；九大Census Divisions；太平洋31城作为一个完整地区。"},
        {"项目": "最终样本", "固定定义": "10座城市；九个地区各至少1座；不设任何城市规模配额。"},
        {"项目": "五项特征", "固定定义": "ln人口、ln人口密度、人均GDP、13类产业结构E、人口CAGR。"},
        {"项目": "连续变量标准化", "固定定义": "ln人口、ln密度、人均GDP、CAGR分别在本地区内做Z-score，ddof=0。"},
        {"项目": "产业份额", "固定定义": "各MSA的13类行业就业人数除以该MSA的13类行业就业合计。"},
        {"项目": "Hellinger", "固定定义": "H_ij=sqrt(0.5*Σ_k(sqrt(s_ik)-sqrt(s_jk))^2)。"},
        {"项目": "RMS校准", "固定定义": "H_RMS=sqrt(mean(H_ij^2,i<j))；K=sqrt(2n/(n-1))/H_RMS；E=K*H。"},
        {"项目": "五指标基准距离", "固定定义": "d=sqrt(0.2*(四个Z差平方之和+E^2))，五项各20%。"},
        {"项目": "V3-P人口权重", "固定定义": "w_i=P_i/ΣP，分母为195个候选MSA的人口总和。"},
        {"项目": "V3-G GDP权重", "固定定义": "GDP总量=P_i*GDPpc_i；w_i=GDP总量_i/ΣGDP总量，分母为195城GDP总和。"},
        {"项目": "GDP口径说明", "固定定义": "程序由数据库2024人均GDP×2024人口推导GDP总量；不联网、不补造；必须先核验人均GDP口径。"},
        {"项目": "人口模型目标函数", "固定定义": "L_P(S)=Σ_i w_i^P min_{j∈S_r(i)} d_ij^(r)。"},
        {"项目": "GDP模型目标函数", "固定定义": "L_G(S)=Σ_i w_i^G min_{j∈S_r(i)} d_ij^(r)。"},
        {"项目": "优化约束", "固定定义": "|S|=10；每个地区至少1座；不设置大型、中小型、特大型名额。"},
        {"项目": "精确求解", "固定定义": "每区枚举单城与双城；总席位10且九区覆盖意味着唯一一个地区获得第二席；将第二席给边际覆盖损失下降最大的地区。"},
        {"项目": "并列规则", "固定定义": f"数值容差={TIE_TOL}；并列数量记录；正式输出按CBSA字典序作可重复的确定性选择。"},
        {"项目": "蒙特卡洛", "固定定义": f"alpha~Dirichlet({DIRICHLET_SHAPE},{DIRICHLET_SHAPE},{DIRICHLET_SHAPE},{DIRICHLET_SHAPE},{DIRICHLET_SHAPE})；默认每模型{mc_count}次；种子{seed}。"},
        {"项目": "敏感性距离", "固定定义": "d(alpha)=sqrt(Σ_k alpha_k*component_k)，alpha五项之和为1。"},
        {"项目": "稳定性", "固定定义": "逐次10城名单与基准名单的Jaccard；报告均值、最低值、第5百分位及城市入选频率。"},
        {"项目": "交叉评价", "固定定义": "人口模型名单分别按人口/GDP权重评价；GDP模型名单也分别按人口/GDP权重评价。"},
        {"项目": "原始数据保护", "固定定义": "输入SQLite只读；输出SQLite为原库副本并追加V3中间结果表。"},
        {"项目": "输入SHA256", "固定定义": db_hash},
    ]
    add_sheet(wb, "模型定义与公式", ["项目", "固定定义"], methodology, [26, 120])

    wb.properties.title = "美国九大地区 V3 人口与GDP双权重 P-Median 代表性城市筛选"
    wb.properties.subject = "195个MSA；九区距离；人口与GDP总量权重独立优化；蒙特卡洛稳定性检验"
    wb.properties.creator = "OpenAI"
    wb.save(path)


# ---------------------------------------------------------------------------
# CSV 与 SQLite 输出
# ---------------------------------------------------------------------------

def write_sqlite_output(input_db: Path, output_db: Path, results: list[dict],
                        national_weights: list[dict], pairwise_rows: list[dict],
                        pop_solution: dict, gdp_solution: dict,
                        pop_assign: list[dict], gdp_assign: list[dict],
                        candidate_rows: list[dict], cross_rows: list[dict],
                        mc_freq_rows: list[dict], mc_run_rows: list[dict],
                        mc_summary_rows: list[dict], metadata_rows: list[dict]):
    if output_db.exists():
        output_db.unlink()
    src = sqlite3.connect(str(input_db))
    dst = sqlite3.connect(str(output_db))
    try:
        src.backup(dst)
    finally:
        src.close()
    dst.row_factory = sqlite3.Row
    try:
        tables = {
            "v3_city_weights": national_weights,
            "v3_pairwise_distances": pairwise_rows,
            "v3_selected_population": pop_solution["selected_rows"],
            "v3_selected_gdp": gdp_solution["selected_rows"],
            "v3_assignments_population": pop_assign,
            "v3_assignments_gdp": gdp_assign,
            "v3_candidate_combinations": candidate_rows,
            "v3_cross_evaluation": cross_rows,
            "v3_mc_frequency": mc_freq_rows,
            "v3_mc_runs": mc_run_rows,
            "v3_mc_summary": mc_summary_rows,
            "v3_metadata": metadata_rows,
        }
        for table, rows in tables.items():
            dst.execute(f'DROP TABLE IF EXISTS "{table}"')
            if not rows:
                continue
            headers = list(rows[0].keys())
            type_map = {}
            for h in headers:
                sample = next((row.get(h) for row in rows if row.get(h) is not None), None)
                if isinstance(sample, bool):
                    type_map[h] = "INTEGER"
                elif isinstance(sample, int):
                    type_map[h] = "INTEGER"
                elif isinstance(sample, float):
                    type_map[h] = "REAL"
                else:
                    type_map[h] = "TEXT"
            col_sql = ", ".join(f'"{h}" {type_map[h]}' for h in headers)
            dst.execute(f'CREATE TABLE "{table}" ({col_sql})')
            placeholders = ",".join("?" for _ in headers)
            sql = f'INSERT INTO "{table}" VALUES ({placeholders})'
            dst.executemany(sql, [tuple(row.get(h) for h in headers) for row in rows])
        dst.commit()
    finally:
        dst.close()


def main():
    parser = argparse.ArgumentParser(
        description="全国195个MSA：人口总量与GDP总量双权重独立P-Median优化及蒙特卡洛检验。"
    )
    parser.add_argument("--db", type=Path, default=DEFAULT_DB, help="全国SQLite数据库路径")
    parser.add_argument("--mc", type=int, default=1000, help="每套模型蒙特卡洛次数，默认1000")
    parser.add_argument("--seed", type=int, default=20261116, help="蒙特卡洛随机种子")
    parser.add_argument("--skip-mc", action="store_true", help="跳过蒙特卡洛，仅用于调试")
    args = parser.parse_args()
    if args.mc < 1 and not args.skip_mc:
        parser.error("--mc 必须大于等于1；如需跳过请使用 --skip-mc")

    started = time.time()
    log_lines = [
        "美国九大地区 V3.0 双权重 P-Median 计算日志",
        f"开始时间：{datetime.now().isoformat(timespec='seconds')}",
        f"输入数据库：{args.db.resolve()}",
        f"蒙特卡洛次数/模型：{0 if args.skip_mc else args.mc}",
        f"随机种子：{args.seed}",
    ]

    try:
        (region_data, five_rows, industry_rows, sector_order, cities_by_cbsa,
         qa_rows, source_audit_rows, db_hash) = load_database(args.db, log_lines)

        results = []
        log("开始九区重新计算 Hellinger → RMS校准E → Z-score → 五指标d……", log_lines)
        for region_id, region_name, short, expected_n in REGIONS:
            r = calculate_region(region_data[region_id])
            results.append(r)
            log(
                f"OK {region_name}：{r['n']}城；配对={r['pair_count']}；"
                f"H_RMS={r['h_rms']:.10f}；K={r['K']:.10f}",
                log_lines,
            )

        log("\n开始 V3-P 人口总量加权精确优化……", log_lines)
        pop_solution = solve_model(results, "w_pop", collect_candidates=True)
        log(
            f"V3-P 最优损失={pop_solution['best_cost']:.12f}；"
            f"第二席地区={pop_solution['extra_region_name']}；"
            f"边际损失下降={pop_solution['extra_region_gain']:.12f}",
            log_lines,
        )
        for row in pop_solution["selected_rows"]:
            log(f"  + {row['地区']} | {row['MSA']} | CBSA {row['CBSA']}", log_lines)

        log("\n开始 V3-G GDP经济总量加权精确优化……", log_lines)
        gdp_solution = solve_model(results, "w_gdp", collect_candidates=True)
        log(
            f"V3-G 最优损失={gdp_solution['best_cost']:.12f}；"
            f"第二席地区={gdp_solution['extra_region_name']}；"
            f"边际损失下降={gdp_solution['extra_region_gain']:.12f}",
            log_lines,
        )
        for row in gdp_solution["selected_rows"]:
            log(f"  + {row['地区']} | {row['MSA']} | CBSA {row['CBSA']}", log_lines)

        pop_assign, pop_assign_sum = build_assignments(results, pop_solution)
        gdp_assign, gdp_assign_sum = build_assignments(results, gdp_solution)
        cross_rows = build_cross_evaluation(results, pop_solution, gdp_solution)
        candidate_rows = pop_solution["candidate_rows"] + gdp_solution["candidate_rows"]
        log(f"人口模型195城贡献和={pop_assign_sum:.12f}（应等于目标值）", log_lines)
        log(f"GDP模型195城贡献和={gdp_assign_sum:.12f}（应等于目标值）", log_lines)

        mc_freq_rows = []
        mc_run_rows = []
        mc_summary_rows = []
        pop_freq = Counter()
        gdp_freq = Counter()
        if args.skip_mc:
            log("按 --skip-mc 指令跳过蒙特卡洛；正式研究报告不应使用此调试结果。", log_lines)
            mc_summary_rows = [
                {"模型": "人口加权", "模拟次数": 0, "说明": "本次调试跳过"},
                {"模型": "GDP加权", "模拟次数": 0, "说明": "本次调试跳过"},
            ]
        else:
            log(f"\n开始蒙特卡洛敏感性检验：每模型{args.mc}次……", log_lines)
            alpha_draws = generate_dirichlet_draws(args.mc, args.seed)
            pop_freq, pop_runs, pop_summary = monte_carlo_model(
                results, "w_pop", pop_solution, alpha_draws, "人口加权"
            )
            log(
                f"人口模型MC完成：平均Jaccard={pop_summary['Jaccard平均值']:.4f}；"
                f"最低={pop_summary['Jaccard最低值']:.4f}；"
                f"P5={pop_summary['Jaccard第5百分位']:.4f}",
                log_lines,
            )
            gdp_freq, gdp_runs, gdp_summary = monte_carlo_model(
                results, "w_gdp", gdp_solution, alpha_draws, "GDP加权"
            )
            log(
                f"GDP模型MC完成：平均Jaccard={gdp_summary['Jaccard平均值']:.4f}；"
                f"最低={gdp_summary['Jaccard最低值']:.4f}；"
                f"P5={gdp_summary['Jaccard第5百分位']:.4f}",
                log_lines,
            )
            mc_run_rows = pop_runs + gdp_runs
            mc_summary_rows = [pop_summary, gdp_summary]
            mc_freq_rows = build_mc_frequency_rows(results, pop_freq, gdp_freq, args.mc)

        if not mc_freq_rows:
            mc_freq_rows = build_mc_frequency_rows(results, pop_freq, gdp_freq, 1)
            for row in mc_freq_rows:
                row["人口模型入选频率"] = None
                row["GDP模型入选频率"] = None

        national_weights = build_national_weight_rows(cities_by_cbsa)
        pairwise_rows = build_pairwise_rows(results)
        metadata_rows = [
            {"键": "model_version", "值": "V3.0"},
            {"键": "input_database", "值": str(args.db.resolve())},
            {"键": "input_sha256", "值": db_hash},
            {"键": "population_weight_formula", "值": "population / sum(population over 195 MSAs)"},
            {"键": "gdp_total_formula", "值": "2024 population * 2024 GDP per capita"},
            {"键": "gdp_weight_formula", "值": "gdp_total / sum(gdp_total over 195 MSAs)"},
            {"键": "base_indicator_weights", "值": "0.2,0.2,0.2,0.2,0.2"},
            {"键": "region_constraint", "值": "9 regions, >=1 selected MSA per region"},
            {"键": "total_selected", "值": "10"},
            {"键": "size_quota", "值": "none"},
            {"键": "monte_carlo_count_per_model", "值": str(0 if args.skip_mc else args.mc)},
            {"键": "monte_carlo_seed", "值": str(args.seed)},
            {"键": "monte_carlo_distribution", "值": "Dirichlet(20,20,20,20,20)"},
            {"键": "timestamp", "值": datetime.now().isoformat(timespec='seconds')},
        ]

        log("\n正在生成 Excel、CSV 和独立中间 SQLite……", log_lines)
        csv_write(OUT_SELECTED_P, pop_solution["selected_rows"])
        csv_write(OUT_SELECTED_G, gdp_solution["selected_rows"])
        csv_write(OUT_ASSIGN, pop_assign + gdp_assign)
        csv_write(OUT_CANDIDATES, candidate_rows)
        csv_write(OUT_CROSS, cross_rows)
        csv_write(OUT_MC_FREQ, mc_freq_rows)
        csv_write(OUT_MC_RUNS, mc_run_rows)

        build_workbook(
            OUT_XLSX, results, five_rows, industry_rows, sector_order, cities_by_cbsa,
            qa_rows, source_audit_rows, pop_solution, gdp_solution,
            pop_assign, gdp_assign, cross_rows, candidate_rows,
            mc_freq_rows, mc_run_rows, mc_summary_rows,
            0 if args.skip_mc else args.mc, args.seed, db_hash,
        )
        write_sqlite_output(
            args.db, OUT_SQLITE, results, national_weights, pairwise_rows,
            pop_solution, gdp_solution, pop_assign, gdp_assign,
            candidate_rows, cross_rows, mc_freq_rows, mc_run_rows,
            mc_summary_rows, metadata_rows,
        )

        elapsed = time.time() - started
        log("\n全部计算完成。", log_lines)
        log(f"V3-P 入选10城：{len(pop_solution['selected_rows'])}；目标值={pop_solution['best_cost']:.12f}", log_lines)
        log(f"V3-G 入选10城：{len(gdp_solution['selected_rows'])}；目标值={gdp_solution['best_cost']:.12f}", log_lines)
        log(f"两模型共同入选城市数：{len({x['CBSA'] for x in pop_solution['selected_rows']} & {x['CBSA'] for x in gdp_solution['selected_rows']})}", log_lines)
        log(f"五指标距离配对数：{len(pairwise_rows)}；区域候选组合记录：{len(candidate_rows)}", log_lines)
        log(f"耗时：{elapsed:.1f}秒", log_lines)
        log(f"Excel：{OUT_XLSX.resolve()}", log_lines)
        log(f"SQLite：{OUT_SQLITE.resolve()}", log_lines)
        log(f"人口模型名单：{OUT_SELECTED_P.resolve()}", log_lines)
        log(f"GDP模型名单：{OUT_SELECTED_G.resolve()}", log_lines)
        log(f"覆盖分配：{OUT_ASSIGN.resolve()}", log_lines)
        log(f"候选组合：{OUT_CANDIDATES.resolve()}", log_lines)
        log(f"交叉评价：{OUT_CROSS.resolve()}", log_lines)
        log(f"蒙特卡洛频率：{OUT_MC_FREQ.resolve()}", log_lines)
        log(f"蒙特卡洛逐次结果：{OUT_MC_RUNS.resolve()}", log_lines)
        log("注意：GDP总量是由数据库内2024人均GDP×2024人口推导；请在论文中说明其来源和边界口径。", log_lines)

        OUT_LOG.write_text("\n".join(log_lines) + "\n", encoding="utf-8-sig")
        print(f"日志：{OUT_LOG.resolve()}")

    except Exception as exc:
        log_lines.append(f"\n计算失败：{type(exc).__name__}: {exc}")
        try:
            OUT_LOG.write_text("\n".join(log_lines) + "\n", encoding="utf-8-sig")
        except Exception:
            pass
        print("\n计算未完成，未将结果标记为正式完成。")
        print(f"原因：{type(exc).__name__}: {exc}")
        sys.exit(2)


if __name__ == "__main__":
    main()
