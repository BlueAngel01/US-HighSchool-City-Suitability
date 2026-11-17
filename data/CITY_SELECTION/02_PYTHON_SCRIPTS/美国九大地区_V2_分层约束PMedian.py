# -*- coding: utf-8 -*-
"""
美国九大地区 V2：分层约束加权 P-Median 联合选城程序
====================================================

冻结的主模型：
1. 全国 195 个 MSA，九大 Census Divisions 分区计算；Pacific 31 城不拆分。
2. 五指标各 20%：
   ln(2024人口)、ln(人口密度)、2024人均GDP、13类产业结构E、2020-2024人口CAGR。
3. 各地区独立计算 Hellinger -> RMS校准E -> 区域内Z-score -> 五指标距离d。
4. 最终选择10座：大型(人口>=1,000,000)恰好5座；
   中小型(250,000<=人口<1,000,000)恰好5座；
   特大型(人口>=5,000,000)至少2座（属于大型组）。
5. 九大地区各至少1座。因总数10、地区9且每区至少1座，
   每区只能选1或2座，且恰有一个地区选2座。
6. 每个地区内部，大型组和中小型组分别占该地区覆盖权重的50%；
   组内按 sqrt(人口) 归一化。
7. 目标：最小化九区加权最近代表城市距离（分层约束P-Median）。
8. 使用动态规划枚举各区所有单城/双城组合，求全局精确最优解；
   不使用随机聚类，不使用20%第二城市门槛。

输入：
  美国九大地区_2024MSA综合数据库.sqlite
  数据表：
    msa_five_indicators
    msa_industry_detail

输出（均保存在程序所在文件夹）：
  美国九大地区_V2_最终结果与中间数据.xlsx
  美国九大地区_V2_最终选城名单.csv
  美国九大地区_V2_195城覆盖分配.csv
  美国九大地区_V2_全部区域候选组合.csv
  美国九大地区_V2_中间数据.sqlite
  美国九大地区_V2_计算日志.txt

依赖：
  py -m pip install openpyxl

运行：
  cd "$env:USERPROFILE\\Downloads"
  py ".\\美国九大地区_V2_分层约束PMedian.py"

数据库不在同一目录时：
  py ".\\美国九大地区_V2_分层约束PMedian.py" --db "D:\\data\\美国九大地区_2024MSA综合数据库.sqlite"

注意：
- 程序只读取原始SQLite数据库，不会修改或覆盖原数据库。
- 正式计算前执行严格数据核验；缺失、重复、类别不完整时停止。
- V2的分层名额、人口权重、区域权重均写入结果文件，便于论文复核。
"""

from __future__ import annotations

import argparse
import csv
import itertools
import math
import sqlite3
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import ColorScaleRule
    from openpyxl.comments import Comment
except ImportError:
    raise SystemExit("缺少 openpyxl。请先运行：py -m pip install openpyxl")


BASE = Path(__file__).resolve().parent
DEFAULT_DB = BASE / "美国九大地区_2024MSA综合数据库.sqlite"

OUT_XLSX = BASE / "美国九大地区_V2_最终结果与中间数据.xlsx"
OUT_SELECTED_CSV = BASE / "美国九大地区_V2_最终选城名单.csv"
OUT_ASSIGN_CSV = BASE / "美国九大地区_V2_195城覆盖分配.csv"
OUT_CONFIG_CSV = BASE / "美国九大地区_V2_全部区域候选组合.csv"
OUT_SQLITE = BASE / "美国九大地区_V2_中间数据.sqlite"
OUT_LOG = BASE / "美国九大地区_V2_计算日志.txt"

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

POP_LARGE = 1_000_000
POP_MEGA = 5_000_000
POP_MIN = 250_000
TOTAL_SELECTED = 10
LARGE_QUOTA = 5
SMALL_QUOTA = 5
MEGA_MIN = 2
POP_WEIGHT_EXPONENT = 0.5
REGION_WEIGHT = 1.0 / 9.0
GROUP_WEIGHT = 0.5
CITY_WEIGHT_DENOMINATOR = 9.0 * 2.0
TIE_TOL = 1e-10

FIVE_TABLE = "msa_five_indicators"
INDUSTRY_TABLE = "msa_industry_detail"

REGION_SHEETS = {
    "五指标": "五指标",
    "产业就业": "产业就业",
    "产业占比": "产业占比",
    "H": "H矩阵",
    "E": "E矩阵",
    "Z": "Z标准化",
    "d": "d矩阵",
    "候选": "候选组合",
}


def log_line(lines, message):
    print(message)
    lines.append(str(message))


def norm_code(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    s = str(value).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def as_float(value, field, msa):
    if value is None or str(value).strip() == "":
        raise ValueError(f"{msa} 的字段“{field}”为空")
    try:
        x = float(str(value).replace(",", "").replace("$", "").strip())
    except Exception:
        raise ValueError(f"{msa} 的字段“{field}”无法转为数字：{value!r}")
    if not math.isfinite(x):
        raise ValueError(f"{msa} 的字段“{field}”不是有限数值：{value!r}")
    return x


def as_int(value, field, msa):
    x = as_float(value, field, msa)
    if abs(x - round(x)) > 1e-8:
        raise ValueError(f"{msa} 的字段“{field}”应为整数，实际为{x}")
    return int(round(x))


def table_columns(conn, table):
    return {row[1] for row in conn.execute(f'PRAGMA table_info("{table}")').fetchall()}


def require_columns(conn, table, required):
    cols = table_columns(conn, table)
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
            "或使用 --db 指定数据库路径。"
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
        )
    if len(industry_rows) != EXPECTED_INDUSTRY_ROWS:
        raise ValueError(
            f"全国产业明细应有{EXPECTED_INDUSTRY_ROWS}条，实际{len(industry_rows)}条。"
        )

    five_by_region = defaultdict(list)
    global_cbsa = set()
    city_by_cbsa = {}
    for row in five_rows:
        region_id = norm_code(row["地区编号"])
        msa = str(row["MSA"]).strip()
        cbsa = norm_code(row["CBSA"])
        if region_id not in REGION_BY_ID:
            raise ValueError(f"{msa} 的地区编号异常：{region_id}")
        if not cbsa:
            raise ValueError(f"{msa} 的CBSA为空")
        if cbsa in global_cbsa:
            raise ValueError(f"全国五指标表出现重复CBSA：{cbsa}")
        global_cbsa.add(cbsa)

        city = {
            "region_id": region_id,
            "region": str(row["地区"]).strip(),
            "no": as_int(row["原始清单序号"], "原始清单序号", msa),
            "name": msa,
            "cbsa": cbsa,
            "population": as_float(row["2024人口"], "2024人口", msa),
            "density": as_float(row["2024人口密度_人每平方英里"], "2024人口密度_人每平方英里", msa),
            "gdp_pc": as_float(row["2024人均GDP_美元"], "2024人均GDP_美元", msa),
            "cagr": as_float(row["2020至2024_CAGR_百分比"], "2020至2024_CAGR_百分比", msa),
        }
        if city["population"] < POP_MIN:
            raise ValueError(f"{msa}人口低于数据库既定门槛25万人：{city['population']}")
        if city["density"] <= 0 or city["gdp_pc"] <= 0:
            raise ValueError(f"{msa} 的人口密度、人均GDP必须大于0")
        if city["region"] != REGION_BY_ID[region_id][1]:
            raise ValueError(
                f"{msa} 的地区文字“{city['region']}”与地区编号{region_id}不一致"
            )
        city["size_group"] = "大型" if city["population"] >= POP_LARGE else "中小型"
        city["is_mega"] = city["population"] >= POP_MEGA
        five_by_region[region_id].append(city)
        city_by_cbsa[cbsa] = city

    # Preserve first-appearance order of the 13 industry categories in the database.
    sector_order = []
    jobs_by_cbsa = defaultdict(dict)
    industry_counts = defaultdict(int)
    industry_region = {}
    industry_names = {}
    for row in industry_rows:
        region_id = norm_code(row["地区编号"])
        cbsa = norm_code(row["CBSA"])
        msa = str(row["MSA"]).strip()
        sector = str(row["行业"]).strip()
        if not sector:
            raise ValueError(f"{msa} 存在空行业名称")
        if sector not in sector_order:
            sector_order.append(sector)
        if cbsa not in city_by_cbsa:
            raise ValueError(f"产业表出现五指标表之外的CBSA：{cbsa} ({msa})")
        if cbsa in jobs_by_cbsa and sector in jobs_by_cbsa[cbsa]:
            raise ValueError(f"{msa} 的行业“{sector}”重复")
        male = as_int(row["男性就业"], "男性就业", msa)
        female = as_int(row["女性就业"], "女性就业", msa)
        total = as_int(row["就业合计"], "就业合计", msa)
        if min(male, female, total) < 0:
            raise ValueError(f"{msa} {sector} 的就业人数不能为负")
        if male + female != total:
            raise ValueError(f"{msa} {sector} 男性+女性就业人数不等于就业合计")
        if norm_code(row["地区编号"]) != city_by_cbsa[cbsa]["region_id"]:
            raise ValueError(f"{msa} 五指标表与产业表地区编号不一致")
        jobs_by_cbsa[cbsa][sector] = total
        industry_counts[cbsa] += 1
        industry_region[cbsa] = region_id
        industry_names[cbsa] = msa

    if len(sector_order) != EXPECTED_SECTORS:
        raise ValueError(f"全国行业类别应为13类，实际{len(sector_order)}类：{sector_order}")
    if set(jobs_by_cbsa) != global_cbsa:
        missing = sorted(global_cbsa - set(jobs_by_cbsa))
        extra = sorted(set(jobs_by_cbsa) - global_cbsa)
        raise ValueError(f"五指标与产业表CBSA不一致。缺产业记录={missing}；多余产业记录={extra}")

    for cbsa in global_cbsa:
        if industry_counts[cbsa] != EXPECTED_SECTORS:
            raise ValueError(
                f"{industry_names[cbsa]} 应有13条产业记录，实际{industry_counts[cbsa]}条"
            )
        if set(jobs_by_cbsa[cbsa]) != set(sector_order):
            raise ValueError(f"{industry_names[cbsa]} 的13类行业不完整")
        if sum(jobs_by_cbsa[cbsa].values()) <= 0:
            raise ValueError(f"{industry_names[cbsa]} 的13类行业就业总数为0")

    region_data = {}
    for region_id, region_name, short, expected_n in REGIONS:
        cities = sorted(five_by_region.get(region_id, []), key=lambda x: x["no"])
        if len(cities) != expected_n:
            raise ValueError(
                f"{region_name} 应有{expected_n}城，实际{len(cities)}城。"
            )
        if len([c for c in cities if c["size_group"] == "大型"]) == 0:
            raise ValueError(f"{region_name} 没有大型MSA，无法执行冻结配额")
        if len([c for c in cities if c["size_group"] == "中小型"]) == 0:
            raise ValueError(f"{region_name} 没有中小型MSA，无法执行冻结配额")
        region_data[region_id] = {
            "region_id": region_id,
            "region": region_name,
            "short": short,
            "expected_n": expected_n,
            "cities": cities,
            "sectors": sector_order,
            "jobs_by_cbsa": jobs_by_cbsa,
        }

    mega_count = sum(c["is_mega"] for c in city_by_cbsa.values())
    large_count = sum(c["size_group"] == "大型" for c in city_by_cbsa.values())
    small_count = sum(c["size_group"] == "中小型" for c in city_by_cbsa.values())
    if (large_count, small_count, mega_count) != (55, 140, 11):
        raise ValueError(
            "规模层人数与V2冻结时核定的全国分布不一致："
            f"大型={large_count}，中小型={small_count}，特大型={mega_count}。"
            "请先确认数据库是否为同一版本，不能静默改变约束。"
        )
    return region_data, five_rows, industry_rows, sector_order, city_by_cbsa


def zscore(values):
    n = len(values)
    mean = sum(values) / n
    sd = math.sqrt(sum((x - mean) ** 2 for x in values) / n)  # population SD, ddof=0
    if sd <= 1e-14:
        raise ValueError("某个指标在地区内标准差为0，无法进行Z-score")
    return [(x - mean) / sd for x in values], mean, sd


def calculate_region(region):
    cities = region["cities"]
    sectors = region["sectors"]
    jobs = region["jobs_by_cbsa"]
    n = len(cities)

    jobs_matrix = []
    shares = []
    job_totals = []
    for city in cities:
        row = [jobs[city["cbsa"]][sector] for sector in sectors]
        total = sum(row)
        if total <= 0:
            raise ValueError(f"{city['name']}的13类行业就业合计为0")
        jobs_matrix.append(row)
        job_totals.append(total)
        shares.append([v / total for v in row])
        if abs(sum(shares[-1]) - 1.0) > 1e-12:
            raise ValueError(f"{city['name']}产业占比合计不为100%")

    # Hellinger matrix
    H = [[0.0] * n for _ in range(n)]
    for i in range(n):
        for j in range(i + 1, n):
            value = math.sqrt(0.5 * sum(
                (math.sqrt(shares[i][k]) - math.sqrt(shares[j][k])) ** 2
                for k in range(EXPECTED_SECTORS)
            ))
            H[i][j] = H[j][i] = value

    pair_count = n * (n - 1) // 2
    h_rms = math.sqrt(
        sum(H[i][j] ** 2 for i in range(n) for j in range(i + 1, n)) / pair_count
    )
    if h_rms <= 1e-14:
        raise ValueError(f"{region['region']}产业结构H_RMS为0，无法校准")
    K = math.sqrt(2.0 * n / (n - 1.0)) / h_rms
    E = [[K * H[i][j] for j in range(n)] for i in range(n)]

    # Four scalar indicators, standardized within this region only.
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
        z, mean, sd = zscore(values)
        z_columns.append(z)
        z_stats.append({
            "指标": label,
            "变换后均值": mean,
            "总体标准差ddof0": sd,
        })
    Z = [[z_columns[k][i] for k in range(4)] for i in range(n)]

    # Five-indicator Euclidean distance, five indicators equally weighted.
    D = [[0.0] * n for _ in range(n)]
    for i in range(n):
        for j in range(i + 1, n):
            scalar_sq = sum((Z[i][k] - Z[j][k]) ** 2 for k in range(4))
            value = math.sqrt(0.2 * (scalar_sq + E[i][j] ** 2))
            D[i][j] = D[j][i] = value

    # Weights: region equal, then two size strata equal, sqrt(population) within stratum.
    group_indices = {
        "大型": [i for i, c in enumerate(cities) if c["size_group"] == "大型"],
        "中小型": [i for i, c in enumerate(cities) if c["size_group"] == "中小型"],
    }
    pi = [0.0] * n
    global_weight = [0.0] * n
    for group, indices in group_indices.items():
        if not indices:
            raise ValueError(f"{region['region']}缺少{group}城市组")
        denom = sum(cities[i]["population"] ** POP_WEIGHT_EXPONENT for i in indices)
        for i in indices:
            pi[i] = cities[i]["population"] ** POP_WEIGHT_EXPONENT / denom
            global_weight[i] = pi[i] / CITY_WEIGHT_DENOMINATOR

    # V1 single-Medoid scores are kept only as a comparison benchmark.
    medoid_scores = [sum(D[i][j] for i in range(n)) / n for j in range(n)]

    result = {
        **region,
        "n": n,
        "pair_count": pair_count,
        "jobs_matrix": jobs_matrix,
        "shares": shares,
        "job_totals": job_totals,
        "H": H,
        "h_rms": h_rms,
        "K": K,
        "E": E,
        "Z": Z,
        "z_stats": z_stats,
        "D": D,
        "group_indices": group_indices,
        "pi": pi,
        "global_weight": global_weight,
        "medoid_scores": medoid_scores,
    }
    result["configs"] = enumerate_region_configs(result)
    return result


def enumerate_region_configs(result):
    """Enumerate every feasible local choice: one or two selected cities."""
    cities = result["cities"]
    D = result["D"]
    n = result["n"]
    group_indices = result["group_indices"]
    pi = result["pi"]

    configs = []
    all_indices = range(n)
    subsets = [(i,) for i in all_indices]
    subsets.extend(itertools.combinations(all_indices, 2))

    for subset in subsets:
        nearest = [
            min(D[i][j] for j in subset)
            for i in range(n)
        ]
        group_losses = {}
        local_cost = 0.0
        for group, indices in group_indices.items():
            weighted_mean_distance = sum(pi[i] * nearest[i] for i in indices)
            group_losses[group] = weighted_mean_distance
            local_cost += (1.0 / 18.0) * weighted_mean_distance

        large_n = sum(cities[j]["size_group"] == "大型" for j in subset)
        small_n = sum(cities[j]["size_group"] == "中小型" for j in subset)
        mega_n = sum(cities[j]["is_mega"] for j in subset)
        key = tuple(sorted(cities[j]["cbsa"] for j in subset))
        configs.append({
            "region_id": result["region_id"],
            "region": result["region"],
            "selected_indices": tuple(subset),
            "selected_cbsa": key,
            "selected_names": tuple(cities[j]["name"] for j in subset),
            "seat_count": len(subset),
            "large_count": int(large_n),
            "small_count": int(small_n),
            "mega_count": int(mega_n),
            "cost": local_cost,
            "large_group_mean_distance": group_losses["大型"],
            "small_group_mean_distance": group_losses["中小型"],
        })

    return configs


def solve_global_dp(results, log_lines):
    """
    Exact dynamic programming:
    Each of 9 regions selects 1 or 2 cities. With 10 total and >=1 per region,
    exactly one region receives a second seat. State tracks total/large/small/mega.
    """
    # state -> {cost, path, tie_count}
    dp = {(0, 0, 0, 0): {"cost": 0.0, "path": tuple(), "tie_count": 1}}
    dp_log = []
    for stage, region in enumerate(results, 1):
        next_dp = {}
        transitions = 0
        for state, prior in dp.items():
            total0, large0, small0, mega0 = state
            for config in region["configs"]:
                total = total0 + config["seat_count"]
                large = large0 + config["large_count"]
                small = small0 + config["small_count"]
                mega = mega0 + config["mega_count"]
                if total > TOTAL_SELECTED or large > LARGE_QUOTA or small > SMALL_QUOTA:
                    continue
                # Remaining regions must be able to provide at least one seat each.
                remaining = len(results) - stage
                if total + remaining > TOTAL_SELECTED:
                    continue
                if total + 2 * remaining < TOTAL_SELECTED:
                    continue
                transitions += 1
                new_state = (total, large, small, mega)
                new_cost = prior["cost"] + config["cost"]
                new_path = prior["path"] + (config,)
                current = next_dp.get(new_state)
                if current is None or new_cost < current["cost"] - TIE_TOL:
                    next_dp[new_state] = {
                        "cost": new_cost,
                        "path": new_path,
                        "tie_count": prior["tie_count"],
                    }
                elif abs(new_cost - current["cost"]) <= TIE_TOL:
                    current["tie_count"] += prior["tie_count"]
                    current_key = tuple(x["selected_cbsa"] for x in current["path"])
                    new_key = tuple(x["selected_cbsa"] for x in new_path)
                    if new_key < current_key:
                        current["path"] = new_path

        dp = next_dp
        if not dp:
            raise ValueError(f"动态规划第{stage}个地区后无可行状态，请检查规模约束")
        dp_log.append({
            "处理地区序号": stage,
            "处理地区": region["region"],
            "保留状态数": len(dp),
            "有效转移次数": transitions,
            "当前最小部分目标值": min(v["cost"] for v in dp.values()),
        })

    feasible = []
    for state, value in dp.items():
        total, large, small, mega = state
        if (
            total == TOTAL_SELECTED
            and large == LARGE_QUOTA
            and small == SMALL_QUOTA
            and mega >= MEGA_MIN
        ):
            feasible.append((value["cost"], state, value))

    if not feasible:
        raise ValueError(
            "不存在满足V2全部约束的可行解。请检查各地区候选城市及人口规模分布。"
        )

    feasible.sort(key=lambda x: (x[0], x[1]))
    best_cost = feasible[0][0]
    tied = [x for x in feasible if abs(x[0] - best_cost) <= TIE_TOL]
    # Canonical deterministic choice among tied states: lexicographically smallest CBSA path.
    chosen = min(
        tied,
        key=lambda x: tuple(config["selected_cbsa"] for config in x[2]["path"])
    )
    total_ties = sum(x[2]["tie_count"] for x in tied)
    best_path = chosen[2]["path"]

    # Global objective check by recomputing all assignment contributions.
    recomputed = sum(config["cost"] for config in best_path)
    if abs(recomputed - best_cost) > 1e-9:
        raise AssertionError("动态规划目标值与所选地区组合成本不一致")

    log_line(log_lines, f"动态规划完成：最终可行规模状态数={len(feasible)}")
    log_line(log_lines, f"最优目标值 L*={best_cost:.12f}")
    log_line(log_lines, f"最优解并列路径数（按DP状态累计）={total_ties}")
    log_line(log_lines, f"采用确定性CBSA字典序规则输出其中一组最优解")

    return {
        "best_cost": best_cost,
        "best_state": chosen[1],
        "path": best_path,
        "tie_count": total_ties,
        "dp_log": dp_log,
        "feasible_state_count": len(feasible),
    }


def build_selected_and_assignments(results, solution):
    selected_rows = []
    assignment_rows = []
    region_summary = []
    selected_by_region = {}
    objective_contributions = []

    for region, config in zip(results, solution["path"]):
        cities = region["cities"]
        D = region["D"]
        chosen_indices = list(config["selected_indices"])
        selected_by_region[region["region_id"]] = chosen_indices

        for seat, idx in enumerate(chosen_indices, 1):
            city = cities[idx]
            selected_rows.append({
                "地区编号": region["region_id"],
                "地区": region["region"],
                "地区席位": seat,
                "MSA": city["name"],
                "CBSA": city["cbsa"],
                "2024人口": city["population"],
                "规模组": city["size_group"],
                "特大型_500万以上": "是" if city["is_mega"] else "否",
                "V2入选": "是",
            })

        large_loss = 0.0
        small_loss = 0.0
        region_loss = 0.0
        for i, city in enumerate(cities):
            nearest_distance = min(D[i][j] for j in chosen_indices)
            nearest_candidates = [
                cities[j]["name"] for j in chosen_indices
                if abs(D[i][j] - nearest_distance) <= TIE_TOL
            ]
            # Deterministic assignment in a distance tie: alphabetically first CBSA;
            # all tied candidates are retained in the output field.
            nearest_idx = min(
                [j for j in chosen_indices if abs(D[i][j] - nearest_distance) <= TIE_TOL],
                key=lambda j: cities[j]["cbsa"]
            )
            contribution = region["global_weight"][i] * nearest_distance
            objective_contributions.append(contribution)
            region_loss += contribution
            if city["size_group"] == "大型":
                large_loss += contribution
            else:
                small_loss += contribution

            assignment_rows.append({
                "地区编号": region["region_id"],
                "地区": region["region"],
                "MSA": city["name"],
                "CBSA": city["cbsa"],
                "2024人口": city["population"],
                "规模组": city["size_group"],
                "sqrt人口组内权重pi": region["pi"][i],
                "全国目标函数权重wi": region["global_weight"][i],
                "最近代表城市": cities[nearest_idx]["name"],
                "代表城市CBSA": cities[nearest_idx]["cbsa"],
                "并列最近代表城市": "；".join(nearest_candidates),
                "最近距离d": nearest_distance,
                "目标函数贡献wi乘距离": contribution,
                "该MSA本身入选": "是" if i in chosen_indices else "否",
            })

        region_summary.append({
            "地区编号": region["region_id"],
            "地区": region["region"],
            "MSA数量": region["n"],
            "入选城市数": len(chosen_indices),
            "大型组入选数": config["large_count"],
            "中小型组入选数": config["small_count"],
            "特大型入选数": config["mega_count"],
            "大型组加权平均覆盖距离": config["large_group_mean_distance"],
            "中小型组加权平均覆盖距离": config["small_group_mean_distance"],
            "地区目标函数贡献": region_loss,
            "入选城市": "；".join(cities[j]["name"] for j in chosen_indices),
            "候选组合成本": config["cost"],
        })

    total_check = sum(objective_contributions)
    if abs(total_check - solution["best_cost"]) > 1e-9:
        raise AssertionError(
            f"195城分配贡献之和({total_check})与DP目标值({solution['best_cost']})不一致"
        )
    return selected_rows, assignment_rows, region_summary, selected_by_region, total_check


def make_v1_benchmark(results, selected_by_region):
    rows = []
    for region in results:
        order = sorted(
            range(region["n"]),
            key=lambda i: (region["medoid_scores"][i], region["cities"][i]["cbsa"])
        )
        for rank, i in enumerate(order, 1):
            c = region["cities"][i]
            rows.append({
                "地区编号": region["region_id"],
                "地区": region["region"],
                "V1单城Medoid排名": rank,
                "MSA": c["name"],
                "CBSA": c["cbsa"],
                "单城平均距离D1": region["medoid_scores"][i],
                "是否V2入选": "是" if i in selected_by_region[region["region_id"]] else "否",
            })
    return rows


def all_region_intermediate_rows(results):
    z_rows, weight_rows, parameter_rows = [], [], []
    distance_rows, share_rows, config_rows = [], [], []
    for region in results:
        for stat in region["z_stats"]:
            parameter_rows.append({
                "地区编号": region["region_id"],
                "地区": region["region"],
                "MSA数量": region["n"],
                "配对数": region["pair_count"],
                "H_RMS": region["h_rms"],
                "K": region["K"],
                "指标": stat["指标"],
                "变换后均值": stat["变换后均值"],
                "总体标准差_ddof0": stat["总体标准差ddof0"],
            })
        for i, city in enumerate(region["cities"]):
            z_rows.append({
                "地区编号": region["region_id"],
                "地区": region["region"],
                "MSA": city["name"],
                "CBSA": city["cbsa"],
                "ln人口": math.log(city["population"]),
                "ln人口密度": math.log(city["density"]),
                "人均GDP原值": city["gdp_pc"],
                "CAGR原值": city["cagr"],
                "Z_ln人口": region["Z"][i][0],
                "Z_ln人口密度": region["Z"][i][1],
                "Z_人均GDP": region["Z"][i][2],
                "Z_CAGR": region["Z"][i][3],
                "规模组": city["size_group"],
                "特大型": "是" if city["is_mega"] else "否",
            })
            weight_rows.append({
                "地区编号": region["region_id"],
                "地区": region["region"],
                "MSA": city["name"],
                "CBSA": city["cbsa"],
                "人口": city["population"],
                "规模组": city["size_group"],
                "sqrt人口": math.sqrt(city["population"]),
                "组内归一化权重pi": region["pi"][i],
                "地区规模组权重": GROUP_WEIGHT,
                "地区权重": REGION_WEIGHT,
                "全国目标函数权重wi": region["global_weight"][i],
            })
            for k, sector in enumerate(region["sectors"]):
                share_rows.append({
                    "地区编号": region["region_id"],
                    "地区": region["region"],
                    "MSA": city["name"],
                    "CBSA": city["cbsa"],
                    "行业": sector,
                    "就业人数": region["jobs_matrix"][i][k],
                    "产业占比": region["shares"][i][k],
                })
        for i in range(region["n"]):
            for j in range(region["n"]):
                distance_rows.append({
                    "地区编号": region["region_id"],
                    "地区": region["region"],
                    "MSA_i": region["cities"][i]["name"],
                    "CBSA_i": region["cities"][i]["cbsa"],
                    "MSA_j": region["cities"][j]["name"],
                    "CBSA_j": region["cities"][j]["cbsa"],
                    "H": region["H"][i][j],
                    "E": region["E"][i][j],
                    "d": region["D"][i][j],
                })
        for cfg_no, config in enumerate(region["configs"], 1):
            config_rows.append({
                "地区编号": region["region_id"],
                "地区": region["region"],
                "候选组合编号": cfg_no,
                "组合类型": "单城" if config["seat_count"] == 1 else "双城",
                "入选城市数": config["seat_count"],
                "候选城市": "；".join(config["selected_names"]),
                "候选CBSA": "；".join(config["selected_cbsa"]),
                "大型入选数": config["large_count"],
                "中小型入选数": config["small_count"],
                "特大型入选数": config["mega_count"],
                "大型组加权平均距离": config["large_group_mean_distance"],
                "中小型组加权平均距离": config["small_group_mean_distance"],
                "区域目标函数成本": config["cost"],
            })
    return z_rows, weight_rows, parameter_rows, distance_rows, share_rows, config_rows


def csv_write(path, rows):
    if not rows:
        return
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def style_sheet(ws, widths=None, freeze="A2", filter_on=True):
    ws.freeze_panes = freeze
    ws.sheet_view.showGridLines = False
    if filter_on and ws.max_row >= 1 and ws.max_column >= 1:
        ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 32
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="164E63")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col in range(1, ws.max_column + 1):
        letter = get_column_letter(col)
        if widths and col <= len(widths):
            width = widths[col - 1]
        else:
            width = 16
        ws.column_dimensions[letter].width = width
    thin = Side(style="hair", color="D9E2F3")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="center")
            cell.border = Border(bottom=thin)


def add_table_sheet(wb, title, headers, rows, widths=None, number_formats=None):
    ws = wb.create_sheet(title)
    ws.append(headers)
    for row in rows:
        if isinstance(row, dict):
            ws.append([row.get(h) for h in headers])
        else:
            ws.append(list(row))
    style_sheet(ws, widths)
    if number_formats:
        for col_idx, fmt in number_formats.items():
            for r in range(2, ws.max_row + 1):
                ws.cell(r, col_idx).number_format = fmt
    return ws


def add_matrix_sheet(wb, title, cities, matrix, fmt="0.000000"):
    headers = ["CBSA", "MSA"] + [f"{c['cbsa']} {c['name']}" for c in cities]
    rows = []
    for i, city in enumerate(cities):
        rows.append([city["cbsa"], city["name"]] + matrix[i])
    ws = add_table_sheet(
        wb, title, headers, rows,
        widths=[12, 42] + [23] * len(cities)
    )
    for row in ws.iter_rows(min_row=2, min_col=3):
        for cell in row:
            cell.number_format = fmt
    if ws.max_row > 1 and ws.max_column > 2:
        ws.conditional_formatting.add(
            f"C2:{get_column_letter(ws.max_column)}{ws.max_row}",
            ColorScaleRule(
                start_type="min", start_color="FEE2E2",
                mid_type="percentile", mid_value=50, mid_color="FEF3C7",
                end_type="max", end_color="DCFCE7"
            )
        )
    return ws


def build_workbook(results, five_rows, industry_rows, sector_order,
                   selected_rows, assignment_rows, region_summary,
                   solution, dp_log, intermediate, v1_rows, qa_rows,
                   total_check):
    wb = Workbook()
    ws = wb.active
    ws.title = "V2最终汇总"

    summary_headers = [
        "地区编号", "地区", "MSA数量", "入选城市数", "大型组入选数", "中小型组入选数",
        "特大型入选数", "大型组加权平均覆盖距离", "中小型组加权平均覆盖距离",
        "地区目标函数贡献", "入选城市"
    ]
    for h in summary_headers:
        ws.cell(1, summary_headers.index(h) + 1, h)
    for r_idx, row in enumerate(region_summary, 2):
        for c_idx, h in enumerate(summary_headers, 1):
            ws.cell(r_idx, c_idx, row[h])
    style_sheet(ws, [10, 30, 12, 12, 16, 17, 16, 24, 24, 22, 90])
    for row in range(2, ws.max_row + 1):
        for col in (8, 9, 10):
            ws.cell(row, col).number_format = "0.0000000000"

    add_table_sheet(
        wb, "全国最终10城",
        ["地区编号", "地区", "地区席位", "MSA", "CBSA", "2024人口",
         "规模组", "特大型_500万以上", "V2入选"],
        selected_rows,
        widths=[10, 30, 12, 48, 12, 18, 14, 22, 12],
        number_formats={6: "#,##0"}
    )
    add_table_sheet(
        wb, "195城覆盖分配",
        ["地区编号", "地区", "MSA", "CBSA", "2024人口", "规模组",
         "sqrt人口组内权重pi", "全国目标函数权重wi",
         "最近代表城市", "代表城市CBSA", "并列最近代表城市",
         "最近距离d", "目标函数贡献wi乘距离", "该MSA本身入选"],
        assignment_rows,
        widths=[10, 30, 48, 12, 18, 14, 22, 22, 48, 14, 70, 18, 24, 16],
        number_formats={5: "#,##0", 7: "0.0000000000", 8: "0.0000000000",
                        12: "0.0000000000", 13: "0.0000000000"}
    )
    add_table_sheet(
        wb, "全国权重",
        ["地区编号", "地区", "MSA", "CBSA", "人口", "规模组",
         "sqrt人口", "组内归一化权重pi", "地区规模组权重",
         "地区权重", "全国目标函数权重wi"],
        intermediate["weight_rows"],
        widths=[10, 30, 48, 12, 18, 14, 18, 22, 20, 16, 22],
        number_formats={5: "#,##0", 7: "0.0000", 8: "0.0000000000",
                        9: "0.0000", 10: "0.0000", 11: "0.0000000000"}
    )
    add_table_sheet(
        wb, "九区模型参数",
        ["地区编号", "地区", "MSA数量", "配对数", "H_RMS", "K",
         "指标", "变换后均值", "总体标准差_ddof0"],
        intermediate["parameter_rows"],
        widths=[10, 30, 12, 12, 18, 18, 22, 22, 24],
        number_formats={5: "0.0000000000", 6: "0.0000000000",
                        8: "0.000000", 9: "0.000000"}
    )
    add_table_sheet(
        wb, "全局动态规划记录",
        ["处理地区序号", "处理地区", "保留状态数", "有效转移次数", "当前最小部分目标值"],
        dp_log,
        widths=[16, 32, 16, 18, 24],
        number_formats={5: "0.0000000000"}
    )

    opt_rows = [
        ["最优目标函数值L*", solution["best_cost"]],
        ["目标值复算（195城权重贡献合计）", total_check],
        ["两者差值", total_check - solution["best_cost"]],
        ["最终状态（总数、大型、中小型、特大型）", str(solution["best_state"])],
        ["并列最优路径数（按DP状态累计）", solution["tie_count"]],
        ["最终可行规模状态数", solution["feasible_state_count"]],
        ["总入选数", TOTAL_SELECTED],
        ["大型城市配额", LARGE_QUOTA],
        ["中小型城市配额", SMALL_QUOTA],
        ["特大型最低数量", MEGA_MIN],
        ["人口权重指数", POP_WEIGHT_EXPONENT],
        ["区域权重", REGION_WEIGHT],
        ["组内权重", GROUP_WEIGHT],
        ["距离并列容差", TIE_TOL],
    ]
    add_table_sheet(wb, "优化目标与约束核验", ["项目", "数值或定义"], opt_rows,
                    widths=[48, 80])

    add_table_sheet(
        wb, "规模分层统计",
        ["地区编号", "地区", "MSA总数", "大型>=100万", "中小型25万至100万",
         "特大型>=500万", "V2大型入选", "V2中小型入选", "V2特大型入选"],
        [
            [
                r["region_id"], r["region"], r["n"],
                sum(c["size_group"] == "大型" for c in r["cities"]),
                sum(c["size_group"] == "中小型" for c in r["cities"]),
                sum(c["is_mega"] for c in r["cities"]),
                next(x["大型组入选数"] for x in region_summary if x["地区编号"] == r["region_id"]),
                next(x["中小型组入选数"] for x in region_summary if x["地区编号"] == r["region_id"]),
                next(x["特大型入选数"] for x in region_summary if x["地区编号"] == r["region_id"]),
            ]
            for r in results
        ],
        widths=[10, 30, 12, 18, 22, 18, 18, 20, 18]
    )

    add_table_sheet(
        wb, "全部区域候选组合",
        ["地区编号", "地区", "候选组合编号", "组合类型", "入选城市数",
         "候选城市", "候选CBSA", "大型入选数", "中小型入选数", "特大型入选数",
         "大型组加权平均距离", "中小型组加权平均距离", "区域目标函数成本"],
        intermediate["config_rows"],
        widths=[10, 30, 16, 14, 14, 85, 24, 16, 18, 16, 24, 24, 24],
        number_formats={11: "0.0000000000", 12: "0.0000000000", 13: "0.0000000000"}
    )

    add_table_sheet(
        wb, "V1单城Medoid对照",
        ["地区编号", "地区", "V1单城排名", "MSA", "CBSA", "单城平均距离D1", "是否V2入选"],
        v1_rows,
        widths=[10, 30, 16, 48, 12, 22, 16],
        number_formats={6: "0.0000000000"}
    )

    add_table_sheet(
        wb, "数据质量检查",
        ["检查项目", "预期值", "实际值", "结果"],
        qa_rows,
        widths=[42, 30, 42, 16]
    )

    method_rows = [
        ["模型名称", "V2：区域分层约束加权P-Median联合选城模型"],
        ["统计单位", "195个MSA；九大Census Divisions；Pacific完整31城，不拆分南北。"],
        ["五指标", "人口规模、人口密度、人均GDP、13类产业结构、人口增长率；每项20%。"],
        ["区域独立性", "H、H_RMS、K、E、四项Z-score和d均在各自地区内部计算，不使用全国195城均值或标准差。"],
        ["产业占比", "p_ik = MSA i 第k类行业就业人数 / 该MSA 13类行业就业人数合计。"],
        ["Hellinger距离", "H_ij = sqrt(0.5 × Σ_k (sqrt(p_ik)-sqrt(p_jk))²)。"],
        ["配对数", "n(n-1)/2；RMS仅使用上三角非对角城市配对。"],
        ["H_RMS", "sqrt(Σ_{i<j} H_ij² / [n(n-1)/2])。"],
        ["产业校准系数", "K = sqrt(2n/(n-1)) / H_RMS。"],
        ["产业结构距离", "E_ij = K × H_ij。"],
        ["标量指标", "ln(2024人口)、ln(2024人口密度)、2024人均GDP、2020-2024人口CAGR。"],
        ["Z-score", "Z=(x-地区均值)/地区总体标准差；ddof=0。"],
        ["五指标距离", "d_ij = sqrt(0.2 × [(ΔZ人口)²+(ΔZ密度)²+(ΔZGDP)²+E_ij²+(ΔZCAGR)²])。"],
        ["人口规模分层", "大型：人口≥1,000,000；中小型：250,000≤人口<1,000,000；特大型：人口≥5,000,000（特大型属于大型子集）。"],
        ["最终名额", "总计10座；大型恰好5座；中小型恰好5座；特大型至少2座。"],
        ["区域约束", "九大地区每区至少1座。因总席位10、地区9且每区至少1座，故每区只能选1或2座，且恰有一个地区选2座。"],
        ["区域与规模组权重", "每个地区总权重1/9；地区内大型组与中小型组各占1/2。"],
        ["组内人口权重", "pi_i|rg = sqrt(P_i) / Σ_{h∈M_rg} sqrt(P_h)。"],
        ["全国目标函数权重", "w_i = pi_i|rg / 18；全部195个MSA权重之和为1。"],
        ["覆盖距离", "d(i,S_r)=min_{j∈S_r} d_ij；只允许同一地区内互相代表。"],
        ["目标函数", "min L(S)=Σ_r Σ_g Σ_{i∈M_rg} [1/18 × pi_i|rg × min_{j∈S_r} d_ij]。"],
        ["求解方法", "对每区枚举所有单城和双城组合；动态规划在九区间联合分配名额，精确求全局最优解。"],
        ["精确性说明", "在固定模型约束下，每个可行解必对应九区各选1或2座的组合路径；DP穷尽所有区域组合状态，不使用随机启发式。"],
        ["并列处理", "目标值在1e-10容差内视为并列；输出CBSA字典序最小的可复现代表解，并记录并列最优路径数。"],
        ["V1对照", "另列各地区单城Medoid排名，仅作诊断对照，不参与V2目标函数。"],
        ["数据源限制", "程序只使用全国SQLite数据库内已有数据；不自动补值、不模拟、不联网。"],
    ]
    add_table_sheet(wb, "V2模型定义与公式", ["项目", "定义、公式或说明"], method_rows,
                    widths=[28, 125])
    for row in wb["V2模型定义与公式"].iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    # National raw five-indicator and industry tables.
    five_headers = [
        "地区编号", "地区", "原始清单序号", "MSA", "CBSA", "2024人口",
        "2024人口密度_人每平方英里", "2024人均GDP_美元", "2020至2024_CAGR_百分比"
    ]
    five_export = []
    for row in five_rows:
        five_export.append([row[h] for h in five_headers])
    add_table_sheet(
        wb, "全国195城原始五指标", five_headers, five_export,
        widths=[10, 30, 14, 48, 12, 18, 26, 22, 24],
        number_formats={6: "#,##0", 7: "0.0000", 8: "$#,##0.00", 9: "0.000000"}
    )

    industry_headers = [
        "地区编号", "地区", "原始清单序号", "MSA", "CBSA", "行业",
        "男性就业", "女性就业", "就业合计"
    ]
    industry_export = [[row[h] for h in industry_headers] for row in industry_rows]
    add_table_sheet(
        wb, "全国2535条产业明细", industry_headers, industry_export,
        widths=[10, 30, 14, 48, 12, 38, 16, 16, 16],
        number_formats={7: "#,##0", 8: "#,##0", 9: "#,##0"}
    )

    # Region-level detail sheets.
    for region in results:
        prefix = f"{region['region_id']}_{region['short']}"
        cities = region["cities"]
        raw_rows = []
        for i, c in enumerate(cities):
            raw_rows.append([
                c["no"], c["name"], c["cbsa"], c["population"], c["density"],
                c["gdp_pc"], c["cagr"], region["job_totals"][i],
                c["size_group"], "是" if c["is_mega"] else "否"
            ])
        add_table_sheet(
            wb, f"{prefix}_五指标",
            ["原始序号", "MSA", "CBSA", "2024人口", "人口密度",
             "2024人均GDP", "2020-2024_CAGR", "13类行业就业合计", "规模组", "特大型"],
            raw_rows,
            widths=[12, 44, 12, 18, 18, 20, 20, 22, 14, 12],
            number_formats={4: "#,##0", 5: "0.0000", 6: "$#,##0.00",
                            7: "0.000000", 8: "#,##0"}
        )

        emp_rows = [
            [c["no"], c["name"], c["cbsa"]] + region["jobs_matrix"][i] + [region["job_totals"][i]]
            for i, c in enumerate(cities)
        ]
        add_table_sheet(
            wb, f"{prefix}_产业就业",
            ["原始序号", "MSA", "CBSA"] + region["sectors"] + ["13类就业合计"],
            emp_rows,
            widths=[12, 44, 12] + [20] * 13 + [20],
            number_formats={j + 4: "#,##0" for j in range(14)}
        )

        share_rows = [
            [c["no"], c["name"], c["cbsa"]] + region["shares"][i] + [sum(region["shares"][i])]
            for i, c in enumerate(cities)
        ]
        add_table_sheet(
            wb, f"{prefix}_产业占比",
            ["原始序号", "MSA", "CBSA"] + region["sectors"] + ["占比合计"],
            share_rows,
            widths=[12, 44, 12] + [20] * 13 + [16],
            number_formats={j + 4: "0.0000%" for j in range(14)}
        )

        add_matrix_sheet(wb, f"{prefix}_H矩阵", cities, region["H"])
        add_matrix_sheet(wb, f"{prefix}_E矩阵", cities, region["E"])

        z_rows = []
        for i, c in enumerate(cities):
            z_rows.append([
                c["no"], c["name"], c["cbsa"],
                math.log(c["population"]), math.log(c["density"]),
                c["gdp_pc"], c["cagr"],
                region["Z"][i][0], region["Z"][i][1],
                region["Z"][i][2], region["Z"][i][3],
                region["pi"][i], region["global_weight"][i]
            ])
        add_table_sheet(
            wb, f"{prefix}_Z标准化",
            ["原始序号", "MSA", "CBSA", "ln人口", "ln人口密度",
             "人均GDP原值", "CAGR原值", "Z_ln人口", "Z_ln密度",
             "Z_GDP", "Z_CAGR", "组内权重pi", "全国权重wi"],
            z_rows,
            widths=[12, 44, 12, 16, 18, 20, 18, 16, 16, 16, 16, 18, 18],
            number_formats={j: "0.0000000000" for j in range(4, 14)}
        )

        add_matrix_sheet(wb, f"{prefix}_d矩阵", cities, region["D"])

        cfg_rows = []
        for k, cfg in enumerate(region["configs"], 1):
            cfg_rows.append([
                k, "单城" if cfg["seat_count"] == 1 else "双城",
                cfg["seat_count"], "；".join(cfg["selected_names"]),
                "；".join(cfg["selected_cbsa"]), cfg["large_count"],
                cfg["small_count"], cfg["mega_count"],
                cfg["large_group_mean_distance"], cfg["small_group_mean_distance"],
                cfg["cost"]
            ])
        add_table_sheet(
            wb, f"{prefix}_候选组合",
            ["组合编号", "类型", "城市数", "候选城市", "候选CBSA",
             "大型数", "中小型数", "特大型数",
             "大型组加权距离", "中小型组加权距离", "区域目标成本"],
            cfg_rows,
            widths=[12, 12, 12, 90, 26, 12, 14, 12, 22, 22, 22],
            number_formats={9: "0.0000000000", 10: "0.0000000000", 11: "0.0000000000"}
        )

    # Add formulas as comments to key column headers.
    wb["全国权重"]["K1"].comment = Comment(
        "wi = pi / 18；每个地区权重1/9，地区内大型组与中小型组各1/2。"
        "组内pi按sqrt(2024人口)归一化。",
        "OpenAI"
    )
    wb["195城覆盖分配"]["L1"].comment = Comment(
        "每个MSA分配给同一地区内与其五指标距离d最小的入选城市。"
        "并列时保留全部并列候选，最近代表城市列采用CBSA字典序作可复现展示。",
        "OpenAI"
    )

    wb.properties.title = "美国九大地区 V2 分层约束加权P-Median最终结果"
    wb.properties.subject = "195个MSA、九区独立五指标距离、10城精确联合优化"
    wb.properties.creator = "OpenAI"
    wb.save(OUT_XLSX)


def write_sqlite_output(db_path, results, five_rows, industry_rows,
                        selected_rows, assignment_rows, region_summary,
                        solution, dp_log, intermediate):
    if OUT_SQLITE.exists():
        OUT_SQLITE.unlink()
    conn = sqlite3.connect(str(OUT_SQLITE))
    try:
        # Preserve source rows in a separate result database for complete provenance.
        src_conn = sqlite3.connect(str(db_path))
        try:
            src_conn.backup(conn)
        finally:
            src_conn.close()

        def replace_table(name, headers, rows):
            conn.execute(f'DROP TABLE IF EXISTS "{name}"')
            if not headers:
                return
            column_types = []
            for col_idx, header in enumerate(headers):
                values = [row[col_idx] for row in rows if row[col_idx] is not None]
                if values and all(isinstance(v, bool) or isinstance(v, int) for v in values):
                    sql_type = "INTEGER"
                elif values and all(isinstance(v, (bool, int, float)) for v in values):
                    sql_type = "REAL"
                else:
                    sql_type = "TEXT"
                column_types.append(f'"{header}" {sql_type}')
            sql_cols = ", ".join(column_types)
            conn.execute(f'CREATE TABLE "{name}" ({sql_cols})')
            placeholders = ",".join(["?"] * len(headers))
            conn.executemany(
                f'INSERT INTO "{name}" VALUES ({placeholders})',
                [tuple(row) for row in rows]
            )

        def rows_from_dicts(headers, dict_rows):
            return [[row.get(h) for h in headers] for row in dict_rows]

        replace_table(
            "v2_selected_cities",
            list(selected_rows[0].keys()),
            rows_from_dicts(list(selected_rows[0].keys()), selected_rows)
        )
        replace_table(
            "v2_city_assignments",
            list(assignment_rows[0].keys()),
            rows_from_dicts(list(assignment_rows[0].keys()), assignment_rows)
        )
        replace_table(
            "v2_region_summary",
            list(region_summary[0].keys()),
            rows_from_dicts(list(region_summary[0].keys()), region_summary)
        )
        for table_name, key in [
            ("v2_city_standardized", "z_rows"),
            ("v2_city_weights", "weight_rows"),
            ("v2_region_parameters", "parameter_rows"),
            ("v2_distance_matrices_long", "distance_rows"),
            ("v2_industry_shares_long", "share_rows"),
            ("v2_local_configurations", "config_rows"),
        ]:
            rows = intermediate[key]
            headers = list(rows[0].keys()) if rows else []
            replace_table(table_name, headers, rows_from_dicts(headers, rows))

        dp_headers = list(dp_log[0].keys()) if dp_log else []
        replace_table("v2_dynamic_programming_log", dp_headers,
                      rows_from_dicts(dp_headers, dp_log))

        conn.execute('CREATE INDEX IF NOT EXISTS idx_v2_assign_cbsa ON v2_city_assignments("CBSA")')
        conn.execute('CREATE INDEX IF NOT EXISTS idx_v2_distance_region ON v2_distance_matrices_long("地区编号")')
        conn.execute('CREATE INDEX IF NOT EXISTS idx_v2_config_region ON v2_local_configurations("地区编号")')
        conn.commit()
    finally:
        conn.close()


def make_qa(results, five_rows, industry_rows, selected_rows,
            assignment_rows, region_summary, solution, total_check):
    qa = []
    def add(name, expected, actual, passed, note=""):
        qa.append({
            "检查项目": name,
            "预期值": expected,
            "实际值": actual,
            "结果": "通过" if passed else "失败",
            "说明": note,
        })

    add("地区数量", 9, len(results), len(results) == 9)
    add("全国MSA数量", 195, len(five_rows), len(five_rows) == 195)
    add("全国产业记录数量", 2535, len(industry_rows), len(industry_rows) == 2535)
    add("最终入选城市数", 10, len(selected_rows), len(selected_rows) == 10)
    add("覆盖分配城市数", 195, len(assignment_rows), len(assignment_rows) == 195)
    add("大型入选数", 5, sum(x["规模组"] == "大型" for x in selected_rows),
        sum(x["规模组"] == "大型" for x in selected_rows) == 5)
    add("中小型入选数", 5, sum(x["规模组"] == "中小型" for x in selected_rows),
        sum(x["规模组"] == "中小型" for x in selected_rows) == 5)
    add("特大型入选数下限", ">=2",
        sum(x["特大型_500万以上"] == "是" for x in selected_rows),
        sum(x["特大型_500万以上"] == "是" for x in selected_rows) >= 2)
    add("九大地区覆盖", "每区至少1座",
        min(x["入选城市数"] for x in region_summary),
        all(x["入选城市数"] >= 1 for x in region_summary))
    add("每区最多2座", "<=2", max(x["入选城市数"] for x in region_summary),
        max(x["入选城市数"] for x in region_summary) <= 2)
    add("目标函数复算", "DP目标值与195城贡献和一致",
        abs(solution["best_cost"] - total_check),
        abs(solution["best_cost"] - total_check) <= 1e-9)
    add("全部城市权重和", 1.0, sum(x["全国目标函数权重wi"] for x in assignment_rows),
        abs(sum(x["全国目标函数权重wi"] for x in assignment_rows) - 1.0) <= 1e-10)
    for r in results:
        shares_ok = all(abs(sum(p) - 1.0) <= 1e-12 for p in r["shares"])
        matrix_ok = True
        for matrix in (r["H"], r["E"], r["D"]):
            for i in range(r["n"]):
                if abs(matrix[i][i]) > 1e-12:
                    matrix_ok = False
                for j in range(r["n"]):
                    if abs(matrix[i][j] - matrix[j][i]) > 1e-12:
                        matrix_ok = False
        add(f'{r["region"]} 13类占比合计', "每城=100%", "通过" if shares_ok else "失败", shares_ok)
        add(f'{r["region"]} H/E/d对称及对角线', "对称且对角线为0",
            "通过" if matrix_ok else "失败", matrix_ok)
    if any(x["结果"] != "通过" for x in qa):
        raise AssertionError("至少一项质量检查失败，拒绝生成正式输出")
    return qa


def main():
    parser = argparse.ArgumentParser(
        description="按冻结V2模型从全国SQLite数据库精确求解九大地区10座代表城市。"
    )
    parser.add_argument("--db", type=Path, default=DEFAULT_DB,
                        help="全国SQLite数据库路径")
    args = parser.parse_args()

    log_lines = []
    log_line(log_lines, "美国九大地区 V2 分层约束加权P-Median")
    log_line(log_lines, f"运行时间：{datetime.now().isoformat(timespec='seconds')}")
    log_line(log_lines, f"输入数据库：{args.db.resolve()}")

    try:
        region_data, five_rows, industry_rows, sector_order, city_by_cbsa = load_database(args.db)
        log_line(log_lines, "数据库完整性检查通过：195个MSA、2535条产业记录、13类行业。")
        log_line(log_lines, "开始分别计算九个地区的H、RMS、K、E、Z和d……")

        results = []
        for region_id, region_name, short, expected_n in REGIONS:
            result = calculate_region(region_data[region_id])
            results.append(result)
            large_n = sum(c["size_group"] == "大型" for c in result["cities"])
            small_n = sum(c["size_group"] == "中小型" for c in result["cities"])
            mega_n = sum(c["is_mega"] for c in result["cities"])
            log_line(
                log_lines,
                f"OK {region_name}: {result['n']}城；大型={large_n}；中小型={small_n}；"
                f"特大型={mega_n}；配对={result['pair_count']}；"
                f"H_RMS={result['h_rms']:.10f}；K={result['K']:.10f}；"
                f"本区候选组合={len(result['configs'])}"
            )

        log_line(log_lines, "开始动态规划联合优化：每区选1或2座，全国合计10座……")
        solution = solve_global_dp(results, log_lines)

        selected_rows, assignment_rows, region_summary, selected_by_region, total_check = (
            build_selected_and_assignments(results, solution)
        )
        v1_rows = make_v1_benchmark(results, selected_by_region)
        z_rows, weight_rows, parameter_rows, distance_rows, share_rows, config_rows = (
            all_region_intermediate_rows(results)
        )
        intermediate = {
            "z_rows": z_rows,
            "weight_rows": weight_rows,
            "parameter_rows": parameter_rows,
            "distance_rows": distance_rows,
            "share_rows": share_rows,
            "config_rows": config_rows,
        }
        dp_log = solution["dp_log"]
        qa_rows = make_qa(
            results, five_rows, industry_rows, selected_rows,
            assignment_rows, region_summary, solution, total_check
        )

        # Save all outputs only after all calculations and checks pass.
        build_workbook(
            results, five_rows, industry_rows, sector_order,
            selected_rows, assignment_rows, region_summary,
            solution, dp_log, intermediate, v1_rows, qa_rows, total_check
        )
        csv_write(OUT_SELECTED_CSV, selected_rows)
        csv_write(OUT_ASSIGN_CSV, assignment_rows)
        csv_write(OUT_CONFIG_CSV, config_rows)
        write_sqlite_output(
            args.db, results, five_rows, industry_rows,
            selected_rows, assignment_rows, region_summary,
            solution, dp_log, intermediate
        )

        log_line(log_lines, "")
        log_line(log_lines, "最终10座城市：")
        for row in selected_rows:
            log_line(
                log_lines,
                f'地区{row["地区编号"]} {row["地区"]} | {row["MSA"]} | '
                f'人口={row["2024人口"]:,.0f} | {row["规模组"]} | '
                f'特大型={row["特大型_500万以上"]}'
            )
        log_line(log_lines, "")
        log_line(log_lines, f"最优目标函数 L* = {solution['best_cost']:.12f}")
        log_line(log_lines, f"195城贡献复算 = {total_check:.12f}")
        log_line(log_lines, f"并列最优路径数 = {solution['tie_count']}")
        log_line(log_lines, "")
        log_line(log_lines, "输出文件：")
        log_line(log_lines, str(OUT_XLSX.resolve()))
        log_line(log_lines, str(OUT_SELECTED_CSV.resolve()))
        log_line(log_lines, str(OUT_ASSIGN_CSV.resolve()))
        log_line(log_lines, str(OUT_CONFIG_CSV.resolve()))
        log_line(log_lines, str(OUT_SQLITE.resolve()))
        log_line(log_lines, str(OUT_LOG.resolve()))
        log_line(log_lines, "质量检查全部通过。原始SQLite数据库未被修改。")

        OUT_LOG.write_text("\n".join(log_lines) + "\n", encoding="utf-8-sig")
        print("\nV2计算全部完成。")

    except Exception as exc:
        log_line(log_lines, "")
        log_line(log_lines, "计算失败，未完成正式输出。")
        log_line(log_lines, f"错误：{exc}")
        OUT_LOG.write_text("\n".join(log_lines) + "\n", encoding="utf-8-sig")
        print("\n计算未完成。请先检查错误信息和日志。")
        sys.exit(2)


if __name__ == "__main__":
    main()
