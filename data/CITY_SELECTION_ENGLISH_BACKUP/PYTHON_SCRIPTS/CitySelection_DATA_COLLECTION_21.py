# -*- coding: utf-8 -*-
"""
从全国 SQLite 数据库中提取指定 Census Division，按太平洋地区 V1.0 算法计算：
产业结构 Hellinger 距离 -> RMS 校准 E -> 五指标综合距离 d -> Medoid / 双城覆盖筛选。

默认计算新英格兰（地区编号 1）。
可通过 --region-id 2 等参数计算其他地区。
依赖：openpyxl（不需要 pandas，不需要联网）

运行：
  py ".\美国九大地区_单区域E_d选城计算.py"
  py ".\美国九大地区_单区域E_d选城计算.py" --region-id 2
"""

from pathlib import Path
import argparse
import sqlite3
import math
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

REGIONS = {
    "1": "新英格兰 New England",
    "2": "中大西洋 Middle Atlantic",
    "3": "东北中部 East North Central",
    "4": "西北中部 West North Central",
    "5": "南大西洋 South Atlantic",
    "6": "东南中部 East South Central",
    "7": "西南中部 West South Central",
    "8": "山地 Mountain",
    "9": "太平洋 Pacific",
}

BASE = Path(__file__).resolve().parent
DEFAULT_DB = BASE / "美国九大地区_2024MSA综合数据库.sqlite"

def zscore(values):
    n = len(values)
    mean = sum(values) / n
    sd = math.sqrt(sum((x - mean) ** 2 for x in values) / n)  # 总体标准差 ddof=0
    if sd <= 1e-12:
        raise ValueError("某个标量指标的总体标准差为0，不能计算Z-score")
    return [(x - mean) / sd for x in values], mean, sd

def write_sheet(wb, name, headers, rows, widths=None, number_formats=None):
    ws = wb.create_sheet(name)
    ws.append(headers)
    for row in rows:
        ws.append(row)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 32
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="155E75")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col in range(1, ws.max_column + 1):
        letter = get_column_letter(col)
        width = widths[col - 1] if widths and col <= len(widths) else 16
        ws.column_dimensions[letter].width = width
    if number_formats:
        for col_idx, fmt in number_formats.items():
            for r in range(2, ws.max_row + 1):
                ws.cell(r, col_idx).number_format = fmt
    return ws

def write_matrix(wb, name, cities, matrix):
    headers = ["原始清单序号", "MSA", "CBSA"] + [
        f"{c['no']:03d} {c['name']}" for c in cities
    ]
    rows = []
    for i, c in enumerate(cities):
        rows.append([c["no"], c["name"], c["cbsa"]] + [float(v) for v in matrix[i]])
    ws = write_sheet(
        wb, name, headers, rows,
        widths=[12, 48, 12] + [25] * len(cities),
        number_formats={j + 4: "0.000000" for j in range(len(cities))}
    )
    for row in ws.iter_rows(min_row=2, min_col=4):
        for cell in row:
            cell.number_format = "0.000000"
    ws.conditional_formatting.add(
        f"D2:{get_column_letter(ws.max_column)}{ws.max_row}",
        __import__("openpyxl").formatting.rule.ColorScaleRule(
            start_type="min", start_color="FEE2E2",
            mid_type="percentile", mid_value=50, mid_color="FEF3C7",
            end_type="max", end_color="DCFCE7"
        )
    )
    return ws

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--db", type=Path, default=DEFAULT_DB, help="全国 SQLite 数据库路径")
    ap.add_argument("--region-id", default="1", choices=list(REGIONS), help="地区编号1至9，默认1新英格兰")
    ap.add_argument("--output", type=Path, default=None, help="输出Excel路径")
    args = ap.parse_args()

    if not args.db.exists():
        raise FileNotFoundError(f"找不到全国数据库：{args.db}")

    region_id = str(args.region_id)
    region_name = REGIONS[region_id]
    conn = sqlite3.connect(args.db)
    conn.row_factory = sqlite3.Row

    five_rows = conn.execute("""
        SELECT * FROM msa_five_indicators
        WHERE CAST(地区编号 AS TEXT)=?
        ORDER BY 原始清单序号
    """, (region_id,)).fetchall()

    industry_rows = conn.execute("""
        SELECT rowid AS _rowid, * FROM msa_industry_detail
        WHERE CAST(地区编号 AS TEXT)=?
        ORDER BY 原始清单序号, rowid
    """, (region_id,)).fetchall()
    conn.close()

    if len(five_rows) < 2:
        raise ValueError(f"{region_name}城市数量不足2个：{len(five_rows)}")
    if len(industry_rows) != len(five_rows) * 13:
        raise ValueError(
            f"{region_name}产业明细应为{len(five_rows)*13}条，实际{len(industry_rows)}条"
        )

    cities = []
    seen_cbsa = set()
    for r in five_rows:
        cbsa = str(r["CBSA"]).strip()
        if cbsa in seen_cbsa:
            raise ValueError(f"重复CBSA：{cbsa}")
        seen_cbsa.add(cbsa)
        pop = float(r["2024人口"])
        density = float(r["2024人口密度_人每平方英里"])
        gdp = float(r["2024人均GDP_美元"])
        cagr = float(r["2020至2024_CAGR_百分比"])
        if min(pop, density, gdp) <= 0:
            raise ValueError(f"{r['MSA']}人口、密度或GDP存在非正值")
        cities.append({
            "no": int(r["原始清单序号"]),
            "name": str(r["MSA"]),
            "cbsa": cbsa,
            "pop": pop,
            "density": density,
            "gdp": gdp,
            "cagr": cagr,
        })

    # 保留数据库产业明细中的行业出现顺序，和太平洋程序“首次出现顺序”一致。
    sectors = []
    jobs = {c["cbsa"]: {} for c in cities}
    for r in industry_rows:
        cbsa = str(r["CBSA"]).strip()
        if cbsa not in jobs:
            raise ValueError(f"产业表出现五指标表之外的CBSA：{cbsa}")
        sector = str(r["行业"]).strip()
        if sector not in sectors:
            sectors.append(sector)
        if sector in jobs[cbsa]:
            raise ValueError(f"{r['MSA']}行业重复：{sector}")
        male = int(r["男性就业"])
        female = int(r["女性就业"])
        total = int(r["就业合计"])
        if male < 0 or female < 0 or total != male + female:
            raise ValueError(f"{r['MSA']} {sector}男女就业人数加总不符")
        jobs[cbsa][sector] = total

    if len(sectors) != 13:
        raise ValueError(f"应有13类行业，实际{len(sectors)}类")
    for c in cities:
        if set(jobs[c["cbsa"]]) != set(sectors):
            raise ValueError(f"{c['name']}行业类别不完整")

    n = len(cities)
    emp = []
    shares = []
    totals = []
    for c in cities:
        row = [jobs[c["cbsa"]][s] for s in sectors]
        total = sum(row)
        if total <= 0:
            raise ValueError(f"{c['name']}13类行业就业总数为0")
        emp.append(row)
        totals.append(total)
        shares.append([x / total for x in row])
        if abs(sum(shares[-1]) - 1.0) > 1e-12:
            raise ValueError(f"{c['name']}行业占比合计不为100%")

    # 1. Hellinger 距离
    H = [[0.0] * n for _ in range(n)]
    for i in range(n):
        for j in range(n):
            H[i][j] = math.sqrt(
                0.5 * sum(
                    (math.sqrt(shares[i][k]) - math.sqrt(shares[j][k])) ** 2
                    for k in range(13)
                )
            )

    pair_h2 = [H[i][j] ** 2 for i in range(n) for j in range(i + 1, n)]
    h_rms = math.sqrt(sum(pair_h2) / len(pair_h2))
    if h_rms <= 1e-12:
        raise ValueError("产业结构距离RMS为0，无法校准")
    K = math.sqrt(2 * n / (n - 1)) / h_rms
    E = [[H[i][j] * K for j in range(n)] for i in range(n)]

    # 2. 四项标量指标：ln人口、ln密度、GDP/人、人口CAGR；区域内总体Z-score
    raw_vectors = [
        [math.log(c["pop"]) for c in cities],
        [math.log(c["density"]) for c in cities],
        [c["gdp"] for c in cities],
        [c["cagr"] for c in cities],
    ]
    z_vectors = []
    z_stats = []
    for vals in raw_vectors:
        z, mean, sd = zscore(vals)
        z_vectors.append(z)
        z_stats.append((mean, sd))
    Z = [[z_vectors[j][i] for j in range(4)] for i in range(n)]

    # 3. 五指标综合距离 d；五项各20%
    D = [[0.0] * n for _ in range(n)]
    for i in range(n):
        for j in range(n):
            scalar_sq = sum((Z[i][k] - Z[j][k]) ** 2 for k in range(4))
            D[i][j] = math.sqrt(0.2 * (scalar_sq + E[i][j] ** 2))

    # 4. 距离矩阵基本核验
    for i in range(n):
        if abs(D[i][i]) > 1e-10 or abs(E[i][i]) > 1e-10 or abs(H[i][i]) > 1e-10:
            raise AssertionError("距离矩阵对角线不为0")
        for j in range(n):
            if abs(D[i][j] - D[j][i]) > 1e-10:
                raise AssertionError("综合距离矩阵不对称")

    # 5. Medoid：平均距离最小
    single_scores = [sum(D[i][k] for i in range(n)) / n for k in range(n)]
    first = min(range(n), key=lambda k: (single_scores[k], cities[k]["no"]))

    # 6. 固定第一城，逐一测试第二城的双城覆盖改善
    second_candidates = []
    for k in range(n):
        if k == first:
            continue
        d2 = sum(min(D[i][first], D[i][k]) for i in range(n)) / n
        improvement = (single_scores[first] - d2) / single_scores[first]
        second_candidates.append((k, d2, improvement))
    second_candidates.sort(key=lambda x: (x[1], cities[x[0]]["no"]))
    best_second, best_d2, best_improvement = second_candidates[0]
    include_second = best_improvement >= 0.20

    # Workbook
    out = args.output or BASE / f"{region_name.split()[0]}_E与d_城市筛选结果.xlsx"
    wb = Workbook()
    ws0 = wb.active
    ws0.title = "结果汇总"
    header = ["地区", "MSA数量", "配对数", "H_RMS", "K校准系数",
              "第一座代表城市", "D1单城平均距离", "第二城候选",
              "D2双城覆盖距离", "改善率", "达到20%门槛", "最终入选城市数"]
    ws0.append(header)
    ws0.append([
        region_name, n, n*(n-1)//2, h_rms, K,
        cities[first]["name"], single_scores[first],
        cities[best_second]["name"], best_d2, best_improvement,
        "是" if include_second else "否", 2 if include_second else 1
    ])
    ws0.freeze_panes = "A2"
    ws0.sheet_view.showGridLines = False
    ws0.row_dimensions[1].height = 32
    for cell in ws0[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="155E75")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col, width in enumerate([25,12,12,16,18,48,20,48,20,16,16,16], 1):
        ws0.column_dimensions[get_column_letter(col)].width = width
    for col in [4,5,7,9,10]:
        ws0.cell(2,col).number_format = "0.000000"
    ws0.cell(2,10).number_format = "0.00%"

    # Raw five indicators
    raw_rows = []
    for c in cities:
        raw_rows.append([c["no"], c["name"], c["cbsa"], c["pop"], c["density"],
                         c["gdp"], c["cagr"], totals[cities.index(c)],
                         math.log(c["pop"]), math.log(c["density"])])
    write_sheet(wb, "原始五指标", [
        "原始清单序号", "MSA", "CBSA", "2024人口", "人口密度_人/平方英里",
        "2024人均GDP_美元/人", "2020-2024人口CAGR_%", "13类产业就业合计",
        "ln(人口)", "ln(人口密度)"
    ], raw_rows, widths=[12,48,12,16,22,22,20,22,16,20],
       number_formats={4:"#,##0",5:"0.0000",6:"$#,##0",7:"0.000000",8:"#,##0",9:"0.000000",10:"0.000000"})

    # Employment and share matrices
    write_sheet(wb, "产业就业人数矩阵",
        ["原始清单序号","MSA","CBSA"] + sectors + ["13类就业合计"],
        [[c["no"],c["name"],c["cbsa"]] + emp[i] + [totals[i]] for i,c in enumerate(cities)],
        widths=[12,48,12] + [20]*13 + [20])
    write_sheet(wb, "产业占比矩阵",
        ["原始清单序号","MSA","CBSA"] + sectors + ["占比合计"],
        [[c["no"],c["name"],c["cbsa"]] + shares[i] + [sum(shares[i])] for i,c in enumerate(cities)],
        widths=[12,48,12] + [20]*13 + [16],
        number_formats={j+4:"0.0000%" for j in range(14)})

    write_matrix(wb, "Hellinger距离H", cities, H)
    write_matrix(wb, "校准产业距离E", cities, E)

    z_rows = []
    for i,c in enumerate(cities):
        z_rows.append([c["no"],c["name"],c["cbsa"]] + Z[i] + [single_scores[i]])
    write_sheet(wb, "五指标Z值与单城距离",
        ["原始清单序号","MSA","CBSA","Z_ln人口","Z_ln人口密度",
         "Z_人均GDP","Z_CAGR","单城平均距离D1"],
        z_rows, widths=[12,48,12,18,22,18,18,22],
        number_formats={j:"0.000000" for j in range(4,9)})

    write_matrix(wb, "综合距离d", cities, D)

    single_rows = []
    ranking = sorted(range(n), key=lambda i:(single_scores[i],cities[i]["no"]))
    for rank,i in enumerate(ranking,1):
        single_rows.append([rank,cities[i]["no"],cities[i]["name"],cities[i]["cbsa"],
                            single_scores[i],single_scores[i]-single_scores[first]])
    write_sheet(wb, "单城Medoid排名",
        ["排名","原始清单序号","MSA","CBSA","单城平均距离D1","与第一名差值"],
        single_rows, widths=[10,16,48,12,22,20],
        number_formats={5:"0.000000",6:"0.000000"})

    two_rows = []
    for rank,(k,d2,imp) in enumerate(second_candidates,1):
        two_rows.append([rank,cities[first]["name"],cities[k]["name"],cities[k]["cbsa"],
                         d2,imp,"是" if imp>=0.20 else "否"])
    write_sheet(wb, "双城覆盖改善排名",
        ["排名","固定第一城","第二城候选","CBSA","双城平均覆盖距离D2","改善率","达到20%门槛"],
        two_rows, widths=[10,48,48,12,24,16,18],
        number_formats={5:"0.000000",6:"0.00%"})

    # Calculation notes and audit
    notes = [
        ["计算范围", f"{region_name}；从全国SQLite数据库按地区编号{region_id}提取"],
        ["数据规模", f"{n}个MSA；{len(industry_rows)}条产业记录；每城13类行业"],
        ["产业就业占比", "p_ij = 城市i第j类行业就业人数 / 城市i的13类行业就业人数合计"],
        ["Hellinger距离", "H(i,k)=sqrt(0.5 × Σ_j (sqrt(p_ij)-sqrt(p_kj))²)"],
        ["配对数", f"n(n-1)/2 = {n}×{n-1}/2 = {n*(n-1)//2}"],
        ["H_RMS", "sqrt(Σ_{i<k} H(i,k)² / [n(n-1)/2])"],
        ["K校准系数", "K=sqrt(2n/(n-1))/H_RMS"],
        ["产业结构距离", "E(i,k)=K×H(i,k)"],
        ["标量指标转换", "ln(2024人口)、ln(人口密度)；人均GDP、2020-2024 CAGR保留原单位"],
        ["Z-score", "Z=(x-区域均值)/区域总体标准差；标准差采用ddof=0"],
        ["五指标权重", "人口20%；人口密度20%；人均GDP20%；产业结构E 20%；人口CAGR20%"],
        ["综合距离", "d(i,k)=sqrt(0.2×[ΔZ_ln人口²+ΔZ_ln密度²+ΔZ_GDP²+E(i,k)²+ΔZ_CAGR²])"],
        ["第一城", "Medoid：D1(c)=mean_i d(i,c)，选择D1最小城市"],
        ["第二城", "固定第一城，最小化D2=mean_i min[d(i,c1),d(i,c2)]"],
        ["第二城门槛", "改善率=(D1-D2)/D1；改善率≥20%才纳入第二座城市"],
        ["全国数据库的作用", "统一数据源和字段；计算仍在本地区n个MSA内部完成，不用全国195城均值或标准差"],
        ["GDP口径提示", "2024人均GDP来自Newconomy对BEA县级GDP估算按2023 OMB都市区边界汇总后除以2024人口的整理值；不是BEA直接发布的MSA原始表。"],
        ["算法一致性", "本程序沿用上传的太平洋地区V1.0脚本：H、RMS、K、Z-score、d、Medoid、双城覆盖及20%门槛公式均相同。"],
        ["质量检查", "检查城市唯一CBSA、每城13行业、就业人数男女加总、占比合计100%、H/E/d对称、对角线为0。"],
    ]
    ws = write_sheet(wb, "计算说明与核验", ["项目","定义、公式或结果"], notes,
                     widths=[25,120])
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws.row_dimensions[18].height = 54

    # Attach provenance note to GDP header.
    raw_ws = wb["原始五指标"]
    raw_ws["F1"].comment = Comment(
        "来源说明：Newconomy 2024 Metro GDP per capita。其方法说明称以BEA县级GDP估算按2023 OMB都市区边界汇总，再除以2024年人口；并非BEA直接发布的MSA GDP原始表。\\n"
        "https://newconomy.media/rankings/metros/gdp-per-capita",
        "OpenAI"
    )
    wb.properties.title = f"{region_name}产业E与五指标d重新计算"
    wb.properties.subject = "按太平洋地区V1.0方法，从全国SQLite数据库独立提取区域数据计算"
    wb.save(out)

    print(f"地区：{region_name}")
    print(f"城市数：{n}；行业记录：{len(industry_rows)}；配对数：{n*(n-1)//2}")
    print(f"H_RMS={h_rms:.8f}；K={K:.8f}")
    print(f"第一城：{cities[first]['name']}；D1={single_scores[first]:.8f}")
    print(f"第二城候选：{cities[best_second]['name']}；D2={best_d2:.8f}；改善率={best_improvement:.4%}；达标={include_second}")
    print(f"最终入选：{2 if include_second else 1}座")
    print(f"完成：{out.resolve()}")

if __name__ == "__main__":
    main()
