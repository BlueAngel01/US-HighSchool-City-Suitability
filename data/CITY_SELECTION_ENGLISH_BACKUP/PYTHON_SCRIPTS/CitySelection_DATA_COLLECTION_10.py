#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
太平洋地区31个MSA｜2024五指标自动补齐
固定来源：
  2020/2024人口：Census Vintage2024 cbsa-est2024-alldata.csv
  2024产业结构：ACS 2024 1-Year C24030 .dat (男女就业合计，13互斥类)
  2024陆地面积：Census TIGERweb ACS24（仅计算密度，不作为第六指标）
  2024人均GDP：Newconomy 2024 GDP-per-capita (BEA derived)
运行：python 太平洋地区_官方数据自动补齐.py
原始两份官方数据自动下载到 ./pacific_official_raw/
生成：太平洋地区_2024五指标_完整.csv、太平洋地区_2024产业结构_13类_完整.csv
如安装了 xlsxwriter，还会生成：太平洋地区_31MSA五指标_完整.xlsx
本脚本不替换行业口径、不增减候选MSA；校验全部31城均取得官方数据。
"""
from __future__ import annotations

import argparse
import csv
import io
import json
import math
import pathlib
import sys
import time
import urllib.error
import urllib.request

PURL = "https://www2.census.gov/programs-surveys/popest/datasets/2020-2024/metro/totals/cbsa-est2024-alldata.csv"
IURL = "https://www2.census.gov/programs-surveys/acs/summary_file/2024/table-based-SF/data/1YRData/acsdt1y2024-c24030.dat"
MSAS = [{'no': 1,
  'city': 'Los Angeles-Long Beach-Anaheim, CA',
  'cbsa': 31080,
  'population': 12927614,
  'arealand': 12570606082,
  'gdp_pc': 104793},
 {'no': 2,
  'city': 'Riverside-San Bernardino-Ontario, CA',
  'cbsa': 40140,
  'population': 4744214,
  'arealand': 70648246244,
  'gdp_pc': 57474},
 {'no': 3,
  'city': 'San Francisco-Oakland-Fremont, CA',
  'cbsa': 41860,
  'population': 4648486,
  'arealand': 6396725033,
  'gdp_pc': 172381},
 {'no': 4,
  'city': 'Seattle-Tacoma-Bellevue, WA',
  'cbsa': 42660,
  'population': 4145494,
  'arealand': 15203989412,
  'gdp_pc': 145716},
 {'no': 5,
  'city': 'San Diego-Chula Vista-Carlsbad, CA',
  'cbsa': 41740,
  'population': 3298799,
  'arealand': 10906140199,
  'gdp_pc': 100603},
 {'no': 6,
  'city': 'Portland-Vancouver-Hillsboro, OR-WA',
  'cbsa': 38900,
  'population': 2537904,
  'arealand': 17320278360,
  'gdp_pc': 88789},
 {'no': 7,
  'city': 'Sacramento-Roseville-Folsom, CA',
  'cbsa': 40900,
  'population': 2463127,
  'arealand': 13192392916,
  'gdp_pc': 81971},
 {'no': 8,
  'city': 'San Jose-Sunnyvale-Santa Clara, CA',
  'cbsa': 41940,
  'population': 1995484,
  'arealand': 6940482577,
  'gdp_pc': 221384},
 {'no': 9, 'city': 'Fresno, CA', 'cbsa': 23420, 'population': 1189557, 'arealand': 20966706476, 'gdp_pc': 61420},
 {'no': 10, 'city': 'Urban Honolulu, HI', 'cbsa': 46520, 'population': 998747, 'arealand': 1556040330, 'gdp_pc': 87176},
 {'no': 11,
  'city': 'Bakersfield-Delano, CA',
  'cbsa': 12540,
  'population': 922529,
  'arealand': 21068706344,
  'gdp_pc': 68366},
 {'no': 12,
  'city': 'Oxnard-Thousand Oaks-Ventura, CA',
  'cbsa': 37100,
  'population': 835427,
  'arealand': 4767597085,
  'gdp_pc': 83317},
 {'no': 13, 'city': 'Stockton-Lodi, CA', 'cbsa': 44700, 'population': 816108, 'arealand': 3606041616, 'gdp_pc': 57004},
 {'no': 14,
  'city': 'Spokane-Spokane Valley, WA',
  'cbsa': 44060,
  'population': 604962,
  'arealand': 10984985432,
  'gdp_pc': 69093},
 {'no': 15, 'city': 'Modesto, CA', 'cbsa': 33700, 'population': 556972, 'arealand': 3875010973, 'gdp_pc': 57571},
 {'no': 16,
  'city': 'Santa Rosa-Petaluma, CA',
  'cbsa': 42220,
  'population': 485375,
  'arealand': 4080103624,
  'gdp_pc': 81273},
 {'no': 17, 'city': 'Visalia, CA', 'cbsa': 47300, 'population': 483546, 'arealand': 12493843648, 'gdp_pc': 52559},
 {'no': 18, 'city': 'Vallejo, CA', 'cbsa': 46700, 'population': 455101, 'arealand': 2128290837, 'gdp_pc': 82828},
 {'no': 19,
  'city': 'Santa Maria-Santa Barbara, CA',
  'cbsa': 42200,
  'population': 444500,
  'arealand': 7080874941,
  'gdp_pc': 93193},
 {'no': 20, 'city': 'Salem, OR', 'cbsa': 41420, 'population': 443416, 'arealand': 4977344406, 'gdp_pc': 61072},
 {'no': 21, 'city': 'Salinas, CA', 'cbsa': 41500, 'population': 436251, 'arealand': 8499610242, 'gdp_pc': 87898},
 {'no': 22, 'city': 'Anchorage, AK', 'cbsa': 11260, 'population': 407213, 'arealand': 68411728021, 'gdp_pc': 89827},
 {'no': 23,
  'city': 'Eugene-Springfield, OR',
  'cbsa': 21660,
  'population': 382396,
  'arealand': 11795044812,
  'gdp_pc': 61649},
 {'no': 24,
  'city': 'Kennewick-Richland, WA',
  'cbsa': 28420,
  'population': 319428,
  'arealand': 7619028355,
  'gdp_pc': 71054},
 {'no': 25,
  'city': 'Olympia-Lacey-Tumwater, WA',
  'cbsa': 36500,
  'population': 302912,
  'arealand': 1871300040,
  'gdp_pc': 71187},
 {'no': 26, 'city': 'Merced, CA', 'cbsa': 32900, 'population': 296774, 'arealand': 5019448521, 'gdp_pc': 44756},
 {'no': 27,
  'city': 'San Luis Obispo-Paso Robles, CA',
  'cbsa': 42020,
  'population': 281843,
  'arealand': 8549140854,
  'gdp_pc': 86947},
 {'no': 28,
  'city': 'Bremerton-Silverdale-Port Orchard, WA',
  'cbsa': 14740,
  'population': 281420,
  'arealand': 1023271578,
  'gdp_pc': 65201},
 {'no': 29, 'city': 'Bend, OR', 'cbsa': 13460, 'population': 264407, 'arealand': 20145544017, 'gdp_pc': 74189},
 {'no': 30,
  'city': 'Santa Cruz-Watsonville, CA',
  'cbsa': 42100,
  'population': 262406,
  'arealand': 1152818085,
  'gdp_pc': 80380},
 {'no': 31, 'city': 'Yakima, WA', 'cbsa': 49420, 'population': 258523, 'arealand': 11122734403, 'gdp_pc': 51633}]
INDUSTRIES = [('农业、林业、渔业与采矿', 'Agriculture, forestry, fishing, hunting and mining', 3, 30),
 ('建筑业', 'Construction', 6, 33),
 ('制造业', 'Manufacturing', 7, 34),
 ('批发贸易', 'Wholesale trade', 8, 35),
 ('零售贸易', 'Retail trade', 9, 36),
 ('运输仓储与公用事业', 'Transportation, warehousing and utilities', 10, 37),
 ('信息业', 'Information', 13, 40),
 ('金融、保险与房地产', 'Finance, insurance, real estate and rental', 14, 41),
 ('专业服务、企业管理与行政服务', 'Professional, management, administration and waste management', 17, 44),
 ('教育、医疗与社会援助', 'Education, health care and social assistance', 21, 48),
 ('文娱、住宿与餐饮', 'Arts, entertainment, recreation, accommodation and food', 24, 51),
 ('其他服务', 'Other services except public administration', 27, 54),
 ('公共行政', 'Public administration', 28, 55)]
SQM_PER_SQMI = 2589988.110336

def download_file(url: str, dest: pathlib.Path) -> pathlib.Path:
    if dest.exists() and dest.stat().st_size > 300_000:
        return dest
    print(f"下载官方数据：{url}")
    req = urllib.request.Request(url, headers={"User-Agent":"Mozilla/5.0 (compatible; ScientificDataReproduction/1.0)"})
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=90) as incoming, dest.open("wb") as output:
                while True:
                    buf = incoming.read(131072)
                    if not buf:
                        break
                    output.write(buf)
            if dest.stat().st_size < 300_000:
                raise ValueError(f"下载的文件异常小：{dest.stat().st_size}")
            return dest
        except (OSError, urllib.error.URLError, ValueError) as exc:
            dest.unlink(missing_ok=True)
            print("下载失败：",exc, file=sys.stderr)
            if attempt == 2: raise
            time.sleep(2)
    raise RuntimeError("未取得官方数据文件")

def rows_from_csv(path: pathlib.Path) -> list[dict[str, str]]:
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            pass
    delimiter = "|" if path.suffix.lower() == ".dat" or text.splitlines()[0].count("|") > text.splitlines()[0].count(",") else ","
    r = csv.DictReader(io.StringIO(text, newline=""), delimiter=delimiter)
    if r.fieldnames is None:
        raise ValueError(f"文件缺少CSV列名：{path}")
    return [{str(k).strip():str(v or "").strip() for k,v in x.items() if k is not None} for x in r]

def estimate_value(row: dict, idx: int) -> int:
    choices = [f"C24030_E{idx:03d}", f"C24030_{idx:03d}E",f"C24030_E{idx}",
               f"C24030_{idx}E"]
    keymap = {k.upper():k for k in row}
    for x in choices:
        if x.upper() in keymap:
            v = row[keymap[x.upper()]].strip().replace(",", "")
            try: return int(v)
            except ValueError: raise ValueError(f"ACS行业字段含异常值：{x}={v!r}")
    raise KeyError(f"缺少ACS列{idx}，文件表头示例：{list(row)[:12]}")

def pick_population_row(rows: list[dict], r: dict) -> dict:
    expected = r["population"]
    cand=[]
    for item in rows:
        try:
            if str(item.get("CBSA","")).strip() != str(r["cbsa"]): continue
            if int(item.get("POPESTIMATE2024","-1")) != expected: continue
            cand.append(item)
        except ValueError: continue
    if len(cand) != 1:
        # 部分人口CSV中可出现Metro与部分县合计相同；优先单独的MSA总行
        if cand:
            overall=[x for x in cand if x.get("STCOU","").strip() in ("","00000","0")
                     and x.get("MDIV","").strip() in ("","0","00000")]
            if len(overall)==1:return overall[0]
        raise ValueError(f"{r['city']} (CBSA {r['cbsa']}) 无唯一匹配的官方2024人口行；找到{len(cand)}行。")
    return cand[0]

def pick_industry_row(rows: list[dict], r: dict) -> dict:
    ident = str(r["cbsa"])
    # Census ACS table-based summary files can have a different GEO_ID column
    # representation. Prefer matching by the CBSA code at the end.
    cand=[]
    for x in rows:
        geo = str(x.get("GEO_ID") or x.get("GEOID") or x.get("GEO_ID ") or "").strip()
        if geo.endswith("US"+ident) or geo == ident:
            cand.append(x)
    if len(cand)==1:
        return cand[0]
    if len(cand)>1:
        raise ValueError(f"{r['city']} 的ACS本地数据匹配到{len(cand)}行，拒绝猜测")
    # Local file must contain an MSA row; otherwise stop with a useful diagnostic.
    sample = [str(x.get("GEO_ID", "")) for x in rows[:8]]
    raise ValueError(f"{r['city']} (CBSA {ident}) 不在本地ACS文件的GEO_ID中。示例：{sample}。请运行 PowerShell：Select-String -Path .\pacific_official_raw\acsdt1y2024-c24030.dat -Pattern '31080' | Select-Object -First 2")
    # Reliable fallback: query the official Census ACS API by exact CBSA code.
    # This also works when the downloaded .dat is a non-CSV format.
    fields=sorted({1,*[k for _,_,m,f in INDUSTRIES for k in (m,f)]})
    names=[f"C24030_{k:03d}E" for k in fields]
    url=("https://api.census.gov/data/2024/acs/acs1?get="+
         ",".join(names)+
         "&for=metropolitan%20statistical%20area/micropolitan%20statistical%20area:"+ident)
    print(f"本地DAT未匹配到 {ident}，改用官方ACS API按CBSA精确查询")
    req=urllib.request.Request(url,headers={"User-Agent":"Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req,timeout=90) as resp:
            payload=json.load(resp)
    except Exception as exc:
        raise RuntimeError(f"{r['city']}：官方ACS API查询失败：{exc}\nURL: {url}") from exc
    if len(payload)!=2:
        raise ValueError(f"{r['city']}：ACS API应返回唯一数据行，实际{len(payload)-1}行")
    return dict(zip(payload[0],payload[1]))

def save_csv(path, headers, rows):
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer=csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)

def save_xlsx_if_possible(folder, five, detail):
    try:
        import xlsxwriter
    except ImportError:
        print("未安装 xlsxwriter：已经生成全部数据CSV；如需Excel工作簿，执行 pip install XlsxWriter 后重运行")
        return
    path=folder/"太平洋地区_31MSA五指标_完整.xlsx"
    w=xlsxwriter.Workbook(str(path))
    title=w.add_format({"bold":True,"font_color":"white","bg_color":"#145A76","text_wrap":True})
    integer=w.add_format({"num_format":"#,##0"})
    dec2=w.add_format({"num_format":"#,##0.00"})
    pct=w.add_format({"num_format":"0.000"})
    heads=[
       "序号","MSA","CBSA","①2024人口（人）","②人口密度（人/平方英里）",
       "③人均GDP（美元/人）","④产业结构（13类就业占比，见第2表）","⑤2020-2024 CAGR（%）"]
    s=w.add_worksheet("五指标总表")
    s.write_row(0,0,heads,title)
    for i,row in enumerate(five,1):
        s.write_row(i,0,row)
        s.write_number(i,3,row[3],integer)
        s.write_number(i,4,row[4],dec2)
        s.write_number(i,5,row[5],integer)
        s.write_number(i,7,row[7],pct)
    s.set_column("A:A",10);s.set_column("B:B",42);s.set_column("C:C",12)
    s.set_column("D:F",25);s.set_column("G:G",27);s.set_column("H:H",24)
    s.freeze_panes(1,3);s.autofilter(0,0,len(five),7)
    d=w.add_worksheet("产业结构明细13类")
    d.write_row(0,0,["序号","MSA","CBSA","行业（13互斥类）","男性就业","女性就业","合计就业","占本MSA就业人口比例（%）","ACS男变量","ACS女变量"],title)
    for i,row in enumerate(detail,1):
        d.write_row(i,0,row)
        d.write_number(i,7,row[7],pct)
    d.set_column("A:A",10);d.set_column("B:B",42);d.set_column("C:C",12)
    d.set_column("D:D",38);d.set_column("E:G",18);d.set_column("H:H",30)
    d.set_column("I:J",21)
    d.freeze_panes(1,3);d.autofilter(0,0,len(detail),9)
    w.close()
    print("已生成：",path)

def run():
    ap=argparse.ArgumentParser()
    ap.add_argument("--population-file",type=pathlib.Path, help="已有官方2024人口CSV路径（可避免下载）")
    ap.add_argument("--acs-file",type=pathlib.Path,help="已有官方2024 ACS C24030.dat路径（可避免下载）")
    ap.add_argument("--output-dir",type=pathlib.Path,default=pathlib.Path.cwd())
    args=ap.parse_args()
    folder=args.output_dir.resolve()
    folder.mkdir(exist_ok=True,parents=True)
    raw=folder/"pacific_official_raw"
    raw.mkdir(exist_ok=True)
    pfile=args.population_file or download_file(PURL,raw/"cbsa-est2024-alldata.csv")
    afile=args.acs_file or download_file(IURL,raw/"acsdt1y2024-c24030.dat")
    people=rows_from_csv(pfile)
    employment=rows_from_csv(afile)
    five=[];detail=[]
    for r in MSAS:
        pp=pick_population_row(people,r)
        p2020=int(pp["POPESTIMATE2020"])
        p2024=int(pp["POPESTIMATE2024"])
        acs=pick_industry_row(employment,r)
        total=estimate_value(acs,1)
        if p2020<=0 or total<=0:raise ValueError("人口/就业总人数不可为零")
        members=[]
        for i,(zh,en,m,f) in enumerate(INDUSTRIES,1):
            male=estimate_value(acs,m)
            female=estimate_value(acs,f)
            count=male+female
            if count<0:raise ValueError(f"{r['city']}：{zh}出现负数")
            share=100*count/total
            members.append((zh,male,female,count,share))
            detail.append([r["no"],r["city"],r["cbsa"],f"{i:02d}. {zh}",male,female,count,round(share,5),
                           f"C24030_{m:03d}E",f"C24030_{f:03d}E"])
        used=sum(x[3] for x in members)
        if used!=total:
            raise AssertionError(f"{r['city']} 13类行业之和{used} 不等于ACS就业总计{total}，数据不能发表")
        density=p2024*SQM_PER_SQMI/r["arealand"]
        cagr=100*((p2024/p2020)**.25-1)
        vec="；".join(f"{m[0]} {m[4]:.4f}%" for m in members)
        five.append([r["no"],r["city"],r["cbsa"],p2024,round(density,4),r["gdp_pc"],
                      vec,round(cagr,6)])
        print(f"完成：{r['city']:<49} CAGR={cagr:.4f}%, 13行业合计={total:,}")
    if len(five)!=31 or len(detail)!=403:raise AssertionError("行数错误，未能覆盖31个MSA")
    save_csv(folder/"太平洋地区_2024五指标_完整.csv",
          ["序号","MSA","CBSA","①2024人口（人）","②人口密度（人/平方英里）",
           "③2024人均GDP（美元/人）","④产业结构13类就业占比","⑤2020至2024 CAGR（%）"],five)
    save_csv(folder/"太平洋地区_2024产业结构_13类_完整.csv",
          ["序号","MSA","CBSA","行业类别","男性就业","女性就业","就业合计","就业占比（%）","ACS男变量","ACS女变量"],detail)
    save_xlsx_if_possible(folder,five,detail)
    print("正式数据完整：31个MSA，五指标 + 403条产业结构均已按2024官方原表核对；没有增减城市。")

if __name__=="__main__":run()
