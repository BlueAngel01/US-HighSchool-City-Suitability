#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
E1-E5 教育资源数据采集与 SQLite 建库程序
目标：10个固定MSA；优先MSA口径，缺失时允许有明确标记的局部区域口径。
重要：本程序不会把城市行政边界数据冒充MSA数据，也不会把缺失值填成0。
运行：
  py -m pip install requests beautifulsoup4 pandas openpyxl
  py ".\E1_E5_教育资源数据采集与SQLite建库.py"
数据文件放在程序生成的 data/raw/ 中；程序会登记来源、年份、文件哈希和数据库结构。
"""

from __future__ import annotations
import csv, hashlib, json, re, sqlite3, sys, time, zipfile
from pathlib import Path
from datetime import datetime, timezone
from urllib.parse import urljoin, urlparse
import requests
from bs4 import BeautifulSoup

ROOT = Path(__file__).resolve().parent
RAW = ROOT / "data" / "raw"
DBDIR = ROOT / "data" / "database"
DB = DBDIR / "US_10_MSA_Education_E1_E5.sqlite"
TIMEOUT = 120
HEADERS = {"User-Agent": "US-MSA-HighSchool-Research/1.0 (academic data collection)"}

# 10个固定研究对象。CBSA代码为2023/2024常用MSA代码；运行时仍须用当年CBSA名录复核。
CITIES = [
    ("波士顿", "Boston-Cambridge-Newton, MA-NH", "14460"),
    ("纽约", "New York-Newark-Jersey City, NY-NJ", "35620"),
    ("费城", "Philadelphia-Camden-Wilmington, PA-NJ-DE-MD", "37980"),
    ("克利夫兰", "Cleveland, OH", "17460"),
    ("堪萨斯城", "Kansas City, MO-KS", "28140"),
    ("杰克逊维尔", "Jacksonville, FL", "27260"),
    ("伯明翰", "Birmingham, AL", "13820"),
    ("休斯敦", "Houston-The Woodlands-Sugar Land, TX", "26420"),
    ("凤凰城", "Phoenix-Mesa-Chandler, AZ", "38060"),
    ("圣迭戈", "San Diego-Chula Vista-Carlsbad, CA", "41740"),
]

# 分母定义：任何比率必须同时保存 numerator、denominator、ratio。
INDICATORS = [
    ("E1", "公立高中毕业率", "graduates", "adjusted_cohort", "graduates / adjusted_cohort * 100", "越高越好"),
    ("E2", "高中生师比", "high_school_enrollment", "high_school_teacher_fte", "enrollment / teacher_FTE", "越低越好"),
    ("E3a", "AP/IB课程开设学校覆盖率", "high_schools_offering_AP_or_IB", "all_reporting_high_schools", "offering_schools / all_schools * 100", "越高越好"),
    ("E3b", "AP/IB课程学生参与率", "AP_IB_student_participants", "high_school_enrollment", "participants / enrollment * 100", "越高越好"),
    ("E4a", "STEM课程开设学校覆盖率", "high_schools_offering_defined_STEM", "all_reporting_high_schools", "offering_schools / all_schools * 100", "越高越好"),
    ("E4b", "STEM课程学生参与率", "students_enrolled_in_defined_STEM", "eligible_high_school_enrollment", "participants / eligible_enrollment * 100", "越高越好"),
    ("E5a", "公共图书馆服务网点密度", "active_public_library_outlets", "resident_population", "outlets / population * 100000", "越高越好"),
    ("E5b", "公共图书馆实体借阅率", "physical_items_circulated", "resident_population", "circulation / population", "越高越好"),
    ("E5c", "公共图书馆电子资源可及性", "library_systems_with_defined_e_resources", "library_systems_in_area", "systems_with_resources / systems * 100", "越高越好"),
]

# 官方入口页；脚本先发现并下载官方页面明确列出的数据包，不猜测文件名。
SOURCE_PAGES = [
    ("NCES_CCD_2023_24", "https://nces.ed.gov/ccd/files.asp",
     ["2023-24", "2023–24"], ["public school", "school universe", "nonfiscal"]),
    ("CRDC_2023_24", "https://ocrdata.ed.gov/data",
     ["2023-24", "2023–24"], ["download"]),
    ("IMLS_PLS_FY2024", "https://www.imls.gov/research-evaluation/surveys/public-libraries-survey-pls",
     ["FY 2024"], ["CSV"]),
    ("EDFACTS_ACGR", "https://www.ed.gov/laws-and-policy/education-policy/data-and-statistics",
     ["graduation", "adjusted cohort"], ["data", "download"]),
]

def now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")

def sha256(path: Path):
    h = hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()

def safe_name(s):
    s = re.sub(r"[^A-Za-z0-9._-]+", "_", s)
    return s[:140] or "download"

def connect():
    DBDIR.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB)
    con.execute("PRAGMA foreign_keys=ON")
    con.executescript("""
    CREATE TABLE IF NOT EXISTS cities(
      city_id INTEGER PRIMARY KEY, city_zh TEXT NOT NULL, msa_name_official TEXT NOT NULL,
      cbsa_code TEXT NOT NULL UNIQUE
    );
    CREATE TABLE IF NOT EXISTS indicator_definitions(
      indicator_code TEXT PRIMARY KEY, indicator_name TEXT NOT NULL,
      numerator_definition TEXT NOT NULL, denominator_definition TEXT NOT NULL,
      formula TEXT NOT NULL, preferred_direction TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS source_audit(
      source_id INTEGER PRIMARY KEY, source_code TEXT NOT NULL, source_page TEXT NOT NULL,
      discovered_url TEXT, local_path TEXT, reference_year TEXT, retrieved_at TEXT NOT NULL,
      http_status INTEGER, sha256 TEXT, status TEXT NOT NULL, notes TEXT
    );
    CREATE TABLE IF NOT EXISTS observations(
      observation_id INTEGER PRIMARY KEY,
      city_id INTEGER NOT NULL REFERENCES cities(city_id),
      indicator_code TEXT NOT NULL REFERENCES indicator_definitions(indicator_code),
      numerator REAL, denominator REAL, ratio REAL,
      unit TEXT, reference_year TEXT,
      geography_level TEXT NOT NULL,
      geography_name TEXT,
      coverage_population REAL,
      coverage_population_share REAL,
      source_id INTEGER REFERENCES source_audit(source_id),
      quality_status TEXT NOT NULL DEFAULT '待核验',
      notes TEXT,
      UNIQUE(city_id, indicator_code, reference_year, geography_level, geography_name)
    );
    CREATE TABLE IF NOT EXISTS school_records(
      record_id INTEGER PRIMARY KEY, city_id INTEGER REFERENCES cities(city_id),
      school_nces_id TEXT, school_name TEXT, state TEXT, county_fips TEXT,
      grades_served TEXT, enrollment REAL, teacher_fte REAL,
      acgr_graduates REAL, acgr_cohort REAL, ap_ib_flag INTEGER,
      stem_course_fields TEXT, school_year TEXT, source_id INTEGER REFERENCES source_audit(source_id),
      geography_match TEXT, notes TEXT
    );
    CREATE TABLE IF NOT EXISTS library_records(
      record_id INTEGER PRIMARY KEY, city_id INTEGER REFERENCES cities(city_id),
      library_system_id TEXT, library_name TEXT, county_fips TEXT,
      outlets REAL, physical_circulation REAL, electronic_resources REAL,
      population_service_area REAL, fiscal_year TEXT,
      source_id INTEGER REFERENCES source_audit(source_id),
      geography_match TEXT, notes TEXT
    );
    CREATE TABLE IF NOT EXISTS data_quality_log(
      log_id INTEGER PRIMARY KEY, logged_at TEXT, city_zh TEXT, indicator_code TEXT,
      severity TEXT, message TEXT
    );
    """)
    for zh, en, cbsa in CITIES:
        con.execute("INSERT OR IGNORE INTO cities(city_zh,msa_name_official,cbsa_code) VALUES(?,?,?)",(zh,en,cbsa))
    for row in INDICATORS:
        con.execute("""INSERT OR REPLACE INTO indicator_definitions
          (indicator_code,indicator_name,numerator_definition,denominator_definition,formula,preferred_direction)
          VALUES(?,?,?,?,?,?)""", row)
    con.commit()
    return con

def discover_links(session, code, page_url, year_terms, link_terms):
    r = session.get(page_url, timeout=TIMEOUT, headers=HEADERS)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")
    found = []
    for a in soup.find_all("a", href=True):
        text = " ".join(a.get_text(" ", strip=True).split())
        href = urljoin(page_url, a["href"])
        hay = (text + " " + href).lower()
        year_ok = any(y.lower() in hay for y in year_terms)
        term_ok = any(t.lower() in hay for t in link_terms)
        ext_ok = bool(re.search(r"\.(zip|csv|xlsx?|sas7bdat|pdf)(\?|$)", urlparse(href).path, re.I))
        if ext_ok and (year_ok or code.startswith("IMLS")):
            found.append((text or code, href))
    # 去重并优先保留压缩包/CSV数据文件
    seen, out = set(), []
    for label, url in found:
        if url not in seen:
            seen.add(url); out.append((label, url))
    return r.status_code, out

def download(session, url, dest):
    with session.get(url, stream=True, timeout=TIMEOUT, headers=HEADERS) as r:
        r.raise_for_status()
        dest.parent.mkdir(parents=True, exist_ok=True)
        with dest.open("wb") as f:
            for chunk in r.iter_content(1024 * 1024):
                if chunk: f.write(chunk)
        return r.status_code

def run_downloads(con):
    session = requests.Session()
    for code, page, years, terms in SOURCE_PAGES:
        print(f"\n检查官方数据入口：{code}")
        try:
            status, links = discover_links(session, code, page, years, terms)
            con.execute("""INSERT INTO source_audit
              (source_code,source_page,retrieved_at,http_status,status,notes)
              VALUES(?,?,?,?,?,?)""",(code,page,now(),status,"页面已访问",f"发现候选文件{len(links)}个"))
            sid = con.execute("SELECT last_insert_rowid()").fetchone()[0]
            # 不盲目下载所有页面附件，只下载符合目标年份且文件名/链接含数据关键词的候选项
            selected = []
            for label, url in links:
                low = (label+" "+url).lower()
                if code.startswith("IMLS"):
                    if any(x in low for x in ["csv","pud24","fy24","data"]): selected.append((label,url))
                elif any(x in low for x in ["school","universe","nonfiscal","crdc","data","zip"]):
                    selected.append((label,url))
            for i,(label,url) in enumerate(selected[:8],1):
                ext = Path(urlparse(url).path).suffix or ".bin"
                fn = safe_name(f"{code}_{i}_{label}") + ext
                dest = RAW / code / fn
                try:
                    st = download(session,url,dest)
                    con.execute("""INSERT INTO source_audit
                      (source_code,source_page,discovered_url,local_path,reference_year,retrieved_at,http_status,sha256,status,notes)
                      VALUES(?,?,?,?,?,?,?,?,?,?)""",
                      (code,page,url,str(dest.relative_to(ROOT)), "2023-24" if "2023" in code or "CRDC" in code else "FY2024",
                       now(),st,sha256(dest),"已下载","官方页面发现的文件；需按用户手册解析字段"))
                    print("  已下载：",dest.name)
                except Exception as e:
                    con.execute("""INSERT INTO source_audit
                      (source_code,source_page,discovered_url,retrieved_at,status,notes)
                      VALUES(?,?,?,?,?,?)""",(code,page,url,now(),"下载失败",str(e)))
                    print("  下载失败：",url,e)
            if not links:
                con.execute("INSERT INTO data_quality_log(logged_at,severity,message) VALUES(?,?,?)",
                            (now(),"WARNING",f"{code}入口页未自动识别到数据文件；需检查网页链接结构。"))
        except Exception as e:
            con.execute("INSERT INTO source_audit(source_code,source_page,retrieved_at,status,notes) VALUES(?,?,?,?,?)",
                        (code,page,now(),"访问失败",str(e)))
            print("  访问失败：",e)
        con.commit()

def export_template(con):
    out = ROOT / "data" / "processed"
    out.mkdir(parents=True, exist_ok=True)
    f = out / "E1_E5_observation_template.csv"
    with f.open("w",newline="",encoding="utf-8-sig") as fp:
        w=csv.writer(fp)
        w.writerow(["城市中文名","指标代码","分子实际人数/数量","分母基础人数/数量",
                    "比率","单位","年份","统计层级","统计区域名称",
                    "统计覆盖人口","覆盖人口占MSA人口比例","来源网址","质量状态","备注"])
        for zh,_,_ in CITIES:
            for code,*_ in INDICATORS:
                w.writerow([zh,code,"","","","","","MSA","","","","","待核验","禁止用0代替缺失值"])
    print("\n已生成待核验观测模板：",f)

def print_summary(con):
    print("\n数据库：",DB)
    print("固定城市数：",con.execute("SELECT COUNT(*) FROM cities").fetchone()[0])
    print("指标字段数：",con.execute("SELECT COUNT(*) FROM indicator_definitions").fetchone()[0])
    print("已登记来源文件：",con.execute("SELECT COUNT(*) FROM source_audit").fetchone()[0])
    print("\n重要说明：程序先下载并登记官方原始文件，不能仅凭下载成功就宣称已取得10城可比指标。")
    print("学校必须通过学校地址/County FIPS映射到MSA；局部区域替代必须记录统计区域和覆盖人口比例。")
    print("分母原则：")
    print(" E1 = 调整后毕业队列人数；分子=该队列毕业人数。")
    print(" E2 = 高中学生总人数；分母=高中教师FTE，另存师生比。")
    print(" E3/E4 = 课程参与学生数/符合条件高中学生数；学校覆盖率另行记录。")
    print(" E5 = 图书馆服务网点/服务区人口；借阅量/服务区人口；有电子资源系统数/系统总数。")

def main():
    RAW.mkdir(parents=True,exist_ok=True)
    con=connect()
    run_downloads(con)
    export_template(con)
    print_summary(con)
    con.close()

if __name__ == "__main__":
    main()
