# -*- coding: utf-8 -*-
"""
E1-E5 真实数据采集与处理 V2
目标：10个固定MSA；下载官方原始文件、保留原件、解析可识别表格、建立SQLite与Excel。
本版修复旧版的核心问题：不再把网页上任意CSV/PDF都当成目标数据；明确记录每个数据集是否真的下载、
解压、解析，以及实际记录数。不会用空表冒充已完成数据。
运行：
  py -m pip install requests beautifulsoup4 pandas openpyxl
  py ".\\E1_E5_真实数据采集处理_V2.py"
"""
from pathlib import Path
from urllib.parse import urljoin, urlparse
from datetime import datetime, timezone
import csv, hashlib, json, re, sqlite3, zipfile, io, sys
import requests
import pandas as pd
from bs4 import BeautifulSoup

ROOT=Path(__file__).resolve().parent
RAW=ROOT/"data"/"raw"
PROC=ROOT/"data"/"processed"
DB=ROOT/"data"/"database"/"US_10_MSA_E1_E5.sqlite"
HEAD={"User-Agent":"Mozilla/5.0 (compatible; AcademicResearchDataBot/1.0)"}
TIMEOUT=180

CITIES=[
("波士顿","Boston-Cambridge-Newton, MA-NH","14460"),
("纽约","New York-Newark-Jersey City, NY-NJ","35620"),
("费城","Philadelphia-Camden-Wilmington, PA-NJ-DE-MD","37980"),
("克利夫兰","Cleveland, OH","17460"),
("堪萨斯城","Kansas City, MO-KS","28140"),
("杰克逊维尔","Jacksonville, FL","27260"),
("伯明翰","Birmingham, AL","13820"),
("休斯敦","Houston-The Woodlands-Sugar Land, TX","26420"),
("凤凰城","Phoenix-Mesa-Chandler, AZ","38060"),
("圣迭戈","San Diego-Chula Vista-Carlsbad, CA","41740"),
]
IND=[
("E1","公立高中毕业率","毕业人数","调整后毕业队列人数","毕业人数/调整后毕业队列人数","%"),
("E2","高中生师比","高中学生人数","高中教师FTE","学生人数/教师FTE","学生/教师"),
("E3a","AP/IB开课学校覆盖率","开设AP/IB学校数","符合条件高中总数","开课学校数/高中总数","%"),
("E3b","AP/IB学生参与率","AP/IB参与学生数","高中学生总人数","参与学生数/高中学生总人数","%"),
("E4a","STEM开课学校覆盖率","开设定义内STEM课程学校数","符合条件高中总数","开课学校数/高中总数","%"),
("E4b","STEM学生参与率","STEM课程参与学生数","符合条件高中学生数","参与学生数/符合条件学生数","%"),
("E5a","公共图书馆服务网点密度","开放服务网点数","服务区常住人口","网点数/人口*100000","个/10万人"),
("E5b","公共图书馆实体借阅率","实体资料借阅量","服务区常住人口","借阅量/人口","次/人"),
("E5c","图书馆电子资源覆盖率","提供电子资源系统数","有效图书馆系统总数","有电子资源系统数/系统总数","%"),
]
SOURCES=[
("IMLS_FY2024","https://www.imls.gov/research-evaluation/surveys/public-libraries-survey-pls","FY2024","IMLS FY2024 Public Libraries Survey"),
("NCES_CCD_2023_24","https://nces.ed.gov/ccd/psu_rev.asp","2023-24","NCES Public School Universe"),
("CRDC_2023_24","https://ocrdata.ed.gov/data","2023-24","CRDC 2023-24 public-use data"),
("CENSUS_CBSA_COUNTY_CROSSWALK","https://www.census.gov/geographies/reference-files/time-series/demo/metro-micro/delineation-files.html","2023/2024","CBSA county delineation crosswalk"),
]
def utc(): return datetime.now(timezone.utc).isoformat(timespec="seconds")
def sha(path): 
 h=hashlib.sha256()
 with open(path,"rb") as f:
  for b in iter(lambda:f.read(1024*1024),b""): h.update(b)
 return h.hexdigest()
def slug(s): return re.sub(r"[^A-Za-z0-9._-]+","_",s)[:120]
def dbopen():
 DB.parent.mkdir(parents=True,exist_ok=True)
 c=sqlite3.connect(DB)
 c.executescript("""
 CREATE TABLE IF NOT EXISTS cities(city_zh TEXT,msa_name TEXT,cbsa TEXT PRIMARY KEY);
 CREATE TABLE IF NOT EXISTS indicator_definitions(code TEXT PRIMARY KEY,name TEXT,numerator TEXT,denominator TEXT,formula TEXT,unit TEXT);
 CREATE TABLE IF NOT EXISTS source_files(id INTEGER PRIMARY KEY,source_code TEXT,url TEXT,local_file TEXT,year TEXT,downloaded_at TEXT,http_status INTEGER,sha256 TEXT,bytes INTEGER,parse_status TEXT,record_count INTEGER,notes TEXT);
 CREATE TABLE IF NOT EXISTS observations(id INTEGER PRIMARY KEY,city_zh TEXT,cbsa TEXT,indicator_code TEXT,numerator REAL,denominator REAL,value REAL,unit TEXT,year TEXT,geography_level TEXT,geography_name TEXT,coverage_population REAL,msa_population REAL,coverage_share REAL,source_file_id INTEGER,quality TEXT,notes TEXT);
 CREATE TABLE IF NOT EXISTS raw_imls_systems(source_file_id INTEGER,row_json TEXT);
 CREATE TABLE IF NOT EXISTS raw_imls_outlets(source_file_id INTEGER,row_json TEXT);
 CREATE TABLE IF NOT EXISTS audit(id INTEGER PRIMARY KEY,logged_at TEXT,stage TEXT,status TEXT,detail TEXT);
 """)
 c.executemany("INSERT OR REPLACE INTO cities VALUES(?,?,?)",CITIES)
 c.executemany("INSERT OR REPLACE INTO indicator_definitions VALUES(?,?,?,?,?,?)",IND)
 c.commit(); return c
def audit(c,stage,status,detail):
 c.execute("INSERT INTO audit(logged_at,stage,status,detail) VALUES(?,?,?,?)",(utc(),stage,status,detail)); c.commit()
def fetch_page(url):
 r=requests.get(url,headers=HEAD,timeout=TIMEOUT); r.raise_for_status()
 return r.text
def links_for(url,html):
 soup=BeautifulSoup(html,"html.parser"); out=[]
 for a in soup.find_all("a",href=True):
  href=urljoin(url,a["href"]); label=" ".join(a.get_text(" ",strip=True).split())
  if re.search(r"\.(zip|csv|xlsx|txt)(?:\?|$)",urlparse(href).path,re.I):
   out.append((label,href))
 return out
def download(url,dest):
 dest.parent.mkdir(parents=True,exist_ok=True)
 with requests.get(url,headers=HEAD,stream=True,timeout=TIMEOUT) as r:
  r.raise_for_status()
  with open(dest,"wb") as f:
   for chunk in r.iter_content(1024*1024):
    if chunk:f.write(chunk)
  return r.status_code
def choose_link(code,links):
 candidates=[]
 for label,url in links:
  x=(label+" "+url).lower()
  if code=="IMLS_FY2024":
   if ("csv" in x or "pud24i" in x) and ("2024" in x or "fy24" in x): candidates.append((label,url))
  elif code=="NCES_CCD_2023_24":
   if "2023-24" in x or "2023_24" in x or "2023%2d24" in x: candidates.append((label,url))
  elif code=="CRDC_2023_24":
   if "2023" in x or "2024" in x or "zip" in x: candidates.append((label,url))
  elif code=="CENSUS_CBSA_COUNTY_CROSSWALK":
   if "2023" in x and ("county" in x or "delineation" in x): candidates.append((label,url))
 return candidates
def get_sources(c):
 s=requests.Session()
 for code,page,year,desc in SOURCES:
  print("\n获取",desc)
  try:
   html=fetch_page(page); links=links_for(page,html); selected=choose_link(code,links)
   if not selected:
    audit(c,code,"入口可访问但未识别下载文件",f"页面链接数={len(links)}；需检查官方页面更新")
    print("  页面已访问，但未自动识别到目标数据附件。")
    continue
   # IMLS CSV ZIP只下载一个系统文件和一个网点文件；其余来源最多下载3个匹配文件
   lim=4 if code=="IMLS_FY2024" else 3
   for label,url in selected[:lim]:
    ext=Path(urlparse(url).path).suffix or ".dat"
    dest=RAW/code/(slug(label or code)+ext)
    try:
     st=download(url,dest)
     cur=c.execute("""INSERT INTO source_files(source_code,url,local_file,year,downloaded_at,http_status,sha256,bytes,parse_status,notes)
      VALUES(?,?,?,?,?,?,?,?,?,?)""",(code,url,str(dest.relative_to(ROOT)),year,utc(),st,sha(dest),dest.stat().st_size,"待解析",desc))
     c.commit(); sid=cur.lastrowid
     print(f"  下载成功 {dest.name} ({dest.stat().st_size:,} bytes)")
     if code=="IMLS_FY2024": parse_imls_zip(c,sid,dest)
    except Exception as e:
     audit(c,code,"下载失败",f"{url} :: {e}"); print("  下载失败",e)
  except Exception as e:
   audit(c,code,"入口访问失败",str(e)); print("  入口访问失败",e)
def read_csv_member(zf,member):
 raw=zf.read(member)
 for enc in ("utf-8-sig","cp1252","latin1"):
  try:return pd.read_csv(io.BytesIO(raw),encoding=enc,low_memory=False)
  except Exception:pass
 raise ValueError("无法识别CSV编码/结构: "+member)
def parse_imls_zip(c,sid,path):
 try:
  with zipfile.ZipFile(path) as z:
   members=[m for m in z.namelist() if m.lower().endswith(".csv")]
   if not members: raise ValueError("ZIP内没有CSV")
   # identify AE(system) vs outlet by filename/member names
   member=members[0]
   df=read_csv_member(z,member)
   c.executemany("INSERT INTO raw_imls_systems(source_file_id,row_json) VALUES(?,?)",
     [(sid,json.dumps(row,ensure_ascii=False,default=str)) for row in df.to_dict("records")])
   c.execute("UPDATE source_files SET parse_status='已解析为系统原始记录',record_count=? WHERE id=?",(len(df),sid))
   c.commit()
   print(f"    CSV解析成功：{len(df):,}条；字段数={len(df.columns)}")
   print("    字段预览：",", ".join(map(str,df.columns[:24])))
 except Exception as e:
  c.execute("UPDATE source_files SET parse_status='解析失败',notes=? WHERE id=?",(str(e),sid));c.commit()
  audit(c,"IMLS解析","失败",str(e))
def export_excel(c):
 from openpyxl import Workbook
 from openpyxl.styles import Font,PatternFill,Alignment
 from openpyxl.utils import get_column_letter
 PROC.mkdir(parents=True,exist_ok=True)
 x=PROC/"E1_E5_真实数据采集结果.xlsx"
 wb=Workbook(); ws=wb.active;ws.title="指标结果"
 ws.append(["城市","CBSA","指标","分子","分母","结果","单位","年份","统计层级","统计区域","覆盖人口","MSA人口","覆盖比例","质量状态","备注"])
 rows=c.execute("SELECT city_zh,cbsa,indicator_code,numerator,denominator,value,unit,year,geography_level,geography_name,coverage_population,msa_population,coverage_share,quality,notes FROM observations ORDER BY city_zh,indicator_code").fetchall()
 for r in rows:ws.append(list(r))
 if not rows: ws.append(["尚无完成核验的指标数据","","","","","","","","","","","","","待采集","空白不代表0；请勿用于论文"])
 w2=wb.create_sheet("来源下载审计");w2.append(["来源","网址","本地文件","年份","下载时间","状态码","SHA256","字节数","解析状态","记录数","说明"])
 for r in c.execute("SELECT source_code,url,local_file,year,downloaded_at,http_status,sha256,bytes,parse_status,record_count,notes FROM source_files"):w2.append(list(r))
 w3=wb.create_sheet("城市与指标");w3.append(["城市","MSA名称","CBSA"])
 for r in CITIES:w3.append(list(r))
 w4=wb.create_sheet("指标定义");w4.append(["代码","指标","分子","分母","公式","单位"])
 for r in IND:w4.append(list(r))
 w5=wb.create_sheet("审计日志");w5.append(["时间","阶段","状态","详情"])
 for r in c.execute("SELECT logged_at,stage,status,detail FROM audit"):w5.append(list(r))
 for sh in wb.worksheets:
  sh.freeze_panes="A2";sh.auto_filter.ref=sh.dimensions;sh.sheet_view.showGridLines=False
  for cell in sh[1]:
   cell.font=Font(bold=True,color="FFFFFF");cell.fill=PatternFill("solid",fgColor="24476B");cell.alignment=Alignment(wrap_text=True)
  for col in sh.columns:
   letter=get_column_letter(col[0].column)
   sh.column_dimensions[letter].width=min(42,max(14,max((len(str(z.value or "")) for z in col[:300]),default=12)+2))
 wb.save(x);print("\nExcel已生成：",x)
def main():
 RAW.mkdir(parents=True,exist_ok=True); c=dbopen()
 get_sources(c)
 print("\n---真实数据采集审计---")
 for r in c.execute("SELECT source_code,COUNT(*),SUM(CASE WHEN parse_status LIKE '已解析%' THEN 1 ELSE 0 END),SUM(COALESCE(record_count,0)) FROM source_files GROUP BY source_code"):
  print(r)
 export_excel(c)
 print("SQLite：",DB)
 print("注意：NCES/CRDC/毕业率的原始文件若下载失败，会明确保留失败记录，不会生成虚假城市指标。")
 c.close()
if __name__=="__main__":main()
