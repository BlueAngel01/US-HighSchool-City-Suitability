# -*- coding: utf-8 -*-
"""
R1 high-school club directory collector (10 MSA pilot)
Purpose: fetch publicly listed club/organization names from official school pages,
deduplicate names, and export an auditable workbook for manual verification.

Install:
    py -m pip install requests beautifulsoup4 pandas openpyxl lxml
Run:
    py R1_collect_school_clubs.py

Output:
    R1_school_clubs_collection/
      R1_club_records_raw.csv
      R1_club_records_review.xlsx
      R1_collection_log.csv

IMPORTANT:
- This is a semi-automated evidence collector, not a complete census.
- It does NOT infer zero clubs when a page has no list.
- It does NOT automatically treat every list item as a valid club.
- Review "candidate_names" and remove sports teams, student government,
  class boards, duplicate entries, and non-club activities according to
  the study's chosen definition.
- The seeded sources below are official school/district pages found online.
  Add more official school URLs to SEEDS as additional schools are identified.
"""

from pathlib import Path
from urllib.parse import urlparse
import re, time, json
import requests
import pandas as pd
from bs4 import BeautifulSoup

OUT = Path("R1_school_clubs_collection")
OUT.mkdir(exist_ok=True)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; R1AcademicResearch/1.0; contact: research)"
}
TIMEOUT = 35

# Seed pages identified from official school/district websites.
# These are verified starting sources across all 10 MSAs, not a representative 15-school sample per MSA.
SEEDS = [
    # New York MSA
    {"msa":"New York","school":"Cristo Rey New York High School","state":"NY","url":"https://www.cristoreyny.org/clubs-and-activities","source_year":"Current page; verify school year"},
    {"msa":"New York","school":"White Plains High School","state":"NY","url":"https://hs.whiteplainspublicschools.org/clubs-activities-at-wphs","source_year":"2025-26"},
    # Philadelphia MSA
    {"msa":"Philadelphia","school":"Cristo Rey Philadelphia High School","state":"PA","url":"https://www.cristoreyphiladelphia.org/student-life/clubs-student-organizations","source_year":"Current page; verify school year"},
    {"msa":"Philadelphia","school":"Kensington Health Sciences Academy","state":"PA","url":"https://khsa.philasd.org/students/","source_year":"Current page; verify school year"},
    {"msa":"Philadelphia","school":"George Washington High School","state":"PA","url":"https://gwhs.philasd.org/ataglance/","source_year":"Current page; verify school year"},
    {"msa":"Philadelphia","school":"Girard Academic Music Program","state":"PA","url":"https://gamp.philasd.org/extracurricular-activities/","source_year":"Current page; verify school year"},
    # Boston MSA
    {"msa":"Boston","school":"Fenway High School","state":"MA","url":"https://www.fenwayhs.org/student-life/student-clubs-and-groups","source_year":"Current page; year varies"},
    {"msa":"Boston","school":"Cathedral High School","state":"MA","url":"https://www.cathedralboston.org/student-life/clubs-and-activities/","source_year":"Current page; year varies"},
    {"msa":"Boston","school":"Boston Prep","state":"MA","url":"https://www.bostonprep.org/student-experience/extracurriculars/","source_year":"SY2026-27 shown; not aligned to 2024-25"},
    {"msa":"Boston","school":"New Mission High School","state":"MA","url":"https://www.bostonpublicschools.org/schools-container/school-profile/new-mission-high-school","source_year":"Current BPS profile; verify year"},
    # Cleveland MSA
    {"msa":"Cleveland","school":"John Hay High School","state":"OH","url":"https://johnhay.clevelandmetroschools.org/jhhs-students/clubs-activities","source_year":"Current page; year varies"},
    {"msa":"Cleveland","school":"Horizon Science Academy Cleveland High School","state":"OH","url":"https://www.hsas.org/apps/pages/index.jsp?pREC_ID=2429764&type=d&uREC_ID=2638845","source_year":"2026-27 booklet referenced; check PDF"},
    {"msa":"Cleveland","school":"Hathaway Brown Upper School","state":"OH","url":"https://www.hb.edu/campus-life/clubs-activities","source_year":"Current page; mixed K-12/grade bands; manual filter required"},
    {"msa":"Cleveland","school":"St. Edward High School","state":"OH","url":"https://www.sehs.net/student-life/clubs-extracurriculars","source_year":"Current page; year varies"},
    # Kansas City MSA
    {"msa":"Kansas City","school":"Winnetonka High School","state":"MO","url":"https://wths.nkcschools.org/activities/clubs-get-involved","source_year":"Current page; year varies"},
    {"msa":"Kansas City","school":"Notre Dame de Sion High School","state":"MO","url":"https://www.ndsion.edu/high-school/student-life/clubs-and-activities","source_year":"Current page; high school section"},
    {"msa":"Kansas City","school":"Cristo Rey Kansas City","state":"MO","url":"https://www.cristoreykc.org/school-life/clubs","source_year":"Current page; year varies"},
    {"msa":"Kansas City","school":"Bishop Ward High School","state":"KS","url":"https://www.wardhigh.org/page/clubs-and-activities","source_year":"Current page; year varies"},
    # Jacksonville MSA
    {"msa":"Jacksonville","school":"Edward H. White High School","state":"FL","url":"https://www.duvalschools.org/o/ewhs/page/activities-and-clubs/","source_year":"Current page; year varies"},
    {"msa":"Jacksonville","school":"Atlantic Coast High School","state":"FL","url":"https://www.duvalschools.org/o/achs/page/activities-and-clubs","source_year":"Current active organizations page; verify year"},
    {"msa":"Jacksonville","school":"Stanton College Preparatory School","state":"FL","url":"https://www.duvalschools.org/o/schs/page/student-activities","source_year":"Current page; official club directory"},
    {"msa":"Jacksonville","school":"Terry Parker High School","state":"FL","url":"https://dcps.duvalschools.org/o/tphs/page/activities-and-clubs/","source_year":"Current page; year varies"},
    # Birmingham MSA
    {"msa":"Birmingham","school":"Spain Park High School","state":"AL","url":"https://sites.google.com/hoover.k12.al.us/sphsclubs2020-2021","source_year":"2026-27"},
    {"msa":"Birmingham","school":"Hoover High School","state":"AL","url":"https://hhs.hoovercityschools.net/83809_2","source_year":"Current page; clubs and organizations"},
    {"msa":"Birmingham","school":"Woodlawn High School","state":"AL","url":"https://www.bhamcityschools.org/o/wohs/page/student-organizations","source_year":"Current page; limited organization list"},
    # Houston MSA
    {"msa":"Houston","school":"Heights High School","state":"TX","url":"https://heights.houstonisd.org/our-athletics-extracurriculars/extracurricular","source_year":"Current page; verify year"},
    {"msa":"Houston","school":"Lamar High School","state":"TX","url":"https://www.houstonisd.org/lamarhs","source_year":"School homepage; locate club page manually if parser finds no list"},
    # Phoenix MSA
    {"msa":"Phoenix","school":"Sunnyslope High School","state":"AZ","url":"https://sunnyslope.guhsdaz.org/clubs","source_year":"Current page; verify year"},
    {"msa":"Phoenix","school":"Glendale High School","state":"AZ","url":"https://glendale.guhsdaz.org/clubs","source_year":"Current page; verify year"},
    {"msa":"Phoenix","school":"Cortez High School","state":"AZ","url":"https://cortez.guhsdaz.org/clubs","source_year":"Current page; verify year"},
    {"msa":"Phoenix","school":"Washington High School","state":"AZ","url":"https://washington.guhsdaz.org/clubs","source_year":"Current page; verify year"},
    {"msa":"Phoenix","school":"Metro Tech High School","state":"AZ","url":"https://www.pxu.org/o/mths/page/clubs-groups","source_year":"2026-27"},
    # San Diego MSA
    {"msa":"San Diego","school":"San Diego High School","state":"CA","url":"https://sdhs.sandiegounified.org/caver_life/student_clubs","source_year":"2024-25"},
]

def clean_text(s):
    return re.sub(r"\s+", " ", str(s or "")).strip(" \t\r\n•·-–—")

def norm(s):
    return re.sub(r"[^a-z0-9]+", " ", clean_text(s).lower()).strip()

def page_to_candidates(html):
    soup = BeautifulSoup(html, "lxml")
    for tag in soup(["script","style","noscript","svg","nav","footer","header"]):
        tag.decompose()

    # Prefer list items under headings mentioning club/activity/organization.
    candidates = []
    headings = soup.find_all(re.compile(r"^h[1-6]$"))
    for h in headings:
        ht = clean_text(h.get_text(" ", strip=True)).lower()
        if not any(k in ht for k in ["club", "extracurricular", "student organization", "activities"]):
            continue
        node = h
        for _ in range(5):
            node = node.find_next()
            if node is None:
                break
            if node.name in ["h1","h2","h3","h4","h5","h6"] and node is not h:
                break
            if node.name == "li":
                t = clean_text(node.get_text(" ", strip=True))
                if 2 <= len(t) <= 140:
                    candidates.append(t)

    # Fallback: list items in page when no suitable heading/list structure found.
    if not candidates:
        for li in soup.find_all("li"):
            t = clean_text(li.get_text(" ", strip=True))
            if 2 <= len(t) <= 140:
                candidates.append(t)

    # Preserve first occurrence, remove exact duplicate strings.
    seen, unique = set(), []
    for item in candidates:
        key = norm(item)
        if key and key not in seen:
            seen.add(key)
            unique.append(item)
    return unique, soup.get_text(" ", strip=True)[:500]

def fetch(seed):
    try:
        r = requests.get(seed["url"], headers=HEADERS, timeout=TIMEOUT)
        r.raise_for_status()
        ctype = r.headers.get("Content-Type","").lower()
        if "pdf" in ctype or seed["url"].lower().endswith(".pdf"):
            return [], "PDF source: download and parse separately; not auto-counted", r.status_code
        names, excerpt = page_to_candidates(r.text)
        return names, ("Fetched; candidate extraction requires review" if names else "No list items extracted"), r.status_code
    except Exception as e:
        return [], f"FAILED: {type(e).__name__}: {e}", ""

def main():
    raw_rows, log_rows = [], []
    for i, seed in enumerate(SEEDS, 1):
        print(f"[{i}/{len(SEEDS)}] {seed['msa']} - {seed['school']}")
        names, note, status = fetch(seed)
        log_rows.append({
            "MSA":seed["msa"],"School":seed["school"],"State":seed["state"],
            "URL":seed["url"],"Source year":seed["source_year"],
            "HTTP status":status,"Extracted candidate count":len(names),"Status/note":note
        })
        for n, name in enumerate(names,1):
            raw_rows.append({
                "MSA":seed["msa"],"School":seed["school"],"State":seed["state"],
                "Candidate club name":name,"Within-page order":n,
                "Official source URL":seed["url"],"Source year":seed["source_year"],
                "Include in final count?":"REVIEW",
                "Exclusion reason":"","Reviewer notes":"Check whether this is a distinct club; exclude sports teams/class boards/student government if following protocol."
            })
        time.sleep(0.5)

    raw = pd.DataFrame(raw_rows)
    log = pd.DataFrame(log_rows)
    raw.to_csv(OUT/"R1_club_records_raw.csv",index=False,encoding="utf-8-sig")
    log.to_csv(OUT/"R1_collection_log.csv",index=False,encoding="utf-8-sig")

    # Review workbook with a summary that deliberately leaves final counts blank.
    if raw.empty:
        raw = pd.DataFrame(columns=["MSA","School","State","Candidate club name","Official source URL","Include in final count?","Reviewer notes"])
    summary = (raw.groupby(["MSA","School"],dropna=False)
               .size().reset_index(name="Extracted candidates (NOT final club count)"))
    with pd.ExcelWriter(OUT/"R1_club_records_review.xlsx",engine="openpyxl") as w:
        log.to_excel(w,sheet_name="采集日志",index=False)
        raw.to_excel(w,sheet_name="逐项审核",index=False)
        summary.to_excel(w,sheet_name="学校候选数量",index=False)
        pd.DataFrame([
            ["统计对象","学校官方网页公布的独立课外社团名称。"],
            ["自动提取","脚本提取社团/活动标题附近的HTML列表项；PDF不自动计数。"],
            ["必须人工审核","体育队、学生会、年级委员会、课程、活动场次、重复名称及非社团组织。"],
            ["缺失值","网页未找到名单、访问失败或PDF未解析均为空缺，不得记为0。"],
            ["代表性","当前SEEDS只是已发现来源的试采样，不是每个MSA随机抽取15校；需先用NCES学校总体建立样本框。"],
            ["建议","把更多官方学校社团目录URL加入脚本SEEDS后重跑，保留每条来源。"],
        ],columns=["项目","说明"]).to_excel(w,sheet_name="统计口径",index=False)

        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter
        for ws in w.book.worksheets:
            ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
            for cell in ws[1]:
                cell.font=Font(bold=True,color="FFFFFF")
                cell.fill=PatternFill("solid",fgColor="34495E")
                cell.alignment=Alignment(wrap_text=True)
            for col in ws.columns:
                letter=get_column_letter(col[0].column)
                width=min(max(max((len(str(c.value or "")) for c in list(col)[:80]),default=10)+2,12),55)
                ws.column_dimensions[letter].width=width
    print("\nFinished.")
    print("Review workbook:", (OUT/"R1_club_records_review.xlsx").resolve())
    print("Raw rows:", len(raw_rows))
    print("Important: extracted candidate counts are not final counts until reviewed.")

if __name__ == "__main__":
    main()
