import os
import re
import csv
import json
import sqlite3
import zipfile
from pathlib import Path
from datetime import datetime

ROOT = Path(r"C:\Users\tsing\Desktop\DATA")
OUT = ROOT / "CH5_AUDIT"
OUT.mkdir(parents=True, exist_ok=True)

try:
    import openpyxl
except ImportError:
    openpyxl = None

EXTENSIONS = {
    ".xlsx", ".xls", ".xlsm", ".csv", ".tsv",
    ".sqlite", ".sqlite3", ".db", ".zip",
    ".py", ".json", ".geojson", ".pdf",
    ".txt", ".md"
}

INDICATORS = [
    "E1", "E2", "E3", "E4", "E5",
    "H1", "H2", "H3",
    "S1", "S2", "S3",
    "L1", "L2", "L3",
    "R1", "R2"
]

KEYWORDS = {
    "E1": ["graduation", "毕业率", "acgr"],
    "E2": ["teacher", "师生比", "fte"],
    "E3": ["ap_ib", "ap-ib", "ap课程", "ib课程"],
    "E4": ["course", "课程", "advanced"],
    "E5": ["library", "libraries", "图书馆", "pls_fy"],
    "H1": ["aqi", "air_quality", "空气质量"],
    "H2": ["park", "parkscore", "公园", "绿地"],
    "H3": ["ahrf", "physician", "医疗", "医生"],
    "S1": ["violent", "暴力犯罪"],
    "S2": ["property_crime", "财产犯罪"],
    "S3": ["fars", "fatalit", "交通死亡"],
    "L1": ["rpp", "price_parit", "生活成本"],
    "L2": ["fare", "票价", "公交费用"],
    "L3": ["gtfs", "transit_stop", "站点密度"],
    "R1": ["cultural", "文化设施"],
    "R2": ["sports", "recreation", "体育设施"]
}

def classify(path):
    text = str(path.relative_to(ROOT)).lower()
    found = []

    for code in INDICATORS:
        if re.search(r"(?<![a-z0-9])" + code.lower() +
                     r"(?![a-z0-9])", text):
            found.append(code)

    if found:
        return ",".join(found)

    for code, words in KEYWORDS.items():
        if any(word.lower() in text for word in words):
            found.append(code)

    return ",".join(found) if found else "UNCLASSIFIED"

def role(path):
    name = path.name.lower()
    if path.suffix.lower() == ".py":
        return "SCRIPT"
    if any(k in name for k in
           ["最终", "汇总", "补全", "结果", "final", "summary"]):
        return "RESULT_CANDIDATE"
    if any(k in name for k in
           ["raw", "原始", "source", "download"]):
        return "RAW_CANDIDATE"
    return "REVIEW"

def inspect_csv(path):
    for enc in ("utf-8-sig", "utf-8", "gb18030", "cp1252"):
        try:
            with path.open("r", encoding=enc, newline="") as f:
                reader = csv.reader(f)
                header = next(reader, [])
                sample = []
                for _, row in zip(range(3), reader):
                    sample.append(row[:8])
            return {
                "encoding": enc,
                "columns": header[:80],
                "sample_rows": sample
            }
        except (UnicodeError, UnicodeDecodeError):
            continue
        except Exception as e:
            return {"error": str(e)}
    return {"error": "Encoding not identified"}

def inspect_sqlite(path):
    result = []
    con = sqlite3.connect(
        "file:" + path.as_posix() + "?mode=ro",
        uri=True
    )
    try:
        tables = con.execute(
            "SELECT name FROM sqlite_master "
            "WHERE type='table' AND name NOT LIKE 'sqlite_%'"
        ).fetchall()

        for (table,) in tables[:60]:
            quoted = '"' + table.replace('"', '""') + '"'
            cols = con.execute(
                "PRAGMA table_info(" + quoted + ")"
            ).fetchall()
            count = con.execute(
                "SELECT COUNT(*) FROM " + quoted
            ).fetchone()[0]
            result.append({
                "table": table,
                "rows": count,
                "columns": [c[1] for c in cols]
            })
    finally:
        con.close()
    return result

def inspect_excel(path):
    if openpyxl is None:
        return {"error": "openpyxl not installed"}
    wb = openpyxl.load_workbook(
        path, read_only=True, data_only=True
    )
    result = []
    try:
        for ws in wb.worksheets[:40]:
            first = next(
                ws.iter_rows(min_row=1, max_row=1,
                             values_only=True), ()
            )
            result.append({
                "sheet": ws.title,
                "rows": ws.max_row,
                "columns": [str(x) if x is not None else ""
                            for x in first[:80]]
            })
    finally:
        wb.close()
    return result

def inspect_zip(path):
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        return {
            "entries": len(names),
            "sample_names": names[:20]
        }

inventory = []
structures = []
errors = []

files = sorted(
    (p for p in ROOT.rglob("*")
     if p.is_file()
     and OUT not in p.parents
     and p.suffix.lower() in EXTENSIONS
     and not p.name.startswith("~$")),
    key=lambda p: str(p).lower()
)

for index, path in enumerate(files, 1):
    code = classify(path)
    kind = role(path)
    rel = str(path.relative_to(ROOT))
    stat = path.stat()

    inventory.append({
        "ID": index,
        "Indicator": code,
        "Role": kind,
        "Extension": path.suffix.lower(),
        "Size_MB": round(stat.st_size / 1048576, 3),
        "Modified": datetime.fromtimestamp(
            stat.st_mtime).isoformat(timespec="seconds"),
        "Relative_Path": rel,
        "Full_Path": str(path)
    })

    try:
        suffix = path.suffix.lower()
        details = None

        if suffix in (".csv", ".tsv"):
            details = inspect_csv(path)
        elif suffix in (".sqlite", ".sqlite3", ".db"):
            details = inspect_sqlite(path)
        elif suffix in (".xlsx", ".xlsm"):
            details = inspect_excel(path)
        elif suffix == ".zip":
            details = inspect_zip(path)

        if details is not None:
            structures.append({
                "ID": index,
                "Indicator": code,
                "File": rel,
                "Structure": details
            })

    except Exception as e:
        errors.append({
            "ID": index,
            "File": rel,
            "Error": str(e)
        })

    if index % 50 == 0:
        print(f"Scanned {index}/{len(files)} files")

with (OUT / "CH5_FILE_INVENTORY.csv").open(
    "w", encoding="utf-8-sig", newline=""
) as f:
    writer = csv.DictWriter(
        f, fieldnames=list(inventory[0].keys())
        if inventory else ["ID", "Indicator", "Role",
                           "Extension", "Size_MB", "Modified",
                           "Relative_Path", "Full_Path"]
    )
    writer.writeheader()
    writer.writerows(inventory)

with (OUT / "CH5_FILE_STRUCTURES.json").open(
    "w", encoding="utf-8"
) as f:
    json.dump(structures, f, ensure_ascii=False, indent=2)

with (OUT / "CH5_SCAN_ERRORS.json").open(
    "w", encoding="utf-8"
) as f:
    json.dump(errors, f, ensure_ascii=False, indent=2)

print("\nSCAN COMPLETE")
print("Files:", len(inventory))
print("Structured files:", len(structures))
print("Errors:", len(errors))
print("\nIndicator coverage:")
for code in INDICATORS:
    count = sum(code in row["Indicator"].split(",")
                for row in inventory)
    print(f"  {code}: {count} candidate files")

print("\nReports saved in:", OUT)
