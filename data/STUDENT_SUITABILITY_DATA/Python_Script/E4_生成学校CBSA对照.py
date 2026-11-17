# -*- coding: utf-8 -*-
"""从 NCES EDGE_GEOCODE_PUBLICSCH_2324.zip 提取学校NCESSCH到CBSA的官方对照。"""
import csv
import zipfile
from pathlib import Path

root = Path.home() / 'Downloads'
source = root / 'EDGE_GEOCODE_PUBLICSCH_2324.zip'
out = root / 'school_cbsa_crosswalk.csv'
audit = root / 'school_cbsa_crosswalk_audit.csv'
if not source.exists():
    raise FileNotFoundError(source)

count = 0
missing = 0
conflicts = 0
seen = {}
with zipfile.ZipFile(source) as z:
    matches = [n for n in z.namelist() if n.lower().endswith('.txt')]
    if len(matches) != 1:
        raise ValueError(f'预期一个TXT，实际找到: {matches}')
    import io
    with z.open(matches[0]) as raw_file:
        with io.TextIOWrapper(raw_file, encoding='utf-8-sig', errors='replace', newline='') as f:
            reader = csv.reader(f, delimiter='|', quotechar='"')
            for parts in reader:
                count += 1
                if len(parts) != 23:
                    raise ValueError(f'第{count}条记录有{len(parts)}列，预期23列；请检查原始文件')
                ncessch, county, cbsa, school, year = (
                    parts[0].strip(), parts[9].strip(), parts[14].strip(),
                    parts[2].strip(), parts[22].strip()
                )
                if not (ncessch.isdigit() and len(ncessch) == 12):
                    raise ValueError(f'第{count}条学校编号异常: {ncessch!r}')
                if year != '2023-2024':
                    raise ValueError(f'第{count}条学年异常: {year!r}')
                if not (cbsa.isdigit() and len(cbsa) == 5):
                    missing += 1
                    continue
                item = (cbsa, county, school)
                if ncessch in seen and seen[ncessch][0] != cbsa:
                    conflicts += 1
                    raise ValueError(f'学校{ncessch}存在不同CBSA: {seen[ncessch][0]} vs {cbsa}')
                seen[ncessch] = item

with out.open('w', newline='', encoding='utf-8-sig') as f:
    w = csv.writer(f)
    w.writerow(['NCESSCH', 'CBSA'])
    for school_id, (cbsa, _, _) in sorted(seen.items()):
        w.writerow([school_id, cbsa])
with audit.open('w', newline='', encoding='utf-8-sig') as f:
    w = csv.writer(f)
    w.writerow(['NCESSCH', 'CBSA', 'COUNTY_FIPS', 'SCHOOL_NAME'])
    for school_id, (cbsa, county, school) in sorted(seen.items()):
        w.writerow([school_id, cbsa, county, school])
print('EDGE记录总数:', count)
print('有效CBSA学校:', len(seen))
print('缺失或非5位CBSA记录:', missing)
print('冲突学校:', conflicts)
print('已生成:', out)
print('审计文件:', audit)
