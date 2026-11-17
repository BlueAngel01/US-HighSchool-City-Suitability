# -*- coding: utf-8 -*-
"""IMLS FY2024: match AE and Outlet records to the fixed 10 MSAs and build SQLite + Excel.

Input: IMLS_PLS_FY2024_1_CSV.zip (official FY2024 public-use data)
Outputs: SQLite database, Excel review workbook, and CSV audit summaries.

Important methodology:
- Match by the CBSA field attached by IMLS to each library/system/outlet's physical address.
- AE = administrative entity/library system; Outlet = individual service outlet record.
- Do not sum POPU_LSA and call it unique population. Legal service areas can overlap.
- POPU_UND is IMLS's state-level prorated unduplicated service-area population; still not identical
  to population residing inside the selected MSA.
- The MSA population denominator is the 2024 population used in the user's frozen city-selection file.
- Status codes 23 (temporary closure) and 25 (survey nonrespondent) are excluded from analytic
  active records; C_FSCS must be Y. All raw matched records are retained separately.
"""
from pathlib import Path
import zipfile, sqlite3, json, re, sys
import pandas as pd

BASE = Path(__file__).resolve().parent
ZIP_PATH = BASE / 'IMLS_PLS_FY2024_1_CSV.zip'
OUT_DIR = BASE / 'data' / 'processed'
DB_DIR = BASE / 'data' / 'database'
OUT_DIR.mkdir(parents=True, exist_ok=True)
DB_DIR.mkdir(parents=True, exist_ok=True)

# Frozen V3 population-weighted 10-city solution. Cleveland CBSA corrected against the
# FY2024 IMLS CBSA records: Cleveland, OH = 17460 (17410 produces no Cleveland records).
CITIES = [
    (1, '新英格兰 New England', '波士顿', 'Boston-Cambridge-Newton, MA-NH', '14460', 5025517),
    (2, '中大西洋 Middle Atlantic', '纽约', 'New York-Newark-Jersey City, NY-NJ', '35620', 19940274),
    (2, '中大西洋 Middle Atlantic', '费城', 'Philadelphia-Camden-Wilmington, PA-NJ-DE-MD', '37980', 6330422),
    (3, '东北中部 East North Central', '克利夫兰', 'Cleveland, OH', '17460', 2171877),
    (4, '西北中部 West North Central', '堪萨斯城', 'Kansas City, MO-KS', '28140', 2253579),
    (5, '南大西洋 South Atlantic', '杰克逊维尔', 'Jacksonville, FL', '27260', 1760548),
    (6, '东南中部 East South Central', '伯明翰', 'Birmingham, AL', '13820', 1192583),
    (7, '西南中部 West South Central', '休斯敦', 'Houston-Pasadena-The Woodlands, TX', '26420', 7796182),
    (8, '山地 Mountain', '凤凰城', 'Phoenix-Mesa-Chandler, AZ', '38060', 5186958),
    (9, '太平洋 Pacific', '圣迭戈', 'San Diego-Chula Vista-Carlsbad, CA', '41740', 3298799),
]
CITY_DF = pd.DataFrame(CITIES, columns=['地区编号','地区','城市中文名','MSA英文名称','CBSA','2024_MSA人口'])


def read_csv_from_zip(zf, suffix):
    names = [n for n in zf.namelist() if n.lower().endswith(suffix.lower())]
    if len(names) != 1:
        raise RuntimeError(f'在压缩包中应找到唯一文件 {suffix}，实际找到：{names}')
    # Official FY2024 CSV includes legacy Windows characters in names; cp1252 preserves them.
    with zf.open(names[0]) as f:
        df = pd.read_csv(f, dtype=str, encoding='cp1252', low_memory=False)
    df.columns = [c.strip().upper() for c in df.columns]
    return df, names[0]


def num(s):
    return pd.to_numeric(s, errors='coerce')


def valid_sum(df, col):
    if col not in df.columns:
        return None, 0, 0
    x = num(df[col])
    valid = x.notna() & (x >= 0)
    return float(x[valid].sum()), int(valid.sum()), int((~valid).sum())


def count_yes_no(df, col):
    if col not in df.columns:
        return None, None, None, None
    s = df[col].fillna('').str.strip().str.upper()
    yes, no = int((s == 'Y').sum()), int((s == 'N').sum())
    missing = int((~s.isin(['Y','N'])).sum())
    denom = yes + no
    return yes, no, missing, (yes / denom if denom else None)


def main():
    if not ZIP_PATH.exists():
        print('找不到输入文件：', ZIP_PATH)
        print('请把 IMLS_PLS_FY2024_1_CSV.zip 放在本脚本同一目录。')
        sys.exit(2)
    with zipfile.ZipFile(ZIP_PATH) as zf:
        ae, ae_member = read_csv_from_zip(zf, 'PLS_FY24_AE_pud24i.csv')
        outlet, outlet_member = read_csv_from_zip(zf, 'pls_fy24_outlet_pud24i.csv')

    targets = set(CITY_DF['CBSA'])
    ae_raw = ae[ae['CBSA'].isin(targets)].copy()
    out_raw = outlet[outlet['CBSA'].isin(targets)].copy()

    # Exclude temporarily closed and nonrespondent records and records outside FSCS definition.
    ae_active = ae_raw[(ae_raw['C_FSCS'].str.upper() == 'Y') & (~ae_raw['STATSTRU'].isin(['23','25']))].copy()
    out_active = out_raw[(out_raw['C_FSCS'].str.upper() == 'Y') & (~out_raw['STATSTRU'].isin(['23','25']))].copy()

    # Add city metadata, preserving CBSA as a 5-character text key.
    def attach(df):
        return df.merge(CITY_DF, on='CBSA', how='left', validate='many_to_one')
    ae_raw, ae_active = attach(ae_raw), attach(ae_active)
    out_raw, out_active = attach(out_raw), attach(out_active)

    summary_rows = []
    resource_cols = ['EBOOK_AE','ESERIAL_AE','EAUDIO_AE','EVIDEO_AE','RESDB_AE','OLP_AE']
    for c in CITIES:
        region_no, region, zh, msa, cbsa, pop = c
        a = ae_active[ae_active.CBSA == cbsa]
        ar = ae_raw[ae_raw.CBSA == cbsa]
        o = out_active[out_active.CBSA == cbsa]
        oraw = out_raw[out_raw.CBSA == cbsa]
        row = {
            '地区编号': region_no, '地区': region, '城市中文名': zh, 'MSA英文名称': msa,
            'CBSA': cbsa, '2024_MSA人口': pop,
            'AE原始匹配记录数': len(ar), 'AE有效系统数_C_FSCS_Y且非23_25': len(a),
            'AE排除记录数': len(ar)-len(a),
            'Outlet原始匹配记录数': len(oraw), 'Outlet有效网点记录数_C_FSCS_Y且非23_25': len(o),
            'Outlet排除记录数': len(oraw)-len(o),
        }
        for col, label in [('POPU_LSA','法定服务人口POPU_LSA_有效值求和'),('POPU_UND','未重复服务人口POPU_UND_有效值求和'),
                           ('TOTCIR','总流通量TOTCIR_有效值求和'),('PHYSCIR','实体流通量PHYSCIR_有效值求和'),
                           ('ELMATCIR','电子资料流通量ELMATCIR_有效值求和'),('EBOOK_CIR','电子书流通量EBOOK_CIR_有效值求和'),
                           ('CENTLIB','中央馆数量CENTLIB_有效值求和'),('BRANLIB','分馆数量BRANLIB_有效值求和'),
                           ('BKMOB','流动图书车网点数BKMOB_有效值求和')]:
            total, nvalid, nbad = valid_sum(a, col)
            row[label] = total
            row[label.replace('_有效值求和','_有效记录数')] = nvalid
            row[label.replace('_有效值求和','_缺失或特殊码记录数')] = nbad
        for col in resource_cols:
            yes,no,miss,rate = count_yes_no(a,col)
            row[f'{col}_Y系统数'] = yes
            row[f'{col}_N系统数'] = no
            row[f'{col}_M或缺失系统数'] = miss
            row[f'{col}_Y占有效回答比例'] = rate
        # Outlet type record counts and actual bookmobiles reported in BS rows.
        types = o['C_OUT_TY'].fillna('').str.upper()
        for typ, label in [('CE','中央馆网点记录数'),('BR','分馆网点记录数'),('BS','流动车记录数'),('BM','仅邮寄服务记录数')]:
            row[label] = int((types == typ).sum())
        if 'L_NUM_BM' in o.columns:
            x = num(o.loc[types == 'BS','L_NUM_BM'])
            row['流动图书车实际车辆数_L_NUM_BM有效值求和'] = float(x[x >= 0].sum())
        else:
            row['流动图书车实际车辆数_L_NUM_BM有效值求和'] = None
        for col,label in [('HOURS','网点年度开放服务小时数'),('SQ_FEET','网点面积平方英尺')]:
            total,nvalid,nbad=valid_sum(o,col)
            row[label+'_有效值求和']=total
            row[label+'_有效记录数']=nvalid
            row[label+'_缺失或特殊码记录数']=nbad
        # Population coverage ratios and per-capita service intensity.
        # POPU_LSA may overlap; POPU_UND is IMLS state-level prorated, not exact MSA-only population.
        lsa_total,_,_=valid_sum(a,'POPU_LSA')
        und_total,_,_=valid_sum(a,'POPU_UND')
        row['POPU_LSA报告服务人口占MSA人口比例_参考值'] = (lsa_total/pop) if lsa_total is not None and pop else None
        row['POPU_UND未重复服务人口占MSA人口比例_参考值'] = (und_total/pop) if und_total is not None and pop else None
        row['每10万MSA居民有效图书馆系统数'] = (len(a)/pop*100000) if pop else None
        row['每10万MSA居民有效服务网点数'] = (len(o)/pop*100000) if pop else None
        for raw_col, rate_label in [('TOTCIR','每千MSA居民总流通量'),('PHYSCIR','每千MSA居民实体流通量'),('ELMATCIR','每千MSA居民电子资料流通量')]:
            total,_,_=valid_sum(a,raw_col)
            row[rate_label] = (total/pop*1000) if total is not None and pop else None
        row['AE匹配口径'] = 'IMLS记录物理地址CBSA；非服务区域边界'
        row['人口口径说明'] = 'MSA人口作统一分母；POPU_LSA可能跨区且重叠，POPU_UND为州级分摊值，不等同MSA内唯一服务人口'
        summary_rows.append(row)

    summary = pd.DataFrame(summary_rows)
    # Matching audit: CBSA field counts before and after analytic filters.
    audit_rows=[]
    for c in CITIES:
        _,region,zh,msa,cbsa,pop=c
        ar=ae_raw[ae_raw.CBSA==cbsa]; a=ae_active[ae_active.CBSA==cbsa]
        rr=out_raw[out_raw.CBSA==cbsa]; o=out_active[out_active.CBSA==cbsa]
        audit_rows.append({'城市中文名':zh,'MSA英文名称':msa,'CBSA':cbsa,'AE原始数':len(ar),'AE有效数':len(a),
                           'Outlet原始数':len(rr),'Outlet有效数':len(o),
                           'AE_C_FSCS非Y':int((ar.C_FSCS.str.upper()!='Y').sum()),
                           'AE_状态23或25':int(ar.STATSTRU.isin(['23','25']).sum()),
                           'Outlet_C_FSCS非Y':int((rr.C_FSCS.str.upper()!='Y').sum()),
                           'Outlet_状态23或25':int(rr.STATSTRU.isin(['23','25']).sum())})
    audit=pd.DataFrame(audit_rows)

    # Write SQLite, preserving both filtered and unfiltered CBSA-matched records.
    db_path=DB_DIR/'US_10_MSA_IMLS_FY2024_E5.sqlite'
    if db_path.exists(): db_path.unlink()
    with sqlite3.connect(db_path) as con:
        CITY_DF.to_sql('target_msa',con,index=False,if_exists='replace')
        ae_raw.to_sql('ae_matched_raw',con,index=False,if_exists='replace')
        ae_active.to_sql('ae_analytic_active',con,index=False,if_exists='replace')
        out_raw.to_sql('outlet_matched_raw',con,index=False,if_exists='replace')
        out_active.to_sql('outlet_analytic_active',con,index=False,if_exists='replace')
        summary.to_sql('msa_e5_summary',con,index=False,if_exists='replace')
        audit.to_sql('matching_audit',con,index=False,if_exists='replace')
        meta=pd.DataFrame([
            {'key':'source','value':'IMLS Public Libraries Survey FY2024 Public Use Data (PLS_FY24_AE_pud24i.csv; pls_fy24_outlet_pud24i.csv)'},
            {'key':'source_zip','value':ZIP_PATH.name},
            {'key':'ae_source_member','value':ae_member}, {'key':'outlet_source_member','value':outlet_member},
            {'key':'matching_key','value':'CBSA field (geocoded physical address of AE/outlet)'},
            {'key':'analytic_filter','value':'C_FSCS=Y and STATSTRU not in (23,25)'},
            {'key':'population_denominator','value':'2024 MSA population in fixed V3 population-weighted 10-city selection file'},
            {'key':'critical_caveat','value':'POPU_LSA may overlap; POPU_UND is state-level prorated estimate; neither is identical to unique population inside MSA.'},
            {'key':'CBSA_correction','value':'Cleveland, OH CBSA corrected to 17460 after 17410 returned zero FY2024 IMLS records and 17460 matched Cleveland-area records.'},
        ])
        meta.to_sql('database_metadata',con,index=False,if_exists='replace')
        for table in ['ae_matched_raw','ae_analytic_active','outlet_matched_raw','outlet_analytic_active']:
            con.execute(f'CREATE INDEX idx_{table}_cbsa ON {table}(CBSA)')

    # Excel review workbook: Chinese-readable sheets; raw data sheets retained for audit.
    xlsx_path=OUT_DIR/'10座城市_IMLS_FY2024_E5_匹配结果.xlsx'
    with pd.ExcelWriter(xlsx_path,engine='openpyxl') as w:
        CITY_DF.to_excel(w,index=False,sheet_name='固定10城与CBSA')
        summary.to_excel(w,index=False,sheet_name='城市汇总')
        audit.to_excel(w,index=False,sheet_name='匹配审计')
        ae_active.to_excel(w,index=False,sheet_name='AE有效系统明细')
        out_active.to_excel(w,index=False,sheet_name='Outlet有效网点明细')
        ae_raw.to_excel(w,index=False,sheet_name='AE原始匹配明细')
        out_raw.to_excel(w,index=False,sheet_name='Outlet原始匹配明细')
        notes=pd.DataFrame({'项目':['数据来源','匹配方法','AE含义','Outlet含义','有效记录筛选','人口统计注意事项','克利夫兰CBSA修正','重要限制'],
            '说明':['IMLS Public Libraries Survey FY2024 public-use files。','按IMLS数据中的CBSA字段匹配图书馆系统/网点的物理地址，不按法律服务区边界匹配。','Administrative Entity，图书馆行政系统；一个系统可有多个服务网点。','具体服务点；类型CE中央馆、BR分馆、BS流动图书车记录、BM仅邮寄服务。','C_FSCS=Y，且STATSTRU不为23（临时关闭）或25（未回应）。原始匹配数据也完整保留。','POPU_LSA可能重叠；POPU_UND为州级分摊去重估计，不等于MSA内唯一服务人口。MSA人口用于统一分母。','人口加权选城CSV曾列17410；FY2024 IMLS中该CBSA没有记录，Cleveland, OH实际匹配CBSA为17460。','这些统计代表物理地址落在MSA内的系统和网点；不保证其服务区完全位于该MSA，也不保证覆盖MSA全部居民。']})
        notes.to_excel(w,index=False,sheet_name='口径与注意事项')
        for ws in w.book.worksheets:
            ws.freeze_panes='A2'
            ws.auto_filter.ref=ws.dimensions
            for col in ws.columns:
                max_len=max((len(str(cell.value)) if cell.value is not None else 0) for cell in list(col)[:300])
                ws.column_dimensions[col[0].column_letter].width=min(max(max_len+2,10),42)

    # CSV summaries to make the process reproducible and GitHub-friendly.
    summary.to_csv(OUT_DIR/'10座城市_IMLS_FY2024_E5_城市汇总.csv',index=False,encoding='utf-8-sig')
    audit.to_csv(OUT_DIR/'10座城市_IMLS_FY2024_E5_匹配审计.csv',index=False,encoding='utf-8-sig')

    print('完成：FY2024 IMLS原始数据匹配与建库')
    print(f'AE原始匹配记录：{len(ae_raw)}；有效系统记录：{len(ae_active)}')
    print(f'Outlet原始匹配记录：{len(out_raw)}；有效网点记录：{len(out_active)}')
    print('\n城市匹配审计：')
    print(audit.to_string(index=False))
    print('\n输出文件：')
    print('SQLite:',db_path)
    print('Excel :',xlsx_path)
    print('CSV   :',OUT_DIR/'10座城市_IMLS_FY2024_E5_城市汇总.csv')
    print('\n重要：本程序完成原始记录匹配与统计，不将POPU_LSA简单相加解释为MSA唯一服务人口，也不直接生成最终E5评分。')

if __name__=='__main__':
    main()
