#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Fetch 2024 current-dollar MSA GDP estimates using BEA's official geo aggregator.

Outputs:
  山地地区_2024GDP.csv  (CBSA,GDP_2024_USD,GDP_PC_2024)
  山地地区_BEAGeoAggregator原始结果.xlsx

BEA API key required. This uses BEA's published county data + BEA's own
aggregation method; the resulting MSA GDP estimates are approximations and
may be missing for some MSAs, as documented by BEA.
"""
import sys, os, io, csv, zipfile, urllib.request, pathlib, subprocess, getpass, time

CITIES = [
    (10,'Phoenix-Mesa-Chandler, AZ','38060'),
    (19,'Denver-Aurora-Centennial, CO','19740'),
    (29,'Las Vegas-Henderson-North Las Vegas, NV','29820'),
    (46,'Salt Lake City-Murray, UT','41620'),
    (52,'Tucson, AZ','46060'),
    (62,'Albuquerque, NM','10740'),
    (74,'Boise City, ID','14260'),
    (79,'Colorado Springs, CO','17820'),
    (81,'Provo-Orem-Lehi, UT','39340'),
    (88,'Ogden, UT','36260'),
    (101,'Reno, NV','39900'),
    (149,'Fort Collins-Loveland, CO','22660'),
    (153,'Greeley, CO','24540'),
    (162,'Boulder, CO','14500'),
    (195,'Prescott Valley-Prescott, AZ','39150'),
]
REPO_ZIP='https://github.com/us-bea/bea_geo_aggregator/archive/refs/heads/main.zip'

def install_and_import(work):
    try:
        import pandas, numpy, tqdm, beaapi  # noqa: F401
    except ImportError:
        subprocess.check_call([sys.executable,'-m','pip','install','pandas','numpy','tqdm','beaapi'])
    try:
        import bea_geo_aggregator as rea
        return rea
    except ImportError:
        zpath=work/'bea_geo_aggregator_main.zip'
        if not zpath.exists():
            print('下载BEA官方地理汇总工具代码:', REPO_ZIP)
            req=urllib.request.Request(REPO_ZIP,headers={'User-Agent':'Mozilla/5.0'})
            with urllib.request.urlopen(req,timeout=120) as r, zpath.open('wb') as f:
                f.write(r.read())
        extract=work/'bea_geo_aggregator_repo'
        extract.mkdir(exist_ok=True)
        with zipfile.ZipFile(zpath) as z: z.extractall(extract)
        roots=list(extract.glob('bea_geo_aggregator-*'))
        if not roots: raise RuntimeError('无法识别BEA官方工具压缩包目录')
        sys.path.insert(0,str(roots[0]))
        import bea_geo_aggregator as rea
        return rea

def extract_2024_gdp(gdp_dfs, cbsa):
    """Find CAGDP2 all-industry total, year 2024; DataValue is $thousands."""
    if not isinstance(gdp_dfs,dict):
        raise TypeError('BEA工具返回的GDP结果不是预期字典')
    df=gdp_dfs.get('CAGDP2')
    if df is None:
        raise KeyError('BEA结果中没有CAGDP2（现价GDP）表；实际表名：'+', '.join(map(str,gdp_dfs.keys())))
    rows=df.copy()
    # Normalize column lookup without changing source values.
    cols={str(c).strip().lower():c for c in rows.columns}
    required=['geofips','timeperiod','linecode','datavalue']
    if not all(k in cols for k in required):
        raise KeyError('CAGDP2字段不完整，实际字段：'+', '.join(map(str,rows.columns)))
    gf=rows[cols['geofips']].astype(str).str.replace(r'\.0$','',regex=True).str.zfill(5)
    yr=rows[cols['timeperiod']].astype(str)
    lc=rows[cols['linecode']].astype(str)
    desc=rows[cols.get('description',cols['linecode'])].astype(str).str.lower() if 'description' in cols else None
    mask=(gf==str(cbsa).zfill(5))&(yr=='2024')&(lc=='1')
    if desc is not None: mask=mask & desc.str.contains('all industry total',na=False)
    found=rows.loc[mask]
    if len(found)!=1:
        raise ValueError(f'CBSA {cbsa}: CAGDP2 2024 All industry total匹配{len(found)}行，拒绝猜测')
    val=float(str(found.iloc[0][cols['datavalue']]).replace(',',''))
    if val<=0: raise ValueError(f'CBSA {cbsa}: GDP数值无效 {val}')
    # BEA CAGDP2 unit: thousands of dollars (UNIT_MULT=3).
    unit=str(found.iloc[0][cols['cl_unit']]).lower() if 'cl_unit' in cols else ''
    mult=str(found.iloc[0][cols['unit_mult']]) if 'unit_mult' in cols else '3'
    if 'thousand' not in unit and mult not in ('3','3.0'):
        raise ValueError(f'CBSA {cbsa}: 未能确认BEA GDP单位，CL_UNIT={unit}, UNIT_MULT={mult}')
    return val*1000.0, found.iloc[0].to_dict()

def main():
    out=pathlib.Path.cwd()
    work=out/'mountain_bea_gdp_raw'; work.mkdir(exist_ok=True)
    api_key=os.environ.get('BEA_API_KEY','').strip()
    if not api_key:
        api_key=getpass.getpass('请输入BEA API Key（输入时不显示）：').strip()
    if not api_key:
        raise SystemExit('未输入BEA API Key，程序停止。可在BEA官网免费申请API Key。')
    rea=install_and_import(work)
    print('BEA官方地理汇总工具版本:',getattr(rea,'__version__','unknown'))
    results=[]; raw_frames=[]; failures=[]
    for rank,name,cbsa in CITIES:
        print(f'查询 {name} | CBSA {cbsa} ...')
        try:
            # BEA's official method estimates MSA aggregates from county GDP
            # and national industry price indexes.
            gdp_dfs=rea.pull_merge_calc_gdp(start_year=2024,last_year=2024,
                geo_agg='MSA',bea_key=api_key,geofips_filter=cbsa,verbosity=0)
            total,raw=extract_2024_gdp(gdp_dfs,cbsa)
            raw_frames.append((cbsa,gdp_dfs))
            results.append([cbsa,name,round(total,2),'BEA CAGDP2 2024 current-dollar MSA aggregate estimate'])
            print(f'  OK GDP=${total:,.0f}')
        except Exception as e:
            failures.append((cbsa,name,str(e)))
            print('  未取得可用值:',e)
        time.sleep(1.0)
    csv_path=out/'山地地区_2024GDP.csv'
    with csv_path.open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f);w.writerow(['CBSA','GDP_2024_USD','GDP_PC_2024','数据来源'])
        for cbsa,name,total,source in results:
            # Per-capita value is calculated by the existing mountain script
            # using its Census 2024 population series; leave it blank here.
            w.writerow([cbsa,f'{total:.2f}','','BEA Geo Aggregator：基于2024县级GDP推算的MSA现价GDP估计'])
    try:
        import pandas as pd
        with pd.ExcelWriter(out/'山地地区_BEAGeoAggregator原始结果.xlsx') as xw:
            for cbsa,dfs in raw_frames:
                for tab,df in dfs.items():
                    df.to_excel(xw,sheet_name=(cbsa+'_'+tab)[:31],index=False)
    except Exception as e:
        print('提示：原始结果Excel未生成（CSV仍已生成）：',e)
    print('\n已生成:',csv_path)
    print(f'成功取得GDP：{len(results)}/15')
    if failures:
        print('以下MSA未能取得完整BEA聚合估计（保留空值，不用县GDP简单相加替代）：')
        for cbsa,name,err in failures: print(f' - {name} [{cbsa}]: {err}')
    print('\n下一步运行升级版主程序：')
    print('py ".\\山地地区_15MSA_官方数据抓取_升级版_修正版.py" --gdp-csv ".\\山地地区_2024GDP.csv"')

if __name__=='__main__':
    main()
