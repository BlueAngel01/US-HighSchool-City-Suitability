#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""东北中部地区24 MSA五指标数据收集（与太平洋、山地、西南中部、东南中部、南大西洋及东北中部地区统一口径）。
运行：py 东北中部_24MSA_官方数据抓取.py
可选：--supplement east_north_central_area_gdp.csv （列：CBSA,AREALAND,GDP_PC_2024）
不填充缺失的GDP或面积；在补齐前只输出原始数据，不宣称五指标完整。
"""
import argparse,csv,io,math,pathlib,urllib.request,zipfile,json,time

POP_URL='https://www2.census.gov/programs-surveys/popest/datasets/2020-2024/metro/totals/cbsa-est2024-alldata.csv'
ACS_URL='https://www2.census.gov/programs-surveys/acs/summary_file/2024/table-based-SF/data/1YRData/acsdt1y2024-c24030.dat'
GAZ_URL='https://www2.census.gov/geo/docs/maps-data/data/gazetteer/2024_Gazetteer/2024_Gaz_cbsa_national.zip'
# 序号严格按用户清单，排名仅作原始清单标识，不参与计算
CITIES=[
(3,'Chicago-Naperville-Elgin, IL-IN',16980),
(14,'Detroit-Warren-Dearborn, MI',19820),
(30,'Cincinnati, OH-KY-IN',17140),
(32,'Columbus, OH',18140),
(33,'Indianapolis-Carmel-Greenwood, IN',26900),
(34,'Cleveland, OH',17460),
(40,'Milwaukee-Waukesha, WI',33340),
(49,'Grand Rapids-Wyoming-Kentwood, MI',24340),
(76,'Dayton-Kettering-Beavercreek, OH',19430),
(85,'Madison, WI',31540),
(87,'Akron, OH',10420),
(98,'Toledo, OH',45780),
(115,'Lansing-East Lansing, MI',29620),
(118,'Fort Wayne, IN',23060),
(131,'Youngstown-Warren, OH',49660),
(138,'Flint, MI',22420),
(139,'Canton-Massillon, OH',15940),
(150,'Ann Arbor, MI',11460),
(156,'Peoria, IL',37900),
(160,'Rockford, IL',40420),
(161,'Green Bay, WI',24580),
(163,'South Bend-Mishawaka, IN-MI',43780),
(186,'Evansville, IN',21780),
(190,'Kalamazoo-Portage, MI',28020),
]
INDUSTRIES=[
('农业、林业、渔业与采矿',3,30),('建筑业',6,33),('制造业',7,34),
('批发贸易',8,35),('零售贸易',9,36),('运输仓储与公用事业',10,37),
('信息业',13,40),('金融、保险与房地产',14,41),
('专业服务、企业管理与行政服务',17,44),('教育、医疗与社会援助',21,48),
('文娱、住宿与餐饮',24,51),('其他服务',27,54),('公共行政',28,55)]
SQM_PER_SQMI=2589988.110336

def fetch(url,path):
    if path.exists() and path.stat().st_size>(1000 if path.suffix=='.zip' else 100000):
        print('使用缓存:',path);return path
    print('下载:',url)
    req=urllib.request.Request(url,headers={'User-Agent':'Mozilla/5.0 Academic research'})
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req,timeout=120) as r,path.open('wb') as f:
                while True:
                    b=r.read(1024*256)
                    if not b:break
                    f.write(b)
            if path.stat().st_size<(1000 if path.suffix=='.zip' else 100000):raise RuntimeError('下载文件异常小')
            return path
        except Exception:
            path.unlink(missing_ok=True)
            if attempt==2:raise
            time.sleep(2)

def read_table(path):
    raw=path.read_bytes()
    try:s=raw.decode('utf-8-sig')
    except UnicodeDecodeError:s=raw.decode('latin-1')
    delimiter='|' if path.suffix=='.dat' else ','
    return list(csv.DictReader(io.StringIO(s,newline=''),delimiter=delimiter))

def val(row,n):
    for key in (f'C24030_E{n:03d}',f'C24030_{n:03d}E',f'C24030_E{n}',f'C24030_{n}E'):
        if key in row:return int(row[key].strip().replace(',',''))
    raise KeyError(f'C24030 index {n} missing')

def write_csv(path,headers,rows):
    with path.open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f);w.writerow(headers);w.writerows(rows)

def load_supplement(path):
    if not path:return {}
    with path.open(encoding='utf-8-sig',newline='') as f:
        data=list(csv.DictReader(f))
    out={}
    for row in data:
        cbsa=str(row['CBSA']).strip()
        if cbsa in out:raise ValueError('补充表CBSA重复: '+cbsa)
        out[cbsa]=row
    return out

def load_gazetteer(zip_path):
    with zipfile.ZipFile(zip_path) as z:
        members=[n for n in z.namelist() if n.lower().endswith('.txt')]
        if len(members)!=1:raise ValueError('Gazetteer ZIP 文件结构异常: '+str(members))
        raw=z.read(members[0]).decode('utf-8-sig')
    reader=csv.DictReader(io.StringIO(raw),delimiter='\t')
    data={}
    for row in reader:
        row={k.strip():v.strip() for k,v in row.items() if k is not None}
        cbsa=row.get('GEOID','')
        if cbsa:
            if cbsa in data:raise ValueError('Gazetteer重复CBSA: '+cbsa)
            data[cbsa]=row
    return data

def load_gdp_csv(path):
    """用户下载的、已核实的2024 MSA名义GDP数据；支持直接GDP或人均GDP。"""
    if not path:return {}
    with path.open(encoding='utf-8-sig',newline='') as f:
        rows=list(csv.DictReader(f))
    out={}
    for row in rows:
        code=str(row.get('CBSA','')).strip()
        if not code:raise ValueError('GDP CSV必须有CBSA列')
        if code in out:raise ValueError('GDP CSV重复CBSA '+code)
        out[code]=row
    return out

def number_or_none(value):
    v=str(value if value is not None else '').strip().replace(',','')
    return float(v) if v else None

def main():
    p=argparse.ArgumentParser()
    p.add_argument('--output-dir',type=pathlib.Path,default=pathlib.Path.cwd())
    p.add_argument('--population-file',type=pathlib.Path)
    p.add_argument('--acs-file',type=pathlib.Path)
    p.add_argument('--supplement',type=pathlib.Path,help='东北中部已核实补充数据CSV，含CBSA,AREALAND,GDP_PC_2024列')
    p.add_argument('--gdp-csv',type=pathlib.Path,help='2024名义GDP核实CSV：CBSA,GDP_2024_USD（美元）或GDP_PC_2024（美元/人）')
    p.add_argument('--gazetteer-file',type=pathlib.Path,help='已下载的2024 Census CBSA Gazetteer ZIP')
    a=p.parse_args()
    out=a.output_dir.resolve();out.mkdir(parents=True,exist_ok=True)
    raw=out/'east_north_central_official_raw';raw.mkdir(exist_ok=True)
    pop=read_table(a.population_file or fetch(POP_URL,raw/'cbsa-est2024-alldata.csv'))
    acs=read_table(a.acs_file or fetch(ACS_URL,raw/'acsdt1y2024-c24030.dat'))
    gaz=load_gazetteer(a.gazetteer_file or fetch(GAZ_URL,raw/'2024_Gaz_cbsa_national.zip'))
    gdp_source=load_gdp_csv(a.gdp_csv)
    # 官方总行须排除县级、分区级重复行
    pops={}
    for row in pop:
        code=str(row.get('CBSA','')).strip()
        if code not in {str(c[2]) for c in CITIES}:continue
        if str(row.get('STCOU','')).strip() not in ('','0','00000'):continue
        if str(row.get('MDIV','')).strip() not in ('','0','00000'):continue
        pops.setdefault(code,[]).append(row)
    acss={}
    for row in acs:
        geo=str(row.get('GEO_ID') or row.get('GEOID') or '').strip()
        for _,_,code in CITIES:
            # 只接受 ACS 310 级 MSA/MSA geography 记录。
            # 不能只用 endswith('US'+CBSA)，否则会把其他地理层级的同尾号记录误匹配进来。
            expected_geo=f'310M700US{code}'
            if geo == expected_geo:
                acss.setdefault(str(code),[]).append(row)
    # 若未指定--supplement，自动读取已有模板，避免覆盖手动填入的GDP
    default_supp=out/'东北中部_面积GDP补充模板.csv'
    supp_path=a.supplement or (default_supp if default_supp.exists() else None)
    extra=load_supplement(supp_path)
    five=[];detail=[];missing=[];supplement_template=[]
    for rank,name,code in CITIES:
        key=str(code)
        if len(pops.get(key,[]))!=1:
            raise ValueError(f'{name} ({code}) 官方人口总行数量={len(pops.get(key,[]))}，拒绝猜测')
        if len(acss.get(key,[]))!=1:
            matched_ids=[str(r.get('GEO_ID') or r.get('GEOID') or '') for r in acss.get(key,[])]
            raise ValueError(f'{name} ({code}) 精确MSA GEO_ID={"310M700US"+key}，匹配记录数量={len(acss.get(key,[]))}；实际匹配={matched_ids}，拒绝猜测')
        pr=pops[key][0]; ar=acss[key][0]
        p20=int(pr['POPESTIMATE2020']);p24=int(pr['POPESTIMATE2024'])
        if min(p20,p24)<=0:raise ValueError(name+'人口异常')
        employed=val(ar,1)
        counts=[]
        for j,(sector,m,f) in enumerate(INDUSTRIES,1):
            male,female=val(ar,m),val(ar,f)
            if male<0 or female<0:raise ValueError(name+'就业人数异常')
            total=male+female;counts.append(total)
            detail.append([rank,name,code,j,sector,male,female,total,round(total/employed*100,6),f'C24030_{m:03d}E',f'C24030_{f:03d}E'])
        if sum(counts)!=employed:
            raise ValueError(f'{name}: 13类就业总数{sum(counts)} != C24030总数{employed}')
        ext=extra.get(key,{})
        gaz_row=gaz.get(key)
        if not gaz_row:raise ValueError(f'{name} CBSA {key} 未在2024 Census Gazetteer找到')
        official_area=str(gaz_row.get('ALAND','')).strip()
        if not official_area:raise ValueError(name+' Gazetteer陆地面积缺失')
        area_raw=str(ext.get('AREALAND','')).strip() or official_area
        # 始终核对2024 Census Gazetteer面积；拒绝无说明地混用其他年份边界
        if abs(float(area_raw)-float(official_area))>1:
            print('注意：补充表面积与2024 Census Gazetteer不同，已采用2024官方面积:',name)
            area_raw=official_area
        gdp_raw=str(ext.get('GDP_PC_2024','')).strip()
        gdp_row=gdp_source.get(key,{})
        if gdp_row:
            direct=number_or_none(gdp_row.get('GDP_PC_2024'))
            total=number_or_none(gdp_row.get('GDP_2024_USD'))
            if direct is None and total is None:
                raise ValueError(name+' GDP CSV中GDP_2024_USD和GDP_PC_2024均为空')
            if direct is not None and total is not None:
                if abs(direct-total/p24)>max(1,0.001*direct):
                    raise ValueError(name+' GDP总额与人均GDP不一致')
            gdp_raw=str(direct if direct is not None else round(total/p24,6))
        area=float(area_raw) if area_raw else None
        gdp=float(gdp_raw) if gdp_raw else None
        if area is not None and area<=0:raise ValueError(name+'陆地面积无效')
        if gdp is not None and gdp<=0:raise ValueError(name+'GDP无效')
        density=round(p24*SQM_PER_SQMI/area,4) if area else ''
        cagr=round(100*((p24/p20)**.25-1),6)
        five.append([rank,name,code,p20,p24,area or '',density,gdp or '',cagr,employed])
        supplement_template.append([code,name,area_raw,gdp_raw])
        if not area or not gdp:missing.append(name)
        print(f'OK {name:<45} 2024人口={p24:,} CAGR={cagr:.4f}% 13行业={employed:,}')
    assert len(five)==24 and len(detail)==312
    write_csv(out/'东北中部_2024五指标_24MSA.csv',
              ['原始清单序号','MSA','CBSA','2020人口','2024人口','陆地面积_平方米','2024人口密度_人每平方英里','2024人均GDP_美元','2020至2024_CAGR_百分比','2024就业总数'],five)
    write_csv(out/'东北中部_2024产业结构_13类_312行.csv',
              ['原始清单序号','MSA','CBSA','产业序号','产业类别','男性就业','女性就业','就业合计','就业占比_百分比','ACS男性变量','ACS女性变量'],detail)
    write_csv(out/'东北中部_面积GDP补充模板.csv',['CBSA','MSA','AREALAND','GDP_PC_2024'],supplement_template)
    print('\n已生成24城五指标CSV、312条行业CSV及面积/GDP补充模板。')
    print('面积来源：2024 U.S. Census Bureau CBSA Gazetteer，24城已自动核实。')
    if missing:
        print('注意：以下城市面积或2024人均GDP尚未提供可核实数据，保持空值：')
        for name in missing:print(' -',name)
        print('GDP缺失：BEA于2026年停止2024年MSA GDP直接发布；需按MSA县域汇总BEA 2024县GDP，或导入已核实的MSA名义GDP。')
        print('可在模板GDP_PC_2024列填值，或使用 --gdp-csv 文件（CBSA,GDP_2024_USD）。')
    else:print('五指标已全部补齐，可进入统一模型计算。')

if __name__=='__main__':main()
