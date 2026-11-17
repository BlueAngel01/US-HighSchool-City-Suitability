#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""山地地区15 MSA五指标数据收集（与太平洋地区统一口径）。
运行：py 山地地区_15MSA_官方数据抓取.py
可选：--supplement mountain_area_gdp.csv （列：CBSA,AREALAND,GDP_PC_2024）
不填充缺失的GDP或面积；在补齐前只输出原始数据，不宣称五指标完整。
"""
import argparse,csv,io,math,pathlib,urllib.request,zipfile,json,time

POP_URL='https://www2.census.gov/programs-surveys/popest/datasets/2020-2024/metro/totals/cbsa-est2024-alldata.csv'
ACS_URL='https://www2.census.gov/programs-surveys/acs/summary_file/2024/table-based-SF/data/1YRData/acsdt1y2024-c24030.dat'
# 序号严格按用户清单，排名仅作原始清单标识，不参与计算
CITIES=[
(10,'Phoenix-Mesa-Chandler, AZ',38060),
(19,'Denver-Aurora-Centennial, CO',19740),
(29,'Las Vegas-Henderson-North Las Vegas, NV',29820),
(46,'Salt Lake City-Murray, UT',41620),
(52,'Tucson, AZ',46060),
(62,'Albuquerque, NM',10740),
(74,'Boise City, ID',14260),
(79,'Colorado Springs, CO',17820),
(81,'Provo-Orem-Lehi, UT',39340),
(88,'Ogden, UT',36260),
(101,'Reno, NV',39900),
(149,'Fort Collins-Loveland, CO',22660),
(153,'Greeley, CO',24540),
(162,'Boulder, CO',14500),
(195,'Prescott Valley-Prescott, AZ',39150),
]
INDUSTRIES=[
('农业、林业、渔业与采矿',3,30),('建筑业',6,33),('制造业',7,34),
('批发贸易',8,35),('零售贸易',9,36),('运输仓储与公用事业',10,37),
('信息业',13,40),('金融、保险与房地产',14,41),
('专业服务、企业管理与行政服务',17,44),('教育、医疗与社会援助',21,48),
('文娱、住宿与餐饮',24,51),('其他服务',27,54),('公共行政',28,55)]
SQM_PER_SQMI=2589988.110336

def fetch(url,path):
    if path.exists() and path.stat().st_size>100000:
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
            if path.stat().st_size<100000:raise RuntimeError('下载文件异常小')
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

def main():
    p=argparse.ArgumentParser()
    p.add_argument('--output-dir',type=pathlib.Path,default=pathlib.Path.cwd())
    p.add_argument('--population-file',type=pathlib.Path)
    p.add_argument('--acs-file',type=pathlib.Path)
    p.add_argument('--supplement',type=pathlib.Path,help='含CBSA,AREALAND,GDP_PC_2024列的已核实补充数据CSV')
    a=p.parse_args()
    out=a.output_dir.resolve();out.mkdir(parents=True,exist_ok=True)
    raw=out/'mountain_official_raw';raw.mkdir(exist_ok=True)
    pop=read_table(a.population_file or fetch(POP_URL,raw/'cbsa-est2024-alldata.csv'))
    acs=read_table(a.acs_file or fetch(ACS_URL,raw/'acsdt1y2024-c24030.dat'))
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
            if geo.endswith('US'+str(code)) or geo==str(code):
                acss.setdefault(str(code),[]).append(row)
    extra=load_supplement(a.supplement)
    five=[];detail=[];missing=[];supplement_template=[]
    for rank,name,code in CITIES:
        key=str(code)
        if len(pops.get(key,[]))!=1:
            raise ValueError(f'{name} ({code}) 官方人口总行数量={len(pops.get(key,[]))}，拒绝猜测')
        if len(acss.get(key,[]))!=1:
            raise ValueError(f'{name} ({code}) ACS记录数量={len(acss.get(key,[]))}，拒绝猜测')
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
        area_raw=str(ext.get('AREALAND','')).strip()
        gdp_raw=str(ext.get('GDP_PC_2024','')).strip()
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
    assert len(five)==15 and len(detail)==195
    write_csv(out/'山地地区_2024五指标_15MSA.csv',
              ['原始清单序号','MSA','CBSA','2020人口','2024人口','陆地面积_平方米','2024人口密度_人每平方英里','2024人均GDP_美元','2020至2024_CAGR_百分比','2024就业总数'],five)
    write_csv(out/'山地地区_2024产业结构_13类_195行.csv',
              ['原始清单序号','MSA','CBSA','产业序号','产业类别','男性就业','女性就业','就业合计','就业占比_百分比','ACS男性变量','ACS女性变量'],detail)
    write_csv(out/'山地地区_面积GDP补充模板.csv',['CBSA','MSA','AREALAND','GDP_PC_2024'],supplement_template)
    print('\n已生成15城五指标CSV、195条行业CSV及面积/GDP补充模板。')
    if missing:
        print('注意：以下城市面积或2024人均GDP尚未提供可核实数据，保持空值：')
        for name in missing:print(' -',name)
        print('用官方来源核实面积与GDP后填入补充模板，再运行 --supplement 山地地区_面积GDP补充模板.csv')
    else:print('五指标已全部补齐，可进入统一模型计算。')

if __name__=='__main__':main()
