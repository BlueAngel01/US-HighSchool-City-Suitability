# -*- coding: utf-8 -*-
"""国家级特大都市补充规则 V1.0（独立于五指标代表性模型）
固定：2024 MSA人口>=10,000,000；仅在原始31个MSA中筛选；已被主模型选中的不重复。
默认主模型31城整体结果：Salem CBSA 41420、San Diego CBSA 41740。
运行：py "太平洋地区_特大城市客观补充筛选.py"
若主模型结果改变：py "太平洋地区_特大城市客观补充筛选.py" --selected 41420,41740
"""
import csv, argparse
from pathlib import Path

SOURCE='太平洋地区_2024五指标_完整.csv'
OUTPUT='太平洋地区_特大城市补充筛选_V1.0.xlsx'
THRESHOLD=10_000_000
EXPECTED=31

def main():
    ap=argparse.ArgumentParser(description='固定人口门槛的国家级特大都市补充筛选')
    ap.add_argument('--folder',type=Path,default=Path.cwd())
    ap.add_argument('--selected',default='41420,41740',help='主模型已入选CBSA，以逗号分隔；默认31城整体结果 Salem, San Diego')
    ap.add_argument('--threshold',type=int,default=THRESHOLD,help='敏感性检验可改阈值；正式主结果固定10000000')
    args=ap.parse_args()
    path=args.folder/SOURCE
    if not path.is_file(): raise SystemExit(f'缺少原始五指标CSV：{path}')
    with path.open(encoding='utf-8-sig',newline='') as f: rows=list(csv.DictReader(f))
    if len(rows)!=EXPECTED:raise SystemExit(f'原清单必须为31城，实际{len(rows)}')
    cities=[]; ids=set()
    for r in rows:
        cbsa=str(r['CBSA']).strip()
        if cbsa in ids:raise SystemExit('重复CBSA：'+cbsa)
        ids.add(cbsa)
        pop=int(float(str(r['①2024人口（人）']).replace(',','')))
        if pop<=0:raise SystemExit('人口无效：'+cbsa)
        cities.append({'no':int(r['序号']),'cbsa':cbsa,'name':r['MSA'],'pop':pop})
    if [c['no'] for c in cities]!=list(range(1,32)):raise SystemExit('原清单序号/顺序不符合31城固定名单')
    selected={s.strip() for s in args.selected.split(',') if s.strip()}
    if not selected or not selected.issubset(ids):raise SystemExit('已选CBSA为空或不属于原始31城：'+str(selected-ids))
    results=[]
    for c in cities:
        over=c['pop']>=args.threshold
        main=c['cbsa'] in selected
        status='主模型已入选，不重复' if main else ('补充入选' if over else '不满足人口门槛')
        results.append({**c,'over':over,'main':main,'status':status,'add':over and not main})
    extra=[r for r in results if r['add']]
    try:import xlsxwriter
    except ImportError:raise SystemExit('请先运行 py -m pip install XlsxWriter')
    out=args.folder/OUTPUT
    w=xlsxwriter.Workbook(str(out))
    h=w.add_format({'bold':True,'font_color':'white','bg_color':'#155A70','text_wrap':True,'border':1})
    num=w.add_format({'num_format':'#,##0'})
    sh=w.add_worksheet('31城逐项判定')
    sh.write_row(0,0,['原清单序号','MSA','CBSA','2024人口','达到1000万人门槛','主模型已入选','补充判定'],h)
    for i,c in enumerate(results,1):
        sh.write_row(i,0,[c['no'],c['name'],c['cbsa']]);sh.write_number(i,3,c['pop'],num)
        sh.write_row(i,4,['是' if c['over'] else '否','是' if c['main'] else '否',c['status']])
    sh.freeze_panes(1,3);sh.autofilter(0,0,len(results),6)
    sh.set_column('A:A',13);sh.set_column('B:B',49);sh.set_column('C:C',12);sh.set_column('D:D',18);sh.set_column('E:G',23)
    sm=w.add_worksheet('补充结果')
    sm.write_row(0,0,['规则','固定阈值','候选总数','主模型已选','符合人口门槛','新增补充数量'],h)
    sm.write_row(1,0,['人口达到门槛且未被主模型选中',args.threshold,len(cities),len(selected),sum(c['over'] for c in results),len(extra)])
    sm.write_row(3,0,['原清单序号','MSA','CBSA','2024人口','入选类型'],h)
    for i,c in enumerate(extra,4):
        sm.write_row(i,0,[c['no'],c['name'],c['cbsa']]);sm.write_number(i,3,c['pop'],num);sm.write(i,4,'特大都市客观补充')
    sm.set_column('A:A',34);sm.set_column('B:B',52);sm.set_column('C:D',20);sm.set_column('E:F',22)
    info=w.add_worksheet('论文方法与来源')
    lines=[('规则编号','国家级特大都市补充筛选 V1.0'),('研究对象','仅原清单31个太平洋MSA；不新增候选MSA'),('数据指标','2024 MSA常住人口（沿用五指标基础数据第①项，不新增第六指标）'),('门槛','2024人口>=10,000,000；研究预设阈值，不是官方“特大城市”定义'),('去重','已由主五指标模型入选的城市不重复纳入'),('结果性质','独立补充名单，不修改五指标20%等权、d矩阵、第一/第二城与20%改善门槛'),('主模型输入','--selected 后填写当前同一候选范围的主模型CBSA；默认41420,41740仅适用31城整体结果'),('分区限制','31城整体结果不等于南北分区正式结论；分区主模型变更后必须同步更新--selected'),('来源','沿用此前R02 Census 2024 MSA人口来源；不新增参考文献'),('敏感性','可使用--threshold 进行800万/1200万等测试，但正式阈值保持1000万且应明确区分敏感性结果')]
    info.write_row(0,0,['项目','说明'],h)
    for i,(a,b) in enumerate(lines,1):info.write_row(i,0,[a,b]);info.set_row(i,30)
    info.set_column('A:A',24);info.set_column('B:B',115)
    w.close()
    print(f'完成：{out.resolve()}')
    print(f'固定清单{len(cities)}城；人口门槛{args.threshold:,}；主模型已选{len(selected)}城；补充{len(extra)}城')
    for c in extra:print(f"  + {c['name']} | CBSA {c['cbsa']} | 人口 {c['pop']:,}")
    if args.threshold!=THRESHOLD:print('注意：当前是阈值敏感性检验，不是正式1000万主结果')

if __name__=='__main__':main()
