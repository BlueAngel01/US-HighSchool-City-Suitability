# -*- coding: utf-8 -*-
"""从上一程序生成的403行CSV构建31×13产业份额矩阵和31×31距离矩阵。
仅需 Python 3.10+ 和 XlsxWriter；不需要网络。
"""
from pathlib import Path
import csv, math, argparse

INDUSTRY_FILE = '太平洋地区_2024产业结构_13类_完整.csv'
OUTPUT = '太平洋地区_产业结构矩阵_31MSA.xlsx'

def read_data(path):
    with path.open('r', encoding='utf-8-sig', newline='') as f:
        rows = list(csv.DictReader(f))
    if len(rows) != 403:
        raise ValueError(f'应有403条行业数据，实际{len(rows)}条')
    cities = {}
    industry_order = []
    for row in rows:
        cbsa = str(row['CBSA']).strip()
        name = row['MSA'].strip()
        industry = row['行业类别'].strip()
        total = int(row['就业合计'].replace(',', ''))
        male = int(row['男性就业'].replace(',', ''))
        female = int(row['女性就业'].replace(',', ''))
        if total != male + female or total < 0:
            raise ValueError(f'{name} {industry}：男女就业人数加总不符')
        if cbsa not in cities:
            cities[cbsa] = {'name': name, 'no': int(row['序号']), 'jobs': {}}
        if industry in cities[cbsa]['jobs']:
            raise ValueError(f'{name} {industry} 重复')
        cities[cbsa]['jobs'][industry] = total
        if industry not in industry_order:
            industry_order.append(industry)
    citylist = sorted(cities.items(), key=lambda kv: kv[1]['no'])
    if len(citylist) != 31 or len(industry_order) != 13:
        raise ValueError(f'预期31个MSA及13类行业，实际{len(citylist)}个和{len(industry_order)}类')
    if sorted(c['no'] for _,c in citylist) != list(range(1,32)):
        raise ValueError('城市序号必须为1至31且无重复')
    shares, totals = [], []
    for cbsa, city in citylist:
        if set(city['jobs']) != set(industry_order):
            raise ValueError(f"{city['name']} 行业类别不完整")
        n = sum(city['jobs'].values())
        if n <= 0: raise ValueError(f"{city['name']} 就业总数为零")
        p = [city['jobs'][industry] / n for industry in industry_order]
        if abs(sum(p)-1) > 1e-12: raise ValueError('产业占比未合计为1')
        shares.append(p); totals.append(n)
    return citylist, industry_order, shares, totals

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input', type=Path, default=Path(INDUSTRY_FILE))
    parser.add_argument('--output', type=Path, default=Path(OUTPUT))
    args = parser.parse_args()
    cities, sectors, shares, totals = read_data(args.input)
    n = len(cities)
    h = [[math.sqrt(0.5 * sum((math.sqrt(a)-math.sqrt(b))**2 for a,b in zip(shares[i],shares[k])))
          for k in range(n)] for i in range(n)]
    rms = math.sqrt(sum(h[i][k]**2 for i in range(n) for k in range(i+1,n)) / (n*(n-1)/2))
    if rms <= 0: raise ValueError('所有城市产业结构相同，无法进行RMS校准')
    factor = math.sqrt(2*n/(n-1)) / rms
    e = [[x * factor for x in row] for row in h]
    try: import xlsxwriter
    except ImportError: raise SystemExit('缺少 XlsxWriter，请执行：py -m pip install XlsxWriter')
    w = xlsxwriter.Workbook(str(args.output))
    head = w.add_format({'bold':True,'font_color':'white','bg_color':'#155E75','text_wrap':True,'border':1})
    pct = w.add_format({'num_format':'0.0000%'})
    dec = w.add_format({'num_format':'0.000000'})
    integer = w.add_format({'num_format':'#,##0'})
    note = w.add_format({'text_wrap':True,'valign':'top'})
    s = w.add_worksheet('31×13就业占比矩阵')
    s.write_row(0,0,['序号','MSA','CBSA'] + sectors + ['就业人数合计','占比合计'],head)
    for i,(cbsa,city) in enumerate(cities):
        s.write_row(i+1,0,[city['no'],city['name'],cbsa])
        for j,p in enumerate(shares[i]): s.write_number(i+1,3+j,p,pct)
        s.write_number(i+1,16,totals[i],integer)
        s.write_formula(i+1,17,f'=SUM(D{i+2}:P{i+2})',pct)
    s.freeze_panes(1,3);s.autofilter(0,0,n,17)
    s.set_column('A:A',8);s.set_column('B:B',49);s.set_column('C:C',12)
    s.set_column('D:P',19);s.set_column('Q:R',18)
    t = w.add_worksheet('31×13就业人数矩阵')
    t.write_row(0,0,['序号','MSA','CBSA'] + sectors + ['就业人数合计'],head)
    for i,(cbsa,city) in enumerate(cities):
        t.write_row(i+1,0,[city['no'],city['name'],cbsa])
        for j,sector in enumerate(sectors): t.write_number(i+1,j+3,city['jobs'][sector],integer)
        t.write_number(i+1,16,totals[i],integer)
    t.freeze_panes(1,3);t.set_column('A:A',8);t.set_column('B:B',49);t.set_column('C:C',12);t.set_column('D:Q',18)
    for sheetname, matrix in [('31×31 Hellinger距离',h),('31×31 校准产业距离E',e)]:
        sh=w.add_worksheet(sheetname)
        sh.write_row(0,0,['序号','MSA','CBSA']+[f'{c[1]["no"]:02d} {c[1]["name"]}' for c in cities],head)
        for i,(cbsa,city) in enumerate(cities):
            sh.write_row(i+1,0,[city['no'],city['name'],cbsa])
            for k,value in enumerate(matrix[i]): sh.write_number(i+1,3+k,value,dec)
        sh.freeze_panes(1,3);sh.set_column('A:A',8);sh.set_column('B:B',49);sh.set_column('C:C',12);sh.set_column(3,33,24)
        sh.conditional_format(1,3,n,33,{'type':'3_color_scale'})
    info=w.add_worksheet('计算说明')
    descriptions=[
        ('统计单位','31个原始清单MSA，顺序与原CSV序号一致'),
        ('原始数据','太平洋地区_2024产业结构_13类_完整.csv，2024 ACS C24030，13类行业'),
        ('产业就业占比','p_ij = 第i个MSA第j行业就业人数 / 第i个MSA的13行业就业总数'),
        ('Hellinger距离','H(i,k) = sqrt(0.5 × SUM_j (sqrt(p_ij)-sqrt(p_kj))^2)'),
        ('配对数','31×30/2 = 465'),
        ('H_RMS','sqrt(所有i<k的H(i,k)^2之和 / 465)'),
        ('E距离','E(i,k) = H(i,k) / H_RMS × sqrt(2n/(n-1))，n=31'),
        ('注意','该E矩阵只适用于31城整体比较。若按北/南太平洋分区筛选，应在各自最终候选集合内重新计算H_RMS和E，不得直接截取本表的E子矩阵。'),
        ('其他指标','本工作簿只处理五指标中的第④项产业结构；未添加新评价指标。'),
        ('质量控制','31城×13行业=403条；各城行业占比之和=100%；对角线距离=0；距离矩阵对称。')]
    info.write_row(0,0,['项目','定义或说明'],head)
    for i,row in enumerate(descriptions,1):
        info.write_row(i,0,row,note)
        info.set_row(i,45 if i in (7,8) else 33)
    info.write(12,0,'31城H_RMS',head);info.write_number(12,1,rms,dec)
    info.write(13,0,'31城E校准系数',head);info.write_number(13,1,factor,dec)
    info.set_column('A:A',25);info.set_column('B:B',110)
    w.close()
    print('完成：',args.output.resolve())
    print(f'31×13占比矩阵；31×13就业人数矩阵；31×31 Hellinger距离；31×31 校准产业距离E')
    print(f'行业记录403条，配对465组，H_RMS={rms:.8f}，E系数={factor:.8f}')

if __name__ == '__main__': main()
