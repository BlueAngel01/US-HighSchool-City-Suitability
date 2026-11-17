# -*- coding: utf-8 -*-
"""Pacific 31 MSA five-indicator V1.0 distance calculation.
Run: py "太平洋地区_五指标综合距离_最终计算.py"
Optional: --groups "分区.csv" (columns: CBSA,分区; every one of 31 MSAs assigned exactly once)
Requires: XlsxWriter; uses only Python standard library otherwise.
"""
import argparse, csv, math
from pathlib import Path

FIVE='太平洋地区_2024五指标_完整.csv'
IND='太平洋地区_2024产业结构_13类_完整.csv'
OUT='太平洋地区_五指标综合距离_V1.0.xlsx'

def read_csv(path):
    with path.open('r', encoding='utf-8-sig', newline='') as f:
        return list(csv.DictReader(f))

def number(s):
    return float(str(s).replace(',','').strip())

def load(folder):
    raw=read_csv(folder/FIVE)
    if len(raw)!=31: raise ValueError(f'五指标总表需要31行，实际{len(raw)}')
    cities=[]; seen=set()
    for r in raw:
        cbsa=str(r['CBSA']).strip()
        if cbsa in seen: raise ValueError('重复CBSA '+cbsa)
        seen.add(cbsa)
        cities.append({'no':int(r['序号']),'name':r['MSA'].strip(),'cbsa':cbsa,
            'pop':number(r['①2024人口（人）']),
            'density':number(r['②人口密度（人/平方英里）']),
            'gdp':number(r['③2024人均GDP（美元/人）']),
            'cagr':number(r['⑤2020至2024 CAGR（%）'])})
    if [c['no'] for c in cities]!=list(range(1,32)):
        raise ValueError('31城序号/顺序与原始清单不一致')
    for c in cities:
        if min(c['pop'],c['density'],c['gdp'])<=0 or not math.isfinite(c['cagr']):
            raise ValueError(c['name']+' 的五指标存在非法值')
    industry=read_csv(folder/IND)
    if len(industry)!=403:raise ValueError(f'行业明细需403条，实际{len(industry)}')
    by={c['cbsa']:{} for c in cities}; sectors=[]
    for r in industry:
        cbsa=str(r['CBSA']).strip(); sec=r['行业类别'].strip()
        if cbsa not in by:raise ValueError('产业表多出CBSA '+cbsa)
        if sec in by[cbsa]:raise ValueError('重复行业 '+cbsa+' '+sec)
        m=int(number(r['男性就业']));f=int(number(r['女性就业']));tot=int(number(r['就业合计']))
        if min(m,f,tot)<0 or m+f!=tot:raise ValueError('就业加总不符 '+cbsa+' '+sec)
        by[cbsa][sec]=tot
        if sec not in sectors:sectors.append(sec)
    if len(sectors)!=13:raise ValueError(f'需要13个行业，实际{len(sectors)}')
    shares=[]
    for c in cities:
        counts=by[c['cbsa']]
        if set(counts)!=set(sectors):raise ValueError('产业缺失 '+c['name'])
        total=sum(counts.values())
        if total<=0:raise ValueError('就业人数为零 '+c['name'])
        shares.append([counts[s]/total for s in sectors])
    return cities,sectors,shares

def calculate(cities, shares, indices):
    n=len(indices)
    if n<2:raise ValueError('分区至少2个MSA才能计算')
    vals=[]
    for k in indices:
        c=cities[k]
        vals.append([math.log(c['pop']),math.log(c['density']),c['gdp'],c['cagr']])
    mean=[sum(row[j] for row in vals)/n for j in range(4)]
    sd=[math.sqrt(sum((row[j]-mean[j])**2 for row in vals)/n) for j in range(4)]
    if any(v<=1e-12 for v in sd):raise ValueError('某指标在该分区总体标准差为0，不能计算Z-score')
    z=[[ (row[j]-mean[j])/sd[j] for j in range(4)] for row in vals]
    h=[[math.sqrt(0.5*sum((math.sqrt(a)-math.sqrt(b))**2 for a,b in zip(shares[indices[i]],shares[indices[k]])))
        for k in range(n)] for i in range(n)]
    rms=math.sqrt(sum(h[i][k]**2 for i in range(n) for k in range(i+1,n))/(n*(n-1)/2))
    if rms<=1e-12:raise ValueError('该分区产业结构距离全为0，无法RMS校准')
    factor=math.sqrt(2*n/(n-1))/rms
    e=[[v*factor for v in row] for row in h]
    d=[[math.sqrt(.2*(sum((z[i][j]-z[k][j])**2 for j in range(4))+e[i][k]**2))
        for k in range(n)] for i in range(n)]
    if any(abs(d[i][k]-d[k][i])>1e-10 or (i==k and abs(d[i][k])>1e-10)
           for i in range(n) for k in range(n)):
        raise AssertionError('综合距离对称性/对角线校验失败')
    scores=[sum(d[i][k] for i in range(n))/n for k in range(n)]
    first=min(range(n),key=lambda k:(scores[k],cities[indices[k]]['no']))
    pairs=[]
    for k in range(n):
        if k==first:continue
        score=sum(min(d[i][first],d[i][k]) for i in range(n))/n
        pairs.append((k,score,(scores[first]-score)/scores[first]))
    pairs.sort(key=lambda t:(t[1],cities[indices[t[0]]]['no']))
    return dict(n=n,idx=indices,z=z,mean=mean,sd=sd,h=h,e=e,d=d,rms=rms,factor=factor,
                scores=scores,first=first,pairs=pairs)

def groups_from_csv(path,cities):
    rows=read_csv(path)
    if not rows or 'CBSA' not in rows[0] or '分区' not in rows[0]:
        raise ValueError('分区CSV需包含CBSA、分区两列')
    ids={c['cbsa'] for c in cities}; assign={}
    for r in rows:
        cbsa=str(r['CBSA']).strip(); group=r['分区'].strip()
        if cbsa not in ids or not group or cbsa in assign:
            raise ValueError('分区出现无效/重复CBSA或空分区 '+cbsa)
        assign[cbsa]=group
    if set(assign)!=ids:raise ValueError('分区必须严格覆盖31城，不能增减。缺少：'+str(ids-set(assign)))
    groups={}
    for i,c in enumerate(cities):groups.setdefault(assign[c['cbsa']],[]).append(i)
    if any(len(v)<2 for v in groups.values()):raise ValueError('每个分区至少需要2城')
    return groups

def write_output(path,cities,sectors,shares,results):
    try:import xlsxwriter
    except ImportError:raise SystemExit('请先执行 py -m pip install XlsxWriter')
    w=xlsxwriter.Workbook(str(path))
    head=w.add_format({'bold':True,'font_color':'white','bg_color':'#145A76','text_wrap':True,'border':1})
    dec=w.add_format({'num_format':'0.000000'}); pct=w.add_format({'num_format':'0.00%'})
    num=w.add_format({'num_format':'#,##0.0000'})
    desc=w.add_format({'text_wrap':True,'valign':'top'})
    overview=w.add_worksheet('结果汇总')
    overview.write_row(0,0,['范围','候选城市数','H_RMS','K校准系数','第一座代表城市','D1','第二座代表城市','D2','改善率','是否≥20%'],head)
    for row,(group,r) in enumerate(results.items(),1):
        f=r['first'];k,score,improve=r['pairs'][0]
        overview.write_row(row,0,[group,r['n'],r['rms'],r['factor'],cities[r['idx'][f]]['name'],r['scores'][f],cities[r['idx'][k]]['name'],score,improve,'是' if improve>=.2 else '否'])
        for col in [2,3,5,7]:overview.write_number(row,col,[r['rms'],r['factor'],r['scores'][f],score][[2,3,5,7].index(col)],dec)
        overview.write_number(row,8,improve,pct)
    overview.set_column('A:A',20);overview.set_column('B:D',18);overview.set_column('E:E',48)
    overview.set_column('F:F',16);overview.set_column('G:G',48);overview.set_column('H:J',18)
    for group,r in results.items():
        # Excel sheet names max 31 chars, keep distinct with prefix
        tag='全体' if group=='31城整体' else '区'+str(list(results).index(group))
        n=r['n']; idx=r['idx']
        sh=w.add_worksheet(tag+'-Z值')
        sh.write_row(0,0,['序号','MSA','CBSA','Z人口ln','Z密度ln','Z人均GDP','Z增长率','平均综合距离'],head)
        for i,k in enumerate(idx,1):
            c=cities[k];sh.write_row(i,0,[c['no'],c['name'],c['cbsa']])
            for j,val in enumerate(r['z'][i-1]):sh.write_number(i,3+j,val,dec)
            sh.write_number(i,7,r['scores'][i-1],dec)
        sh.freeze_panes(1,3);sh.set_column('A:A',9);sh.set_column('B:B',48);sh.set_column('C:C',12);sh.set_column('D:H',20)
        for label,matrix in [('H',r['h']),('E',r['e']),('d',r['d'])]:
            sh=w.add_worksheet(tag+'-'+label+'矩阵')
            sh.write_row(0,0,['序号','MSA','CBSA']+[f"{cities[k]['no']:02d} {cities[k]['name']}" for k in idx],head)
            for i,k in enumerate(idx,1):
                c=cities[k];sh.write_row(i,0,[c['no'],c['name'],c['cbsa']])
                for j,v in enumerate(matrix[i-1]):sh.write_number(i,j+3,v,dec)
            sh.freeze_panes(1,3);sh.set_column('A:A',9);sh.set_column('B:B',48);sh.set_column('C:C',12);sh.set_column(3,2+n,26)
            sh.conditional_format(1,3,n,2+n,{'type':'3_color_scale'})
        sh=w.add_worksheet(tag+'-候选排名')
        sh.write_row(0,0,['排名','第一城候选','单城覆盖距离D({c})','固定第一城','第二城候选','双城覆盖距离','改善率'],head)
        ranking=sorted(range(n),key=lambda i:(r['scores'][i],cities[idx[i]]['no']))
        for j in range(n):
            i=ranking[j];sh.write_row(j+1,0,[j+1,cities[idx[i]]['name'],r['scores'][i]])
            if j<len(r['pairs']):
                k,score,imp=r['pairs'][j]
                sh.write(j+1,3,cities[idx[r['first']]]['name']);sh.write(j+1,4,cities[idx[k]]['name'])
                sh.write_number(j+1,5,score,dec);sh.write_number(j+1,6,imp,pct)
        sh.freeze_panes(1,1);sh.set_column('A:A',10);sh.set_column('B:B',48);sh.set_column('C:C',24)
        sh.set_column('D:E',48);sh.set_column('F:G',22)
    info=w.add_worksheet('方法说明')
    lines=[
        ('数据输入',FIVE+'；'+IND),
        ('固定候选清单','31个原始MSA，严格使用原始序号与CBSA；不增减'),
        ('五指标','人口、人口密度、人均GDP、13类行业就业结构、2020-2024 CAGR；各权重20%'),
        ('数值指标','ln(人口)、ln(人口密度)、人均GDP、CAGR百分数；各组内Z=(x-均值)/总体标准差'),
        ('行业结构','2024 ACS C24030：13类就业人数 / 各城13类就业人数之和'),
        ('Hellinger','H(i,k)=sqrt(0.5*sum_j (sqrt(p_ij)-sqrt(p_kj))^2)'),
        ('RMS','H_RMS=sqrt(sum_{i<k} H(i,k)^2 / (n*(n-1)/2))'),
        ('产业校准','E(i,k)=H(i,k)/H_RMS * sqrt(2*n/(n-1))'),
        ('综合距离','d(i,k)=sqrt(0.2*[ΔZ人口²+ΔZ密度²+ΔZ人均GDP²+E²+ΔZ增长率²])'),
        ('单城代表','D({c})=mean_i d(i,c)，取最小者'),
        ('双城代表','固定第一城，D({c1,c2})=mean_i min(d(i,c1),d(i,c2))，取最小者'),
        ('第二城阈值','(D1-D2)/D1>=20% 才满足事先固定的第二城纳入标准'),
        ('分区提醒','31城整体的Z、E和d不可直接截取用于南北分区；提供--groups时各组全部重新计算'),
        ('数据来源提醒','GDP、人口密度采用此前已整理的五指标CSV数据；本程序不重新验证其原始外部来源'),
        ('结果解释','d越小五指标越相似；筛选结果依赖分组，31城整体结果不等于南北分区最终结果')]
    info.write_row(0,0,['项目','计算定义'],head)
    for i,(a,b) in enumerate(lines,1):info.write_row(i,0,[a,b],desc);info.set_row(i,33)
    info.set_column('A:A',23);info.set_column('B:B',115)
    w.close()

def main():
    ap=argparse.ArgumentParser(description='计算五指标V1.0综合距离矩阵及代表城市')
    ap.add_argument('--folder',type=Path,default=Path.cwd(),help='包含两份CSV的文件夹')
    ap.add_argument('--groups',type=Path,help='可选分区CSV，列名为CBSA,分区，31城必须全部分配')
    ap.add_argument('--output',type=Path,default=None)
    args=ap.parse_args()
    cities,sectors,shares=load(args.folder)
    groups={'31城整体':list(range(31))}
    if args.groups:
        groups.update(groups_from_csv(args.groups,cities))
    results={name:calculate(cities,shares,idx) for name,idx in groups.items()}
    out=args.output or args.folder/OUT
    write_output(out,cities,sectors,shares,results)
    for group,r in results.items():
        f=r['first'];k,d2,imp=r['pairs'][0]
        print(f'{group}: {r["n"]}城；H_RMS={r["rms"]:.8f}；K={r["factor"]:.8f}')
        print(f'  第一城: {cities[r["idx"][f]]["name"]}, D1={r["scores"][f]:.6f}')
        print(f'  第二城: {cities[r["idx"][k]]["name"]}, D2={d2:.6f}, 改善={imp:.2%}, 达标={imp>=.2}')
    print('完成：',out.resolve())

if __name__=='__main__':main()
