# -*- coding: utf-8 -*-
"""
2023 Census CBP cultural facilities for 10 MSAs + TOTAL population normalization.

Install:
  py -m pip install pandas openpyxl requests

Run:
  py R1_Cultural_Facilities_CBP_10MSA_全人口标准化.py

Inputs downloaded automatically:
  cbp23msa.zip
  cbsa-est2023-alldata.csv

Output:
  R1_Cultural_Facilities_2023_10MSA_全人口标准化.xlsx

Method:
- Population denominator is TOTAL MSA resident population in 2023, not ages 14-18.
- Rate = establishments / total population * 10,000.
- CBP MSA field is MSA; establishment count field is EST.
- Suppressed / absent records remain missing, never converted to zero.
- Cleveland (CBSA 17460) is explicitly checked and reported; if CBP has no
  publishable record, this script does not fabricate a value.
"""

from pathlib import Path
import zipfile, requests
import pandas as pd

YEAR=2023
CBP_URL="https://www2.census.gov/programs-surveys/cbp/datasets/2023/cbp23msa.zip"
POP_URL="https://www2.census.gov/programs-surveys/popest/datasets/2020-2023/metro/totals/cbsa-est2023-alldata.csv"
CBP_ZIP=Path("cbp23msa.zip")
POP_CSV=Path("cbsa-est2023-alldata.csv")
OUT=Path("R1_Cultural_Facilities_2023_10MSA_全人口标准化.xlsx")

MSAS={
 "New York":"35620","Philadelphia":"37980","Boston":"14460","Cleveland":"17460",
 "Kansas City":"28140","Jacksonville":"27260","Birmingham":"13820",
 "Houston":"26420","Phoenix":"38060","San Diego":"41740"
}
INDUSTRIES={
 "512131":"电影院（普通电影院）",
 "512132":"汽车影院",
 "711110":"剧院公司及晚餐剧场",
 "711310":"有场地的表演艺术活动主办机构（宽口径，单列）",
 "712110":"博物馆"
}

def get_file(url,path):
    if path.exists() and path.stat().st_size>1000:
        print("Using cached:",path)
        return
    print("Downloading:",url)
    r=requests.get(url,timeout=180,headers={"User-Agent":"Mozilla/5.0"})
    r.raise_for_status()
    path.write_bytes(r.content)
    print("Saved:",path,"bytes:",len(r.content))

def read_cbp():
    get_file(CBP_URL,CBP_ZIP)
    with zipfile.ZipFile(CBP_ZIP) as z:
        names=[n for n in z.namelist() if n.lower().endswith((".txt",".csv"))]
        if not names: raise RuntimeError("No TXT/CSV in CBP ZIP: "+str(z.namelist()))
        name=next((n for n in names if "msa" in Path(n).name.lower()),names[0])
        with z.open(name) as f:
            sample=f.read(4096).decode("utf-8-sig",errors="replace")
            f.seek(0)
            line=sample.splitlines()[0]
            sep=max([",","\t","|"],key=lambda x:line.count(x))
            d=pd.read_csv(f,sep=sep,dtype=str,low_memory=False,encoding="utf-8-sig")
    d.columns=[c.strip().upper() for c in d.columns]
    # Official CBP 2023 MSA file schema
    geo="MSA" if "MSA" in d.columns else "CBSA"
    industry="NAICS" if "NAICS" in d.columns else "NAICS2017"
    estab="EST" if "EST" in d.columns else "ESTAB"
    d["GEO"]=d[geo].astype(str).str.replace(r"\.0$","",regex=True).str.zfill(5)
    d["IND"]=d[industry].astype(str).str.replace(r"\.0$","",regex=True)
    d["EST_NUM"]=pd.to_numeric(d[estab],errors="coerce")
    d["STATUS"]="reported"
    d.loc[d[estab].isna() | d[estab].astype(str).str.strip().isin(["","N","S"]), "STATUS"]="suppressed_or_not_reported"
    return d,estab

def read_population():
    get_file(POP_URL,POP_CSV)
    # Census population CSV may contain non-UTF-8 characters in NAME
    # (for example accented place names). Read with Windows-1252/Latin-1
    # fallback; CBSA and population fields themselves are numeric/text ASCII.
    try:
        p=pd.read_csv(POP_CSV,dtype={"CBSA":str},low_memory=False,encoding="utf-8-sig")
    except UnicodeDecodeError:
        try:
            p=pd.read_csv(POP_CSV,dtype={"CBSA":str},low_memory=False,encoding="cp1252")
        except UnicodeDecodeError:
            p=pd.read_csv(POP_CSV,dtype={"CBSA":str},low_memory=False,encoding="latin1")
    p.columns=[c.strip().upper() for c in p.columns]
    cb="CBSA"
    popcol="POPESTIMATE2023"
    namecol="NAME"
    p["GEO"]=p[cb].astype(str).str.zfill(5)
    p["POP2023"]=pd.to_numeric(p[popcol],errors="coerce")
    return p[["GEO",namecol,"POP2023"]].drop_duplicates("GEO")

def main():
    cbp,estabcol=read_cbp()
    pop=read_population()
    rev={v:k for k,v in MSAS.items()}
    sub=cbp[cbp["GEO"].isin(MSAS.values()) & cbp["IND"].isin(INDUSTRIES)].copy()
    sub["MSA_NAME"]=sub["GEO"].map(rev)
    sub["Industry"]=sub["IND"].map(INDUSTRIES)
    sub["NAICS"]=sub["IND"]

    records=[]
    for name,code in MSAS.items():
        pr=pop[pop["GEO"]==code]
        population=pr.iloc[0]["POP2023"] if len(pr) else pd.NA
        for naics,label in INDUSTRIES.items():
            rr=sub[(sub["GEO"]==code)&(sub["IND"]==naics)]
            if len(rr):
                count=rr.iloc[0]["EST_NUM"]
                status=rr.iloc[0]["STATUS"]
            else:
                count=pd.NA
                status="no_publishable_record_or_suppressed; NOT zero"
            rate=(count/population*10000) if pd.notna(count) and pd.notna(population) and population else pd.NA
            records.append({
                "MSA":name,"CBSA":code,"2023总人口":population,
                "NAICS":naics,"文化设施类别":label,
                "CBP经营场所数量":count,
                "每万人口经营场所数":rate,
                "数据状态":status
            })
    long=pd.DataFrame(records)
    wide=long.pivot(index=["MSA","CBSA","2023总人口"],columns="文化设施类别",
                    values="CBP经营场所数量").reset_index()
    ratewide=long.pivot(index=["MSA","CBSA","2023总人口"],columns="文化设施类别",
                    values="每万人口经营场所数").reset_index()
    flags=long[["MSA","NAICS","文化设施类别","数据状态"]]
    # Verify Cleveland presence at geography level independent of target industries.
    geo_check=cbp[cbp["GEO"]=="17460"].copy()
    cleveland_check=pd.DataFrame([
      ["Cleveland CBSA code","17460"],
      ["Any Cleveland MSA records in CBP file", "Yes" if len(cbp[cbp["GEO"]=="17460"]) else "No"],
      ["Cleveland rows for selected five NAICS", str(len(sub[sub["GEO"]=="17460"]))],
      ["Selected-industry rows in all 10 MSAs",str(len(sub))]
    ],columns=["检查项目","结果"])
    method=pd.DataFrame([
      ["年份","2023，CBP行业数据与Census Vintage 2023 MSA总人口估计保持同年。"],
      ["人口分母","MSA 2023年7月1日总常住人口；不是14—18岁人口。"],
      ["标准化公式","每万人口设施数 = CBP establishments / MSA total population × 10,000。"],
      ["CBP统计单位","有雇员经营场所（establishments），不是建筑物、银幕数、座位数或全部公共设施。"],
      ["电影院","512131普通电影院；512132汽车影院单独列示。汽车影院缺失时不擅自当0。"],
      ["剧院","711110剧院公司及晚餐剧场；711310是更宽的表演艺术/体育/类似活动主办机构，单独列示，不合并冒充剧场数。"],
      ["博物馆","712110 Museums。"],
      ["缺失处理","未发布、抑制、无记录均保留缺失，不得填0。2017年后CBP对少于3家机构的详细单元可能不发布。"],
      ["Cleveland","CBSA 17460单独检查。若对应行业未发布，需用独立场馆名录或县级可发布数据补充，不能用0替代。"],
      ["CBP官方数据","https://www.census.gov/data/datasets/2023/econ/cbp/2023-cbp.html"],
      ["人口官方数据","https://www2.census.gov/programs-surveys/popest/datasets/2020-2023/metro/totals/cbsa-est2023-alldata.csv"]
    ],columns=["项目","统计口径与限制"])
    with pd.ExcelWriter(OUT,engine="openpyxl") as w:
        long.to_excel(w,sheet_name="10MSA长表_数量和率",index=False)
        wide.to_excel(w,sheet_name="设施数量汇总",index=False)
        ratewide.to_excel(w,sheet_name="每万人口设施率",index=False)
        flags.to_excel(w,sheet_name="缺失状态核查",index=False)
        cleveland_check.to_excel(w,sheet_name="Cleveland专项检查",index=False)
        method.to_excel(w,sheet_name="论文口径与来源",index=False)
        from openpyxl.styles import Font,PatternFill,Alignment
        from openpyxl.utils import get_column_letter
        for ws in w.book.worksheets:
            ws.freeze_panes="A2";ws.auto_filter.ref=ws.dimensions
            for c in ws[1]:
                c.font=Font(bold=True,color="FFFFFF")
                c.fill=PatternFill("solid",fgColor="34495E")
                c.alignment=Alignment(wrap_text=True)
            for col in ws.columns:
                letter=get_column_letter(col[0].column)
                ws.column_dimensions[letter].width=min(max(max(len(str(x.value or "")) for x in list(col)[:100])+2,12),55)
    print("\n2023 MSA cultural facilities, normalized by TOTAL population")
    print(long.to_string(index=False))
    print("\nCleveland check:")
    print(cleveland_check.to_string(index=False))
    print("\nSaved:",OUT.resolve())

if __name__=="__main__":
    main()
