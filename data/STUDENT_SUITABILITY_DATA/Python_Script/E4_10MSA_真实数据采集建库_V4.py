# -*- coding: utf-8 -*-
"""E4 10个MSA真实数据采集建库 V3.0
需要：pandas、openpyxl
地理匹配：必须提供 school_cbsa_crosswalk.csv (NCESSCH,CBSA)
或 zip_cbsa_crosswalk.csv (ZIP,CBSA)。NCES preliminary directory只有州FIPST，
不能据此准确判定MSA；程序拒绝用州数据替代MSA数据。
"""
import argparse, csv, re, sqlite3
from pathlib import Path
import pandas as pd

MSAS = {
    "波士顿":"14460", "纽约":"35620", "费城":"37980", "克利夫兰":"17410",
    "堪萨斯城":"28140", "杰克逊维尔":"27260", "伯明翰":"13820",
    "休斯敦":"26420", "凤凰城":"38060", "圣迭戈":"41740"
}
COURSES = {
 "Advanced Mathematics":("Advanced Mathematics.csv","SCH_MATHCLASSES_ADVM",["TOT_MATHENR_ADVM"]),
 "Algebra I":("Algebra I.csv","SCH_MATHCLASSES_ALG",["TOT_ALGENR_GS0910","TOT_ALGENR_GS1112"]),
 "Algebra II":("Algebra II.csv","SCH_MATHCLASSES_ALG2",["TOT_MATHENR_ALG2"]),
 "Geometry":("Geometry.csv","SCH_MATHCLASSES_GEOM",["TOT_MATHENR_GEOM"]),
 "Calculus":("Calculus.csv","SCH_MATHCLASSES_CALC",["TOT_MATHENR_CALC"]),
 "Biology":("Biology.csv","SCH_SCICLASSES_BIOL",["TOT_SCIENR_BIOL"]),
 "Chemistry":("Chemistry.csv","SCH_SCICLASSES_CHEM",["TOT_SCIENR_CHEM"]),
 "Physics":("Physics.csv","SCH_SCICLASSES_PHYS",["TOT_SCIENR_PHYS"]),
 "Computer Science":("Computer Science.csv","SCH_COMPCLASSES_CSCI",["TOT_COMPENR_CSCI"]),
 "Data Science":("Data Science.csv","SCH_DATACLASSES_DSCI",["TOT_DATAENR_DSCI"])
}

def header(p):
    with open(p,encoding="cp1252",newline="",errors="replace") as f:
        return next(csv.reader(f))
def norm_id(v):
    return "" if pd.isna(v) else re.sub(r"\.0$","",str(v).strip())
def norm_zip(v):
    if pd.isna(v): return ""
    s=re.sub(r"\D","",str(v))
    return s[:5].zfill(5) if s else ""

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("--downloads",default=str(Path.home()/"Downloads"))
    ap.add_argument("--school-cbsa",default="")
    ap.add_argument("--zip-cbsa",default="")
    ap.add_argument("--out-dir",default="")
    a=ap.parse_args()
    dl=Path(a.downloads)
    crdc=dl/"2023-24-crdc-data"/"SCH"
    directory=dl/"ccd_sch_029_2324_w_1a_073124"/"ccd_sch_029_2324_w_1a_073124.csv"
    membership=dl/"ccd_sch_052_2324_l_1a_073124"/"ccd_sch_052_2324_l_1a_073124.csv"
    out=Path(a.out_dir) if a.out_dir else dl/"E4_10MSA_V4_结果"
    out.mkdir(parents=True,exist_ok=True)
    if not directory.exists(): raise FileNotFoundError(f"学校目录不存在：{directory}")
    if not membership.exists(): raise FileNotFoundError(f"Membership不存在：{membership}")
    if not crdc.exists(): raise FileNotFoundError(f"CRDC目录不存在：{crdc}")

    print("1/6 读取NCES学校目录",flush=True)
    wanted=["NCESSCH","LEAID","SCH_NAME","STATENAME","ST","FIPST","SCH_TYPE_TEXT","SCH_TYPE",
            "SY_STATUS_TEXT","G_9_OFFERED","G_10_OFFERED","G_11_OFFERED","G_12_OFFERED",
            "GSLO","GSHI","LZIP","LCITY","LSTATE"]
    h=header(directory); use=[c for c in wanted if c in h]
    sch=pd.read_csv(directory,usecols=use,dtype=str,encoding="cp1252",low_memory=False)
    sch["NCESSCH"]=sch.NCESSCH.map(norm_id)
    if a.school_cbsa:
        g=pd.read_csv(a.school_cbsa,dtype=str)
        if not {"NCESSCH","CBSA"}.issubset(g.columns): raise ValueError("学校对照文件必须有 NCESSCH,CBSA")
        g.NCESSCH=g.NCESSCH.map(norm_id)
        sch=sch.merge(g[["NCESSCH","CBSA"]].drop_duplicates("NCESSCH"),on="NCESSCH",how="left")
        geo_method="NCESSCH直接匹配"
    elif a.zip_cbsa:
        g=pd.read_csv(a.zip_cbsa,dtype=str)
        zcol="ZIP" if "ZIP" in g.columns else "LZIP" if "LZIP" in g.columns else None
        if not zcol or "CBSA" not in g.columns: raise ValueError("邮编对照文件必须有 ZIP,CBSA")
        g["ZIP5"]=g[zcol].map(norm_zip)
        g=g[["ZIP5","CBSA"]].drop_duplicates()
        # 一个邮编对应多个CBSA时排除，避免猜测
        n=g.groupby("ZIP5").CBSA.nunique()
        g=g[g.ZIP5.isin(n[n==1].index)].drop_duplicates("ZIP5")
        sch["ZIP5"]=sch.LZIP.map(norm_zip)
        sch=sch.merge(g,left_on="ZIP5",right_on="ZIP5",how="left")
        geo_method="实际地址LZIP匹配；跨CBSA邮编排除"
    else:
        raise RuntimeError("缺少地理对照文件。请提供 school_cbsa_crosswalk.csv (NCESSCH,CBSA) 或 zip_cbsa_crosswalk.csv (ZIP,CBSA)。NCES目录没有县FIPS，不能只凭州代码分配MSA。")
    sch["CBSA"]=sch.CBSA.astype("string").str.replace(r"\.0$","",regex=True).str.zfill(5)
    sch=sch[sch.CBSA.isin(MSAS.values())].copy()
    def yes(x): return str(x).strip().lower() in {"1","yes","y","true","t"}
    gcols=[c for c in ["G_9_OFFERED","G_10_OFFERED","G_11_OFFERED","G_12_OFFERED"] if c in sch]
    sch["HS_ELIGIBLE"]=sch[gcols].apply(lambda r:any(yes(x) for x in r),axis=1) if gcols else False
    if "SY_STATUS_TEXT" in sch:
        sch=sch[~sch.SY_STATUS_TEXT.fillna("").str.lower().str.contains("closed|inactive|not operational")]
    sch=sch[sch.HS_ELIGIBLE].drop_duplicates("NCESSCH").copy()
    if sch.empty: raise RuntimeError("没有匹配到目标MSA高中，请检查CBSA对照文件和年级字段编码。")
    sch["MSA_NAME"]=sch.CBSA.map({v:k for k,v in MSAS.items()})
    ids=set(sch.NCESSCH)
    print(f"   匹配高中 {len(sch):,} 所；地理方法：{geo_method}",flush=True)

    print("2/6 分块读取Membership，只提取目标学校9-12年级Subtotal 4",flush=True)
    use=["NCESSCH","GRADE","TOTAL_INDICATOR","DMS_FLAG","STUDENT_COUNT"]
    rows=[]; chunk_size=400000; chunks=0
    for ch in pd.read_csv(membership,usecols=use,dtype=str,encoding="cp1252",chunksize=chunk_size,low_memory=False):
        chunks+=1
        ch["NCESSCH"]=ch.NCESSCH.map(norm_id)
        ch=ch[ch.NCESSCH.isin(ids)&(ch.TOTAL_INDICATOR=="Subtotal 4 - By Grade")&
              ch.GRADE.isin(["Grade 9","Grade 10","Grade 11","Grade 12"])&
              ch.DMS_FLAG.isin(["Reported","Manual adjustment"])].copy()
        if not ch.empty:
            ch["COUNT"]=pd.to_numeric(ch.STUDENT_COUNT,errors="coerce").where(lambda x:x>=0)
            ch["GRADE_NO"]=ch.GRADE.str.extract(r"(\d+)").astype(int)
            rows.append(ch[["NCESSCH","GRADE","GRADE_NO","DMS_FLAG","COUNT"]])
        if chunks%10==0: print(f"   已扫描{chunks*chunk_size:,}行，命中{sum(map(len,rows)):,}条",flush=True)
    if rows:
        m=pd.concat(rows,ignore_index=True)
        m["P"]=m.DMS_FLAG.map({"Manual adjustment":0,"Reported":1})
        m=m.sort_values(["NCESSCH","GRADE_NO","P"]).drop_duplicates(["NCESSCH","GRADE_NO"])
        p=m.pivot(index="NCESSCH",columns="GRADE",values="COUNT").reset_index()
        p.columns.name=None
        for grade in ["Grade 9","Grade 10","Grade 11","Grade 12"]:
            if grade not in p: p[grade]=pd.NA
        sch=sch.merge(p,on="NCESSCH",how="left")
        offered_cols=[f"G_{g}_OFFERED" for g in range(9,13)]
        for g in range(9,13):
            grade=f"Grade {g}"
            flag=f"G_{g}_OFFERED"
            sch[grade]=pd.to_numeric(sch[grade],errors="coerce")
            # Only offered grades require membership data. Unoffered grades contribute zero.
            sch[f"VALID_{g}"]=~sch[flag].map(yes) | sch[grade].ge(0)
            sch[f"COUNT_{g}"]=sch[grade].where(sch[flag].map(yes),0)
        sch["HS_GRADE_DATA_COMPLETE"]=sch[[f"VALID_{g}" for g in range(9,13)]].all(axis=1)
        sch["HS_STUDENT_COUNT"]=sch[[f"COUNT_{g}" for g in range(9,13)]].sum(axis=1,min_count=4).where(sch.HS_GRADE_DATA_COMPLETE)

    else:
        for c in ["Grade 9","Grade 10","Grade 11","Grade 12","HS_STUDENT_COUNT"]: sch[c]=pd.NA
        sch["HS_GRADE_DATA_COMPLETE"]=False

    print("3/6 读取CRDC课程数据",flush=True)
    details=[]
    for course,(filename,classcol,prefixes) in COURSES.items():
        p=crdc/filename
        if not p.exists():
            print("   缺少文件："+filename,flush=True)
            continue
        h=header(p)
        if "COMBOKEY" not in h: raise ValueError(f"{filename}缺少COMBOKEY")
        enroll=[f"{prefix}_{sex}" for prefix in prefixes for sex in ("M","F","X")]
        absent=[c for c in enroll if c not in h]
        if absent: raise ValueError(f"{filename}缺少官方TOT列：{absent}")
        if classcol not in h: raise ValueError(f"{filename}缺少课程班级数列：{classcol}")
        use=["COMBOKEY",classcol]+enroll
        d=pd.read_csv(p,usecols=use,dtype=str,encoding="utf-8-sig",low_memory=False)
        d["NCESSCH"]=d.COMBOKEY.map(norm_id)
        d=d[d.NCESSCH.isin(ids)].copy()
        if d.empty: continue
        if d.NCESSCH.duplicated().any(): raise ValueError(f"{course}出现重复COMBOKEY，不能直接汇总")
        d["COURSE"]=course
        d["CLASS_COUNT"]=pd.to_numeric(d[classcol],errors="coerce").where(lambda x:x>=0)
        vals=d[enroll].apply(pd.to_numeric,errors="coerce")
        # CRDC -10 is treated as an explicitly flagged X-code, not a numeric zero.
        # Official code meaning must be checked before publication.
        # M/F must be valid nonnegative; X may be valid nonnegative or -10.
        # For Algebra I, both grade bands must meet this rule.
        parts=[]; states=[]
        for prefix in prefixes:
            mc,fc,xc=(f"{prefix}_{sex}" for sex in ("M","F","X"))
            mf_ok=vals[mc].ge(0)&vals[fc].ge(0)
            x_num=vals[xc].ge(0)
            x_flag=vals[xc].eq(-10)
            ok=mf_ok&(x_num|x_flag)
            part=(vals[mc]+vals[fc]+vals[xc].where(x_num,0)).where(ok)
            parts.append(part)
            states.append(pd.DataFrame({"ok":ok,"x_flag":x_flag,"mf_ok":mf_ok}))
        all_ok=pd.concat([v["ok"] for v in states],axis=1).all(axis=1)
        d["STUDENT_PARTICIPANTS"]=pd.concat(parts,axis=1).sum(axis=1,min_count=len(parts)).where(all_ok)
        d["ENROLL_VALID_FIELD_COUNT"]=vals.ge(0).sum(axis=1)
        d["ENROLL_FIELD_COUNT"]=len(enroll)
        d["X_MINUS10_COUNT"]=sum(v["x_flag"].astype(int) for v in states)
        d["ENROLL_STATUS"]="MISSING_OR_OTHER_CODE"
        d.loc[all_ok,"ENROLL_STATUS"]="COMPLETE_NUMERIC"
        x_any=d["X_MINUS10_COUNT"].gt(0)
        d.loc[all_ok & x_any,"ENROLL_STATUS"]="MF_VALID_X_MINUS10_UNVERIFIED"
        d["COURSE_CLASS_STATUS"]="VALID"
        d.loc[d["CLASS_COUNT"].isna(),"COURSE_CLASS_STATUS"]="MISSING_OR_NEGATIVE"
        details.append(d[["NCESSCH","COURSE","CLASS_COUNT","STUDENT_PARTICIPANTS","ENROLL_VALID_FIELD_COUNT","ENROLL_FIELD_COUNT","X_MINUS10_COUNT","ENROLL_STATUS","COURSE_CLASS_STATUS"]])
        print(f"   {course}: {len(d):,}所学校；官方TOT列{len(enroll)}个；可计算记录{d.STUDENT_PARTICIPANTS.notna().sum():,}；其中X=-10标记{(d.ENROLL_STATUS=="MF_VALID_X_MINUS10_UNVERIFIED").sum():,}",flush=True)
    cd=pd.concat(details,ignore_index=True) if details else pd.DataFrame(columns=["NCESSCH","COURSE","CLASS_COUNT","STUDENT_PARTICIPANTS","ENROLL_VALID_FIELD_COUNT","ENROLL_FIELD_COUNT","X_MINUS10_COUNT","ENROLL_STATUS","COURSE_CLASS_STATUS"])
    cd=cd.merge(sch[["NCESSCH","CBSA","MSA_NAME"]],on="NCESSCH",how="left")
    cd["COURSE_OFFERED"]=cd.CLASS_COUNT.gt(0)
    print("4/6 汇总10个MSA",flush=True)
    summary=cd.groupby(["CBSA","MSA_NAME","COURSE"],dropna=False).agg(
        schools_with_course=("COURSE_OFFERED","sum"),
        course_student_participants=("STUDENT_PARTICIPANTS",lambda x:x.sum(min_count=1)),
        schools_with_course_data=("CLASS_COUNT","count"),
        schools_with_participant_data=("STUDENT_PARTICIPANTS","count"),
        total_class_sections=("CLASS_COUNT",lambda x:x.sum(min_count=1))).reset_index()
    base=sch.groupby(["CBSA","MSA_NAME"]).agg(
        high_schools=("NCESSCH","nunique"),
        high_school_students=("HS_STUDENT_COUNT",lambda x:x.sum(min_count=1)),
        schools_with_complete_grade_data=("HS_GRADE_DATA_COMPLETE","sum"),
        schools_with_any_grade_data=("HS_STUDENT_COUNT","count")).reset_index()
    summary=summary.merge(base,on=["CBSA","MSA_NAME"],how="left")
    summary["school_course_coverage_pct"]=summary.schools_with_course/summary.high_schools*100
    # Comparable participation rates require matched, valid denominator schools.
    eligible=cd.merge(sch[["NCESSCH","HS_STUDENT_COUNT"]],on="NCESSCH",how="left")
    eligible["VALID_BOTH"]=eligible.STUDENT_PARTICIPANTS.notna() & eligible.HS_STUDENT_COUNT.notna() & eligible.HS_STUDENT_COUNT.gt(0)
    eligible["MATCHED_DENOM"]=eligible.HS_STUDENT_COUNT.where(eligible.VALID_BOTH)
    eligible["MATCHED_NUM"]=eligible.STUDENT_PARTICIPANTS.where(eligible.VALID_BOTH)
    matched=eligible.groupby(["CBSA","MSA_NAME","COURSE"]).agg(
        matched_students=("MATCHED_DENOM",lambda x:x.sum(min_count=1)),
        matched_course_enrollments=("MATCHED_NUM",lambda x:x.sum(min_count=1)),
        matched_schools=("VALID_BOTH","sum")
    ).reset_index()
    summary=summary.merge(matched,on=["CBSA","MSA_NAME","COURSE"],how="left")
    flags=cd.groupby(["CBSA","MSA_NAME","COURSE"]).agg(
        schools_x_minus10_unverified=("ENROLL_STATUS",lambda x:(x=="MF_VALID_X_MINUS10_UNVERIFIED").sum()),
        schools_course_missing=("ENROLL_STATUS",lambda x:(x=="MISSING_OR_OTHER_CODE").sum())
    ).reset_index()
    summary=summary.merge(flags,on=["CBSA","MSA_NAME","COURSE"],how="left")
    summary["student_participation_pct"]=summary.matched_course_enrollments/summary.matched_students*100
    summary["crdc_school_record_coverage_pct"]=summary.schools_with_course_data/summary.high_schools*100
    summary["matched_school_coverage_pct"]=summary.matched_schools/summary.high_schools*100
    # Summed course enrollment is NOT unique-student participation across courses.
    # Rates may exceed 100% if a school reports multiple course enrollments per student.

    # Explicit denominator and numerator coverage audit, never substitute missing as zero.
    summary["matched_student_coverage_pct"]=summary.matched_students/summary.high_school_students*100
    summary["x_code_verified_by_dictionary"]=False
    print("5/6 写入SQLite",flush=True)
    db=out/"E4_10MSA_V4.sqlite"
    with sqlite3.connect(db) as con:
        sch.to_sql("school_master",con,if_exists="replace",index=False)
        cd.to_sql("course_school_detail",con,if_exists="replace",index=False)
        summary.to_sql("msa_course_summary",con,if_exists="replace",index=False)
        base.to_sql("msa_school_enrollment_summary",con,if_exists="replace",index=False)
    print("6/6 导出Excel和CSV",flush=True)
    xlsx=out/"E4_10MSA_V4_真实数据结果.xlsx"
    with pd.ExcelWriter(xlsx,engine="openpyxl") as w:
        summary.to_excel(w,sheet_name="MSA课程指标",index=False)
        base.to_excel(w,sheet_name="MSA高中及学生基数",index=False)
        sch.to_excel(w,sheet_name="高中学校明细",index=False)
        cd.to_excel(w,sheet_name="学校课程明细",index=False)
    sch[["NCESSCH","SCH_NAME","CBSA","MSA_NAME","LZIP"]].to_csv(out/"E4_学校MSA匹配审计.csv",index=False,encoding="utf-8-sig")
    summary.to_csv(out/"E4_MSA课程指标.csv",index=False,encoding="utf-8-sig")
    print("\n完成。输出目录："+str(out),flush=True)
    print("SQLite："+str(db)); print("Excel："+str(xlsx))
    print("注意：X=-10官方语义尚未核验，当前M+F结果是暂定统计值，不可直接用于论文。")

if __name__=="__main__":
    main()
