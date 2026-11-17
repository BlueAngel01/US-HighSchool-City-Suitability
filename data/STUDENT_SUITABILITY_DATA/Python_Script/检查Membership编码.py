import pandas as pd

f = "ccd_sch_052_2324_l_1a_073124.csv"

cols = ["GRADE", "RACE_ETHNICITY", "SEX",
        "TOTAL_INDICATOR", "DMS_FLAG"]

values = {c:set() for c in cols}

print("开始检查2.35GB文件，请稍候……")

for chunk in pd.read_csv(
    f, usecols=cols, dtype=str,
    encoding="cp1252", chunksize=300000
):
    for c in cols:
        values[c].update(chunk[c].dropna().unique())

for c in cols:
    print("\n", c)
    print(sorted(values[c]))

print("\n字段编码检查完成")
