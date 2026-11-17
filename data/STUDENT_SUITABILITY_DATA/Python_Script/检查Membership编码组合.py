import pandas as pd

f = "ccd_sch_052_2324_l_1a_073124.csv"

usecols = ["GRADE", "TOTAL_INDICATOR", "DMS_FLAG"]
counts = {}

print("开始扫描2.35GB文件，请稍候……")

for chunk in pd.read_csv(
    f,
    usecols=usecols,
    dtype=str,
    encoding="cp1252",
    chunksize=300000
):
    grouped = chunk.groupby(
        ["TOTAL_INDICATOR", "DMS_FLAG", "GRADE"],
        dropna=False
    ).size()

    for key, value in grouped.items():
        counts[key] = counts.get(key, 0) + int(value)

result = pd.DataFrame([
    {
        "TOTAL_INDICATOR": k[0],
        "DMS_FLAG": k[1],
        "GRADE": k[2],
        "记录数量": v
    }
    for k, v in counts.items()
])

result = result.sort_values(
    ["TOTAL_INDICATOR", "DMS_FLAG", "GRADE"]
)

print(result.to_string(index=False))

result.to_csv(
    "Membership_编码组合审计.csv",
    index=False,
    encoding="utf-8-sig"
)

print("\n审计完成！")
print("已生成 Membership_编码组合审计.csv")
