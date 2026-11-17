
import io
import zipfile
import requests
import pandas as pd

YEAR = 2024

FARS_URL = (
    "https://static.nhtsa.gov/nhtsa/downloads/FARS/"
    "2024/National/FARS2024NationalCSV.zip"
)

CENSUS_URL = (
    "https://www2.census.gov/programs-surveys/popest/"
    "datasets/2020-2024/metro/totals/"
    "cbsa-est2024-alldata.csv"
)

CITIES = {
    "14460": "Boston",
    "35620": "New York",
    "37980": "Philadelphia",
    "17460": "Cleveland",
    "28140": "Kansas City",
    "27260": "Jacksonville",
    "13820": "Birmingham",
    "26420": "Houston",
    "38060": "Phoenix",
    "41740": "San Diego",
}

def download(url):
    r = requests.get(url, timeout=180)
    r.raise_for_status()
    return r.content

def clean_columns(df):
    df.columns = [
        str(c).strip().upper().replace(" ", "_")
        for c in df.columns
    ]
    return df

print("1. Downloading Census 2024 data...")
census_bytes = download(CENSUS_URL)

census = pd.read_csv(
    io.BytesIO(census_bytes),
    dtype=str,
    encoding="latin1"
)
census = clean_columns(census)

required = {
    "CBSA", "STCOU", "NAME", "LSAD", "POPESTIMATE2024"
}
if not required.issubset(census.columns):
    raise ValueError(
        f"Census columns missing: {required - set(census.columns)}"
    )

census["CBSA"] = census["CBSA"].fillna("").str.strip().str.zfill(5)
census["STCOU"] = census["STCOU"].fillna("").str.strip()

# MSA population: official Census metropolitan-area total rows
msa_rows = census[
    census["CBSA"].isin(CITIES)
    & census["LSAD"].str.contains(
        "Metropolitan Statistical Area",
        case=False, na=False
    )
].copy()

pop = {}
for cbsa in CITIES:
    rows = msa_rows[msa_rows["CBSA"] == cbsa]
    if len(rows) != 1:
        raise ValueError(
            f"Cannot uniquely identify population for {cbsa}"
        )
    pop[cbsa] = int(
        float(rows.iloc[0]["POPESTIMATE2024"])
    )

# County components belonging to the ten MSAs
counties = census[
    census["CBSA"].isin(CITIES)
    & census["STCOU"].str.fullmatch(r"\d{5}", na=False)
][["CBSA", "STCOU", "NAME"]].drop_duplicates()

if counties.empty:
    raise ValueError("No MSA county components found.")

print("2. Downloading NHTSA FARS 2024...")
fars_bytes = download(FARS_URL)

with zipfile.ZipFile(io.BytesIO(fars_bytes)) as z:
    person_file = next(
        n for n in z.namelist()
        if n.split("/")[-1].upper() == "PERSON.CSV"
    )
    person = pd.read_csv(
        z.open(person_file),
        dtype=str,
        low_memory=False,
        encoding="latin1"
    )

person = clean_columns(person)

required = {
    "STATE", "COUNTY", "PER_TYP", "INJ_SEV"
}
if not required.issubset(person.columns):
    raise ValueError(
        f"FARS columns missing: {required - set(person.columns)}"
    )

# Build five-digit County FIPS
person["COUNTY_FIPS"] = (
    person["STATE"].str.strip().str.zfill(2)
    + person["COUNTY"].str.strip().str.zfill(3)
)

# S3: fatal drivers and passengers only
# PER_TYP 1 = Driver; 2 = Passenger
# INJ_SEV 4 = Fatal injury
occupants = person[
    person["PER_TYP"].str.strip().isin(["1", "2"])
    & (person["INJ_SEV"].str.strip() == "4")
].copy()

print("3. Aggregating fatalities by county...")
county_deaths = (
    occupants.groupby("COUNTY_FIPS")
    .size()
    .rename("S3_FATALITIES")
    .reset_index()
)

counties["COUNTY_FIPS"] = counties["STCOU"].str.zfill(5)

county_detail = counties.merge(
    county_deaths,
    on="COUNTY_FIPS",
    how="left"
)
county_detail["S3_FATALITIES"] = (
    county_detail["S3_FATALITIES"].fillna(0).astype(int)
)

print("4. Calculating MSA fatality rates...")

summary = (
    county_detail.groupby("CBSA")["S3_FATALITIES"]
    .sum()
    .reset_index()
)

summary["CITY"] = summary["CBSA"].map(CITIES)
summary["MSA_POPULATION_2024"] = (
    summary["CBSA"].map(pop)
)

summary["S3_DEATHS_PER_100K"] = (
    summary["S3_FATALITIES"]
    / summary["MSA_POPULATION_2024"]
    * 100000
)

summary["YEAR"] = YEAR
summary["DEFINITION"] = (
    "Fatal motor vehicle drivers and passengers; "
    "PER_TYP 1 and 2; INJ_SEV 4; pedestrians excluded"
)

summary = summary[
    [
        "CITY", "CBSA", "YEAR",
        "S3_FATALITIES",
        "MSA_POPULATION_2024",
        "S3_DEATHS_PER_100K",
        "DEFINITION"
    ]
].sort_values("S3_DEATHS_PER_100K")

# Export results
summary.to_csv(
    "S3_2024_MSA_Summary.csv",
    index=False,
    encoding="utf-8-sig"
)

with pd.ExcelWriter(
    "S3_2024_MSA_Results.xlsx",
    engine="openpyxl"
) as writer:
    summary.to_excel(
        writer, sheet_name="S3_Summary", index=False
    )
    county_detail.to_excel(
        writer, sheet_name="County_Detail", index=False
    )

print("\nCOMPLETED")
print(summary.to_string(index=False))
print("\nOutput:")
print("S3_2024_MSA_Summary.csv")
print("S3_2024_MSA_Results.xlsx")
