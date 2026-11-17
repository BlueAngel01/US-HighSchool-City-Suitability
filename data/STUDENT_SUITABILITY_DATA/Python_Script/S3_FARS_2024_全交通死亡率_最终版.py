# -*- coding: utf-8 -*-
"""
S3 - 2024 FARS all road-user fatalities for 10 MSAs
Sources:
  NHTSA FARS 2024 National CSV
  U.S. Census Bureau CBSA-EST2024-ALLDATA

Definition:
  S3 counts all FARS fatal persons (INJ_SEV=4), regardless of PER_TYP,
  assigned to the MSA by STATE + COUNTY FIPS.

Outputs:
  S3_2024_AllRoadUsers_Summary.csv
  S3_2024_AllRoadUsers_10MSA.xlsx
  S3_2024_County_Detail.csv
  S3_2024_FARS_Occupant_Fatal_Records.csv
"""

import io
import zipfile
from pathlib import Path

import pandas as pd
import requests

YEAR = 2024
OUTDIR = Path(__file__).resolve().parent

FARS_URL = (
    "https://static.nhtsa.gov/nhtsa/downloads/FARS/"
    "2024/National/FARS2024NationalCSV.zip"
)
CENSUS_URL = (
    "https://www2.census.gov/programs-surveys/popest/"
    "datasets/2020-2024/metro/totals/cbsa-est2024-alldata.csv"
)

# The ten MSA / CBSA codes used in the current project
CITIES = {
    "14460": "Boston",
    "35620": "New York",
    "37980": "Philadelphia",
    "17410": "Cleveland",
    "28140": "Kansas City",
    "27260": "Jacksonville",
    "13820": "Birmingham",
    "26420": "Houston",
    "38060": "Phoenix",
    "41740": "San Diego",
}


def download(url, filename):
    """Download once and cache the original official source file."""
    path = OUTDIR / filename
    if path.exists() and path.stat().st_size > 0:
        print(f"Using cached file: {path.name}")
        return path.read_bytes()

    print(f"Downloading: {filename}")
    response = requests.get(
        url,
        timeout=(30, 300),
        headers={"User-Agent": "Mozilla/5.0"}
    )
    response.raise_for_status()
    path.write_bytes(response.content)
    return response.content


def normalize_columns(df):
    df.columns = [
        str(c).replace("\ufeff", "").strip().upper()
        for c in df.columns
    ]
    return df


def main():
    print("1. Downloading Census 2024 MSA population file...")
    census_bytes = download(
        CENSUS_URL, "cbsa-est2024-alldata.csv"
    )
    census = pd.read_csv(
        io.BytesIO(census_bytes),
        dtype=str,
        encoding="latin1",
        low_memory=False
    )
    census = normalize_columns(census)

    needed = {
        "CBSA", "MDIV", "STCOU", "NAME",
        "LSAD", "POPESTIMATE2024"
    }
    missing = needed - set(census.columns)
    if missing:
        raise RuntimeError(
            "Census file layout differs from expected layout. "
            f"Missing columns: {sorted(missing)}"
        )

    for col in ["CBSA", "MDIV", "STCOU"]:
        census[col] = census[col].fillna("").str.strip()

    census["CBSA"] = census["CBSA"].str.zfill(5)
    census["STCOU"] = census["STCOU"].str.zfill(5)

    # IMPORTANT:
    # In CBSA-EST2024-ALLDATA, the MSA-level total row is identified by:
    #   STCOU = 00000 AND MDIV is blank.
    # Metropolitan-division rows also have STCOU=00000, but MDIV is populated.
    # County component rows have actual county FIPS in STCOU.
    total_rows = census[
        census["CBSA"].isin(CITIES)
        & (census["STCOU"] == "00000")
        & (census["MDIV"] == "")
    ].copy()

    population = {}
    msa_names = {}

    for cbsa, city in CITIES.items():
        rows = total_rows[total_rows["CBSA"] == cbsa]

        if len(rows) != 1:
            debug = census[
                census["CBSA"] == cbsa
            ][["CBSA", "MDIV", "STCOU", "NAME", "LSAD"]]
            raise RuntimeError(
                f"{city} ({cbsa}): expected exactly one MSA total row "
                f"with STCOU=00000 and blank MDIV; found {len(rows)}.\n"
                f"Rows for diagnosis:\n{debug.to_string(index=False)}"
            )

        population[cbsa] = int(
            pd.to_numeric(
                rows.iloc[0]["POPESTIMATE2024"],
                errors="raise"
            )
        )
        msa_names[cbsa] = rows.iloc[0]["NAME"]

    # County components: actual county FIPS only; exclude the
    # MSA/division total rows (STCOU=00000) and any non-county rows.
    county_rows = census[
        census["CBSA"].isin(CITIES)
        & census["STCOU"].str.fullmatch(r"\d{5}", na=False)
        & (census["STCOU"] != "00000")
    ][["CBSA", "STCOU", "NAME"]].drop_duplicates()

    if county_rows.empty:
        raise RuntimeError("No county components found in Census file.")

    county_rows = county_rows.rename(
        columns={"STCOU": "COUNTY_FIPS", "NAME": "CENSUS_COUNTY_NAME"}
    )

    # Check each selected MSA has county components.
    missing_counties = [
        cbsa for cbsa in CITIES
        if county_rows.loc[county_rows["CBSA"] == cbsa].empty
    ]
    if missing_counties:
        raise RuntimeError(
            "No county components found for CBSA(s): "
            + ", ".join(missing_counties)
        )

    print("2. Downloading NHTSA FARS 2024 national ZIP...")
    fars_bytes = download(
        FARS_URL, "FARS2024NationalCSV.zip"
    )

    with zipfile.ZipFile(io.BytesIO(fars_bytes)) as archive:
        person_candidates = [
            name for name in archive.namelist()
            if name.replace("\\", "/").split("/")[-1].upper()
            == "PERSON.CSV"
        ]
        if len(person_candidates) != 1:
            raise RuntimeError(
                "Could not uniquely locate PERSON.CSV in FARS ZIP. "
                f"Candidates: {person_candidates}"
            )

        with archive.open(person_candidates[0]) as f:
            person = pd.read_csv(
                f,
                dtype=str,
                low_memory=False,
                encoding="latin1"
            )

    person = normalize_columns(person)

    required_person = {
        "STATE", "COUNTY", "PER_TYP", "INJ_SEV"
    }
    missing = required_person - set(person.columns)
    if missing:
        raise RuntimeError(
            "FARS PERSON.CSV is missing required columns: "
            f"{sorted(missing)}"
        )

    # Build 5-digit county FIPS from FARS state and county fields.
    person["STATE"] = person["STATE"].fillna("").str.strip()
    person["COUNTY"] = person["COUNTY"].fillna("").str.strip()
    person["COUNTY_FIPS"] = (
        person["STATE"].str.zfill(2)
        + person["COUNTY"].str.zfill(3)
    )

    # Count every FARS person with fatal injury severity, regardless of
    # road-user type. Do not use an occupant-only PER_TYP filter.
    fatal = person[person["INJ_SEV"].fillna("").str.strip() == "4"].copy()
    selected_fips = set(county_rows["COUNTY_FIPS"])
    fatal_msa = fatal[fatal["COUNTY_FIPS"].isin(selected_fips)].copy()
    fatal_msa["PER_TYP"] = fatal_msa["PER_TYP"].fillna("").str.strip()

    # The 'Other' category deliberately includes unknown and unusual
    # PER_TYP codes; these remain in the total rather than being dropped.
    category_map = {
        "1": "D1_Driver", "2": "D2_Passenger",
        "5": "D3_Pedestrian", "6": "D4_Pedalcyclist"
    }
    fatal_msa["S3_CATEGORY"] = fatal_msa["PER_TYP"].map(
        category_map
    ).fillna("D5_Other")
    categories = [
        "D1_Driver", "D2_Passenger", "D3_Pedestrian",
        "D4_Pedalcyclist", "D5_Other"
    ]

    print("3. Matching all fatal road users to MSA counties...")
    counts = pd.crosstab(
        fatal_msa["COUNTY_FIPS"], fatal_msa["S3_CATEGORY"]
    ).reindex(columns=categories, fill_value=0)
    counts["S3_FATALITIES"] = counts[categories].sum(axis=1)
    counts = counts.reset_index()

    county_detail = county_rows.merge(
        counts, on="COUNTY_FIPS", how="left"
    )
    for col in categories + ["S3_FATALITIES"]:
        county_detail[col] = county_detail[col].fillna(0).astype(int)
    county_detail["CITY"] = county_detail["CBSA"].map(CITIES)

    # The county list must contain unique MSA/county pairs.
    if county_detail.duplicated(["CBSA", "COUNTY_FIPS"]).any():
        raise RuntimeError("Duplicate MSA/county pair in Census components")
    if county_detail["COUNTY_FIPS"].eq("00000").any():
        raise RuntimeError("Invalid MSA aggregate row in county detail")

    print("4. Calculating total fatalities and category rates...")
    summary = county_detail.groupby("CBSA", as_index=False)[
        categories + ["S3_FATALITIES"]
    ].sum()
    summary["CITY"] = summary["CBSA"].map(CITIES)
    summary["MSA_NAME_CENSUS"] = summary["CBSA"].map(msa_names)
    summary["MSA_POPULATION_2024"] = summary["CBSA"].map(population)
    summary["YEAR"] = YEAR
    summary["S3_DEATHS_PER_100K"] = (
        summary["S3_FATALITIES"] / summary["MSA_POPULATION_2024"] * 100000
    )
    summary["DEFINITION"] = (
        "All FARS persons with INJ_SEV=4, all PER_TYP categories; "
        "MSA county FIPS, 2024 Census population"
    )
    summary = summary[[
        "CITY", "CBSA", "MSA_NAME_CENSUS", "YEAR",
        *categories, "S3_FATALITIES", "MSA_POPULATION_2024",
        "S3_DEATHS_PER_100K", "DEFINITION"
    ]].sort_values("S3_DEATHS_PER_100K")

    # Independent checks: no fatality is dropped by category grouping.
    if int(summary["S3_FATALITIES"].sum()) != len(fatal_msa):
        raise RuntimeError("Fatal-person total does not reconcile")
    if not (summary[categories].sum(axis=1) == summary["S3_FATALITIES"]).all():
        raise RuntimeError("Road-user categories do not sum to S3 total")

    summary.to_csv(
        OUTDIR / "S3_2024_AllRoadUsers_Summary.csv",
        index=False, encoding="utf-8-sig"
    )
    county_detail.to_csv(
        OUTDIR / "S3_2024_AllRoadUsers_County.csv",
        index=False, encoding="utf-8-sig"
    )
    fatal_msa.to_csv(
        OUTDIR / "S3_2024_AllRoadUsers_FARS_Person.csv",
        index=False, encoding="utf-8-sig"
    )

    # Audit table reports all PER_TYP codes, including unknown categories.
    code_audit = fatal_msa.groupby(
        ["PER_TYP", "S3_CATEGORY"], dropna=False
    ).size().rename("FATALITIES").reset_index()
    with pd.ExcelWriter(
        OUTDIR / "S3_2024_AllRoadUsers_10MSA.xlsx", engine="openpyxl"
    ) as writer:
        summary.to_excel(writer, sheet_name="S3_Summary", index=False)
        county_detail.to_excel(writer, sheet_name="County_Detail", index=False)
        code_audit.to_excel(writer, sheet_name="PER_TYP_Audit", index=False)
        fatal_msa.to_excel(writer, sheet_name="FARS_Person_Records", index=False)

    print("\nSUCCESS: 10 MSAs, all FARS fatal road users")
    print(summary[["CITY", *categories, "S3_FATALITIES",
                   "S3_DEATHS_PER_100K"]].to_string(index=False))
    print("\nFiles saved in:", OUTDIR)
    print("S3_2024_AllRoadUsers_10MSA.xlsx")
    print("S3_2024_AllRoadUsers_Summary.csv")
    print("S3_2024_AllRoadUsers_County.csv")
    print("S3_2024_AllRoadUsers_FARS_Person.csv")



if __name__ == "__main__":
    main()
