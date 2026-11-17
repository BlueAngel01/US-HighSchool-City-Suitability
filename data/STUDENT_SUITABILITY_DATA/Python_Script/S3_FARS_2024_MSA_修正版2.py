# -*- coding: utf-8 -*-
"""
S3 - 2024 FARS motor-vehicle occupant fatalities for 10 MSAs
Sources:
  NHTSA FARS 2024 National CSV
  U.S. Census Bureau CBSA-EST2024-ALLDATA

Definition:
  S3 counts fatal drivers (PER_TYP=1) and passengers (PER_TYP=2)
  with fatal injury severity (INJ_SEV=4), assigned to the MSA
  by the person's STATE + COUNTY FIPS.
  Pedestrians (PER_TYP=5) and pedalcyclists (PER_TYP=6) are excluded.

Outputs:
  S3_2024_MSA_Summary.csv
  S3_2024_MSA_Results.xlsx
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
    "17460": "Cleveland",
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
                f"with STCOU=99999 and MDIV=0/blank; found {len(rows)}.\n"
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
    # synthetic MSA-total row STCOU=99999 and any non-county rows.
    county_rows = census[
        census["CBSA"].isin(CITIES)
        & census["STCOU"].str.fullmatch(r"\d{5}", na=False)
        & (census["STCOU"] != "99999")
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

    # FARS codes:
    # PER_TYP 1 = driver of motor vehicle in-transport
    # PER_TYP 2 = passenger of motor vehicle in-transport
    # INJ_SEV 4 = fatal injury
    occupant_fatal = person[
        person["PER_TYP"].str.strip().isin(["1", "2"])
        & (person["INJ_SEV"].str.strip() == "4")
    ].copy()

    print("3. Matching FARS fatal occupants to MSA counties...")
    selected_fips = set(county_rows["COUNTY_FIPS"])
    occupant_fatal_msa = occupant_fatal[
        occupant_fatal["COUNTY_FIPS"].isin(selected_fips)
    ].copy()

    # County-level counts; retain zero-fatality counties.
    deaths_by_county = (
        occupant_fatal_msa.groupby("COUNTY_FIPS")
        .size()
        .rename("S3_FATALITIES")
        .reset_index()
    )

    county_detail = county_rows.merge(
        deaths_by_county,
        on="COUNTY_FIPS",
        how="left"
    )
    county_detail["S3_FATALITIES"] = (
        county_detail["S3_FATALITIES"]
        .fillna(0)
        .astype(int)
    )
    county_detail["CITY"] = county_detail["CBSA"].map(CITIES)

    print("4. Calculating MSA rates...")
    summary = (
        county_detail.groupby("CBSA", as_index=False)
        ["S3_FATALITIES"].sum()
    )
    summary["CITY"] = summary["CBSA"].map(CITIES)
    summary["MSA_NAME_CENSUS"] = summary["CBSA"].map(msa_names)
    summary["MSA_POPULATION_2024"] = summary["CBSA"].map(population)
    summary["YEAR"] = YEAR
    summary["S3_DEATHS_PER_100K"] = (
        summary["S3_FATALITIES"]
        / summary["MSA_POPULATION_2024"]
        * 100000
    )
    summary["DEFINITION"] = (
        "FARS fatal drivers and passengers: PER_TYP 1 or 2, "
        "INJ_SEV 4; pedestrians and pedalcyclists excluded"
    )

    summary = summary[
        [
            "CITY", "CBSA", "MSA_NAME_CENSUS", "YEAR",
            "S3_FATALITIES", "MSA_POPULATION_2024",
            "S3_DEATHS_PER_100K", "DEFINITION"
        ]
    ].sort_values(
        "S3_DEATHS_PER_100K", ascending=True
    )

    # Output 1: concise summary
    summary_path = OUTDIR / "S3_2024_MSA_Summary.csv"
    summary.to_csv(
        summary_path, index=False, encoding="utf-8-sig"
    )

    # Output 2: county detail
    county_path = OUTDIR / "S3_2024_County_Detail.csv"
    county_detail.to_csv(
        county_path, index=False, encoding="utf-8-sig"
    )

    # Output 3: matched FARS person-level fatal occupant records
    # Keep original FARS columns and add county FIPS.
    person_path = OUTDIR / "S3_2024_FARS_Occupant_Fatal_Records.csv"
    occupant_fatal_msa.to_csv(
        person_path, index=False, encoding="utf-8-sig"
    )

    # Output 4: workbook
    xlsx_path = OUTDIR / "S3_2024_MSA_Results.xlsx"
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        summary.to_excel(
            writer, sheet_name="S3_Summary", index=False
        )
        county_detail.to_excel(
            writer, sheet_name="County_Detail", index=False
        )
        occupant_fatal_msa.to_excel(
            writer, sheet_name="FARS_Person_Records", index=False
        )

    print("\nSUCCESS - all ten MSA records processed.")
    print(summary.to_string(index=False))
    print("\nFiles created in:")
    print(OUTDIR)
    print(" - S3_2024_MSA_Summary.csv")
    print(" - S3_2024_County_Detail.csv")
    print(" - S3_2024_FARS_Occupant_Fatal_Records.csv")
    print(" - S3_2024_MSA_Results.xlsx")
    print("\nNote: The Census MSA total row uses STCOU=99999.")
    print("County components use actual county FIPS and exclude STCOU=00000.")


if __name__ == "__main__":
    main()
