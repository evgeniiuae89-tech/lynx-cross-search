import re
from pathlib import Path

import duckdb
import pandas as pd


# ============================================================
# SETTINGS
# ============================================================

EXCEL_FILE = "LYNXauto_Cross.xlsx"
DATABASE_FILE = "lynx_cross.duckdb"

SHEETS = {
    "Aftermarket": "Aftermarket",
    "OEM": "OEM"
}


# ============================================================
# NORMALIZE PART NUMBER
# ============================================================

def normalize_number(value):
    """
    Convert a part number to a normalized search key.

    Examples:
    48510-02150 -> 4851002150
    UM51-33-251 -> UM5133251
    """

    if pd.isna(value):
        return ""

    return re.sub(
        r"[^A-Za-z0-9]",
        "",
        str(value)
    ).upper()


# ============================================================
# LOAD SHEET
# ============================================================

def load_sheet(excel_file, sheet_name, source_type):

    print(f"Reading sheet: {sheet_name}...")

    df = pd.read_excel(
        excel_file,
        sheet_name=sheet_name,
        dtype=str,
        engine="openpyxl"
    )

    # Remove spaces from column names
    df.columns = [
        str(column).strip()
        for column in df.columns
    ]

    required_columns = [
        "Groups",
        "LYNXauto number",
        "Description",
        "Cross type",
        "Cross clear",
        "Cross number",
        "Brand"
    ]

    missing_columns = [
        column
        for column in required_columns
        if column not in df.columns
    ]

    if missing_columns:

        raise ValueError(
            f"Missing columns in {sheet_name}: "
            f"{missing_columns}"
        )

    # Keep only required columns
    df = df[required_columns].copy()

    # Rename columns for database
    df = df.rename(
        columns={
            "Groups": "groups",
            "LYNXauto number": "lynx_number",
            "Description": "description",
            "Cross type": "cross_type",
            "Cross clear": "cross_clear",
            "Cross number": "cross_number",
            "Brand": "brand"
        }
    )

    # Clean text values
    for column in df.columns:

        df[column] = (
            df[column]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    # Create our own normalized search key.
    #
    # We do NOT rely only on Excel Cross clear.
    # This makes the database more robust.
    df["search_number"] = (
        df["cross_number"]
        .apply(normalize_number)
    )

    # If Cross number is empty,
    # use Cross clear as fallback
    empty_search = (
        df["search_number"] == ""
    )

    df.loc[
        empty_search,
        "search_number"
    ] = (
        df.loc[
            empty_search,
            "cross_clear"
        ]
        .apply(normalize_number)
    )

    # Add source
    df["source_type"] = source_type

    # Remove rows without searchable number
    df = df[
        df["search_number"] != ""
    ].copy()

    print(
        f"{sheet_name}: "
        f"{len(df):,} rows prepared."
    )

    return df


# ============================================================
# BUILD DATABASE
# ============================================================

def build_database():

    excel_path = Path(EXCEL_FILE)

    if not excel_path.exists():

        raise FileNotFoundError(
            f"\nFile not found: {EXCEL_FILE}\n"
            f"Put the Excel file in the same folder "
            f"as build_database.py."
        )

    print()
    print("=" * 60)
    print("LYNX DATABASE BUILDER")
    print("=" * 60)
    print()

    # --------------------------------------------------------
    # LOAD AFTERMARKET
    # --------------------------------------------------------

    aftermarket = load_sheet(
        excel_path,
        SHEETS["Aftermarket"],
        "Aftermarket"
    )

    # --------------------------------------------------------
    # LOAD OEM
    # --------------------------------------------------------

    oem = load_sheet(
        excel_path,
        SHEETS["OEM"],
        "OEM"
    )

    # --------------------------------------------------------
    # COMBINE
    # --------------------------------------------------------

    print()
    print("Combining databases...")

    combined = pd.concat(
        [
            aftermarket,
            oem
        ],
        ignore_index=True
    )

    # Remove exact duplicates
    combined = combined.drop_duplicates(
        subset=[
            "search_number",
            "lynx_number",
            "source_type",
            "cross_number",
            "brand"
        ]
    )

    print(
        f"Total rows after cleanup: "
        f"{len(combined):,}"
    )

    # --------------------------------------------------------
    # CREATE DUCKDB
    # --------------------------------------------------------

    database_path = Path(
        DATABASE_FILE
    )

    # Delete old database before rebuild
    if database_path.exists():
        database_path.unlink()

    print()
    print("Creating DuckDB database...")

    connection = duckdb.connect(
        DATABASE_FILE
    )

    connection.register(
        "combined_df",
        combined
    )

    connection.execute(
        """
        CREATE TABLE crosses AS

        SELECT
            search_number,
            lynx_number,
            description,
            source_type,
            cross_type,
            cross_number,
            brand,
            groups

        FROM combined_df
        """
    )

    # --------------------------------------------------------
    # INDEX
    # --------------------------------------------------------

    print(
        "Creating search index..."
    )

    connection.execute(
        """
        CREATE INDEX idx_search_number
        ON crosses(search_number)
        """
    )

    # --------------------------------------------------------
    # CHECK DATABASE
    # --------------------------------------------------------

    total_rows = connection.execute(
        """
        SELECT COUNT(*)
        FROM crosses
        """
    ).fetchone()[0]

    oem_rows = connection.execute(
        """
        SELECT COUNT(*)
        FROM crosses
        WHERE source_type = 'OEM'
        """
    ).fetchone()[0]

    aftermarket_rows = connection.execute(
        """
        SELECT COUNT(*)
        FROM crosses
        WHERE source_type = 'Aftermarket'
        """
    ).fetchone()[0]

    connection.close()

    # --------------------------------------------------------
    # FINISHED
    # --------------------------------------------------------

    print()
    print("=" * 60)
    print("DATABASE CREATED SUCCESSFULLY")
    print("=" * 60)

    print(
        f"Database: {DATABASE_FILE}"
    )

    print(
        f"Total rows: {total_rows:,}"
    )

    print(
        f"OEM rows: {oem_rows:,}"
    )

    print(
        f"Aftermarket rows: {aftermarket_rows:,}"
    )

    print("=" * 60)


# ============================================================
# START
# ============================================================

if __name__ == "__main__":
    build_database()