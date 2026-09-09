import re
from io import BytesIO
from pathlib import Path

import duckdb
import pandas as pd
import streamlit as st


# ============================================================
# PAGE SETTINGS
# ============================================================

st.set_page_config(
    page_title="LYNX Cross Search",
    page_icon="🔎",
    layout="wide",
    initial_sidebar_state="collapsed"
)

DATABASE_FILE = "lynx_cross.duckdb"


# ============================================================
# CSS
# ============================================================

st.html(
    """
    <style>

    .stApp {
        background: #f3f3f3;
    }

    .block-container {
        max-width: 1500px;
        padding-top: 0rem;
        padding-bottom: 2rem;
    }

    #MainMenu {
        visibility: hidden;
    }

    footer {
        visibility: hidden;
    }

    header[data-testid="stHeader"] {
        background: transparent;
    }

    html, body {
        font-family: Arial, Helvetica, sans-serif;
    }


    /* ======================================================
       HEADER
    ====================================================== */

    .lynx-header {
        background: linear-gradient(
            90deg,
            #3d3d3a 0%,
            #4d4d48 100%
        );

        min-height: 76px;

        display: flex;
        align-items: center;
        justify-content: space-between;

        padding: 0 28px;

        margin-left: -1rem;
        margin-right: -1rem;

        border-bottom: 3px solid #ef433d;
    }

    .lynx-header-left {
        display: flex;
        align-items: center;
        gap: 22px;
        min-width: 0;
    }

    .lynx-logo {
        height: 48px;
        width: auto;
        object-fit: contain;
        flex-shrink: 0;
    }

    .lynx-divider {
        width: 1px;
        height: 36px;
        background: #73736e;
        flex-shrink: 0;
    }

    .lynx-app-name {
        color: white;
        font-size: 16px;
        font-weight: 700;
        letter-spacing: 1.8px;
        white-space: nowrap;
    }

    .lynx-header-right {
        display: flex;
        align-items: center;
        gap: 14px;
        flex-shrink: 0;
    }

    .database-status {
        color: #e6e6e6;
        font-size: 12px;
        font-weight: 600;
        letter-spacing: .4px;
        white-space: nowrap;
    }

    .database-dot {
        display: inline-block;
        width: 8px;
        height: 8px;
        background: #69b96b;
        border-radius: 50%;
        margin-right: 7px;
    }

    .database-count {
        background: #ef433d;
        color: white;
        padding: 8px 14px;
        border-radius: 4px;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: .5px;
        white-space: nowrap;
    }


    /* ======================================================
       HERO
    ====================================================== */

    .lynx-hero {
        background: linear-gradient(
            110deg,
            #292927 0%,
            #3f3f3c 64%,
            #4a4a46 100%
        );

        padding: 32px 44px 34px 44px;

        margin-left: -1rem;
        margin-right: -1rem;

        position: relative;
        overflow: hidden;
    }

    .lynx-hero::after {
        content: "";
        position: absolute;
        width: 280px;
        height: 280px;
        right: -80px;
        top: -145px;
        border: 1px solid rgba(255,255,255,.06);
        border-radius: 50%;
    }

    .hero-small {
        color: #bcbcb9;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 2.2px;
        margin-bottom: 10px;
    }

    .hero-title {
        color: white;
        font-size: 34px;
        line-height: 1.08;
        font-weight: 700;
        margin-bottom: 10px;
    }

    .hero-title-red {
        color: #ef433d;
    }

    .hero-description {
        color: #d3d3d0;
        font-size: 15px;
        max-width: 760px;
        line-height: 1.45;
    }

    .hero-line {
        background: #ef433d;
        width: 150px;
        height: 3px;
        margin-top: 20px;
    }


    /* ======================================================
       WORKSPACE
    ====================================================== */

    .workspace-title {
        margin-top: 24px;
        margin-bottom: 4px;
        color: #333330;
        font-size: 23px;
        font-weight: 700;
    }

    .workspace-subtitle {
        color: #777773;
        font-size: 13px;
        margin-bottom: 14px;
    }


    /* ======================================================
       INPUTS
    ====================================================== */

    label {
        color: #454541 !important;
        font-weight: 600 !important;
    }

    .stTextArea textarea {
        background: white !important;
        color: #252525 !important;
        border: 1px solid #d0d0cc !important;
        border-radius: 4px !important;
        font-size: 14px !important;
        padding: 12px !important;
    }

    .stTextArea textarea:focus {
        border-color: #ef433d !important;
        box-shadow: 0 0 0 1px #ef433d !important;
    }

    div[data-baseweb="select"] > div {
        background: white !important;
        border-color: #d0d0cc !important;
        color: #333330 !important;
        border-radius: 4px !important;
    }

    .stTextInput input {
        background: white !important;
        color: #252525 !important;
        border: 1px solid #d0d0cc !important;
        border-radius: 4px !important;
    }

    .stTextInput input:focus {
        border-color: #ef433d !important;
        box-shadow: 0 0 0 1px #ef433d !important;
    }


    /* ======================================================
       RADIO
    ====================================================== */

    div[role="radiogroup"] {
        gap: 10px;
    }

    div[role="radiogroup"] label {
        background: white;
        border: 1px solid #d1d1cd;
        border-radius: 4px;
        padding: 8px 16px;
        transition: .15s ease;
    }

    div[role="radiogroup"] label:hover {
        border-color: #ef433d;
    }


    /* ======================================================
       BUTTONS
    ====================================================== */

    .stButton > button {
        background: linear-gradient(
            90deg,
            #ef433d,
            #f05a50
        ) !important;

        color: white !important;
        border: none !important;
        border-radius: 4px !important;

        min-height: 46px;

        font-size: 13px;
        font-weight: 700;
        letter-spacing: .8px;

        transition: .15s ease;
    }

    .stButton > button:hover {
        background: #dc3833 !important;
        transform: translateY(-1px);
        box-shadow: 0 5px 12px rgba(0,0,0,.12);
    }

    .stDownloadButton > button {
        background: #454541 !important;
        color: white !important;
        border: none !important;
        border-radius: 4px !important;
        min-height: 42px;
        font-weight: 700;
        font-size: 12px;
        letter-spacing: .6px;
    }

    .stDownloadButton > button:hover {
        background: #ef433d !important;
        color: white !important;
    }


    /* ======================================================
       SEARCH TIPS
    ====================================================== */

    .search-tips {
        background: #454541;
        border-left: 4px solid #ef433d;
        color: #d7d7d4;
        padding: 14px 16px;
        border-radius: 3px;
        margin-top: 14px;
        font-size: 12px;
        line-height: 1.55;
    }

    .search-tips-title {
        color: white;
        font-size: 13px;
        font-weight: 700;
        margin-bottom: 6px;
    }


    /* ======================================================
       METRICS
    ====================================================== */

    div[data-testid="stMetric"] {
        background: white;
        border: 1px solid #dededb;
        border-top: 3px solid #ef433d;
        border-radius: 4px;
        padding: 14px 18px;
        box-shadow: 0 2px 7px rgba(0,0,0,.04);
    }

    div[data-testid="stMetricLabel"] {
        color: #777773 !important;
        font-size: 12px;
    }

    div[data-testid="stMetricValue"] {
        color: #333330 !important;
        font-weight: 700;
    }


    /* ======================================================
       DATAFRAME
    ====================================================== */

    div[data-testid="stDataFrame"] {
        background: white;
        border: 1px solid #d9d9d5;
        border-radius: 4px;
        overflow: hidden;
        box-shadow: 0 2px 7px rgba(0,0,0,.04);
    }


    /* ======================================================
       RESULTS TOOLBAR
    ====================================================== */

    .results-toolbar-title {
        color: #777773;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 1px;
        margin-top: 5px;
        margin-bottom: 2px;
    }

    .results-count {
        color: #777773;
        font-size: 12px;
        margin-top: 5px;
        margin-bottom: 8px;
    }


    /* ======================================================
       FILE UPLOADER
    ====================================================== */

    section[data-testid="stFileUploaderDropzone"] {
        background: white !important;
        border: 1px dashed #bdbdb8 !important;
        border-radius: 4px !important;
    }


    /* ======================================================
       UPLOAD INFO
    ====================================================== */

    .upload-ready {
        background: white;
        border-left: 4px solid #ef433d;
        border-top: 1px solid #ddd;
        border-right: 1px solid #ddd;
        border-bottom: 1px solid #ddd;
        padding: 11px 15px;
        border-radius: 3px;
        color: #454541;
        margin-top: 10px;
        margin-bottom: 12px;
        font-size: 13px;
    }


    /* ======================================================
       SECTION TITLE
    ====================================================== */

    .section-title {
        display: flex;
        align-items: center;
        gap: 10px;
        margin-top: 24px;
        margin-bottom: 12px;
        color: #333330;
        font-size: 20px;
        font-weight: 700;
    }

    .section-red-line {
        width: 4px;
        height: 23px;
        background: #ef433d;
        border-radius: 2px;
    }


    /* ======================================================
       FOOTER
    ====================================================== */

    .lynx-footer {
        margin-top: 32px;
        padding: 18px 0;
        border-top: 1px solid #d5d5d1;
        color: #8a8a86;
        text-align: center;
        font-size: 11px;
    }


    /* ======================================================
       RESPONSIVE
    ====================================================== */

    @media (max-width: 900px) {

        .lynx-header {
            padding: 12px 18px;
            flex-wrap: wrap;
            gap: 12px;
        }

        .lynx-header-right {
            gap: 8px;
        }

        .database-count {
            display: none;
        }

        .hero-title {
            font-size: 30px;
        }
    }

    </style>
    """
)


# ============================================================
# NORMALIZE NUMBER
# ============================================================

def normalize_number(number):

    if number is None:
        return ""

    text = str(number).strip()

    if not text or text.lower() == "nan":
        return ""

    return re.sub(
        r"[^A-Za-z0-9]",
        "",
        text
    ).upper()


# ============================================================
# SPLIT NUMBERS
# ============================================================

def split_numbers(text):

    if text is None:
        return []

    text = str(text).strip()

    if not text or text.lower() == "nan":
        return []

    parts = re.split(
        r"[,;/\n\r]+",
        text
    )

    result = []

    for part in parts:

        part = part.strip()

        if part:
            result.append(part)

    return result


# ============================================================
# FORMAT CUSTOMER NUMBERS
# ============================================================

def format_customer_numbers(numbers):

    return " / ".join(numbers)


# ============================================================
# PREPARE CUSTOMER ROWS
# ============================================================

def prepare_customer_rows(values):

    prepared = []

    for index, value in enumerate(values):

        if pd.isna(value):
            raw_value = ""
        else:
            raw_value = str(value).strip()

        prepared.append({
            "row_no": index + 1,
            "raw_value": raw_value
        })

    return prepared


# ============================================================
# BUILD SEARCH TABLE
# ============================================================

def build_search_table(prepared_rows):

    search_rows = []
    original_rows = []

    for item in prepared_rows:

        row_no = item["row_no"]
        raw_value = item["raw_value"]

        numbers = split_numbers(
            raw_value
        )

        customer_display = (
            format_customer_numbers(numbers)
            if numbers
            else raw_value
        )

        original_rows.append({
            "row_no": row_no,
            "customer_display": customer_display,
            "has_numbers": bool(numbers)
        })

        for original_number in numbers:

            search_number = normalize_number(
                original_number
            )

            if search_number:

                search_rows.append({
                    "row_no": row_no,
                    "customer_display": customer_display,
                    "search_number": search_number
                })

    search_df = pd.DataFrame(
        search_rows,
        columns=[
            "row_no",
            "customer_display",
            "search_number"
        ]
    )

    original_df = pd.DataFrame(
        original_rows,
        columns=[
            "row_no",
            "customer_display",
            "has_numbers"
        ]
    )

    return search_df, original_df


# ============================================================
# SEARCH DATABASE
# ============================================================

def search_crosses(
    prepared_rows,
    source_filter
):

    search_df, original_df = build_search_table(
        prepared_rows
    )

    connection = duckdb.connect(
        DATABASE_FILE,
        read_only=True
    )

    try:

        if search_df.empty:

            matches = pd.DataFrame(
                columns=[
                    "row_no",
                    "customer_display",
                    "lynx_number",
                    "description",
                    "source_type"
                ]
            )

        else:

            connection.register(
                "search_inputs",
                search_df
            )

            if source_filter == "All":

                query = """
                    SELECT
                        s.row_no,
                        s.customer_display,
                        c.lynx_number,
                        c.description,
                        c.source_type
                    FROM search_inputs s
                    INNER JOIN crosses c
                        ON c.search_number = s.search_number
                    ORDER BY
                        s.row_no,
                        c.lynx_number
                """

                matches = connection.execute(
                    query
                ).fetchdf()

            else:

                query = """
                    SELECT
                        s.row_no,
                        s.customer_display,
                        c.lynx_number,
                        c.description,
                        c.source_type
                    FROM search_inputs s
                    INNER JOIN crosses c
                        ON c.search_number = s.search_number
                    WHERE c.source_type = ?
                    ORDER BY
                        s.row_no,
                        c.lynx_number
                """

                matches = connection.execute(
                    query,
                    [source_filter]
                ).fetchdf()

    finally:

        connection.close()


    results = []


    for _, original_row in original_df.iterrows():

        row_no = int(
            original_row["row_no"]
        )

        customer_display = (
            original_row["customer_display"]
        )

        has_numbers = bool(
            original_row["has_numbers"]
        )


        # EMPTY ROW
        if not has_numbers:

            results.append({
                "No.": row_no,
                "Customer Part Number": customer_display,
                "LYNXauto Number": "",
                "Description": "",
                "Cross Type": "",
                "Status": ""
            })

            continue


        # MATCHES
        if matches.empty:

            row_matches = pd.DataFrame()

        else:

            row_matches = matches[
                matches["row_no"] == row_no
            ].copy()


        # NOT FOUND
        if row_matches.empty:

            results.append({
                "No.": row_no,
                "Customer Part Number": customer_display,
                "LYNXauto Number": "",
                "Description": "",
                "Cross Type": "",
                "Status": "Not Found"
            })

            continue


        # REMOVE DUPLICATES
        row_matches = row_matches.drop_duplicates(
            subset=[
                "lynx_number",
                "description",
                "source_type"
            ]
        )


        # FOUND
        for _, match in row_matches.iterrows():

            results.append({
                "No.":
                    row_no,

                "Customer Part Number":
                    customer_display,

                "LYNXauto Number":
                    str(match["lynx_number"]).strip(),

                "Description":
                    str(match["description"]).strip(),

                "Cross Type":
                    str(match["source_type"]).strip(),

                "Status":
                    "Found"
            })


    return pd.DataFrame(
        results
    )


# ============================================================
# EXCEL EXPORT
# ============================================================

def create_excel_download(df):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            sheet_name="Cross Results",
            index=False
        )

        worksheet = writer.book[
            "Cross Results"
        ]

        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        worksheet.column_dimensions["A"].width = 9
        worksheet.column_dimensions["B"].width = 48
        worksheet.column_dimensions["C"].width = 22
        worksheet.column_dimensions["D"].width = 36
        worksheet.column_dimensions["E"].width = 18
        worksheet.column_dimensions["F"].width = 16

    output.seek(0)

    return output.getvalue()


# ============================================================
# RESULT FILTER
# ============================================================

def filter_results(
    df,
    status_filter,
    search_text
):

    filtered = df.copy()


    if status_filter != "All":

        filtered = filtered[
            filtered["Status"] == status_filter
        ]


    search_text = str(
        search_text
    ).strip().lower()


    if search_text:

        searchable_columns = [
            "Customer Part Number",
            "LYNXauto Number",
            "Description",
            "Cross Type"
        ]

        mask = pd.Series(
            False,
            index=filtered.index
        )

        for column in searchable_columns:

            mask = mask | (
                filtered[column]
                .fillna("")
                .astype(str)
                .str.lower()
                .str.contains(
                    search_text,
                    regex=False
                )
            )

        filtered = filtered[
            mask
        ]


    return filtered


# ============================================================
# TABLE STYLE
# ============================================================

def style_results_table(df):

    styled = df.style


    def status_style(value):

        if value == "Found":

            return (
                "background-color:#eaf6ec;"
                "color:#237a37;"
                "font-weight:700;"
            )

        if value == "Not Found":

            return (
                "background-color:#fdecec;"
                "color:#c53932;"
                "font-weight:700;"
            )

        return ""


    def cross_type_style(value):

        if value == "OEM":

            return (
                "background-color:#eeeeec;"
                "color:#454541;"
                "font-weight:700;"
            )

        if value == "Aftermarket":

            return (
                "background-color:#fff0ee;"
                "color:#c7443e;"
                "font-weight:700;"
            )

        return ""


    styled = styled.map(
        status_style,
        subset=["Status"]
    )

    styled = styled.map(
        cross_type_style,
        subset=["Cross Type"]
    )

    styled = styled.map(
        lambda value: (
            "font-weight:700;"
            "color:#252525;"
            if str(value).strip()
            else ""
        ),
        subset=["LYNXauto Number"]
    )

    return styled


# ============================================================
# DATABASE CHECK
# ============================================================

database_path = Path(
    DATABASE_FILE
)

if not database_path.exists():

    st.error(
        "Database file lynx_cross.duckdb was not found."
    )

    st.stop()


# ============================================================
# HEADER
# ============================================================

st.html(
    """
    <div class="lynx-header">

        <div class="lynx-header-left">

            <img
                class="lynx-logo"
                src="https://lynxauto.jp/assets/img/logo.png"
            >

            <div class="lynx-divider"></div>

            <div class="lynx-app-name">
                CROSS SEARCH
            </div>

        </div>

        <div class="lynx-header-right">

            <div class="database-status">
                <span class="database-dot"></span>
                DATABASE ONLINE
            </div>

            <div class="database-count">
                1.37M REFERENCES
            </div>

        </div>

    </div>

    <div class="lynx-hero">

        <div class="hero-small">
            LYNXAUTO PRODUCT DATABASE
        </div>

        <div class="hero-title">
            Find Your
            <span class="hero-title-red">
                Cross Reference
            </span>
        </div>

        <div class="hero-description">
            Fast search across OEM and Aftermarket cross references
            in the LYNXauto product database.
        </div>

        <div class="hero-line"></div>

    </div>
    """
)


# ============================================================
# WORKSPACE
# ============================================================

st.html(
    """
    <div class="workspace-title">
        Cross Reference Search
    </div>

    <div class="workspace-subtitle">
        Enter customer part numbers manually or upload an Excel customer request.
    </div>
    """
)


# ============================================================
# INPUT MODE
# ============================================================

input_mode = st.radio(
    "Input method",
    [
        "Manual Search",
        "Upload Excel"
    ],
    horizontal=True,
    label_visibility="collapsed"
)


prepared_rows = []


# ============================================================
# MANUAL SEARCH
# ============================================================

if input_mode == "Manual Search":

    left, right = st.columns(
        [3.4, 1]
    )


    with left:

        customer_input = st.text_area(
            "Customer Part Numbers",
            value="L42508\n96337133",
            height=175,
            placeholder=(
                "Enter one customer position per line..."
            ),
            help=(
                "Several part numbers inside one customer "
                "position can be separated by commas, "
                "semicolons or slashes."
            )
        )

        manual_values = (
            customer_input.splitlines()
        )

        prepared_rows = prepare_customer_rows(
            manual_values
        )


    with right:

        source_filter = st.selectbox(
            "Cross Type",
            [
                "All",
                "OEM",
                "Aftermarket"
            ]
        )

        st.html(
            """
            <div class="search-tips">

                <div class="search-tips-title">
                    SEARCH TIPS
                </div>

                One customer position per line.<br>
                Multiple references can stay in the same line.<br>
                Spaces and hyphens are normalized automatically.

            </div>
            """
        )


# ============================================================
# EXCEL UPLOAD
# ============================================================

else:

    source_filter = st.selectbox(
        "Cross Type",
        [
            "All",
            "OEM",
            "Aftermarket"
        ]
    )

    uploaded_file = st.file_uploader(
        "Upload Customer Excel File",
        type=["xlsx"]
    )


    if uploaded_file is not None:

        try:

            excel_file = pd.ExcelFile(
                uploaded_file,
                engine="openpyxl"
            )

            col1, col2 = st.columns(2)


            with col1:

                selected_sheet = st.selectbox(
                    "1. Select Sheet",
                    excel_file.sheet_names
                )


            customer_excel = pd.read_excel(
                excel_file,
                sheet_name=selected_sheet,
                dtype=str
            )


            customer_excel.columns = [
                str(column).strip()
                for column in customer_excel.columns
            ]


            if customer_excel.empty:

                st.warning(
                    "The selected sheet is empty."
                )


            else:

                with col2:

                    selected_column = st.selectbox(
                        "2. Customer Part Number Column",
                        customer_excel.columns.tolist()
                    )


                selected_values = (
                    customer_excel[
                        selected_column
                    ].tolist()
                )


                prepared_rows = prepare_customer_rows(
                    selected_values
                )


                non_empty_count = sum(
                    1
                    for item in prepared_rows
                    if item["raw_value"]
                )


                st.html(
                    f"""
                    <div class="upload-ready">
                        <strong>{non_empty_count:,}</strong>
                        customer positions ready for search
                    </div>
                    """
                )


                st.dataframe(
                    customer_excel[
                        [selected_column]
                    ].head(8),
                    width="stretch",
                    hide_index=True
                )


        except Exception as error:

            st.error(
                f"Unable to read Excel file: {error}"
            )


# ============================================================
# SEARCH BUTTON
# ============================================================

st.write("")


search_button = st.button(
    "SEARCH CROSS REFERENCES",
    type="primary",
    width="stretch"
)


# ============================================================
# STORE RESULTS
# ============================================================

if "search_results" not in st.session_state:

    st.session_state.search_results = None


if search_button:

    if not prepared_rows:

        st.warning(
            "No customer data is available for search."
        )

    else:

        with st.spinner(
            "Searching LYNXauto database..."
        ):

            st.session_state.search_results = (
                search_crosses(
                    prepared_rows,
                    source_filter
                )
            )


# ============================================================
# RESULTS
# ============================================================

if st.session_state.search_results is not None:

    df = st.session_state.search_results


    found_count = len(
        df[
            df["Status"] == "Found"
        ]
    )


    not_found_count = len(
        df[
            df["Status"] == "Not Found"
        ]
    )


    customer_rows_count = (
        df[
            df["Status"] != ""
        ]["No."]
        .nunique()
    )


    # ========================================================
    # TITLE
    # ========================================================

    st.html(
        """
        <div class="section-title">
            <div class="section-red-line"></div>
            Search Results
        </div>
        """
    )


    # ========================================================
    # METRICS
    # ========================================================

    m1, m2, m3 = st.columns(3)


    m1.metric(
        "Customer Positions",
        customer_rows_count
    )


    m2.metric(
        "Crosses Found",
        found_count
    )


    m3.metric(
        "Not Found",
        not_found_count
    )


    # ========================================================
    # TOOLBAR
    # ========================================================

    filter_col, search_col, download_col = st.columns(
        [1.1, 2.4, 1]
    )


    with filter_col:

        result_status_filter = st.selectbox(
            "Result Status",
            [
                "All",
                "Found",
                "Not Found"
            ],
            key="result_status_filter"
        )


    with search_col:

        result_search = st.text_input(
            "Search Results",
            placeholder=(
                "Customer PN, LYNX number, description..."
            ),
            key="result_search"
        )


    # ========================================================
    # FILTER DISPLAY DATA
    # ========================================================

    display_df = filter_results(
        df,
        result_status_filter,
        result_search
    )


    # ========================================================
    # DOWNLOAD FULL RESULTS
    # ========================================================

    excel_data = create_excel_download(
        df
    )


    with download_col:

        st.write("")

        st.download_button(
            label="DOWNLOAD XLSX",
            data=excel_data,
            file_name="LYNX_Cross_Results.xlsx",
            mime=(
                "application/"
                "vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            width="stretch"
        )


    # ========================================================
    # RESULT COUNT
    # ========================================================

    st.html(
        f"""
        <div class="results-count">
            Showing <strong>{len(display_df):,}</strong>
            of <strong>{len(df):,}</strong> result rows
        </div>
        """
    )


    # ========================================================
    # DYNAMIC TABLE HEIGHT
    # ========================================================

    table_height = min(
        max(
            180,
            42 + len(display_df) * 35
        ),
        520
    )


    # ========================================================
    # TABLE
    # ========================================================

    styled_df = style_results_table(
        display_df
    )


    st.dataframe(
        styled_df,
        width="stretch",
        height=table_height,
        hide_index=True,

        column_order=[
            "No.",
            "Customer Part Number",
            "LYNXauto Number",
            "Description",
            "Cross Type",
            "Status"
        ],

        column_config={

            "No.":
                st.column_config.NumberColumn(
                    "No.",
                    width="small"
                ),

            "Customer Part Number":
                st.column_config.TextColumn(
                    "Customer Part Number",
                    width="large"
                ),

            "LYNXauto Number":
                st.column_config.TextColumn(
                    "LYNXauto Number",
                    width="medium"
                ),

            "Description":
                st.column_config.TextColumn(
                    "Description",
                    width="large"
                ),

            "Cross Type":
                st.column_config.TextColumn(
                    "Cross Type",
                    width="small"
                ),

            "Status":
                st.column_config.TextColumn(
                    "Status",
                    width="small"
                )
        }
    )


# ============================================================
# FOOTER
# ============================================================

st.html(
    """
    <div class="lynx-footer">
        LYNX Cross Search • Internal Product Tool
    </div>
    """
)