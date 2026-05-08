import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font

st.set_page_config(page_title="Seller Converter", layout="wide")

st.title("Diamond Seller Format Converter")

uploaded_file = st.file_uploader(
    "Upload Seller Excel File",
    type=["xlsx", "xls"]
)

# ================= CUSTOM SR NO OPTION =================

use_custom_sr = st.checkbox(
    "Use Custom SR Number Start"
)

# DEFAULT VALUES
cvd_start = 1
hpht_start = 1

# CUSTOM VALUES
if use_custom_sr:

    cvd_start = st.number_input(
        "CVD Start Number",
        min_value=1,
        value=1,
        step=1
    )

    hpht_start = st.number_input(
        "HPHT Start Number",
        min_value=1,
        value=1,
        step=1
    )

# ================= COMPANY OUTPUT FORMAT =================

final_columns = [
    "VENDOR",
    "SR NO.",
    "SHAPE",
    "COLOR",
    "CLARITY",
    "CARATS",
    "CERT#",
    "AMT/CTS $",
    "TOTAL AMT $",
    "$ RATE",
    "AMT/CTS RS",
    "TOTAL AMT RS",
    "BRAND",
    "ES CODE#",
    "ACTUAL SHAPE",
    "VIDEO"
]

# ================= SHAPE SHORT NAME =================

shape_mapping = {
    "EM": "EMERALD",
    "LR": "RADIANT",
    "CMB": "CUSHION MODIFIED",
    "TRI": "TRILLIANT"
}

# ================= AUTO COLUMN FINDER =================

def find_column(df, keywords):

    for col in df.columns:

        col_name = str(col).strip().lower()

        for key in keywords:

            if key.lower() in col_name:
                return col

    return None

# ================= BRAND FINDER =================

def detect_brand(row):

    text = " ".join(
        [str(x) for x in row.values]
    ).upper()

    # HPHT CHECK
    if (
        "HPHT" in text or
        "HIGH PRESSURE HIGH TEMPERATURE" in text
    ):
        return "HPHT"

    # CVD CHECK
    if (
        "CVD" in text or
        "CHEMICAL VAPOR DEPOSITION" in text
    ):
        return "CVD"

    # DEFAULT
    return ""

# ================= MAIN =================

if uploaded_file:

    try:

        # ================= READ FILE =================

        if uploaded_file.name.endswith(".xls"):

            temp_df = pd.read_excel(
                uploaded_file,
                engine="xlrd",
                header=None
            )

        else:

            temp_df = pd.read_excel(
                uploaded_file,
                header=None
            )

        # ================= AUTO FIND HEADER ROW =================

        header_row = 0

        possible_headers = [
            "company",
            "vendor",
            "seller",
            "shape",
            "shp",
            "color",
            "col",
            "clr",
            "clarity",
            "cl",
            "carat",
            "cts",
            "stone no",
            "cert"
        ]

        for i in range(min(20, len(temp_df))):

            row_text = " ".join(
                [str(x).lower() for x in temp_df.iloc[i].values]
            )

            if any(
                header in row_text
                for header in possible_headers
            ):
                header_row = i
                break

        # ================= RE-READ USING DETECTED HEADER =================

        uploaded_file.seek(0)

        if uploaded_file.name.endswith(".xls"):

            df = pd.read_excel(
                uploaded_file,
                engine="xlrd",
                header=header_row
            )

        else:

            df = pd.read_excel(
                uploaded_file,
                header=header_row
            )

        # ================= CLEAN COLUMNS =================

        df.columns = [str(c).strip() for c in df.columns]

        st.success("File Uploaded Successfully ✅")

        # ================= AUTO DETECT COLUMNS =================

        vendor_col = find_column(df, [
            "company",
            "vendor",
            "seller",
            "supplier",
            "owner",
            "party"
        ])

        shape_col = find_column(df, [
            "shape",
            "shp",
            "cut",
            "diamond shape",
            "stone shape"
        ])

        color_col = find_column(df, [
            "color",
            "col",
            "clr",
            "diamond color",
            "stone color"
        ])

        clarity_col = find_column(df, [
            "clarity",
            "cla",
            "cl",
            "diamond clarity",
            "stone clarity"
        ])

        cts_col = find_column(df, [
            "cts",
            "cts.",
            "ct",
            "cts weight",
            "carat",
            "carats",
            "weight",
            "size",
            "stone size",
            "diamond weight"
        ])

        # ================= CREATE OUTPUT =================

        output = pd.DataFrame()

        # ================= CERT# AUTO FIND =================

        cert_values = []
        valid_rows = []

        for index, row in df.iterrows():

            cert_found = ""

            for value in row:

                value = str(value).replace(".0", "").strip()

                # CHECK 9 DIGIT NUMBER
                if value.isdigit() and len(value) == 9:

                    cert_found = "LG" + value
                    break

            # KEEP ONLY VALID CERT ROWS
            if cert_found != "":

                cert_values.append(cert_found)
                valid_rows.append(index)

        # FILTER ONLY VALID ROWS
        df = df.loc[valid_rows].reset_index(drop=True)

        # ================= AUTO DETECT COLUMNS =================

        vendor_col = find_column(df, [
            "company",
            "vendor",
            "seller",
            "supplier",
            "owner",
            "party"
        ])

        shape_col = find_column(df, [
            "shape",
            "shp",
            "cut",
            "diamond shape",
            "stone shape"
        ])

        color_col = find_column(df, [
            "color",
            "col",
            "clr",
            "diamond color",
            "stone color"
        ])

        clarity_col = find_column(df, [
            "clarity",
            "cla",
            "cl",
            "diamond clarity",
            "stone clarity"
        ])

        cts_col = find_column(df, [
            "cts",
            "cts.",
            "ct",
            "cts weight",
            "carat",
            "carats",
            "weight",
            "size",
            "stone size",
            "diamond weight"
        ])

        # CERT#
        output["CERT#"] = cert_values

        # VENDOR
        if vendor_col:
            output["VENDOR"] = df[vendor_col].reset_index(drop=True)
        else:
            output["VENDOR"] = "GOLDEN"

        # BRAND
        output["BRAND"] = df.apply(
            detect_brand,
            axis=1
        ).reset_index(drop=True)

        # SHAPE
        if shape_col:
            output["SHAPE"] = df[shape_col].reset_index(drop=True)
        else:
            output["SHAPE"] = ""

        output["SHAPE"] = output["SHAPE"].replace(shape_mapping)

        # COLOR
        if color_col:
            output["COLOR"] = df[color_col].reset_index(drop=True)
        else:
            output["COLOR"] = ""

        # CLARITY
        if clarity_col:
            output["CLARITY"] = df[clarity_col].reset_index(drop=True)
        else:
            output["CLARITY"] = ""

        # CARATS
        if cts_col:
            output["CARATS"] = df[cts_col].reset_index(drop=True)
        else:
            output["CARATS"] = ""

        # ================= EMPTY COLUMNS =================

        output["AMT/CTS $"] = ""
        output["TOTAL AMT $"] = ""
        output["$ RATE"] = ""
        output["AMT/CTS RS"] = ""
        output["TOTAL AMT RS"] = ""
        output["ES CODE#"] = ""
        output["ACTUAL SHAPE"] = ""
        output["VIDEO"] = ""


        # REMOVE BLANK BRAND ROWS
        output = output[
            output["BRAND"] != ""
        ].reset_index(drop=True)
        # ================= CREATE SEPARATE FILES =================

        cvd_df = output[
            output["BRAND"] == "CVD"
        ].reset_index(drop=True)

        hpht_df = output[
            output["BRAND"] == "HPHT"
        ].reset_index(drop=True)

        # ================= CVD SR NO =================

        cvd_df["SR NO."] = [
            f"C{i}"
            for i in range(
                cvd_start,
                cvd_start + len(cvd_df)
            )
        ]

        # ================= HPHT SR NO =================

        hpht_df["SR NO."] = [
            f"H{i}"
            for i in range(
                hpht_start,
                hpht_start + len(hpht_df)
            )
        ]

        # ================= MERGE FINAL =================

        output = pd.concat(
            [cvd_df, hpht_df],
            ignore_index=True
        )

        # ================= FINAL FORMAT =================

        output = output[final_columns]

        cvd_df = cvd_df[final_columns]
        hpht_df = hpht_df[final_columns]

        # ================= SHOW DATA =================

        st.subheader("Converted Data")

        st.dataframe(
            output,
            use_container_width=True
        )

        # ================= COUNTS =================

        st.markdown(
            f"### Total Diamonds: {len(output)}"
        )

        st.markdown(
            f"### CVD Diamonds: {len(cvd_df)}"
        )

        st.markdown(
            f"### HPHT Diamonds: {len(hpht_df)}"
        )

        # ================= EXCEL DOWNLOAD FUNCTION =================

        def create_excel(dataframe):

            buffer = BytesIO()

            with pd.ExcelWriter(
                buffer,
                engine="openpyxl"
            ) as writer:

                dataframe.to_excel(
                    writer,
                    index=False,
                    sheet_name="Output"
                )

                worksheet = writer.sheets["Output"]

                # BOLD HEADER
                for cell in worksheet[1]:
                    cell.font = Font(bold=True)

            buffer.seek(0)

            return buffer

        # ================= DOWNLOAD BUTTONS =================

        st.download_button(
            label="Download Full File",
            data=create_excel(output),
            file_name="Final_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            label="Download CVD File",
            data=create_excel(cvd_df),
            file_name="CVD_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            label="Download HPHT File",
            data=create_excel(hpht_df),
            file_name="HPHT_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:

        if str(e) != "None":
            st.error(f"Error: {e}")
