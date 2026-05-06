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

    text = " ".join([str(x) for x in row.values]).upper()

    if "HPHT" in text:
        return "HPHT"

    return "CVD"

# ================= MAIN =================

if uploaded_file:

    try:

        # ================= READ FILE =================

        if uploaded_file.name.endswith(".xls"):
            df = pd.read_excel(uploaded_file, engine="xlrd")

        else:
            df = pd.read_excel(uploaded_file)

        # ================= FIX SOME SELLER FILES =================

        if all("Unnamed" in str(c) for c in df.columns):
            df = pd.read_excel(uploaded_file, header=5)

        # ================= CLEAN COLUMNS =================

        df.columns = [str(c).strip() for c in df.columns]

        st.success("File Uploaded Successfully ✅")

        # ================= AUTO DETECT COLUMNS =================

        shape_col = find_column(df, ["shape"])
        color_col = find_column(df, ["color"])
        clarity_col = find_column(df, ["clarity"])
        cts_col = find_column(df, ["cts", "carat", "weight", "size"])
        cert_col = find_column(df, ["cert", "certificate"])
        rate_col = find_column(df, ["amt/cts", "price/ct", "rate"])
        total_col = find_column(df, ["total amt", "amount"])
        vendor_col = find_column(df, ["vendor", "company"])
        video_col = find_column(df, ["video"])

        # ================= CREATE OUTPUT =================

        output = pd.DataFrame()

        # VENDOR
        if vendor_col:
            output["VENDOR"] = df[vendor_col]
        else:
            output["VENDOR"] = "GOLDEN"

        # SR NO
        output["SR NO."] = [f"C{i+1}" for i in range(len(df))]

        # SHAPE
        if shape_col:
            output["SHAPE"] = df[shape_col]
        else:
            output["SHAPE"] = ""

        output["SHAPE"] = output["SHAPE"].replace(shape_mapping)

        # COLOR
        if color_col:
            output["COLOR"] = df[color_col]
        else:
            output["COLOR"] = ""

        # CLARITY
        if clarity_col:
            output["CLARITY"] = df[clarity_col]
        else:
            output["CLARITY"] = ""

        # CARATS
        if cts_col:
            output["CARATS"] = df[cts_col]
        else:
            output["CARATS"] = ""

                # CERT#
                # ================= CERT# AUTO FIND =================

        cert_values = []

        for index, row in df.iterrows():

            cert_found = ""

            for value in row:

                value = str(value).replace(".0", "").strip()

                # CHECK 9 DIGIT NUMBER
                if value.isdigit() and len(value) == 9:

                    cert_found = "LG" + value
                    break

            cert_values.append(cert_found)

        output["CERT#"] = cert_values

        # AMT/CTS $
        if rate_col:
            output["AMT/CTS $"] = df[rate_col]
        else:
            output["AMT/CTS $"] = ""

        # TOTAL AMT $
        if total_col:
            output["TOTAL AMT $"] = df[total_col]
        else:
            output["TOTAL AMT $"] = ""

        # EMPTY COLUMNS
        output["$ RATE"] = ""
        output["AMT/CTS RS"] = 0
        output["TOTAL AMT RS"] = 0

        # BRAND
        output["BRAND"] = df.apply(detect_brand, axis=1)

        # ES CODE#
        output["ES CODE#"] = range(79720, 79720 + len(output))

        # ACTUAL SHAPE
        output["ACTUAL SHAPE"] = output["SHAPE"]

        # VIDEO
        if video_col:
            output["VIDEO"] = df[video_col]
        else:
            output["VIDEO"] = ""

        # ================= FINAL FORMAT =================

        output = output[final_columns]

        # ================= SPLIT FILES =================

        cvd_df = output[
            output["BRAND"] == "CVD"
        ]

        hpht_df = output[
            output["BRAND"] == "HPHT"
        ]

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

        st.error(f"Error: {e}")
