import pandas as pd
from io import BytesIO
from lxml import etree
import xml.sax.saxutils as saxutils
import streamlit as st

# Column mapping for 13F Filing
COLUMN_MAPPING_13F = {
    "Name of Issuer": "nameOfIssuer",
    "Title of Class": "titleOfClass",
    "CUSIP": "cusip",
    "FIGI": "figi",
    "Value (to the nearest dollar)": "value",
    "Shares or Principal Amount": "sshPrnamt",
    "Shares/Principal": "sshPrnamtType",
    "Put/Call": "putCall",
    "Investment Discretion": "investmentDiscretion",
    "Other Managers": "otherManager",
    "Sole": "Sole",
    "Shared": "Shared",
    "None": "None"
}

# Utility to validate 13F Excel data
def validate_13f_excel_data(df):
    required_columns = set(COLUMN_MAPPING_13F.keys())
    if not required_columns.issubset(df.columns):
        raise ValueError(f"Excel file must contain these columns: {required_columns}")

    if not all(df["CUSIP"].astype(str).str.len() == 9):
        raise ValueError("CUSIP must be exactly 9 characters.")
    if "FIGI" in df.columns:
        if not df["FIGI"].astype(str).str.len().eq(12).all() and not df["FIGI"].isna().all():
            raise ValueError("FIGI must be 12 characters or left blank.")
    return df

# Utility to create 13F XML
def create_13f_xml(df):
    NSMAP = {None: "http://www.sec.gov/edgar/document/thirteenf/informationtable"}
    root = etree.Element("informationTable", nsmap=NSMAP)

    for _, row in df.iterrows():
        info_table_entry = etree.SubElement(root, "infoTable")
        for col, xml_tag in COLUMN_MAPPING_13F.items():
            if col in ["Shares or Principal Amount", "Shares/Principal", "Sole", "Shared", "None"]:
                continue
            value = saxutils.escape(str(row[col])) if pd.notna(row[col]) else ""
            etree.SubElement(info_table_entry, xml_tag).text = value

        shrs_or_prn_amt = etree.SubElement(info_table_entry, "shrsOrPrnAmt")
        etree.SubElement(shrs_or_prn_amt, "sshPrnamt").text = str(row["Shares or Principal Amount"])
        etree.SubElement(shrs_or_prn_amt, "sshPrnamtType").text = str(row["Shares/Principal"])

        voting_authority = etree.SubElement(info_table_entry, "votingAuthority")
        for vote_col in ["Sole", "Shared", "None"]:
            etree.SubElement(voting_authority, vote_col).text = str(row[vote_col])

    return etree.tostring(root, pretty_print=True, xml_declaration=True, encoding="UTF-8")

# Utility to generate sample 13F Excel file
def generate_sample_13f_excel():
    data = {
        "Name of Issuer": ["123 Co", "ABC Inc"],
        "Title of Class": ["COM", "COM SER A"],
        "CUSIP": ["00206R102", "030420103"],
        "FIGI": ["BBG001560KQ0", "BBGA115608Q1"],
        "Value (to the nearest dollar)": [1234567, 2345678],
        "Shares or Principal Amount": [123, 234],
        "Shares/Principal": ["SH", "PRN"],
        "Put/Call": ["Put", "Call"],
        "Investment Discretion": ["SOLE", "DFND"],
        "Other Managers": ["12", "1,34,56,13"],
        "Sole": [123, 25],
        "Shared": [0, 30],
        "None": [123, 179]
    }
    df = pd.DataFrame(data)
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="SampleData")
    buffer.seek(0)
    return buffer

# Utility to create SHO XML
def create_sho_xml(df):
    nsmap = {"com": "http://www.sec.gov/edgar/common"}
    root = etree.Element("edgarSubmission", nsmap=nsmap)
    form_data = etree.SubElement(root, "formData")
    sho_elem = etree.SubElement(form_data, "shoDetails")

    for _, row in df.iterrows():
        entry = etree.SubElement(sho_elem, "shoEntry")
        etree.SubElement(entry, "settlementDate").text = row["settlementDate"]
        etree.SubElement(entry, "issuerName").text = row["issuerName"]
        etree.SubElement(entry, "shares").text = str(row["shares"])
        etree.SubElement(entry, "value").text = str(row["value"])

    return etree.tostring(root, pretty_print=True, xml_declaration=True, encoding="UTF-8")

# Streamlit app with tabs
st.title("Filing Application")

tab1, tab2 = st.tabs(["13F Filing", "SHO Filing"])

# 13F Filing Tab
with tab1:
    st.header("13F Filing")
    uploaded_file_13f = st.file_uploader("Upload 13F Excel File", type=["xlsx"], key="13f")
    if uploaded_file_13f:
        try:
            df_13f = pd.read_excel(uploaded_file_13f)
            df_13f = validate_13f_excel_data(df_13f)
            xml_13f = create_13f_xml(df_13f)
            st.success("13F XML successfully generated!")
            st.download_button("Download 13F XML", data=xml_13f, file_name="form13F.xml", mime="application/xml")
        except Exception as e:
            st.error(f"Error: {e}")

    st.markdown("### Need a Sample Excel File?")
    st.download_button(
        label="Download Sample 13F Excel File",
        data=generate_sample_13f_excel(),
        file_name="Sample_13F.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# SHO Filing Tab
with tab2:
    st.header("SHO Filing")
    uploaded_file_sho = st.file_uploader("Upload SHO Excel File", type=["xlsx"], key="sho")
    if uploaded_file_sho:
        try:
            df_sho = pd.read_excel(uploaded_file_sho)
            xml_sho = create_sho_xml(df_sho)
            st.success("SHO XML successfully generated!")
            st.download_button("Download SHO XML", data=xml_sho, file_name="formSHO.xml", mime="application/xml")
        except Exception as e:
            st.error(f"Error: {e}")
