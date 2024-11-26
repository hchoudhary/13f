import pandas as pd
from io import BytesIO
from lxml import etree
import xml.sax.saxutils as saxutils
import streamlit as st
import base64

# Column mapping to match the XML structure
COLUMN_MAPPING = {
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

def create_header():
    # Add a header with Broadridge logo
    logo_path = "BR-Logo-RGB-Blue.png"  # Update path if logo is present
    with open(logo_path, "rb") as logo_file:
        logo_base64 = base64.b64encode(logo_file.read()).decode("utf-8")

    header_html = f"""
    <div style="background-color:#f8f9fa; padding:10px; display: flex; justify-content: space-between; align-items: center;">
        <h1 style="color: #2d2d2d; margin: 0; font-size: 1.5rem;">Excel to Form 13F XML Converter</h1>
        <img src="data:image/png;base64,{logo_base64}" alt="Broadridge Logo" style="height:50px;">
    </div>
    """
    st.markdown(header_html, unsafe_allow_html=True)

def validate_excel_data(df):
    # Basic validation for required fields
    required_columns = set(COLUMN_MAPPING.keys())
    if not required_columns.issubset(df.columns):
        raise ValueError(f"Excel file must contain these columns: {required_columns}")

    if not all(df["CUSIP"].astype(str).str.len() == 9):
        raise ValueError("CUSIP must be exactly 9 characters.")
    if "FIGI" in df.columns:
        if not df["FIGI"].astype(str).str.len().eq(12).all() and not df["FIGI"].isna().all():
            raise ValueError("FIGI must be 12 characters or left blank.")

    return df

def create_xml(df):
    # Define the namespace map
    NSMAP = {
        None: "http://www.sec.gov/edgar/document/thirteenf/informationtable"
    }

    # Create the root element
    root = etree.Element("informationTable", nsmap=NSMAP)

    for _, row in df.iterrows():
        # Create an infoTable entry
        info_table_entry = etree.SubElement(root, "infoTable")

        # Add main elements
        for col, xml_tag in COLUMN_MAPPING.items():
            if col in ["Shares or Principal Amount", "Shares/Principal", "Sole", "Shared", "None"]:
                continue  # Skip these for nested structures
            value = saxutils.escape(str(row[col])) if pd.notna(row[col]) else ""
            etree.SubElement(info_table_entry, xml_tag).text = value

        # Add shrsOrPrnAmt nested structure
        shrs_or_prn_amt = etree.SubElement(info_table_entry, "shrsOrPrnAmt")
        sshPrnamt = saxutils.escape(str(row["Shares or Principal Amount"])) if pd.notna(row["Shares or Principal Amount"]) else ""
        sshPrnamtType = saxutils.escape(str(row["Shares/Principal"])) if pd.notna(row["Shares/Principal"]) else ""
        etree.SubElement(shrs_or_prn_amt, "sshPrnamt").text = sshPrnamt
        etree.SubElement(shrs_or_prn_amt, "sshPrnamtType").text = sshPrnamtType

        # Add votingAuthority nested structure
        voting_authority = etree.SubElement(info_table_entry, "votingAuthority")
        for vote_col in ["Sole", "Shared", "None"]:
            value = saxutils.escape(str(row[vote_col])) if pd.notna(row[vote_col]) else "0"
            etree.SubElement(voting_authority, vote_col).text = value

    # Convert XML to string and return
    return etree.tostring(root, pretty_print=True, xml_declaration=True, encoding="UTF-8")

def generate_sample_excel():
    # Create a sample DataFrame
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

    # Save the DataFrame to an Excel file in memory
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="SampleData")
    buffer.seek(0)
    return buffer

# Streamlit UI
create_header()

st.markdown("### Instructions for Preparing the Information Table in Excel")

st.markdown(
    """
    - The information table should have exactly 13 columns with the following headers: Name of Issuer, Title of Class, CUSIP, FIGI, Value (to the nearest dollar), Shares or Principal Amount, Shares/Principal, Put/Call, Investment Discretion, Other Managers, Sole, Shared, None.
    - No blank rows above or below the table.
    - Header information must be on one row (e.g., Name of Issuer, Title of Class, etc.).
    - **CUSIP** number must be nine (9) characters.
    - **FIGI** column can be blank if there is no value. If filled, it must be twelve (12) characters.
    - The table must contain only whole numbers—no decimals. Manually correct rounded values in Excel to avoid validation issues.
    - The only acceptable values in the **Shares/Principal** column are `SH` or `PRN`.
    - The **Put/Call** column must be either blank or contain `Put` or `Call` (case-sensitive).
    - **Investment Discretion** must be `SOLE`, `DFND`, or `OTR` (case-sensitive).
    - If more than one value exists in **Other Managers**, use commas (e.g., `1,10,23`).
    - Enter `0` in the Sole, Shared, and None columns if no values exist.
    """
)

st.markdown("### Download a Sample Excel File")
sample_excel = generate_sample_excel()
st.download_button(
    label="Download Sample Excel File",
    data=sample_excel,
    file_name="Sample_Form13F.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
if uploaded_file:
    try:
        # Read the uploaded Excel file
        df = pd.read_excel(uploaded_file)
        df = validate_excel_data(df)

        # Generate the XML file
        xml_content = create_xml(df)

        # Provide download link for the XML
        st.success("XML file successfully generated!")
        st.download_button(
            label="Download XML File",
            data=xml_content,
            file_name="form13F.xml",
            mime="application/xml"
        )
    except Exception as e:
        st.error(f"An error occurred: {e}")
