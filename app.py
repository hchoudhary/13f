import pandas as pd
from io import BytesIO
from lxml import etree
import xml.sax.saxutils as saxutils
import streamlit as st
import base64

# Utility function to create a header
def create_header():
    logo_path = "/mnt/data/BR-Logo-RGB-Blue.png"  # Replace with the correct path to the logo
    with open(logo_path, "rb") as image_file:
        logo_base64 = base64.b64encode(image_file.read()).decode("utf-8")

    logo_html = f"""
    <div style="background-color: #f8f9fa; padding: 10px; display: flex; justify-content: space-between; align-items: center;">
        <h1 style="color: #2d2d2d; margin: 0; font-size: 1.5rem;">Filing Application</h1>
        <img src="data:image/png;base64,{logo_base64}" alt="Broadridge Logo" style="height: 50px;">
    </div>
    """
    st.markdown(logo_html, unsafe_allow_html=True)

# Utility function to generate SHO sample Excel
def generate_sho_sample_excel():
    header_data = {
        "submissionType": ["SHO"],
        "cik": ["9999999999"],
        "ccc": ["test#1"],
        "liveTestFlag": ["LIVE"],
        "overrideInternetFlag": ["false"],
        "contactName": ["Tex"],
        "contactPhoneNumber": ["5554443333"],
        "contactEmailAddress": ["johndoe@example.com"],
        "notificationEmailAddress": ["johndoe@example.com"],
    }

    cover_page_data = {
        "reportPeriodEnded": ["09/26/2024"],
        "reporterName": ["John Doe"],
        "reporterStreet1": ["100 STREET NE"],
        "reporterStreet2": ["Suite 300"],
        "reporterCity": ["WASHINGTON"],
        "reporterStateOrCountry": ["DC"],
        "reporterZipCode": ["20549"],
        "reporterPhoneNumber": ["5555555555"],
        "reporterEmail": ["johndoe@example.com"],
        "nonLapsedLEI": ["76867876876867867867"],
        "employeeContactName": ["Jane Doe"],
        "employeeContactTitle": ["Jane Doe title"],
        "employeeContactEmail": ["janedoe@example.com"],
        "employeeContactPhoneNumber": ["5555555556"],
        "dateFiled": ["10/01/2024"],
        "reportType": ["FORM SHO COMBINATION REPORT"],
    }

    gross_short_table1 = {
        "settlementDate": ["09/26/2024"],
        "issuerName": ["Issuer"],
        "leiNumber": ["34534DE4564564564564"],
        "securitiesClassTitle": ["Test Class"],
        "issuerCusip": ["5645654FSD34"],
        "figiNumber": ["34545345435D"],
        "numberOfShares": [5000],
        "valueOfShares": [500000],
    }

    daily_gross_short_table2 = {
        "shortIssuerName": ["Issuer", "Issuer", "Issuer", "Issuer", "Issuer"],
        "leiNumber": ["34534DE4564564564564"] * 5,
        "securitiesClassTitle": ["Test Class"] * 5,
        "issuerCusip": ["5645654FSD34"] * 5,
        "figiNumber": ["34545345435D"] * 5,
        "settlementDate": [
            "09/03/2024",
            "09/04/2024",
            "09/05/2024",
            "09/06/2024",
            "09/09/2024",
        ],
        "netChangeOfShares": [35435, 54646, None, -299, 436],
    }

    df_header_data = pd.DataFrame(header_data)
    df_cover_page = pd.DataFrame(cover_page_data)
    df_gross_short_table1 = pd.DataFrame(gross_short_table1)
    df_daily_gross_short_table2 = pd.DataFrame(daily_gross_short_table2)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_header_data.to_excel(writer, index=False, sheet_name="HeaderData")
        df_cover_page.to_excel(writer, index=False, sheet_name="CoverPage")
        df_gross_short_table1.to_excel(writer, index=False, sheet_name="GrossShortTable1")
        df_daily_gross_short_table2.to_excel(writer, index=False, sheet_name="DailyGrossShortTable2")

    output.seek(0)
    return output

# Streamlit UI
create_header()

# Create tabs for 13F Filing and SHO Filing
tab1, tab2 = st.tabs(["13F Filing", "SHO Filing"])

# 13F Filing Tab
with tab1:
    st.header("13F Filing")
    st.write("Upload your 13F Excel file and generate XML.")
    uploaded_file_13f = st.file_uploader("Upload 13F Excel File", type=["xlsx"], key="13f")
    if uploaded_file_13f:
        st.success("File uploaded for 13F Filing.")
        st.write("Processing 13F filing is under development.")

# SHO Filing Tab
with tab2:
    st.header("SHO Filing")
    st.write("Download a sample SHO Excel file for data entry.")
    sho_sample_excel = generate_sho_sample_excel()
    st.download_button(
        label="Download SHO Sample Excel",
        data=sho_sample_excel,
        file_name="SHO_Sample.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
