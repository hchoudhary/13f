def generate_sho_xml_from_excel(excel_file):
    # Read the Excel file into DataFrames
    excel_data = pd.ExcelFile(excel_file)
    header_data = excel_data.parse("HeaderData")
    cover_page = excel_data.parse("CoverPage")
    gross_short_table1 = excel_data.parse("GrossShortTable1")
    daily_gross_short_table2 = excel_data.parse("DailyGrossShortTable2")

    # Create the XML root with namespaces
    nsmap = {"com": "http://www.sec.gov/edgar/common"}
    root = etree.Element("edgarSubmission", nsmap=nsmap)
    root.attrib["xmlns"] = "http://www.sec.gov/edgar/formSho"

    # Add HeaderData
    header = etree.SubElement(root, "headerData")
    submission_type = etree.SubElement(header, "submissionType")
    submission_type.text = header_data.at[0, "submissionType"]
    filer_info = etree.SubElement(header, "filerInfo")
    filer = etree.SubElement(filer_info, "filer")
    filer_credentials = etree.SubElement(filer, "filerCredentials")
    etree.SubElement(filer_credentials, "cik").text = header_data.at[0, "cik"]
    etree.SubElement(filer_credentials, "ccc").text = header_data.at[0, "ccc"]
    etree.SubElement(filer_info, "liveTestFlag").text = header_data.at[0, "liveTestFlag"]
    flags = etree.SubElement(filer_info, "flags")
    etree.SubElement(flags, "overrideInternetFlag").text = header_data.at[0, "overrideInternetFlag"]
    contact = etree.SubElement(filer_info, "contact")
    etree.SubElement(contact, "contactName").text = header_data.at[0, "contactName"]
    etree.SubElement(contact, "contactPhoneNumber").text = header_data.at[0, "contactPhoneNumber"]
    etree.SubElement(contact, "contactEmailAddress").text = header_data.at[0, "contactEmailAddress"]
    notifications = etree.SubElement(filer_info, "notifications")
    etree.SubElement(notifications, "notificationEmailAddress").text = header_data.at[0, "notificationEmailAddress"]

    # Add FormData Cover Page
    form_data = etree.SubElement(root, "formData")
    cover_page_elem = etree.SubElement(form_data, "coverPage")
    etree.SubElement(cover_page_elem, "reportPeriodEnded").text = cover_page.at[0, "reportPeriodEnded"]
    reporter_info = etree.SubElement(cover_page_elem, "reporterInfo")
    etree.SubElement(reporter_info, "name").text = cover_page.at[0, "reporterName"]
    reporter_address = etree.SubElement(reporter_info, "reporterAddress")
    etree.SubElement(reporter_address, "{http://www.sec.gov/edgar/common}street1").text = cover_page.at[0, "reporterStreet1"]
    etree.SubElement(reporter_address, "{http://www.sec.gov/edgar/common}street2").text = cover_page.at[0, "reporterStreet2"]
    etree.SubElement(reporter_address, "{http://www.sec.gov/edgar/common}city").text = cover_page.at[0, "reporterCity"]
    etree.SubElement(reporter_address, "{http://www.sec.gov/edgar/common}stateOrCountry").text = cover_page.at[0, "reporterStateOrCountry"]
    etree.SubElement(reporter_address, "{http://www.sec.gov/edgar/common}zipCode").text = cover_page.at[0, "reporterZipCode"]
    etree.SubElement(reporter_info, "phoneNumber").text = cover_page.at[0, "reporterPhoneNumber"]
    etree.SubElement(reporter_info, "email").text = cover_page.at[0, "reporterEmail"]
    etree.SubElement(reporter_info, "nonLapsedLEI").text = cover_page.at[0, "nonLapsedLEI"]
    employee_contact = etree.SubElement(cover_page_elem, "employeeContact")
    etree.SubElement(employee_contact, "name").text = cover_page.at[0, "employeeContactName"]
    etree.SubElement(employee_contact, "title").text = cover_page.at[0, "employeeContactTitle"]
    etree.SubElement(employee_contact, "email").text = cover_page.at[0, "employeeContactEmail"]
    etree.SubElement(employee_contact, "phoneNumber").text = cover_page.at[0, "employeeContactPhoneNumber"]
    etree.SubElement(employee_contact, "dateFiled").text = cover_page.at[0, "dateFiled"]
    etree.SubElement(cover_page_elem, "reportType").text = cover_page.at[0, "reportType"]

    # Add Managers Gross Short Table 1
    gross_short_table_elem = etree.SubElement(form_data, "managersGrossShortTable1")
    for _, row in gross_short_table1.iterrows():
        gross_short_info = etree.SubElement(gross_short_table_elem, "managersGrossShortTable1Info")
        etree.SubElement(gross_short_info, "settlementDate").text = row["settlementDate"]
        etree.SubElement(gross_short_info, "issuerName").text = row["issuerName"]
        etree.SubElement(gross_short_info, "leiNumber").text = row["leiNumber"]
        etree.SubElement(gross_short_info, "securitiesClassTitle").text = row["securitiesClassTitle"]
        etree.SubElement(gross_short_info, "issuerCusip").text = row["issuerCusip"]
        etree.SubElement(gross_short_info, "figiNumber").text = row["figiNumber"]
        etree.SubElement(gross_short_info, "numberOfShares").text = str(row["numberOfShares"])
        etree.SubElement(gross_short_info, "valueOfShares").text = str(row["valueOfShares"])

    # Add Managers Daily Gross Short Table 2
    daily_gross_table_elem = etree.SubElement(form_data, "managersDailyGrossShortTable2")
    for _, row in daily_gross_short_table2.iterrows():
        issuer_elem = etree.SubElement(daily_gross_table_elem, "table2IssuerList")
        etree.SubElement(issuer_elem, "shortIssuerName").text = row["shortIssuerName"]
        etree.SubElement(issuer_elem, "leiNumber").text =
