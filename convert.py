import pandas as pd
from lxml import etree
import xml.sax.saxutils as saxutils

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

def validate_excel_data(file_path):
    # Load the Excel file
    df = pd.read_excel(file_path)

    # Check if there are required columns
    required_columns = set(COLUMN_MAPPING.keys())
    if not required_columns.issubset(df.columns):
        raise ValueError(f"Excel file must contain these columns: {required_columns}")

    # Basic validation for required fields
    if not all(df["CUSIP"].astype(str).str.len() == 9):
        raise ValueError("CUSIP must be exactly 9 characters.")
    if "FIGI" in df.columns:
        if not df["FIGI"].astype(str).str.len().eq(12).all() and not df["FIGI"].isna().all():
            raise ValueError("FIGI must be 12 characters or left blank.")

    return df

def create_xml(df, output_file):
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

    # Write the XML to a file
    tree = etree.ElementTree(root)
    with open(output_file, "wb") as f:
        tree.write(f, encoding="UTF-8", xml_declaration=True, pretty_print=True)

    print(f"XML file successfully created: {output_file}")

def main():
    # Input Excel and output XML paths
    input_excel = "information_table.xlsx"  # Replace with your input file path
    output_xml = "form13F.xml"

    try:
        # Validate the Excel file
        df = validate_excel_data(input_excel)

        # Create the XML file
        create_xml(df, output_xml)
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
