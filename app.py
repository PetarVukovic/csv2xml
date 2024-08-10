import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
from io import BytesIO
import time
import html
import re


def clean_column_name(name):
    """Clean the column name to be a valid XML tag."""
    # Replace spaces with underscores and remove any non-alphanumeric characters
    name = re.sub(r"\W+", "_", name)
    return name


def clean_data(value):
    """Cleans the data to ensure it is safe for XML conversion."""
    if isinstance(value, str):
        # Replace or remove any characters that could cause issues
        value = value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return value


def convert_df_to_pretty_xml(df):
    # Clean the DataFrame column names to prevent XML errors
    df.columns = [clean_column_name(col) for col in df.columns]

    # Clean the DataFrame values to prevent XML errors
    df = df.applymap(clean_data)

    # Create root element of the XML
    root = ET.Element("Root")

    # Iterate through the rows of the DataFrame and add them to the XML
    for i, row in df.iterrows():
        # Create a new XML element for each row
        item = ET.SubElement(root, "Item")

        # Add data from each column as sub-elements
        for col_name in df.columns:
            col_element = ET.SubElement(item, col_name)
            try:
                # Escape special characters using html.escape
                col_element.text = html.escape(str(row[col_name]))
            except Exception as e:
                st.error(f"Error converting column '{col_name}' in row {i}: {e}")
                st.write(f"Original Value: {row[col_name]}")

    # Prepare the XML structure
    xml_string = ET.tostring(root, encoding="utf-8", method="xml")
    try:
        xml_pretty = minidom.parseString(xml_string).toprettyxml(indent="  ")
    except Exception as e:
        st.error(f"Error formatting XML: {e}")
        st.write(f"Problem detected with the following raw XML snippet:")
        # Show the specific part of XML that might be causing the issue
        st.write(xml_string.decode("utf-8")[:500])  # Display the first 500 characters
        raise e  # Rethrow the exception after logging

    # Prepare XML file for download
    xml_io = BytesIO()
    xml_io.write(xml_pretty.encode("utf-8"))
    xml_io.seek(0)  # Reset pointer to the start
    return xml_io


def test_conversion_accuracy(df, xml_file):
    """
    Tests the accuracy of the conversion.
    df - Original DataFrame from Excel/CSV file
    xml_file - Generated XML file as BytesIO object
    """
    # Parse the XML file
    xml_file.seek(0)  # Reset pointer to the start
    tree = ET.parse(xml_file)
    root = tree.getroot()

    # Convert XML data to a list of dictionaries
    xml_data = []
    for item in root.findall("Item"):
        xml_row = {}
        for child in item:
            xml_row[child.tag] = child.text
        xml_data.append(xml_row)

    # Compare data row by row
    for i, row in df.iterrows():
        for key in df.columns:
            if str(row[key]) != xml_data[i].get(key):
                return (
                    False,
                    f"Mismatch found in record {i+1} for column '{key}': '{row[key]}' != '{xml_data[i].get(key)}'",
                )

    return True, "All data has been accurately converted!"


def main():
    st.set_page_config(page_title="Excel to XML Converter", page_icon="📄")

    # Welcome users
    st.title("📄 Excel/CSV to XML Converter")
    st.write(
        "Welcome to the Excel or CSV file to XML format conversion application! 🚀"
    )
    st.info(
        "📂 Upload a file using the form in the sidebar and download the generated XML file."
    )

    st.sidebar.title("Application Settings")
    st.sidebar.info(
        "This application allows you to convert Excel or CSV files to XML format. "
        "Upload the file using the form below and download the generated XML file."
    )

    # Start measuring time
    start_time = time.time()

    # File upload
    uploaded_file = st.sidebar.file_uploader(
        "Upload Excel or CSV file", type=["xlsx", "xls", "csv"]
    )

    if uploaded_file is not None:
        try:
            # Extract data from Excel or CSV file
            if uploaded_file.name.endswith(".csv"):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)

            st.success("✅ File successfully loaded!")

            # Display DataFrame
            st.write("Data Preview:")
            st.dataframe(df)

            # Convert DataFrame to XML
            xml_file = convert_df_to_pretty_xml(df)

            # Display option to download XML
            st.download_button(
                label="💾 Download XML file",
                data=xml_file,
                file_name="output.xml",
                mime="application/xml",
            )

            # Test conversion accuracy
            test_result, test_message = test_conversion_accuracy(df, xml_file)
            if test_result:
                st.success(f"✅ {test_message}")
            else:
                st.error(f"❌ {test_message}")

        except Exception as e:
            st.error(f"❌ An error occurred: {e}")

    # End measuring time
    end_time = time.time()
    execution_time = end_time - start_time
    st.sidebar.write(f"⏱ Execution time: {execution_time:.2f} seconds")


if __name__ == "__main__":
    main()
