import datetime
import os
import pandas as pd
import geopandas as gpd
import zipfile
import xml.etree.ElementTree as ET
import re
from datetime import datetime

def find_year_in_filename(filename):
    # Define the range of years, assuming it's from 1900 to 2099
    pattern = r'(19\d{2}|20\d{2})'
    # Use regular expression to search for the year in the filename
    match = re.search(pattern, filename)
    if match:
        return match.group()  # Return the matched year
    else:
        return None  # Return None if no year is found

# Analyze information from an Excel file
def excel_EL(file_path):
    # Read Excel file
    df = pd.read_excel(file_path, header=None, nrows=1000)
    # Count rows with data
    row_counts = df.count(axis=1)
    # Check if the DataFrame is empty
    if row_counts.empty:
        raise ValueError("DataFrame is empty or contains no valid data.")
    # Find the row with the fewest null values
    max_row_count = row_counts.max()
    if max_row_count == 0:
        raise ValueError("All rows are empty.")
    # Set this row as the header and retrieve the table below it
    first_row = row_counts.idxmax() + 1
    # Read data from the title line
    df = pd.read_excel(file_path, header=first_row)
    return df



# Analyze information from a CSV file
def csv_EL(file_path):
    # Read CSV with various encodings and delimiters
    try:
        df = pd.read_csv(file_path, low_memory=False, on_bad_lines='skip')
    except:
        df = pd.read_csv(file_path, encoding='latin1', low_memory=False, on_bad_lines='skip')
    if df.shape[1] == 1:
        try:
            df = pd.read_csv(file_path, sep=';', low_memory=False, on_bad_lines='skip')
        except:
            df = pd.read_csv(file_path, encoding='latin1', sep=';', low_memory=False, on_bad_lines='skip')
    return df

# Analyze information from a GeoJSON file
def geojson_EL(file_path):
    # Read GeoJSON data
    gdf = gpd.read_file(file_path)
    df = gdf.to_pandas()
    return df


# Analyze information from a Shapefile (.shp)
def shapefile_EL(zip_file_path):
    extracted_folder = os.getcwd()
    # Unzip the file
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        zip_ref.extractall(extracted_folder)

    # Find Shapefile within extracted files
    shapefile_path = None
    for root, dirs, files in os.walk(extracted_folder):
        for file in files:
            if file.endswith('.shp'):
                shapefile_path = os.path.join(root, file)
                break

    if shapefile_path:
        gdf = gpd.read_file(shapefile_path)
        df = gdf.to_pandas()
    return df

# Analyze information from a JSON file

def json_EL(file_path):
    df = pd.read_json(file_path, lines=True)
    return df

# Recursively traverse through XML nodes
def recursive_traverse(node, level=0):
    indent = " " * (level * 4)  # Set indent for visual representation
    print(f"{indent}Tag: {node.tag}, Attributes: {node.attrib}, Text: {node.text.strip() if node.text else 'None'}")

    # Traverse child nodes
    for child in node:
        recursive_traverse(child, level + 1)

def xml_EL(file_path):
    tree = ET.parse('example.xml')
    root = tree.getroot()
    # Create an empty list to store the data
    data = []
    # Iterate over each element in the XML tree
    for elem in root:
        # Store each child element's tag and text content as a dictionary
        record = {}
        for child in elem:
            record[child.tag] = child.text
        # Add the dictionary to the data list
        data.append(record)
    # Convert the data list to a DataFrame
    df = pd.DataFrame(data)
    return df

# Determine file type and process accordingly
def find_type(path):
    file_extension = os.path.splitext(path)[1].lower()
    match file_extension:
        case '.xlsx':
            return excel_EL(path), "structured"
        case '.csv':
            return csv_EL(path), "structured"
        case '.geojson':
            return geojson_EL(path), "semi-structured"
        case '.zip':
            return shapefile_EL(path), "semi-structured"
        case '.json':
            return json_EL(path), "semi-structured"
        case '.xml':
            return xml_EL(path),
        # case _:
        #     return False, False, False, False, False, False, False, False, False, False, False, False, False

# Retrieve all files within the specified path
def get_all_files(file_path):
    file_list = []
    # Traverse directory
    for root, dirs, files in os.walk(file_path):
        for file in files:
            if not file.endswith('.ini'):
                file_list.append(os.path.join(root, file))
    return file_list