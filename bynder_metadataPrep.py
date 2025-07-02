import warnings
import pandas as pd
from datetime import datetime
import os
import re
import sys
import argparse
import tkinter as tk
from tkinter import Tk, filedialog 

# Suppress specific UserWarning about the workbook style
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# --- Hardcoded STEP Export Paths (Platform Specific) ---
# These paths point to the Excel files containing SKU, Vendor Code, and Asset Paths.
# They are adjusted based on the operating system to ensure correct file access.
if sys.platform == "darwin":  # macOS
    HARDCODED_STEP_ONE = "/Volumes/Work_In_Progress/2024_DAM_DI_Projects/_STEP Export/STEP_Export_one.xlsx"
    HARDCODED_STEP_TWO = "/Volumes/Work_In_Progress/2024_DAM_DI_Projects/_STEP Export/STEP_Export_two.xlsx" # CORRECTED PATH
elif os.name == "nt":  # Windows
    HARDCODED_STEP_ONE = r"W:\2024_DAM_DI_Projects\_STEP Export\STEP_Export_one.xlsx"
    HARDCODED_STEP_TWO = r"W:\2024_DAM_DI_Projects\_STEP Export\STEP_Export_two.xlsx"
else:
    # Fallback for unsupported OS. User might need to manually adjust these paths.
    print(f"Warning: Unsupported OS '{sys.platform}'. Please update hardcoded STEP paths in bynder_metadataPrep.py.", file=sys.stderr)
    HARDCODED_STEP_ONE = "STEP_Export_one.xlsx"
    HARDCODED_STEP_TWO = "STEP_Export_two.xlsx"


# --- Function to get folder path (Supports command-line or UI selection) ---
def get_folder_path(input_folder_arg=None):
    """
    Prompts the user to select an input folder using a Tkinter dialog,
    or uses a provided command-line argument if available and valid.

    Args:
        input_folder_arg (str, optional): A folder path provided as a command-line argument.

    Returns:
        str: The absolute path to the selected or provided input folder.
    
    Exits:
        sys.exit(1) if no folder is selected or an error occurs.
    """
    if input_folder_arg:
        # Check if the provided command-line argument is a valid directory
        if os.path.isdir(input_folder_arg):
            print(f"Script: Using input folder from argument: {input_folder_arg}")
            return input_folder_arg
        else:
            # If argument is invalid, inform user and fall back to interactive selection
            print(f"Script: Error: Provided path is not a valid directory: {input_folder_arg}", file=sys.stderr)
            print("Script: Attempting interactive folder selection...", file=sys.stderr)

    # If no argument or invalid argument, proceed with interactive folder selection via Tkinter
    try:
        root = tk.Tk()
        root.withdraw() # Hide the main Tkinter window to only show the dialog
        folder_path = filedialog.askdirectory(title="Select the folder containing your assets")
        root.destroy() # Destroy the temporary root window after selection to clean up resources

        if not folder_path:
            # If user closes the dialog without selecting a folder
            print("Script: No folder selected. Exiting.", file=sys.stderr)
            sys.exit(1)
        return folder_path
    except ImportError:
        # Handle cases where Tkinter might not be installed or a display is not available
        print("Script: Tkinter not found or display not available. Please provide the input folder as a command-line argument, e.g., 'python script.py --input /path/to/assets'", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        # Catch any other unexpected errors during folder selection
        print(f"Script: An unexpected error occurred during interactive folder selection: {e}", file=sys.stderr)
        sys.exit(1)

# --- Function to extract SKU and Vendor from filename ---
def extract_sku_and_vendor_from_filename(filename):
    """
    Extracts vendor code and SKU from a given filename based on predefined patterns.
    It handles filenames starting with an optional 'FW_' prefix, followed by
    the vendor code (alphanumeric), a 9-character SKU (alphanumeric), and then
    either an underscore or a dot, followed by other characters.

    Args:
        filename (str): The name of the file (e.g., "FW_VENDOR_SKU987654321_3000.jpg").

    Returns:
        tuple: A tuple containing (vendor_code, sku) in uppercase, or (None, None) if no match.
    """
    # Regex for patterns like FW_VENDOR_SKU_... or VENDOR_SKU_...
    match = re.search(r'^(?:FW_)?([A-Z0-9]+)_([A-Z0-9]{9})_.*', filename, re.IGNORECASE)
    if match:
        vendor_code = match.group(1).upper() # Convert to uppercase for consistency
        sku = match.group(2).upper()         # Convert to uppercase for consistency
        return vendor_code, sku
    else:
        # Alternative regex for patterns like FW_VENDOR_SKU.ext or VENDOR_SKU.ext
        # This handles cases where the SKU is directly followed by the file extension.
        match_alt = re.search(r'^(?:FW_)?([A-Z0-9]+)_([A-Z0-9]{9})\..*', filename, re.IGNORECASE)
        if match_alt:
            vendor_code = match_alt.group(1).upper()
            sku = match_alt.group(2).upper()
            return vendor_code, sku
        else:
            # If no pattern matches, return None for both
            return None, None

# --- Function to generate rows based on template ---
def generate_rows(vendor, sku, step_path, template):
    """
    Generates a list of dictionaries (rows) for the metadata importer CSV.
    Each row is based on the provided template and filled with product-specific
    data (vendor, SKU, STEP path) and other default values.

    Args:
        vendor (str): The vendor code for the product.
        sku (str): The SKU for the product.
        step_path (str): The STEP path obtained from the reference Excel files.
        template (list): A list of dictionaries, where each dictionary defines
                         a template for a specific asset type (e.g., main image, alt image).

    Returns:
        list: A list of dictionaries, each representing a row for the output CSV.
    """
    rows = []
    # Define all possible column headers for the output CSV.
    # This ensures a consistent structure regardless of which template items are used.
    column_headers = [
        "filename", "name", "description", "Asset Type", "Asset Sub-Type", "Deliverable", "Product SKU",
        "Product SKU Position", "Asset Status", "Usage Rights", "tags", "File Type", "STEP Path",
        "Link to Wrike Project", "Sync to Site", "Generic Dimension Diagram With Measurements", "Admin Status",
        "Product Status", "Product Category", "Product Sub-Category", "Product Collection",
        "Component SKUs", "Stock Level (only relevant for Inline products)", "Restock Date (only relevant for Inline products)",
        "Link to Print Materials", "Link to Lifestyle Images", "Link to Store Images", "Initiative", "Sub-Initiative",
        "Print Tracking Code", "Print Tracking - Start Date", "Print Tracking - End Date", "Year", "Video Expiration",
        "Audio Licensing Expiration", "Ad ID", "Lead Offer Message", "Lead Finance Message", "Video Focus",
        "Video Objective", "Video Type", "Total Run Time (TRT)", "Spot Running (MM/DD/YYYY)", "Language",
        "Season", "Holiday/Special Occasion", "Talent", "Sunset Date (MM/DD/YYYY)", "Location Name", "Store Code",
        "Location Status", "Location Address", "Location Town", "Location State", "Location Zip Code",
        "Location Phone Number", "Location Type", "Location", "Inactive Product", "Partner", "Notes",
        "Sign Facade Color", "Sign Location", "Sign Color", "Sign Text",
        "Reviewed products in lifestyle", "Reviewed Studio Uploads", "Featured SKU", "Image Type",
        "scratchpad", "3D Model Source Files Acquired", "Visible to", "BynderTest", "dim_Length",
        "Bynder Report", "Dimensions", "dim_Height", "Figmage doc id", "dim_Width", "Figmage image extension",
        "Figmage node id", "Figmage page id", "Performance Metric", "DNUCampaign", "DNUFeatures",
        "DNUMaterials", "DNUStyle", "DNUPattern", "DNUPackage SKUs", "DNUSign Size",
        "DNUDistribution Channel", "Dim diagram re-cropped", "Embedded Instructions (for updating existing metadata based on automations)",
        "Mattress Size", "Asset Identifier", "Sync Batch", "Marked for Deletion from Site", "scene7 folder",
        "Variant Type", "Source", "PSA Image Type", "Rights Notes", "Workflow", "Workflow Status",
        "Product Name (STEP)", "Vendor Code", "Family Code", "Hero SKU", "Product Color", "Dropped",
        "Visible on Website", "Sales Channel", "Associated Materials Status", "Product in Studio",
        "DNU_PromoUpdate2", "Additional Files Upload Scratchpad", "Bump", "Carousel Dimensions Diagram Audit", "User Status", "Reviewed for Site Content Refresh", "Image Type Pre-Classification"
    ]

    for item in template:
        # Initialize a new row dictionary with all column headers set to empty string
        new_row_dict = {col: "" for col in column_headers}

        # Populate the specific fields based on the template item and product data
        new_row_dict["filename"] = item["filename"].format(vendor=vendor, sku=sku)
        new_row_dict["name"] = item["name"].format(vendor=vendor, sku=sku)
        new_row_dict["Deliverable"] = item["Deliverable"]
        new_row_dict["Product SKU"] = sku
        new_row_dict["Product SKU Position"] = item["Product SKU Position"].format(sku=sku)
        new_row_dict["File Type"] = item["File Type"]
        new_row_dict["Asset Identifier"] = item["name"].format(vendor=vendor, sku=sku)

        # Set common metadata fields
        new_row_dict["Asset Type"] = "Product Materials"
        new_row_dict["Asset Sub-Type"] = "Product Site Asset"
        new_row_dict["Asset Status"] = "Final"
        new_row_dict["Usage Rights"] = "Approved for External Usage"
        new_row_dict["tags"] = sku
        new_row_dict["STEP Path"] = step_path
        new_row_dict["Link to Wrike Project"] = "No link"
        new_row_dict["Sync to Site"] = "Do sync to site"
        new_row_dict["Vendor Code"] = vendor
        new_row_dict["Product Name (STEP)"] = "" # This column typically needs to be filled later via lookup

        rows.append(new_row_dict)
    return rows


# --- Main Script Execution Logic ---
def main():
    # 1. Setup Argument Parser for command-line input
    parser = argparse.ArgumentParser(description="Prepare metadata for Bynder upload from a folder of assets.")
    # Add an optional argument for the input folder path
    parser.add_argument('--input', '-i', type=str, help='Path to the folder containing assets (optional). If not provided, a folder selection dialog will appear.')
    args = parser.parse_args()

    # 2. Get Input Folder Path, using argument if provided, otherwise using UI dialog
    input_folder = get_folder_path(args.input)

    # 3. Use hardcoded STEP file paths defined at the top of the script
    ref_file_one = HARDCODED_STEP_ONE
    ref_file_two = HARDCODED_STEP_TWO

    print(f"Script: Using STEP Export One from: {ref_file_one}")
    print(f"Script: Using STEP Export Two from: {ref_file_two}")

    # Validate existence and format of the hardcoded STEP files
    if not os.path.exists(ref_file_one) or not ref_file_one.lower().endswith('.xlsx'):
        print(f"Error: STEP_Export_one.xlsx not found or invalid at '{ref_file_one}'. Please verify path and file.", file=sys.stderr)
        sys.exit(1)
    if not os.path.exists(ref_file_two) or not ref_file_two.lower().endswith('.xlsx'):
        print(f"Error: STEP_Export_two.xlsx not found or invalid at '{ref_file_two}'. Please verify path and file.", file=sys.stderr)
        sys.exit(1)

    # 4. Define output file path in the user's Downloads folder
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S') # Create a unique timestamp for the filename
    output_file = os.path.join(downloads_folder, f"newBatch_metadataImporter_{timestamp}.csv")

    try:
        # Load data from the two STEP Excel reference files
        df_one = pd.read_excel(ref_file_one, usecols=['SKU', 'Vendor Code', 'Path'])
        df_two = pd.read_excel(ref_file_two, usecols=['SKU', 'Vendor Code', 'Path'])
        
        # Concatenate the two DataFrames into one master reference DataFrame
        df_refs = pd.concat([df_one, df_two], ignore_index=True)
        
        # Convert all SKUs in the reference DataFrame to uppercase for consistent matching
        df_refs['SKU'] = df_refs['SKU'].str.upper()
    except Exception as e:
        # Handle errors during reading or processing of Excel files
        print(f"Script: Error reading reference Excel files from '{ref_file_one}' and '{ref_file_two}': {e}", file=sys.stderr)
        sys.exit(1)

    # Define the template for generating metadata rows. This specifies the expected
    # filenames, names, download URLs, file types, product SKU positions, and deliverables.
    template = [
        {"filename": "FW_{vendor}_{sku}_3000.jpg", "name": "FW_{vendor}_{sku}_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_100", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt1_3000.jpg", "name": "FW_{vendor}_{sku}_Alt1_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt1_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_200", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt2_3000.jpg", "name": "FW_{vendor}_{sku}_Alt2_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt2_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_300", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt3_3000.jpg", "name": "FW_{vendor}_{sku}_Alt3_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt3_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_400", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt4_3000.jpg", "name": "FW_{vendor}_{sku}_Alt4_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt4_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_500", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt5_3000.jpg", "name": "FW_{vendor}_{sku}_Alt5_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt5_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_600", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt6_3000.jpg", "name": "FW_{vendor}_{sku}_Alt6_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt6_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_700", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt7_3000.jpg", "name": "FW_{vendor}_{sku}_Alt7_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt7_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_800", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt8_3000.jpg", "name": "FW_{vendor}_{sku}_Alt8_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt8_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_900", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_1000", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9a_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9a_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9a_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_1100", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9b_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9b_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9b_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_1200", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9c_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9c_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9c_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_1300", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9d_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9d_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9d_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_1400", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9e_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9e_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9e_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_1500", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9f_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9f_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9f_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_1600", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9g_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9g_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9g_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_1700", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9h_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9h_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9h_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_1800", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9i_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9i_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9i_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_1900", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9j_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9j_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9j_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_2000", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9k_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9k_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9k_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_2100", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9l_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9l_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9l_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_2200", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9m_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9m_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9m_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_2300", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9n_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9n_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9n_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_2400", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9o_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9o_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9o_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_2500", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9p_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9p_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9p_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_2600", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9q_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9q_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9q_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_2700", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9r_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9r_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9r_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_2800", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9s_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9s_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9s_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_2900", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9t_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9t_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9t_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_3000", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9u_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9u_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9u_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_3100", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9v_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9v_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9v_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_3200", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9w_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9w_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9w_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_3300", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9x_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9x_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9x_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_3400", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9y_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9y_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9y_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_3500", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9z_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9z_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9z_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_3600", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9za_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9za_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9za_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_3700", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9zb_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9zb_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9zb_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_3800", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9zc_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9zc_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9zc_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_3900", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9zd_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9zd_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9zd_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_4000", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9ze_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9ze_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9ze_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_4100", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9zf_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9zf_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9zf_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_4200", "Deliverable": "Product Carousel Image"},
        {"filename": "FW_{vendor}_{sku}_Alt9zg_3000.jpg", "name": "FW_{vendor}_{sku}_Alt9zg_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_Alt9zg_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_4300", "Deliverable": "Product Carousel Image"},    
        {"filename": "FW_{vendor}_{sku}_s1_3000.jpg", "name": "FW_{vendor}_{sku}_s1_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_s1_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_5000", "Deliverable": "Product Swatch Detail Image"},
        {"filename": "FW_{vendor}_{sku}_s2_3000.jpg", "name": "FW_{vendor}_{sku}_s2_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_s2_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_5100", "Deliverable": "Product Swatch Detail Image"},
        {"filename": "FW_{vendor}_{sku}_s3_3000.jpg", "name": "FW_{vendor}_{sku}_s3_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_s3_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_5200", "Deliverable": "Product Swatch Detail Image"},
        {"filename": "FW_{vendor}_{sku}_s4_3000.jpg", "name": "FW_{vendor}_{sku}_s4_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_s4_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_5300", "Deliverable": "Product Swatch Detail Image"},
        {"filename": "FW_{vendor}_{sku}_s5_3000.jpg", "name": "FW_{vendor}_{sku}_s5_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_s5_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_5400", "Deliverable": "Product Swatch Detail Image"},
        {"filename": "FW_{vendor}_{sku}_s6_3000.jpg", "name": "FW_{vendor}_{sku}_s6_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/FW_{vendor}_{sku}_s6_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_5500", "Deliverable": "Product Swatch Detail Image"},
        {"filename": "FW_{vendor}_{sku}_Alt1m_Vid1.mp4", "name": "FW_{vendor}_{sku}_Alt1m_Vid1", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_Alt1m_Vid1", "File Type": "MP4", "Product SKU Position": "{sku}_210", "Deliverable": "Product Feature Video"},
        {"filename": "FW_{vendor}_{sku}_Alt1m_Vid2.mp4", "name": "FW_{vendor}_{sku}_Alt1m_Vid2", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_Alt1m_Vid2", "File Type": "MP4", "Product SKU Position": "{sku}_220", "Deliverable": "Product Feature Video"},
        {"filename": "FW_{vendor}_{sku}_Alt1m_Vid3.mp4", "name": "FW_{vendor}_{sku}_Alt1m_Vid3", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_Alt1m_Vid3", "File Type": "MP4", "Product SKU Position": "{sku}_230", "Deliverable": "Product Feature Video"},
        {"filename": "FW_{vendor}_{sku}_Alt1m_Vid4.mp4", "name": "FW_{vendor}_{sku}_Alt1m_Vid4", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_Alt1m_Vid4", "File Type": "MP4", "Product SKU Position": "{sku}_240", "Deliverable": "Product Feature Video"},
        {"filename": "FW_{vendor}_{sku}_Alt1m_Vid5.mp4", "name": "FW_{vendor}_{sku}_Alt1m_Vid5", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_Alt1m_Vid5", "File Type": "MP4", "Product SKU Position": "{sku}_250", "Deliverable": "Product Feature Video"},
        {"filename": "FW_{vendor}_{sku}_Alt1m_Vid6.mp4", "name": "FW_{vendor}_{sku}_Alt1m_Vid6", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_Alt1m_Vid6", "File Type": "MP4", "Product SKU Position": "{sku}_260", "Deliverable": "Product Feature Video"},
        {"filename": "FW_{vendor}_{sku}_Alt1m_Vid7.mp4", "name": "FW_{vendor}_{sku}_Alt1m_Vid7", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_Alt1m_Vid7", "File Type": "MP4", "Product SKU Position": "{sku}_270", "Deliverable": "Product Feature Video"},
        {"filename": "FW_{vendor}_{sku}_Alt1m_Vid8.mp4", "name": "FW_{vendor}_{sku}_Alt1m_Vid8", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_Alt1m_Vid8", "File Type": "MP4", "Product SKU Position": "{sku}_280", "Deliverable": "Product Feature Video"},
        {"filename": "FW_{vendor}_{sku}_Alt1m_Vid9.mp4", "name": "FW_{vendor}_{sku}_Alt1m_Vid9", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_Alt1m_Vid9", "File Type": "MP4", "Product SKU Position": "{sku}_290", "Deliverable": "Product Feature Video"},
        {"filename": "FW_{vendor}_{sku}_VidInstruction1.mp4", "name": "FW_{vendor}_{sku}_VidInstruction1", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_VidInstruction1", "File Type": "MP4", "Product SKU Position": "{sku}_6000", "Deliverable": "Product Instruction Video"},
        {"filename": "FW_{vendor}_{sku}_VidInstruction2.mp4", "name": "FW_{vendor}_{sku}_VidInstruction2", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_VidInstruction2", "File Type": "MP4", "Product SKU Position": "{sku}_6100", "Deliverable": "Product Instruction Video"},
        {"filename": "FW_{vendor}_{sku}_VidInstruction3.mp4", "name": "FW_{vendor}_{sku}_VidInstruction3", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_VidInstruction3", "File Type": "MP4", "Product SKU Position": "{sku}_6200", "Deliverable": "Product Instruction Video"},
        {"filename": "FW_{vendor}_{sku}_VidInstruction4.mp4", "name": "FW_{vendor}_{sku}_VidInstruction4", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_VidInstruction4", "File Type": "MP4", "Product SKU Position": "{sku}_6300", "Deliverable": "Product Instruction Video"},
        {"filename": "FW_{vendor}_{sku}_VidInstruction5.mp4", "name": "FW_{vendor}_{sku}_VidInstruction5", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_VidInstruction5", "File Type": "MP4", "Product SKU Position": "{sku}_6400", "Deliverable": "Product Instruction Video"},
        {"filename": "FW_{vendor}_{sku}_VidInstruction6.mp4", "name": "FW_{vendor}_{sku}_VidInstruction6", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_VidInstruction6", "File Type": "MP4", "Product SKU Position": "{sku}_6500", "Deliverable": "Product Instruction Video"},
        {"filename": "FW_{vendor}_{sku}_VidInstruction7.mp4", "name": "FW_{vendor}_{sku}_VidInstruction7", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_VidInstruction7", "File Type": "MP4", "Product SKU Position": "{sku}_6600", "Deliverable": "Product Instruction Video"},
        {"filename": "FW_{vendor}_{sku}_VidInstruction8.mp4", "name": "FW_{vendor}_{sku}_VidInstruction8", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_VidInstruction8", "File Type": "MP4", "Product SKU Position": "{sku}_6700", "Deliverable": "Product Instruction Video"},
        {"filename": "FW_{vendor}_{sku}_VidInstruction9.mp4", "name": "FW_{vendor}_{sku}_VidInstruction9", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_VidInstruction9", "File Type": "MP4", "Product SKU Position": "{sku}_6800", "Deliverable": "Product Instruction Video"},
        {"filename": "FW_{vendor}_{sku}_A_Vid1.mp4", "name": "FW_{vendor}_{sku}_A_Vid1", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_A_Vid1", "File Type": "MP4", "Product SKU Position": "{sku}_110", "Deliverable": "Product Feature Video"},
        {"filename": "FW_{vendor}_{sku}_A_Vid2.mp4", "name": "FW_{vendor}_{sku}_A_Vid2", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_A_Vid2", "File Type": "MP4", "Product SKU Position": "{sku}_120", "Deliverable": "Product Feature Video"},
        {"filename": "FW_{vendor}_{sku}_A_Vid3.mp4", "name": "FW_{vendor}_{sku}_A_Vid3", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_A_Vid3", "File Type": "MP4", "Product SKU Position": "{sku}_130", "Deliverable": "Product Feature Video"},
        {"filename": "FW_{vendor}_{sku}_A_Vid4.mp4", "name": "FW_{vendor}_{sku}_A_Vid4", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_A_Vid4", "File Type": "MP4", "Product SKU Position": "{sku}_140", "Deliverable": "Product Feature Video"},
        {"filename": "FW_{vendor}_{sku}_A_Vid5.mp4", "name": "FW_{vendor}_{sku}_A_Vid5", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_A_Vid5", "File Type": "MP4", "Product SKU Position": "{sku}_150", "Deliverable": "Product Feature Video"},
        {"filename": "FW_{vendor}_{sku}_A_Vid6.mp4", "name": "FW_{vendor}_{sku}_A_Vid6", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_A_Vid6", "File Type": "MP4", "Product SKU Position": "{sku}_160", "Deliverable": "Product Feature Video"},
        {"filename": "FW_{vendor}_{sku}_A_Vid7.mp4", "name": "FW_{vendor}_{sku}_A_Vid7", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_A_Vid7", "File Type": "MP4", "Product SKU Position": "{sku}_170", "Deliverable": "Product Feature Video"},
        {"filename": "FW_{vendor}_{sku}_A_Vid8.mp4", "name": "FW_{vendor}_{sku}_A_Vid8", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_A_Vid8", "File Type": "MP4", "Product SKU Position": "{sku}_180", "Deliverable": "Product Feature Video"},
        {"filename": "FW_{vendor}_{sku}_A_Vid9.mp4", "name": "FW_{vendor}_{sku}_A_Vid9", "Download URL": "https://raymourflanigan.scene7.com/is/content/RaymourandFlanigan/FW_{vendor}_{sku}_A_Vid9", "File Type": "MP4", "Product SKU Position": "{sku}_190", "Deliverable": "Product Feature Video"},
        {"filename": "{vendor}_{sku}_3000.jpg", "name": "{vendor}_{sku}_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid", "Deliverable": "Product Grid Image"},
        {"filename": "{vendor}_{sku}_Alt1_3000.jpg", "name": "{vendor}_{sku}_Alt1_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt1_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid200", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt2_3000.jpg", "name": "{vendor}_{sku}_Alt2_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt2_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid300", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt3_3000.jpg", "name": "{vendor}_{sku}_Alt3_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt3_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid400", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt4_3000.jpg", "name": "{vendor}_{sku}_Alt4_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt4_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid500", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt5_3000.jpg", "name": "{vendor}_{sku}_Alt5_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt5_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid600", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt6_3000.jpg", "name": "{vendor}_{sku}_Alt6_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt6_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid700", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt7_3000.jpg", "name": "{vendor}_{sku}_Alt7_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt7_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid800", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt8_3000.jpg", "name": "{vendor}_{sku}_Alt8_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt8_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid900", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt9_3000.jpg", "name": "{vendor}_{sku}_Alt9_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt9_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid1000", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt9a_3000.jpg", "name": "{vendor}_{sku}_Alt9a_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt9a_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid1100", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt9b_3000.jpg", "name": "{vendor}_{sku}_Alt9b_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt9b_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid1200", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt9c_3000.jpg", "name": "{vendor}_{sku}_Alt9c_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt9c_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid1300", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt9d_3000.jpg", "name": "{vendor}_{sku}_Alt9d_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt9d_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid1400", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt9e_3000.jpg", "name": "{vendor}_{sku}_Alt9e_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt9e_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid1500", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt9f_3000.jpg", "name": "{vendor}_{sku}_Alt9f_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt9f_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid1600", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt9g_3000.jpg", "name": "{vendor}_{sku}_Alt9g_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt9g_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid1700", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt9h_3000.jpg", "name": "{vendor}_{sku}_Alt9h_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt9h_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid1800", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt9i_3000.jpg", "name": "{vendor}_{sku}_Alt9i_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt9i_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid1900", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt9j_3000.jpg", "name": "{vendor}_{sku}_Alt9j_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt9j_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid2000", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt9k_3000.jpg", "name": "{vendor}_{sku}_Alt9k_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt9k_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid2100", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt9l_3000.jpg", "name": "{vendor}_{sku}_Alt9l_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt9l_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid2200", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt9m_3000.jpg", "name": "{vendor}_{sku}_Alt9m_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt9m_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid2300", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt9n_3000.jpg", "name": "{vendor}_{sku}_Alt9n_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt9n_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid2400", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt9o_3000.jpg", "name": "{vendor}_{sku}_Alt9o_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt9o_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid2500", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt9p_3000.jpg", "name": "{vendor}_{sku}_Alt9p_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt9p_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid2600", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt9q_3000.jpg", "name": "{vendor}_{sku}_Alt9q_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt9q_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid2700", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt9r_3000.jpg", "name": "{vendor}_{sku}_Alt9r_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt9r_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid2800", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt9s_3000.jpg", "name": "{vendor}_{sku}_Alt9s_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt9s_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid2900", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt9t_3000.jpg", "name": "{vendor}_{sku}_Alt9t_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt9t_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid3000", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_Alt9u_3000.jpg", "name": "{vendor}_{sku}_Alt9u_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_Alt9u_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid3100", "Deliverable": "Product Grid Additional Image"},
        {"filename": "{vendor}_{sku}_swatch.jpg", "name": "{vendor}_{sku}_swatch", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_swatch", "File Type": "JPEG", "Product SKU Position": "{sku}_swatch", "Deliverable": "Product Swatch Image"},
        {"filename": "{vendor}_{sku}_s1_3000.jpg", "name": "{vendor}_{sku}_s1_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_s1_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid5000", "Deliverable": "Product Grid Swatch Detail Image"},
        {"filename": "{vendor}_{sku}_s2_3000.jpg", "name": "{vendor}_{sku}_s2_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_s2_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid5100", "Deliverable": "Product Grid Swatch Detail Image"},
        {"filename": "{vendor}_{sku}_s3_3000.jpg", "name": "{vendor}_{sku}_s3_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_s3_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid5200", "Deliverable": "Product Grid Swatch Detail Image"},
        {"filename": "{vendor}_{sku}_s4_3000.jpg", "name": "{vendor}_{sku}_s4_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_s4_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid5300", "Deliverable": "Product Grid Swatch Detail Image"},
        {"filename": "{vendor}_{sku}_s5_3000.jpg", "name": "{vendor}_{sku}_s5_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_s5_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid5400", "Deliverable": "Product Grid Swatch Detail Image"},
        {"filename": "{vendor}_{sku}_s6_3000.jpg", "name": "{vendor}_{sku}_s6_3000", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_s6_3000", "File Type": "JPEG", "Product SKU Position": "{sku}_grid5500", "Deliverable": "Product Grid Swatch Detail Image"},
        {"filename": "{vendor}_{sku}_dimension.jpg", "name": "{vendor}_{sku}_dimension", "Download URL": "https://raymourflanigan.scene7.com/is/image/RaymourandFlanigan/{vendor}_{sku}_dimension", "File Type": "JPEG", "Product SKU Position": "{sku}_dimension", "Deliverable": "Product Dimensions Diagram"}
    ]

    # Initialize lists to hold processed data and SKUs not found in STEP exports
    parsed_skus = set()  # Use a set to store unique SKUs that have been processed
    output_data = []     # List to accumulate all generated metadata rows
    missing_skus = []    # List to track SKUs not found in the STEP reference files

    # Get a list of all supported asset files (images and videos) in the input folder
    all_assets_in_folder = [f for f in os.listdir(input_folder) if f.lower().endswith(('.jpg', '.jpeg', '.png', '.mp4'))]
    total_assets = len(all_assets_in_folder)
    
    # Process each asset file in the input folder with progress reporting for UI
    print(f"Script: Processing {total_assets} assets in '{input_folder}'...")
    for i, filename in enumerate(all_assets_in_folder):
        # Extract vendor code and SKU from the current filename
        vendor_code, sku = extract_sku_and_vendor_from_filename(filename)
        
        # If SKU could not be extracted, skip this file
        if not sku:
            print(f"Script: Warning: Could not extract SKU/Vendor from filename: {filename}. Skipping.", file=sys.stderr)
            continue

        # Process metadata only once per unique SKU to avoid duplicate row sets
        if sku not in parsed_skus:
            parsed_skus.add(sku) # Add SKU to the set of processed SKUs

            # Look up the SKU in the combined reference DataFrame
            ref_row = df_refs[df_refs['SKU'] == sku]

            step_path = "" # Initialize STEP path as blank

            if not ref_row.empty:
                # If SKU is found in STEP data, use its Vendor Code and Path
                vendor_code_from_step = ref_row.iloc[0]['Vendor Code']
                # Prefer vendor code from STEP if it's not NaN, otherwise keep the one extracted from filename
                vendor_code = vendor_code_from_step if pd.notna(vendor_code_from_step) else vendor_code
                # Assign STEP path if it's not NaN, otherwise keep it blank
                step_path = ref_row.iloc[0]['Path'] if pd.notna(ref_row.iloc[0]['Path']) else ""
            else:
                # If SKU is not found in STEP data, add it to missing_skus list
                missing_skus.append(sku)
                print(f"Script: Warning: SKU '{sku}' not found in STEP exports. 'STEP Path' and 'Product Name (STEP)' will be blank in CSV.", file=sys.stderr)

            # Generate the set of metadata rows for the current SKU based on the template
            generated_rows = generate_rows(vendor_code, sku, step_path, template)
            for row in generated_rows:
                output_data.append(row)
        
        # --- Send Progress Update to UI ---
        # This line is critical for external UI applications to track progress.
        progress_percentage = (i + 1) / total_assets * 100
        print(f"PROGRESS:{progress_percentage:.2f}", flush=True) # print and flush immediately

    # Check if any data was processed before attempting to create CSV
    if not output_data:
        print("Script: No supported asset files found or processed with valid SKUs in the selected folder. No CSV will be generated.", file=sys.stderr)
        sys.exit(0)

    # Define the final ordered list of columns for the output CSV
    final_columns = [
        "filename", "name", "description", "Asset Type", "Asset Sub-Type", "Deliverable", "Product SKU",
        "Product SKU Position", "Asset Status", "Usage Rights", "tags", "File Type", "STEP Path",
        "Link to Wrike Project", "Sync to Site", "Generic Dimension Diagram With Measurements", "Admin Status",
        "Product Status", "Product Category", "Product Sub-Category", "Product Collection",
        "Component SKUs", "Stock Level (only relevant for Inline products)", "Restock Date (only relevant for Inline products)",
        "Link to Print Materials", "Link to Lifestyle Images", "Link to Store Images", "Initiative", "Sub-Initiative",
        "Print Tracking Code", "Print Tracking - Start Date", "Print Tracking - End Date", "Year", "Video Expiration",
        "Audio Licensing Expiration", "Ad ID", "Lead Offer Message", "Lead Finance Message", "Video Focus",
        "Video Objective", "Video Type", "Total Run Time (TRT)", "Spot Running (MM/DD/YYYY)", "Language",
        "Season", "Holiday/Special Occasion", "Talent", "Sunset Date (MM/DD/YYYY)", "Location Name", "Store Code",
        "Location Status", "Location Address", "Location Town", "Location State", "Location Zip Code",
        "Location Phone Number", "Location Type", "Location", "Inactive Product", "Partner", "Notes",
        "Sign Facade Color", "Sign Location", "Sign Color", "Sign Text",
        "Reviewed products in lifestyle", "Reviewed Studio Uploads", "Featured SKU", "Image Type",
        "scratchpad", "3D Model Source Files Acquired", "Visible to", "BynderTest", "dim_Length",
        "Bynder Report", "Dimensions", "dim_Height", "Figmage doc id", "dim_Width", "Figmage image extension",
        "Figmage node id", "Figmage page id", "Performance Metric", "DNUCampaign", "DNUFeatures",
        "DNUMaterials", "DNUStyle", "DNUPattern", "DNUPackage SKUs", "DNUSign Size",
        "DNUDistribution Channel", "Dim diagram re-cropped", "Embedded Instructions (for updating existing metadata based on automations)", "Mattress Size", "Asset Identifier", "Sync Batch", "Marked for Deletion from Site", "scene7 folder", "Variant Type", "Source", "PSA Image Type", "Rights Notes", "Workflow", "Workflow Status",
        "Product Name (STEP)", "Vendor Code", "Family Code", "Hero SKU", "Product Color", "Dropped",
        "Visible on Website", "Sales Channel", "Associated Materials Status", "Product in Studio",
        "DNU_PromoUpdate2", "Additional Files Upload Scratchpad", "Bump", "Carousel Dimensions Diagram Audit", "User Status", "Reviewed for Site Content Refresh", "Image Type Pre-Classification"
    ]
    # Create the final pandas DataFrame from the accumulated data
    output_df = pd.DataFrame(output_data, columns=final_columns)

    try:
        # Save the DataFrame to a semicolon-separated CSV file
        output_df.to_csv(output_file, sep=';', index=False)
        print(f"Script: Your metadata importer sheet is ready in your Downloads folder! File saved as: {output_file}")
    except Exception as e:
        # Handle errors during saving the CSV file
        print(f"Script: Error saving CSV file to {output_file}: {e}", file=sys.stderr)
        sys.exit(1)

    # Report any SKUs that were found in filenames but not in the STEP exports
    if missing_skus:
        print("\n--- Script: SKUs not found in the STEP exports: ---")
        # Sort and print unique missing SKUs for readability
        for sku in sorted(list(set(missing_skus))):
            print(sku)
        print("Script: These SKUs will have a blank 'STEP Path' and 'Product Name (STEP)' in the CSV.")

# Entry point for the script
if __name__ == "__main__":
    main()
