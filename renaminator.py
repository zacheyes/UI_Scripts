import os
import shutil
import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import csv
import argparse

def print_progress(message, is_stderr=False):
    """
    Prints messages that the UI can capture.
    For progress updates, prepend with "PROGRESS:".
    For errors, direct to sys.stderr.
    """
    if is_stderr:
        print(message, file=sys.stderr)
    else:
        print(message)
    sys.stdout.flush() # Ensure output is sent immediately

def prompt_for_file():
    """Prompt the user to select an Excel file and return its path."""
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select Excel Spreadsheet",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not file_path:
        print_progress("No spreadsheet selected. Exiting.", is_stderr=True)
        sys.exit(1)
    return file_path

def prompt_for_folder():
    """Prompt the user to select a folder and return its path."""
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="Select Folder Containing JPG Files")
    if not folder_path:
        print_progress("No folder selected. Exiting.", is_stderr=True)
        sys.exit(1)
    return folder_path

def is_jpg_folder(folder_path):
    """
    Check that every file in the given folder ends with .jpg (case-insensitive)
    and ignore hidden files (starting with '.').
    Return False if any non-JPG, non-hidden file is found.
    """
    for fname in os.listdir(folder_path):
        if fname.startswith('.'): # Ignore hidden files
            continue
        if not fname.lower().endswith('.jpg'):
            return False
    return True

def build_sku_data(df):
    """
    Parse the DataFrame to build a mapping from SKU to:
      - main images (columns C–H across all rows for that SKU) as base names (no extension)
      - swatch (column I, first non-empty per SKU) as base name
      - s_images (column J, one per row for that SKU) as base names
      - dimensions (column K, possibly multiple, as base names)
    Returns:
      sku_order: list of SKUs in order of first appearance
      sku_data: dict mapping SKU to {
          'images': [(base_name, cell), ...],
          'swatch': (base_name, cell) or None,
          's_images': [(base_name, cell), ...],
          'dimensions': [(base_name, cell), ...]
      }
    """
    from collections import OrderedDict
    sku_data = OrderedDict()

    # Mapping of DataFrame column index to Excel column letter
    col_letters = {
        2: 'C', 3: 'D', 4: 'E', 5: 'F', 6: 'G', 7: 'H',
        8: 'I', 9: 'J', 10: 'K'
    }

    for idx, row in df.iterrows():
        excel_row = idx + 2  # header is row 1
        raw_sku = row.iloc[1]  # column B
        if pd.isna(raw_sku) or str(raw_sku).strip() == '':
            continue
        sku = str(raw_sku).strip()

        if sku not in sku_data:
            sku_data[sku] = {
                'images': [],
                'swatch': None,
                's_images': [],
                'dimensions': []
            }
        entry = sku_data[sku]

        # Process columns C–H for main images
        for col_idx in range(2, 8):
            raw_val = row.iloc[col_idx]
            if pd.isna(raw_val) or str(raw_val).strip() == '':
                continue
            # Build base name: strip whitespace, remove ".jpg" if present, then rstrip any leftover spaces
            base = str(raw_val)
            base = base.strip()
            if base.lower().endswith('.jpg'):
                base = base[:-4]
            base = base.rstrip()
            entry['images'].append((base, f"{col_letters[col_idx]}{excel_row}"))

        # Process column I (swatch) – only first non-empty per SKU
        raw_swatch = row.iloc[8]
        if (not pd.isna(raw_swatch)) and str(raw_swatch).strip() != '' and entry['swatch'] is None:
            base = str(raw_swatch).strip()
            if base.lower().endswith('.jpg'):
                base = base[:-4]
            base = base.rstrip()
            entry['swatch'] = (base, f"I{excel_row}")

        # Process column J (s_images) – can be multiple per SKU
        raw_s = row.iloc[9]
        if (not pd.isna(raw_s)) and str(raw_s).strip() != '':
            base = str(raw_s).strip()
            if base.lower().endswith('.jpg'):
                base = base[:-4]
            base = base.rstrip()
            entry['s_images'].append((base, f"J{excel_row}"))

        # Process column K (dimensions) – possibly multiple per SKU
        raw_dim = row.iloc[10]
        if (not pd.isna(raw_dim)) and str(raw_dim).strip() != '':
            base = str(raw_dim).strip()
            if base.lower().endswith('.jpg'):
                base = base[:-4]
            base = base.rstrip()
            entry['dimensions'].append((base, f"K{excel_row}"))

    return list(sku_data.keys()), sku_data

def check_and_resolve_filenames(input_folder, sku_order, sku_data):
    """
    For every (base, cell) in sku_data, attempt to find an actual filename in input_folder:
      - Try "base.jpg"
      - If not found, try "base .jpg"
    If found, replace (base, cell) with (actual_filename, cell) in sku_data.
    Otherwise, collect missing = [(cell, base+".jpg"), ...].
    """
    missing = []
    for sku in sku_order:
        entry = sku_data[sku]

        # Resolve main images
        for i, (base, cell) in enumerate(entry['images']):
            fname1 = f"{base}.jpg"
            fname2 = f"{base} .jpg"
            path1 = os.path.join(input_folder, fname1)
            path2 = os.path.join(input_folder, fname2)
            if os.path.isfile(path1):
                entry['images'][i] = (fname1, cell)
            elif os.path.isfile(path2):
                entry['images'][i] = (fname2, cell)
            else:
                missing.append((cell, f"{base}.jpg"))

        # Resolve swatch (if present)
        if entry['swatch'] is not None:
            base, cell = entry['swatch']
            fname1 = f"{base}.jpg"
            fname2 = f"{base} .jpg"
            path1 = os.path.join(input_folder, fname1)
            path2 = os.path.join(input_folder, fname2)
            if os.path.isfile(path1):
                entry['swatch'] = (fname1, cell)
            elif os.path.isfile(path2):
                entry['swatch'] = (fname2, cell)
            else:
                missing.append((cell, f"{base}.jpg"))

        # Resolve s_images
        for i, (base, cell) in enumerate(entry['s_images']):
            fname1 = f"{base}.jpg"
            fname2 = f"{base} .jpg"
            path1 = os.path.join(input_folder, fname1)
            path2 = os.path.join(input_folder, fname2)
            if os.path.isfile(path1):
                entry['s_images'][i] = (fname1, cell)
            elif os.path.isfile(path2):
                entry['s_images'][i] = (fname2, cell)
            else:
                missing.append((cell, f"{base}.jpg"))

        # Resolve dimensions
        for i, (base, cell) in enumerate(entry['dimensions']):
            fname1 = f"{base}.jpg"
            fname2 = f"{base} .jpg"
            path1 = os.path.join(input_folder, fname1)
            path2 = os.path.join(input_folder, fname2)
            if os.path.isfile(path1):
                entry['dimensions'][i] = (fname1, cell)
            elif os.path.isfile(path2):
                entry['dimensions'][i] = (fname2, cell)
            else:
                missing.append((cell, f"{base}.jpg"))

    return missing

def write_missing_csv(missing_list):
    """
    Write the missing entries to a CSV in the user's Downloads folder.
    """
    user_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    if not os.path.isdir(user_downloads):
        user_downloads = os.getcwd()

    report_path = os.path.join(user_downloads, "missing_files_report.csv")
    with open(report_path, mode='w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(["Cell", "Expected File"])
        for cell, fname in missing_list:
            writer.writerow([cell, fname])
    print_progress(f"Missing files report exported to: {report_path}")

def generate_alt_suffix(k):
    """
    Given k (1-based index for additional images beyond the first), return the Alt suffix.
    For k = 1..9: return 'Alt{1..9}'
    For k >= 10: calculate n = k - 10, build letter_sequence:
      letter_sequence = ('z' * (floor(n/26))) + chr(ord('a') + (n % 26))
      suffix = 'Alt9' + letter_sequence
    """
    if k <= 9:
        return f"Alt{k}"
    else:
        n = k - 10
        repeat = (n // 26) + 1
        letter_seq = 'z' * (repeat - 1) + chr(ord('a') + (n % 26))
        return f"Alt9{letter_seq}"

def main():
    parser = argparse.ArgumentParser(description="Rename image files based on an Excel spreadsheet.")
    parser.add_argument('--matrix', type=str, help='Path to the Excel Renamer Matrix file.')
    parser.add_argument('--input', type=str, help='Path to the folder containing input JPG files.')
    parser.add_argument('--vendor_code', type=str, help='Vendor Code for the project.')
    parser.add_argument('--force_continue', action='store_true',
                        help='Force the script to continue even if non-JPG files are found or files are missing.')
    args = parser.parse_args()

    # Determine input paths and vendor code based on arguments or interactive prompts
    excel_path = args.matrix
    if not excel_path:
        print_progress("No matrix file provided. Prompting user...")
        excel_path = prompt_for_file()
        print_progress(f"Using matrix file: {excel_path}")

    image_folder = args.input
    if not image_folder:
        print_progress("No input folder provided. Prompting user...")
        image_folder = prompt_for_folder()
        print_progress(f"Using image folder: {image_folder}")

    vendor_code = args.vendor_code
    if not vendor_code:
        print_progress("No vendor code provided. Prompting user...")
        vendor_code = input("Enter Vendor Code for this project: ").strip()
        if vendor_code == "":
            print_progress("Vendor Code cannot be blank. Exiting.", is_stderr=True)
            sys.exit(1)
        print_progress(f"Using vendor code: {vendor_code}")

    # 4. Check that all files in the folder are JPGs
    print_progress("Checking if all files in the input folder are JPGs...")
    if not is_jpg_folder(image_folder):
        msg = "Not all the files in this folder are JPGs. Please fix."
        if not args.force_continue:
            print_progress(msg, is_stderr=True)
            sys.exit(1)
        else:
            print_progress(f"WARNING: {msg} Continuing due to --force_continue.", is_stderr=True)
    print_progress("All files are JPGs (or force_continue enabled).")

    # 5. Read the Excel file into a DataFrame
    print_progress(f"Reading Excel file: {excel_path}...")
    try:
        df = pd.read_excel(excel_path, engine='openpyxl')
        if df.empty:
            print_progress("The Excel spreadsheet is empty. Exiting.", is_stderr=True)
            sys.exit(1)
    except Exception as e:
        print_progress(f"Error reading Excel file: {e}", is_stderr=True)
        sys.exit(1)
    print_progress("Excel file loaded successfully.")

    # 6. Build SKU data structures (still storing base names, no extension)
    print_progress("Building SKU data from matrix...")
    sku_order, sku_data = build_sku_data(df)
    if not sku_order:
        print_progress("No valid SKUs found in the spreadsheet. Exiting.", is_stderr=True)
        sys.exit(1)
    print_progress(f"Found {len(sku_order)} SKUs.")

    # 7. Attempt to resolve each base name to an actual filename in the folder
    print_progress("Checking for missing files...")
    missing = check_and_resolve_filenames(image_folder, sku_order, sku_data)
    if missing:
        print_progress("The following referenced files could not be found:", is_stderr=True)
        for cell, fname in missing:
            print_progress(f"  Cell {cell}: expected '{fname}'", is_stderr=True)
        write_missing_csv(missing)
        if not args.force_continue:
            sys.exit(1)
        else:
            print_progress("WARNING: Missing files detected. Continuing due to --force_continue.", is_stderr=True)
    else:
        print_progress("All referenced files found.")

    # 8. All checks passed: prompt user to proceed (only if not forced by UI)
    if not args.matrix: # Only prompt if running standalone
        yn = input(
            "All files in your folders are JPGs and I have found all the files from your matrix.\n"
            "Ready to proceed to copying and renaming? (Y/N): "
        ).strip().lower()
        if not yn.startswith('y'):
            print_progress("Operation cancelled by user. Exiting.")
            sys.exit(0)
    else:
        print_progress("Proceeding with copying and renaming (launched from UI).")

    # 9. Create output folder "RenamedImages" inside the image folder
    output_folder = os.path.join(image_folder, "RenamedImages")
    print_progress(f"Creating output folder: {output_folder}")
    os.makedirs(output_folder, exist_ok=True)

    # 10. Iterate through each SKU and perform copying/renaming
    total_skus = len(sku_order)
    for i, sku in enumerate(sku_order):
        print_progress(f"PROGRESS: {((i+1)/total_skus) * 100:.2f}") # Send progress to UI
        print_progress(f"Processing SKU: {sku}")
        entry = sku_data[sku]
        images_list = entry['images']
        if not images_list:
            print_progress(f"  No main images found for SKU {sku}. Skipping.", is_stderr=True)
            continue  # no main images for this SKU

        # 10a. First image: copy & rename as "FW_[Vendor]_[SKU]_3000.jpg"
        first_image_name, _ = images_list[0]
        src_first = os.path.join(image_folder, first_image_name)
        dest_first = os.path.join(output_folder, f"FW_{vendor_code}_{sku}_3000.jpg")
        print_progress(f"  Copying {first_image_name} to {os.path.basename(dest_first)}")
        shutil.copy2(src_first, dest_first)

        # 10b. FS check: build FS filename and copy additional
        base, _ = os.path.splitext(first_image_name)
        fs_candidate = f"{base}_FS.jpg"
        fs_path = os.path.join(image_folder, fs_candidate)
        if os.path.isfile(fs_path):
            src_for_fs = fs_path
            print_progress(f"  Found FS image: {fs_candidate}")
        else:
            src_for_fs = src_first
            print_progress(f"  No dedicated FS image found, using main image for FS: {first_image_name}")
        dest_fs = os.path.join(output_folder, f"{vendor_code}_{sku}_3000.jpg")
        print_progress(f"  Copying to {os.path.basename(dest_fs)}")
        shutil.copy2(src_for_fs, dest_fs)

        # 10c. Additional images: rename with Alt suffixes
        for idx in range(1, len(images_list)):
            img_name, _ = images_list[idx]
            src_img = os.path.join(image_folder, img_name)
            k = idx  # k=1 for second image, etc.
            alt_suffix = generate_alt_suffix(k)
            dest_name = f"FW_{vendor_code}_{sku}_{alt_suffix}_3000.jpg"
            print_progress(f"  Copying {img_name} to {dest_name}")
            shutil.copy2(src_img, os.path.join(output_folder, dest_name))

        # 10d. Swatch: copy & rename "[Vendor]_[SKU]_swatch.jpg" (if present)
        if entry['swatch'] is not None:
            swatch_name, _ = entry['swatch']
            src_swatch = os.path.join(image_folder, swatch_name)
            dest_swatch = os.path.join(output_folder, f"{vendor_code}_{sku}_swatch.jpg")
            print_progress(f"  Copying swatch {swatch_name} to {os.path.basename(dest_swatch)}")
            shutil.copy2(src_swatch, dest_swatch)

        # 10e. s_images: copy each and rename "FW_[Vendor]_[SKU]_s{n}_3000.jpg"
        for s_idx, (s_name, _) in enumerate(entry['s_images'], start=1):
            src_s = os.path.join(image_folder, s_name)
            dest_s = os.path.join(output_folder, f"FW_{vendor_code}_{sku}_s{s_idx}_3000.jpg")
            print_progress(f"  Copying s-image {s_name} to {os.path.basename(dest_s)}")
            shutil.copy2(src_s, dest_s)

        # 10f. Dimensions: copy each and rename
        #      First dimension → "[Vendor]_[SKU]_dimension.jpg"
        #      Subsequent → "[Vendor]_[SKU]_dimension2.jpg", etc.
        for d_idx, (dim_name, _) in enumerate(entry['dimensions'], start=1):
            src_dim = os.path.join(image_folder, dim_name)
            if d_idx == 1:
                dest_dim = os.path.join(output_folder, f"{vendor_code}_{sku}_dimension.jpg")
            else:
                dest_dim = os.path.join(output_folder, f"{vendor_code}_{sku}_dimension{d_idx}.jpg")
            print_progress(f"  Copying dimension {dim_name} to {os.path.basename(dest_dim)}")
            shutil.copy2(src_dim, dest_dim)

    print_progress(f"Done! All files copied and renamed into:\n  {output_folder}")
    print_progress("PROGRESS: 100.00") # Final progress update

if __name__ == "__main__":
    main()