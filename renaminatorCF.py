import os
import sys
import pandas as pd
import shutil
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog
import argparse # New import for command-line arguments

def print_progress(message, is_stderr=False):
    """
    Prints messages that the UI can capture.
    For errors, direct to sys.stderr.
    Includes a special format for progress percentage updates: "PROGRESS: <float_percentage>"
    """
    if message.startswith("PROGRESS:"):
        print(message, flush=True) # Ensure progress updates are sent immediately
    elif is_stderr:
        print(message, file=sys.stderr, flush=True) # Flush stderr immediately too
    else:
        print(message, flush=True) # Flush all regular messages immediately

def prompt_for_file_tk(prompt_message, filetypes):
    """Prompt the user to select a file using a Tkinter dialog."""
    root = tk.Tk()
    root.withdraw() # Hide the main window
    file_path = filedialog.askopenfilename(title=prompt_message, filetypes=filetypes)
    root.destroy() # Close the Tkinter root window
    return file_path

def prompt_for_folder_tk(prompt_message):
    """Prompt the user to select a folder using a Tkinter dialog."""
    root = tk.Tk()
    root.withdraw() # Hide the main window
    folder_path = filedialog.askdirectory(title=prompt_message)
    root.destroy() # Close the Tkinter root window
    return folder_path

def get_downloads_folder():
    """Returns the user's Downloads folder, or home directory as a fallback."""
    home = Path.home()
    downloads = home / "Downloads"
    return str(downloads) if downloads.exists() else str(home)

def build_stem_map(root_folder):
    """
    Walk through root_folder (including subfolders) once,
    and build a dict: { lowercased stripped stem : full_path_to_file }.
    Strips leading/trailing spaces from the stem so that both
    "N9008000CK-134 " and "N9008000CK-134" map to the same key.
    """
    stem_map = {}
    for dirpath, _, filenames in os.walk(root_folder):
        for fname in filenames:
            # Split off extension, then strip whitespace from the stem
            raw_stem = os.path.splitext(fname)[0]
            stem = raw_stem.strip().lower()
            if stem and stem not in stem_map:
                stem_map[stem] = os.path.join(dirpath, fname)
    return stem_map

def main():
    parser = argparse.ArgumentParser(description="Renamer File Copier Script")
    parser.add_argument('--matrix', help='Path to the Excel spreadsheet containing filenames.')
    parser.add_argument('--input', help='Path to the input folder containing source files (will search subfolders).')
    parser.add_argument('--output', help='Path to the output folder where files will be copied.')
    
    args = parser.parse_args()

    spreadsheet_path = None
    source_folder = None
    output_folder = None

    print_progress("PROGRESS: 0.0") # Initial progress

    # Determine if arguments were provided (from GUI) or if we need to prompt the user (standalone)
    if not args.matrix or not args.input or not args.output:
        print_progress("\n--- Running Renamer File Copier in Standalone Mode ---")
        print_progress("Please provide the required paths using graphical dialogs:")
        
        spreadsheet_path = prompt_for_file_tk(
            "Select the Excel spreadsheet (header row will be ignored):",
            [("Excel files", "*.xlsx *.xls")]
        )
        if not spreadsheet_path:
            print_progress("No spreadsheet selected. Exiting.", is_stderr=True)
            print_progress("PROGRESS: 0.0", is_stderr=True) # Indicate failure
            sys.exit(1)

        source_folder = prompt_for_folder_tk(
            "Select folder containing all source files (will search subfolders):"
        )
        if not source_folder:
            print_progress("No source folder selected. Exiting.", is_stderr=True)
            print_progress("PROGRESS: 0.0", is_stderr=True) # Indicate failure
            sys.exit(1)
        
        output_folder = prompt_for_folder_tk(
            "Select output folder for copied files:"
        )
        if not output_folder:
            print_progress("No output folder selected. Exiting.", is_stderr=True)
            print_progress("PROGRESS: 0.0", is_stderr=True) # Indicate failure
            sys.exit(1)
    else:
        # Arguments were provided, likely from the GUI
        spreadsheet_path = args.matrix
        source_folder = args.input
        output_folder = args.output

    print_progress(f"\nRenamer File Copier script starting with:")
    print_progress(f"  Matrix: {spreadsheet_path}")
    print_progress(f"  Source Folder: {source_folder}")
    print_progress(f"  Output Folder: {output_folder}")

    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Read the spreadsheet, skipping the header row (use header=None and start at row_idx=1)
    try:
        print_progress("Loading spreadsheet...")
        df = pd.read_excel(spreadsheet_path, header=None)
    except FileNotFoundError:
        print_progress(f"ERROR: Spreadsheet not found at {spreadsheet_path}. Exiting.", is_stderr=True)
        print_progress("PROGRESS: 0.0", is_stderr=True)
        sys.exit(1)
    except Exception as e:
        print_progress(f"ERROR: Failed to read spreadsheet: {e}. Exiting.", is_stderr=True)
        print_progress("PROGRESS: 0.0", is_stderr=True)
        sys.exit(1)

    print_progress("PROGRESS: 10.0") # After loading spreadsheet

    # Build a lookup map from “lowercased, stripped stem” → full path
    print_progress(f"Building file lookup map from: {source_folder} (this may take a moment for large folders)...")
    stem_map = build_stem_map(source_folder)
    print_progress(f"Lookup map built. Found {len(stem_map)} unique file stems.")

    print_progress("PROGRESS: 40.0") # After building stem map (can be a large step)

    errors = []
    files_copied_count = 0
    total_entries_to_process = 0
    entries_processed_count = 0

    # First, count all relevant entries to determine total work
    for row_idx in range(1, df.shape[0]): # Start from row 1 (second row in Excel, assuming header is row 0)
        for col_idx in range(2, 11): # Columns C through K (indices 2-10)
            cell_value = df.iat[row_idx, col_idx]
            if pd.notna(cell_value) and str(cell_value).strip():
                total_entries_to_process += 1
    
    if total_entries_to_process == 0:
        print_progress("No filenames found in the specified columns (C-K) of the spreadsheet to copy. Exiting.")
        print_progress("PROGRESS: 100.0") # Complete immediately if nothing to do
        sys.exit(0)

    print_progress(f"Found {total_entries_to_process} entries to process for copying.")

    # Iterate over each cell in columns C (idx=2) through K (idx=10), skipping header at idx=0
    # Excel columns C-K correspond to pandas indices 2-10
    for row_idx in range(1, df.shape[0]): # Start from row 1 (second row in Excel, assuming header is row 0)
        for col_idx in range(2, 11): # Columns C through K
            cell_value = df.iat[row_idx, col_idx]
            if pd.notna(cell_value):
                raw_name = str(cell_value)
                # Strip whitespace around the matrix entry
                name = raw_name.strip()
                if not name: # Skip empty strings after stripping
                    continue

                entries_processed_count += 1 # Increment counter for every non-empty, non-whitespace cell

                # Compute the Excel‐style cell coordinate, e.g. "H2"
                excel_row = row_idx + 1 # +1 because pandas is 0-indexed, Excel is 1-indexed
                col_letter = chr(ord('A') + col_idx)
                matrix_location = f"{col_letter}{excel_row}"

                # Strip any extension from the matrix entry, then strip whitespace again, and lowercase
                base_stem = os.path.splitext(name)[0].strip().lower()

                print_progress(f"Processing matrix entry {matrix_location}: '{name}' (looking for stem '{base_stem}')")

                if base_stem in stem_map:
                    source_path = stem_map[base_stem]
                    dest_path = os.path.join(output_folder, os.path.basename(source_path))
                    try:
                        shutil.copy2(source_path, dest_path)
                        files_copied_count += 1
                        print_progress(f"  Copied: '{os.path.basename(source_path)}' to '{output_folder}'")
                    except Exception as e:
                        print_progress(f"  ERROR copying '{source_path}': {e}", is_stderr=True)
                        errors.append({
                            "Matrix Location": matrix_location,
                            "File Name": raw_name,
                            "Error": str(e)
                        })
                else:
                    print_progress(f"  File not found for stem '{base_stem}' (from '{name}')", is_stderr=True)
                    errors.append({
                        "Matrix Location": matrix_location,
                        "File Name": raw_name,
                        "Error": "File not found in source folder"
                    })
                
                # Update progress based on entries processed
                current_progress = 40.0 + (entries_processed_count / total_entries_to_process) * 50.0
                print_progress(f"PROGRESS: {min(current_progress, 99.9):.1f}") # Cap at 99.9 before final 100

    # Write a report if any files were missing or couldn’t be copied
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_name = f"renaminatorCF_report_{timestamp}.xlsx"
    downloads_folder = get_downloads_folder()
    report_path = os.path.join(downloads_folder, report_name)

    print_progress(f"\n--- File Copy Summary ---")
    print_progress(f"Total entries processed in matrix: {total_entries_to_process}")
    print_progress(f"Files successfully copied: {files_copied_count}")
    print_progress(f"Files with errors/not found: {len(errors)}")

    if errors:
        report_df = pd.DataFrame(errors)
        try:
            # Define columns explicitly to ensure order and consistency in report
            report_df.to_excel(report_path, index=False, columns=["Matrix Location", "File Name", "Error"])
            print_progress(f"Completed with errors/missing files. See the detailed report at: {report_path}")
        except Exception as e:
            print_progress(f"ERROR: Failed to write error report to '{report_path}': {e}. Exiting.", is_stderr=True)
            print_progress("PROGRESS: 100.0", is_stderr=True) # Final progress update on error
            sys.exit(1) # Exit with an error code if there were failures
        print_progress("PROGRESS: 100.0") # Final progress update after report generation
        sys.exit(1) # Exit with an error code if there were failures
    else:
        print_progress(f"All files specified in the matrix were found and copied to {output_folder}!")
        print_progress("PROGRESS: 100.0") # Final progress update on success
        sys.exit(0) # Exit with success code

if __name__ == "__main__":
    main()
