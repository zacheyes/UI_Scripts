import os
import shutil
import pandas as pd
import argparse
import sys
import tkinter as tk
from tkinter import filedialog, TclError # Added TclError

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

def prompt_for_path(is_file=True, title_msg="Select File", file_types=None):
    """
    Prompt the user to select a file or folder using a Tkinter dialog.
    Returns the selected path or None if cancelled.
    """
    try:
        root = tk.Tk()
        root.withdraw() # Hide the main Tkinter window
        if is_file:
            path = filedialog.askopenfilename(title=title_msg, filetypes=file_types)
        else:
            path = filedialog.askdirectory(title=title_msg)
        root.destroy() # Destroy the hidden root window
        return path
    except TclError:
        print_progress("Warning: Tkinter GUI not available. Cannot prompt. Script might require direct arguments.", is_stderr=True)
        return None
    except Exception as e:
        print_progress(f"Error opening dialog: {e}", is_stderr=True)
        return None

def _get_filenames_from_input_file(filenames_file_path):
    """
    Reads filenames from an Excel spreadsheet file (.xlsx) or a plain text file (.txt).
    Returns a list of unique filenames.
    """
    filenames = []
    if not filenames_file_path or not os.path.exists(filenames_file_path):
        print_progress(f"Error: Filenames file not found or path is empty: {filenames_file_path}", is_stderr=True)
        return None

    file_extension = os.path.splitext(filenames_file_path)[1].lower()

    try:
        if file_extension in ('.xlsx', '.xls'):
            print_progress(f"Reading filenames from Excel file: {filenames_file_path}")
            # Load the Excel file and get filenames from the first column (iloc[:, 0])
            # Use header=None to ensure the first row is read as data, not as a header.
            df = pd.read_excel(filenames_file_path, header=None, dtype=str)
            filenames = df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
        elif file_extension == '.txt':
            print_progress(f"Reading filenames from text file: {filenames_file_path}")
            with open(filenames_file_path, 'r', encoding='utf-8') as f:
                filenames = [line.strip() for line in f if line.strip()]
        else:
            print_progress(f"Error: Unsupported filenames file type: {file_extension}. Please provide an .xlsx or .txt file.", is_stderr=True)
            return None
    except Exception as e:
        print_progress(f"Error reading filenames from '{filenames_file_path}': {e}", is_stderr=True)
        return None
    
    if not filenames:
        print_progress("No valid filenames found in the provided input file.", is_stderr=True)
        return None
    
    return list(pd.Series(filenames).unique()) # Return unique filenames


def main():
    parser = argparse.ArgumentParser(description="Move files based on a list of filenames.")
    parser.add_argument('--filenames_file', type=str, help='Path to a file (Excel .xlsx or Text .txt) containing filenames in the first column or one per line.')
    # Removed: --filenames_list
    parser.add_argument('--source_folder', type=str, help='Path to the source folder.')
    parser.add_argument('--destination_folder', type=str, help='Path to the destination folder.')

    args = parser.parse_args()

    # --- Determine Source Folder ---
    source_folder = args.source_folder
    if not source_folder:
        print_progress("No source folder provided. Prompting user to select source folder...")
        source_folder = prompt_for_path(is_file=False, title_msg="Select Source Folder")
        if not source_folder:
            print_progress("Source folder selection cancelled. Exiting.", is_stderr=True)
            sys.exit(0) # Exit gracefully on cancel

    # --- Determine Destination Folder ---
    destination_folder = args.destination_folder
    if not destination_folder:
        print_progress("No destination folder provided. Prompting user to select destination folder...")
        destination_folder = prompt_for_path(is_file=False, title_msg="Select Destination Folder")
        if not destination_folder:
            print_progress("Destination folder selection cancelled. Exiting.", is_stderr=True)
            sys.exit(0) # Exit gracefully on cancel

    # Validate source folder existence
    if not os.path.exists(source_folder):
        print_progress(f"Error: Source folder not found: {source_folder}", is_stderr=True)
        sys.exit(1)

    # Ensure the destination folder exists, create it if it doesn't
    if not os.path.exists(destination_folder):
        try:
            os.makedirs(destination_folder)
            print_progress(f"Created destination folder: {destination_folder}")
        except OSError as e:
            print_progress(f"Error creating destination folder {destination_folder}: {e}", is_stderr=True)
            sys.exit(1)

    # --- Determine Filenames Input Source ---
    filenames_file_path = args.filenames_file
    if not filenames_file_path:
        print_progress("No filenames file provided. Prompting user to select a file with filenames...")
        filenames_file_path = prompt_for_path(
            is_file=True,
            title_msg="Select File with Filenames (.xlsx or .txt)",
            file_types=[("Excel files", "*.xlsx *.xls"), ("Text files", "*.txt"), ("All files", "*.*")]
        )
        if not filenames_file_path:
            print_progress("Filenames file selection cancelled. Exiting.", is_stderr=True)
            sys.exit(0) # Exit gracefully on cancel

    files_to_move = _get_filenames_from_input_file(filenames_file_path)

    if files_to_move is None or not files_to_move:
        print_progress("No valid filenames found in the input file. Exiting.", is_stderr=True)
        sys.exit(1)

    total_files = len(files_to_move)
    moved_count = 0

    print_progress(f"Attempting to move {total_files} files from '{source_folder}' to '{destination_folder}'.")

    for i, file_name in enumerate(files_to_move):
        # Strip any leading/trailing whitespace from the filename
        file_name = file_name.strip()
        source_path = os.path.join(source_folder, file_name)
        destination_path = os.path.join(destination_folder, file_name)

        print_progress(f"Processing: '{file_name}'")
        # print_progress(f"  Source Path Attempt: '{source_path}'") # Too much debug for regular run
        # print_progress(f"  Destination Path Attempt: '{destination_path}'") # Too much debug for regular run

        if os.path.exists(source_path):
            try:
                shutil.move(source_path, destination_path)
                print_progress(f"  Moved: '{file_name}'")
                moved_count += 1
            except shutil.Error as se:
                print_progress(f"  ERROR: Failed to move '{file_name}' (shutil error): {se}", is_stderr=True)
            except OSError as oe:
                print_progress(f"  ERROR: Failed to move '{file_name}' (OS error): {oe}", is_stderr=True)
            except Exception as e:
                print_progress(f"  ERROR: Failed to move '{file_name}' (unexpected error): {e}", is_stderr=True)
        else:
            print_progress(f"  File NOT FOUND at source: '{source_path}'", is_stderr=True)

        # Report progress to stdout, which the GUI will capture
        progress_percentage = (i + 1) / total_files * 100
        print_progress(f"PROGRESS:{progress_percentage:.2f}")

    print_progress(f"\n--- File Movement Summary ---")
    print_progress(f"Total files attempted: {total_files}")
    print_progress(f"Files successfully moved: {moved_count}")
    print_progress(f"Files not found at source or failed to move: {total_files - moved_count}")
    print_progress("--- End Script Run ---\n")

    # The cleanup of any temporary file generated by the GUI is handled by the GUI.
    # This script should not delete `args.filenames_file` as it might be a user's original file.

if __name__ == "__main__":
    main()
