import os
import shutil
import pandas as pd
import argparse
import sys
import tkinter as tk
from tkinter import filedialog

def move_files_based_on_excel(excel_file=None, filenames_list=None, source_folder=None, destination_folder=None):
    """
    Moves files from a source folder to a destination folder based on a list of filenames.
    The filenames can be provided via an Excel file (first column) or a direct list.
    Source and destination folders, and the Excel file, can be provided as arguments.
    If run standalone without these arguments, it will prompt via GUI windows.
    Progress updates are printed to stdout.
    Includes enhanced logging for debugging purposes.
    """

    # Check if any necessary argument (source_folder, destination_folder, excel_file, filenames_list) is missing.
    # If so, assume standalone execution and prompt via GUI.
    if source_folder is None or destination_folder is None or (excel_file is None and filenames_list is None):
        # Initialize Tkinter root, but hide the main window to only show file dialogs
        root = tk.Tk()
        root.withdraw() # Hide the main window

        # Prompt for source folder if not provided
        if source_folder is None:
            source_folder = filedialog.askdirectory(title="Select Source Folder")
            if not source_folder:
                print("Source folder selection cancelled. Exiting.", file=sys.stderr)
                root.destroy()
                return

        # Prompt for destination folder if not provided
        if destination_folder is None:
            destination_folder = filedialog.askdirectory(title="Select Destination Folder")
            if not destination_folder:
                print("Destination folder selection cancelled. Exiting.", file=sys.stderr)
                root.destroy()
                return

        # Prompt for Excel file if neither excel_file nor filenames_list was provided
        if excel_file is None and filenames_list is None:
            excel_file = filedialog.askopenfilename(
                title="Select Excel file with filenames in the first column",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            if not excel_file:
                print("Excel file selection cancelled. Exiting.", file=sys.stderr)
                root.destroy()
                return

        root.destroy() # Destroy the hidden Tkinter root once all dialogs are dismissed

    # --- Start Debugging Output ---
    print("\n--- Debugging Information ---")
    print(f"Resolved Source Folder: '{source_folder}'")
    print(f"Resolved Destination Folder: '{destination_folder}'")
    print(f"Using Excel File: '{excel_file}' (if applicable)")
    print(f"Using Filenames List (from textbox): {len(filenames_list) if filenames_list else 0} entries (if applicable)")
    print("---------------------------\n")
    # --- End Debugging Output ---

    # Validate source folder existence
    if not os.path.exists(source_folder):
        print(f"Error: Source folder not found: {source_folder}", file=sys.stderr)
        return

    # Ensure the destination folder exists, create it if it doesn't
    if not os.path.exists(destination_folder):
        try:
            os.makedirs(destination_folder)
            print(f"Created destination folder: {destination_folder}")
        except OSError as e:
            print(f"Error creating destination folder {destination_folder}: {e}", file=sys.stderr)
            return

    files_to_move = []
    if excel_file:
        try:
            # Load the Excel file and get filenames from the first column (iloc[:, 0])
            # Use header=None to ensure the first row is read as data, not as a header.
            df = pd.read_excel(excel_file, header=None)
            print(f"DEBUG: Raw DataFrame head:\n{df.head()}", file=sys.stderr)
            print(f"DEBUG: DataFrame columns: {df.columns.tolist()}", file=sys.stderr)
            
            # Get the list of file names from the first column, regardless of header
            # Ensure names are strings and handle NaN values
            raw_column_data = df.iloc[:, 0].tolist()
            print(f"DEBUG: Raw data from first column: {raw_column_data}", file=sys.stderr)

            files_to_move = [str(f).strip() for f in raw_column_data if pd.notna(f) and str(f).strip()]
            print(f"DEBUG: Files extracted from Excel after processing: {files_to_move}", file=sys.stderr)

        except IndexError:
            print(f"Error: No columns found in Excel file {excel_file}. "
                  "Please ensure the Excel file contains data.", file=sys.stderr)
            return
        except Exception as e:
            print(f"Error reading Excel file {excel_file}: {e}", file=sys.stderr)
            return
    elif filenames_list:
        # Use the list of filenames provided directly (e.g., from the GUI's textbox)
        files_to_move = [str(f).strip() for f in filenames_list if str(f).strip()] # Ensure strings and clean up whitespace
        print(f"DEBUG: Files from provided list (after processing): {files_to_move}", file=sys.stderr)
    else: # This 'else' will only be hit if folders were provided/selected, but no file source was.
        print("No filenames provided via Excel file or list. Exiting.", file=sys.stderr)
        return

    if not files_to_move:
        print("No files to move found in the spreadsheet or provided list.", file=sys.stderr)
        return

    total_files = len(files_to_move)
    moved_count = 0

    print(f"Attempting to move {total_files} files from '{source_folder}' to '{destination_folder}'.")

    for i, file_name in enumerate(files_to_move):
        # Strip any leading/trailing whitespace from the filename
        file_name = file_name.strip()
        source_path = os.path.join(source_folder, file_name)
        destination_path = os.path.join(destination_folder, file_name)

        print(f"Processing: '{file_name}'")
        print(f"  Source Path Attempt: '{source_path}'")
        print(f"  Destination Path Attempt: '{destination_path}'")

        if os.path.exists(source_path):
            print(f"  File EXISTS at source: '{source_path}'")
            try:
                shutil.move(source_path, destination_path)
                print(f"  Moved: '{file_name}'")
                moved_count += 1

                # Verification after move attempt
                if not os.path.exists(source_path) and os.path.exists(destination_path):
                    print(f"  Verification: '{file_name}' successfully moved.")
                else:
                    print(f"  WARNING: Verification failed for '{file_name}'. Still at source: {os.path.exists(source_path)}, In dest: {os.path.exists(destination_path)}")

            except shutil.Error as se:
                print(f"  SHUTIL_ERROR: Failed to move '{file_name}': {se}", file=sys.stderr)
            except OSError as oe:
                print(f"  OS_ERROR: Failed to move '{file_name}': {oe}", file=sys.stderr)
            except Exception as e:
                print(f"  UNEXPECTED_ERROR: Failed to move '{file_name}': {e}", file=sys.stderr)
        else:
            print(f"  File NOT FOUND at source: '{source_path}'")

        # Report progress to stdout, which the GUI will capture
        progress_percentage = (i + 1) / total_files * 100
        print(f"PROGRESS:{progress_percentage:.2f}")

    print(f"\n--- File Movement Summary ---")
    print(f"Total files attempted: {total_files}")
    print(f"Files successfully moved: {moved_count}")
    print(f"Files not found at source or failed to move: {total_files - moved_count}")
    print("--- End Debugging Information ---\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Move files based on an Excel spreadsheet or a list of filenames.")
    parser.add_argument('--excel_file', type=str, help='Path to the Excel file containing filenames in the first column.')
    parser.add_argument('--filenames_list', type=str, help='Comma-separated list of filenames.')
    parser.add_argument('--source_folder', type=str, help='Path to the source folder.')
    parser.add_argument('--destination_folder', type=str, help='Path to the destination folder.')

    args = parser.parse_args()

    # Convert comma-separated string to list if provided
    filenames = args.filenames_list.split(',') if args.filenames_list else None

    # Call the main function. If --source_folder, --destination_folder,
    # --excel_file, or --filenames_list are not provided via command line,
    # the function will trigger GUI prompts for the missing pieces.
    move_files_based_on_excel(args.excel_file, filenames, args.source_folder, args.destination_folder)
