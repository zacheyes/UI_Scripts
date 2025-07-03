import os
import csv
import sys
from datetime import datetime
import tkinter as tk # Import tkinter
from tkinter import filedialog # Import filedialog for the graphical picker

def export_directory_list_to_csv(directory_path, progress_callback=None):
    """
    Exports a list of files in the specified directory to a CSV file in the user's Downloads folder.
    Each row contains the full file path and the filename.

    Parameters:
        directory_path (str): The path to the directory.
        progress_callback (callable, optional): A function (value, total) to call for progress updates.
                                                This function should accept two arguments:
                                                the number of files processed so far, and the total number of files.
    Returns:
        tuple: (success (bool), output_message (str), output_csv_path (str))
    """
    if not os.path.isdir(directory_path):
        return False, f"Error: Directory not found at '{directory_path}'", None

    try:
        # Get current timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        # Define the output CSV file path in the Downloads folder
        output_csv_filename = f"Directory_List_{timestamp}.csv"
        output_csv_path = os.path.join(os.path.expanduser('~'), 'Downloads', output_csv_filename)
        
        # Ensure the Downloads directory exists
        os.makedirs(os.path.dirname(output_csv_path), exist_ok=True)

        all_files = []
        # First pass: count all files to determine total for progress bar
        for root, _, files in os.walk(directory_path):
            for filename in files:
                all_files.append(os.path.join(root, filename))
        
        total_files = len(all_files)
        processed_files = 0

        with open(output_csv_path, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(["Full Path", "Filename"])  # Write the header

            # Second pass: write files and update progress
            for full_path in all_files:
                filename = os.path.basename(full_path)
                writer.writerow([full_path, filename])
                processed_files += 1
                if progress_callback:
                    # Report progress as value/total
                    progress_callback(processed_files, total_files)

        return True, f"Directory list has been exported to: {output_csv_path}", output_csv_path

    except Exception as e:
        return False, f"An unexpected error occurred during directory listing: {e}", None

if __name__ == "__main__":
    # Initialize Tkinter
    root = tk.Tk()
    root.withdraw() # Hide the main Tkinter window

    # Open a directory selection dialog
    print("Please select a folder to list its contents...") # Print to console for clarity
    cli_dir_path = filedialog.askdirectory(title="Select a Folder to List Contents Of")

    # If the user cancels the dialog, cli_dir_path will be an empty string
    if not cli_dir_path:
        print("Folder selection cancelled. Exiting.")
        sys.exit(1) # Exit with an error code

    def cli_progress_callback(value, total):
        """A simple progress callback for CLI output."""
        if total > 0:
            percent = (value / total) * 100
            # Use carriage return '\r' to update the same line in terminal
            sys.stdout.write(f"\rPROGRESS: {percent:.1f}% ({value}/{total})")
            sys.stdout.flush()
        else:
            # Handle cases where there are no files (e.g., only empty subfolders)
            sys.stdout.write(f"\rPROGRESS: {'Finished (No files found or progress not applicable)'}")
            sys.stdout.flush()

    print(f"Starting directory list export for: {cli_dir_path}")
    success, message, csv_path = export_directory_list_to_csv(cli_dir_path, progress_callback=cli_progress_callback)
    sys.stdout.write("\n") # Ensure a new line after the progress bar finishes

    if success:
        print(f"Success: {message}")
    else:
        print(f"Error: {message}", file=sys.stderr)

    # Destroy the Tkinter root window after use
    root.destroy()