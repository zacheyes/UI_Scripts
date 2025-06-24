import pandas as pd
import os
import sys
import tkinter as tk # Keep tkinter import for standalone mode
from tkinter import filedialog # Keep filedialog import for standalone mode

# Function to replace commas within cells with |
def replace_commas(cell_value):
    if isinstance(cell_value, str):  # Check if the cell value is a string
        return cell_value.replace(',', '|')
    return cell_value  # Return the value unchanged if it's not a string

def convert_bynder_metadata_csv_to_xlsx(input_csv_path, output_folder):
    if not os.path.exists(input_csv_path):
        print(f"Error: Input CSV file not found at {input_csv_path}", file=sys.stderr)
        return False

    # Ensure output_folder exists
    if not output_folder or not os.path.isdir(output_folder):
        try:
            os.makedirs(output_folder, exist_ok=True)
            print(f"Created output directory: {output_folder}")
        except Exception as e:
            print(f"Error creating output directory {output_folder}: {e}", file=sys.stderr)
            return False

    print("PROGRESS: 5")
    try:
        # Read the CSV file with semicolon as the delimiter
        df = pd.read_csv(input_csv_path, delimiter=';')
        print("PROGRESS: 30")

        # Apply the replace_commas function to all cells in each column
        df = df.map(replace_commas)
        print("PROGRESS: 70")

        # Get the base name of the input file without extension
        base_name = os.path.splitext(os.path.basename(input_csv_path))[0]
        output_file_name = f"{base_name}.xlsx"
        output_file_path = os.path.join(output_folder, output_file_name)

        # Save the DataFrame to an Excel file
        df.to_excel(output_file_path, index=False, engine='openpyxl')
        print("PROGRESS: 95")

        print(f"File has been converted and saved to {output_file_path}")
        print("PROGRESS: 100")
        return True

    except pd.errors.EmptyDataError:
        print(f"Error: The input CSV file is empty: {input_csv_path}", file=sys.stderr)
        return False
    except FileNotFoundError:
        print(f"Error: The file {input_csv_path} was not found.", file=sys.stderr)
        return False
    except Exception as e:
        print(f"An error occurred during conversion: {e}", file=sys.stderr)
        return False

if __name__ == "__main__":
    # Check if arguments are provided (when run from GUI)
    if len(sys.argv) > 2:
        input_csv_path_arg = sys.argv[1]
        output_folder_arg = sys.argv[2]
        convert_bynder_metadata_csv_to_xlsx(input_csv_path_arg, output_folder_arg)
    else:
        # Fallback to tkinter filedialog if run standalone
        root = tk.Tk()
        root.withdraw() # Hide the main window

        file_path = filedialog.askopenfilename(title="Select a CSV file", filetypes=[("CSV files", "*.csv")])

        if file_path:
            # Default to user's downloads folder for standalone execution
            downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
            os.makedirs(downloads_folder, exist_ok=True)
            print(f"Converting file: {file_path}")
            print(f"Saving to: {downloads_folder}")
            success = convert_bynder_metadata_csv_to_xlsx(file_path, downloads_folder)
            if success:
                print("Conversion completed successfully.")
            else:
                print("Conversion failed.")
        else:
            print("File selection cancelled.")