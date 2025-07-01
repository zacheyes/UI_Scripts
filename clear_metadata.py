# Save this as clear_metadata.py
import os
import subprocess
import sys
import argparse
import shutil
import tkinter as tk # NO LONGER NEEDED FOR THIS SCRIPT'S INTERNAL LOGIC WHEN CALLED WITH ARGS
from tkinter import filedialog # NO LONGER NEEDED FOR THIS SCRIPT'S INTERNAL LOGIC WHEN CALLED WITH ARGS

# --- Configuration for Metadata Fields ---

# Mapping of user-friendly names to ExifTool tags.
# Each value is a single ExifTool tag string.
# These are the *only* properties the user will be prompted about or
# can specify via the --clear_properties command-line argument.
METADATA_PROPERTIES = {
    "Description": "-Description=",
    "ImageDescription": "-ImageDescription=",
    "Caption-Abstract": "-Caption-Abstract=",
    "Keywords": "-Keywords=",
    "Subject": "-Subject=",
    "Title": "-Title=",
    "Headline": "-Headline=",
    "ObjectName": "-ObjectName=",
    "Event": "-Event=",
}


def get_exiftool_executable_path():
    """
    Attempts to find the ExifTool executable.
    1. Checks a local 'tools/exiftool' directory (for bundled/portable distribution).
    2. Checks if 'exiftool' is in the system's PATH.
    Returns the full path to the executable, or None if not found.
    """
    # Path relative to the script's directory
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Check for Windows executable (exiftool.exe)
    portable_windows_path = os.path.join(script_dir, 'tools', 'exiftool', 'exiftool.exe')
    if os.path.exists(portable_windows_path) and os.path.isfile(portable_windows_path):
        return portable_windows_path
    
    # Check for Linux/macOS executable (plain 'exiftool' binary/script)
    # This covers both the Perl script and potential native binaries if you bundled one
    portable_unix_path = os.path.join(script_dir, 'tools', 'exiftool', 'exiftool')
    if os.path.exists(portable_unix_path) and os.path.isfile(portable_unix_path):
        return portable_unix_path

    # If not found in bundled location, check system PATH
    try:
        # shutil.which is Python 3.3+ and is the robust way to check PATH
        exiftool_in_path = shutil.which("exiftool")
        if exiftool_in_path:
            return exiftool_in_path
    except Exception as e:
        # Fallback for older Python versions or unexpected issues with shutil.which
        for path_dir in os.environ.get("PATH", "").split(os.pathsep):
            full_path = os.path.join(path_dir, "exiftool")
            if os.path.exists(full_path) and os.path.isfile(full_path):
                return full_path
            full_path_exe = os.path.join(path_dir, "exiftool.exe") # Windows might have it in PATH too
            if os.path.exists(full_path_exe) and os.path.isfile(full_path_exe):
                return full_path_exe

    return None # ExifTool not found


def clear_metadata_from_image_exiftool(file_path, exiftool_tags_to_clear_final):
    """
    Clears specified metadata fields from a single image file using ExifTool.
    Supports JPEG, PNG, TIFF. Overwrites the original file.

    Args:
        file_path (str): The path to the image file.
        exiftool_tags_to_clear_final (list): List of ExifTool tag strings to clear.
    """
    print(f"\n  Processing: {os.path.basename(file_path)}")
    
    exiftool_path = get_exiftool_executable_path()
    if not exiftool_path:
        sys.stderr.write(f"  Error: ExifTool not found in bundled 'tools/exiftool' directory or system PATH.\n")
        sys.stderr.write(f"  Please ensure 'exiftool.exe' (Windows) or 'exiftool' (macOS/Linux) is in the '{os.path.join(os.path.dirname(os.path.abspath(__file__)), 'tools', 'exiftool')}' folder, or installed and in your system's PATH.\n")
        sys.stderr.write(f"  Download from: https://exiftool.org/\n")
        return # Exit function if ExifTool is not found

    print(f"  Using ExifTool from: {exiftool_path}")

    command = [exiftool_path] # Use the found path

    if not exiftool_tags_to_clear_final:
        print(f"  No metadata fields selected for clearing for {os.path.basename(file_path)}. Skipping ExifTool operation.")
        return # Exit function if no tags to clear

    command.extend(exiftool_tags_to_clear_final)
    command.append("-overwrite_original")
    command.append(file_path)

    try:
        # Run ExifTool command
        process = subprocess.run(command, capture_output=True, text=True, check=True, encoding='utf-8')
        print(f"  ExifTool stdout: {process.stdout.strip()}")
        if process.stderr:
            sys.stderr.write(f"  ExifTool stderr: {process.stderr.strip()}\n")
        print(f"  Metadata processed by ExifTool for: {os.path.basename(file_path)}")
    except subprocess.CalledProcessError as e:
        sys.stderr.write(f"  Error running ExifTool for {os.path.basename(file_path)}: {e}\n")
        sys.stderr.write(f"  ExifTool stdout: {e.stdout.strip()}\n")
        sys.stderr.write(f"  ExifTool stderr: {e.stderr.strip()}\n")
    except Exception as e:
        sys.stderr.write(f"  An unexpected error occurred for {os.path.basename(file_path)}: {e}\n")


def get_user_choices_interactive():
    """
    Interactively prompts the user for which metadata properties to clear.
    Returns a list of ExifTool tags that should be cleared based on user input.
    """
    print("\n--- Metadata Configuration (Interactive Mode) ---")
    print("For each property, do you want to CLEAR it?")
    print("Type 'y' for YES or 'n' for NO.")
    print("-------------------------------------------------")

    final_tags_to_clear_exiftool = []
    
    for prop_display_name, exiftool_tag in METADATA_PROPERTIES.items():
        user_input = ""
        while user_input.lower() not in ['y', 'n']:
            prompt_text = f"Do you want to CLEAR '{prop_display_name}'? (y/n): "
            user_input = input(prompt_text).strip().lower()

            if user_input == 'y':
                final_tags_to_clear_exiftool.append(exiftool_tag)
                print(f"  -> '{prop_display_name}' will be CLEARED.")
            elif user_input == 'n':
                print(f"  -> '{prop_display_name}' will be KEPT.")
            else:
                print("Invalid input. Please type 'y' or 'n'.")
    
    final_tags_to_clear_exiftool = list(set(final_tags_to_clear_exiftool))

    return final_tags_to_clear_exiftool


def main():
    # Dynamically build the list of available properties for argparse help text
    available_properties_str = ", ".join(METADATA_PROPERTIES.keys())

    parser = argparse.ArgumentParser(
        description="Clear specific embedded metadata from image files in a folder using ExifTool.\n"
                    "If --clear_properties is NOT used and --input_folder is NOT used, the script will prompt interactively.",
        formatter_class=argparse.RawTextHelpFormatter
    )
    # Make input_folder optional, but expected if not interactive
    parser.add_argument("--input_folder", help="Path to the folder containing image files.")
    parser.add_argument("--clear_properties", nargs='*',
                        help=f"Optional: Space-separated list of metadata properties to CLEAR. "
                             f"Options (case-sensitive as listed): {available_properties_str}. "
                             f"If this argument is provided, the script will NOT prompt interactively for metadata; "
                             f"only the specified properties will be cleared, others will be kept.")

    args = parser.parse_args()

    tags_to_clear_for_all_files = []
    input_folder = None # Initialize input_folder

    # Logic for determining input_folder and tags_to_clear
    if args.input_folder and args.clear_properties is not None:
        # GUI mode: input_folder and clear_properties provided via CLI
        input_folder = args.input_folder
        print("\n--- Metadata Configuration (from command line) ---")
        for prop_name in args.clear_properties:
            if prop_name in METADATA_PROPERTIES:
                tags = METADATA_PROPERTIES[prop_name]
                tags_to_clear_for_all_files.append(tags)
                print(f"  -> '{prop_name}' will be CLEARED.")
            else:
                sys.stderr.write(f"Warning: Unknown metadata property '{prop_name}' specified. Skipping.\n")
        tags_to_clear_for_all_files = list(set(tags_to_clear_for_all_files))

    elif args.input_folder and args.clear_properties is None:
        # CLI mode: input_folder provided, but no specific clear properties. Assume interactive for properties.
        input_folder = args.input_folder
        print(f"Input folder specified: '{input_folder}'. Prompting for metadata options...")
        tags_to_clear_for_all_files = get_user_choices_interactive()

    elif not args.input_folder and args.clear_properties is None:
        # Pure interactive mode: neither input_folder nor clear_properties provided
        # Use Tkinter for folder selection in this standalone interactive mode
        print("\nPure interactive mode: Launching folder selector and prompting for metadata options...")
        root = tk.Tk()
        root.withdraw() # Hide the main window
        input_folder = filedialog.askdirectory(title="Select Input Folder")
        root.destroy()

        if not input_folder:
            print("No folder selected. Exiting.")
            sys.exit(0)

        print(f"Selected folder: {input_folder}")
        tags_to_clear_for_all_files = get_user_choices_interactive()
    else:
        # This case should ideally not happen if argparse is set up correctly,
        # but handles if only --clear_properties is given without --input_folder.
        sys.stderr.write("Error: --clear_properties requires --input_folder when not running fully interactively.\n")
        sys.exit(1)


    if not os.path.isdir(input_folder):
        sys.stderr.write(f"Error: Input folder not found at '{input_folder}'\n")
        sys.exit(1)

    image_files = [f for f in os.listdir(input_folder) if f.lower().endswith(('.jpg', '.jpeg', '.png', '.tif', '.tiff'))]
    
    if not image_files:
        print(f"No supported image files found in '{input_folder}'.")
        sys.exit(0)

    if not tags_to_clear_for_all_files:
        print("\nNo metadata fields selected for clearing. Exiting.")
        sys.exit(0)
    
    print("\n--- Starting File Processing ---")
    print(f"The following ExifTool commands will be used to clear metadata: {tags_to_clear_for_all_files}")

    total_files = len(image_files)
    processed_count = 0

    print(f"\nStarting metadata processing for {total_files} images in '{input_folder}'...")

    for filename in image_files:
        file_path = os.path.join(input_folder, filename)
        clear_metadata_from_image_exiftool(
            file_path,
            exiftool_tags_to_clear_final=tags_to_clear_for_all_files
        )
        processed_count += 1
        # Report progress to stdout, which the GUI will capture
        sys.stdout.write(f"PROGRESS:{processed_count / total_files * 100:.2f}\n")
        sys.stdout.flush()

    print(f"\nFinished processing metadata for {processed_count} files.")
    sys.stdout.write(f"PROGRESS:100.00\n")
    sys.stdout.flush()
    sys.exit(0)

if __name__ == "__main__":
    main()