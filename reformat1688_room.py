import os
import sys
import argparse
import tkinter as tk
from tkinter import filedialog
from PIL import Image, UnidentifiedImageError
import shutil
import concurrent.futures

def prompt_for_folder_tk(prompt_message):
    """
    Prompts the user to select a folder using a Tkinter dialog.
    """
    root = tk.Tk()
    root.withdraw() # Hide the main window
    folder_path = filedialog.askdirectory(title=prompt_message)
    root.destroy() # Close the Tkinter root window
    return folder_path

def process_image(file_path, output_dir, problematic_dir, canvas_size=(3000, 1688), margin=0):
    """
    Processes a single image:
    1. Opens it with alpha to detect transparency, handling errors by moving to problematic_dir.
    2. Finds the bounding box of non-white and non-transparent pixels.
    3. Crops the image to its content.
    4. Scales the cropped content to fit within the canvas dimensions (preserving aspect ratio).
    5. Pastes the resized content onto a white background canvas, centered.
    6. Saves the result to the output directory.
    """
    try:
        # Open image with alpha so we can detect transparency
        # Added UnidentifiedImageError handling similar to the second script
        with Image.open(file_path) as im:
            im = im.convert("RGBA")
    except (OSError, UnidentifiedImageError) as e:
        problem_path = os.path.join(problematic_dir, os.path.basename(file_path))
        try:
            shutil.move(file_path, problem_path)
            return f"  ✖ Moved problematic file: {os.path.basename(file_path)} → Problematic Images ({e})"
        except Exception as move_error:
            return f"  ✖ Failed to move {os.path.basename(file_path)}: {move_error}"

    width, height = im.size
    pixels = im.load()

    # Find bounding box of all non-white (and non-transparent) pixels
    min_x, min_y = width, height
    max_x, max_y = 0, 0
    found_content = False
    for y in range(height):
        for x in range(width):
            r, g, b, a = pixels[x, y]
            # Check if pixel is not fully transparent AND not pure white
            if a != 0 and not (r == 255 and g == 255 and b == 255):
                found_content = True
                min_x = min(min_x, x)
                min_y = min(min_y, y)
                max_x = max(max_x, x)
                max_y = max(max_y, y)

    # If no non-white/non-transparent pixels found, skip
    # Changed return value to a string for consistency with threaded output
    if not found_content or min_x > max_x or min_y > max_y:
        return f"  ➔ Skipped (no discernible content): {os.path.basename(file_path)}"

    # Crop to content
    cropped = im.crop((min_x, min_y, max_x + 1, max_y + 1))

    # --- MODIFICATION START: Adopt scaling logic from the second script ---
    canvas_w, canvas_h = canvas_size
    content_w = canvas_w - 2 * margin
    content_h = canvas_h - 2 * margin
    
    crop_w, crop_h = cropped.size

    # Prevent division by zero if crop_w or crop_h is 0
    if crop_w == 0 or crop_h == 0:
        return f"  ➔ Skipped (zero width or height after crop): {os.path.basename(file_path)}"

    # Scale to fit entirely within the canvas (preserve aspect ratio, potentially add borders)
    scale = min(content_w / crop_w, content_h / crop_h)
    # --- MODIFICATION END ---

    # Resize with high-quality LANCZOS
    new_w, new_h = int(crop_w * scale), int(crop_h * scale)
    resized = cropped.resize((new_w, new_h), Image.LANCZOS)

    # Create white background canvas and center the resized crop
    canvas = Image.new("RGB", canvas_size, (255, 255, 255))
    paste_x = (canvas_w - new_w) // 2
    paste_y = (canvas_h - new_h) // 2
    canvas.paste(resized, (paste_x, paste_y), mask=resized) # Use mask for proper alpha blending

    # Save to output folder
    out_path = os.path.join(output_dir, os.path.basename(file_path))
    canvas.save(out_path)
    return f"  ✔ Processed and saved: {os.path.basename(file_path)}"

def main():
    parser = argparse.ArgumentParser(description="Process and crop images to 3000x1688, scaling to fit.")
    parser.add_argument('--input', help='Path to the input folder containing images to process.')
    
    args = parser.parse_args()

    input_folder = None

    # Determine if input path was provided (from GUI) or if we need to prompt (standalone)
    if not args.input:
        print("\n--- Running Cropping Script (3000x1688 - Fit Within) in Standalone Mode ---")
        input_folder = prompt_for_folder_tk("Select folder with images to process")
        if not input_folder:
            sys.exit("No folder selected. Exiting.")
    else:
        input_folder = args.input
        if not os.path.isdir(input_folder):
            sys.exit(f"Error: Input path '{input_folder}' is not a valid directory. Exiting.")

    # Prepare output directory and problematic directory
    output_dir = os.path.join(input_folder, "Processed_1688_Fit") # Changed folder name to reflect new logic
    problematic_dir = os.path.join(input_folder, "Problematic_Images") # Added problematic images directory
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(problematic_dir, exist_ok=True)

    # Supported extensions (Pillow handles many, but explicitly list common ones)
    exts = (".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".gif")

    print(f"\nProcessing images in: {input_folder}")
    print(f"Output will be saved to: {output_dir}")
    print(f"Problematic files will be moved to: {problematic_dir}\n")

    image_files = [f for f in sorted(os.listdir(input_folder)) if f.lower().endswith(exts)]
    image_paths = [os.path.join(input_folder, fname) for fname in image_files]

    if not image_paths:
        print("No supported image files found in the specified folder. Exiting.")
        sys.exit(0)

    # Use ThreadPoolExecutor for parallel processing, similar to the second script
    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = [
            executor.submit(process_image, file_path, output_dir, problematic_dir)
            for file_path in image_paths
        ]
        for future in concurrent.futures.as_completed(futures):
            result = future.result()
            if result:
                print(result) # Print results from the threaded tasks

    print(f"\n--- Processing Complete ---")
    print(f"All done! Check '{output_dir}' for processed images and '{problematic_dir}' for any issues.")
    sys.exit(0)

if __name__ == "__main__":
    main()
