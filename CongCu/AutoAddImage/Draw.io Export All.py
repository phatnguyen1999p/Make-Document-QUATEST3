import os
import subprocess
import xml.etree.ElementTree as ET
from urllib.parse import unquote  # Fixes web-encoded spaces like %20

def clean_filename(filename):
    """Removes characters that Windows doesn't allow in file names."""
    # Decode web encoding first (e.g., converts 'My%20Tab' to 'My Tab')
    filename = unquote(filename)
    
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '')
    return filename.strip()

def export_drawio_pages():
    print("=== Draw.io Multi-Page Exporter ===")
    
    # 1. Get Inputs
    raw_input_file = input("Paste the path to your .drawio file: ").strip()
    raw_output_dir = input("Paste the destination folder path: ").strip()
    
    input_file = raw_input_file.strip('"').strip("'")
    output_dir = raw_output_dir.strip('"').strip("'")
    
    drawio_exe = r"C:\Program Files\draw.io\draw.io.exe"

    if not os.path.exists(input_file):
        print(f"\n[Error] The input file could not be found at: {input_file}")
        return

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 2. Parse XML to read the explicit tab names
    try:
        tree = ET.parse(input_file)
        root = tree.getroot()
        
        # Draw.io files can be structured with <mxfile> or <mxGraphModel> at the root
        diagrams = root.findall('.//diagram')
        total_pages = len(diagrams)
    except Exception as e:
        print(f"\n[Error] Failed to read or parse the Draw.io file: {e}")
        return

    if total_pages == 0:
        print("\nNo pages found inside the Draw.io file.")
        return

    print(f"\nFound exactly {total_pages} pages. Starting export...\n")

    # 3. Loop and Export
    for index in range(total_pages):
        # Fetch the 'name' attribute from the <diagram> tag
        raw_name = diagrams[index].get('name')
        
        # Clean the name to ensure Windows handles spaces and special characters perfectly
        if raw_name:
            page_name = clean_filename(raw_name)
        else:
            page_name = f"Tab_{index + 1}" # Fallback if a tab somehow has no name
        
        output_file_path = os.path.join(output_dir, f"{page_name}.png")
        print(f"Processing Page {index + 1} of {total_pages}: [{page_name}.png] ...")

        command = [
            drawio_exe,
            "--export",
            "--page-index", str(index + 1),
            "--format", "png",
            "--output", output_file_path,
            input_file
        ]

        # Run background export
        subprocess.run(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    print(f"\nSuccess! All {total_pages} pages have been exported to: {output_dir}")

if __name__ == "__main__":
    export_drawio_pages()
    print("-" * 40)
    input("Press ENTER to close...")