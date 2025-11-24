padding = 10
stroke_width_ratio = 1

import sys
import json
import urllib.parse
import urllib.request
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

def convert(input_path, output_path):
    tree = ET.parse(input_path)
    root = tree.getroot()

    lines = []
    min_x, min_y = float('inf'), float('inf')
    max_x, max_y = float('-inf'), float('-inf')

    for ink in root.findall('.//Ink'):

        width_str = ink.findtext('Thickness')
        if width_str is None:
            width = 1.0
        else:
            width = float(width_str)
        width *= stroke_width_ratio

        points = ink.find('Points')
        if points is None:
            continue
        stroke = []
        for pt in points.findall('StylusPoint'):
            # Format: "x,y,pressure"
            parts = pt.text.split(',')
            x = float(parts[0])
            y = float(parts[1])
            stroke.append((x, y))
            
            if x < min_x: min_x = x
            if x > max_x: max_x = x
            if y < min_y: min_y = y
            if y > max_y: max_y = y

        points_str = " ".join([f"{x},{y}" for x, y in stroke])
        lines.append(
            f'<polyline points="{points_str}" '
            f'fill="none" stroke="black" stroke-width="{width}" '
            f'stroke-linecap="round" stroke-linejoin="round"/>\n'
        )

    viewbox_width = max_x - min_x + (padding * 2)
    viewbox_height = max_y - min_y + (padding * 2)
    viewbox_x = min_x - padding
    viewbox_y = min_y - padding

    print(f"x from {min_x:.1f} to {max_x:.1f}, y from {min_y:.1f} to {max_y:.1f}")

    # Write to SVG
    with open(output_path, 'w') as f:
        f.write(
            f'<svg xmlns="http://www.w3.org/2000/svg" version="1.1" viewBox="{viewbox_x} {viewbox_y} {viewbox_width} {viewbox_height}">\n'
        )
        f.writelines(lines)
        f.write('</svg>\n')

    return

def extract_sid(url):
    parsed_url = urllib.parse.urlparse(url)
    query_params = urllib.parse.parse_qs(parsed_url.query)
    
    if 's_id' not in query_params:
        raise ValueError("URL does not contain an 's_id' parameter.")
    
    return query_params['s_id'][0]

def fetch_file_url(sid):
    """Requests the API to get the actual file download URL."""
    api_url = f"https://res.maxhub.com/v3/clientairdisk/api/share/v2/{sid}/resources.json"
    req = urllib.request.Request(api_url, headers={'User-Agent': 'Mozilla/5.0'})
    
    with urllib.request.urlopen(req) as response:
        data = json.loads(response.read().decode('utf-8'))
        if not data or 'url' not in data[0]:
            raise ValueError("API response did not contain the expected file URL.")
        return data[0]['url']

def download_and_extract(url, dst_folder):
    """Downloads the file and unzips it into the destination folder."""
    dst_path = Path(dst_folder)
    dst_path.mkdir(parents=True, exist_ok=True)
    
    local_filename = dst_path / "source_archive.bin"
    
    try:
        with urllib.request.urlopen(url) as response, local_filename.open('wb') as out_file:
            out_file.write(response.read())
    except Exception as e:
        raise RuntimeError(f"Failed to download file: {e}")

    print("Unzipping archive...")
    try:
        with zipfile.ZipFile(local_filename, 'r') as zip_ref:
            zip_ref.extractall(dst_path)
    except zipfile.BadZipFile:
        raise RuntimeError("The downloaded file is not a valid zip archive.")

def process_slides(temp_dir):
    """Finds .xml files in tmp/Slides/ and converts them."""
    slides_dir = Path(temp_dir) / "Slides"
    
    if not slides_dir.exists():
        print(f"Warning: {slides_dir} does not exist. Skipping conversion.")
        return

    xml_files = list(slides_dir.glob("*.xml"))
    
    if not xml_files:
        print("No .xml files found in Slides directory.")
        return

    print(f"Found {len(xml_files)} XML files. processing...")

    for xml_path in xml_files:
        output_path = f"{xml_path.stem}.svg"
        convert(str(xml_path), output_path)

if __name__ == "__main__":
    if len(sys.argv) == 1:
        print(f"Usage: python {sys.argv[0]} <url>")
        sys.exit(1)
        
    url = sys.argv[1]
    tmp_dir = Path("tmp")

    # Step 1: Extract s_id
    sid = extract_sid(url)
    print(f"Extracted SID: {sid}")

    # Step 2: Get File URL
    file_url = fetch_file_url(sid)
    print(f"File URL: {file_url}")

    # Step 3: Download and Unzip
    download_and_extract(file_url, tmp_dir)

    # Step 4: Process XMLs
    process_slides(tmp_dir)

