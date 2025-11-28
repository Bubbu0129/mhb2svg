import json
import urllib.parse
import urllib.request
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import argparse
import tempfile

def convert(input_path, output_path, padding, ratio):
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
        width *= ratio
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
        raise ValueError("URL does not contain the 's_id' parameter.")
    return query_params['s_id'][0]

def fetch_file_url(sid):
    api_url = f"https://res.maxhub.com/v3/clientairdisk/api/share/v2/{sid}/resources.json"
    req = urllib.request.Request(api_url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req) as response:
        data = json.loads(response.read().decode('utf-8'))
        if not data or 'url' not in data[0]:
            raise ValueError("API response did not contain the expected file URL.")
    return data[0]['url']

def download_archive(url, dst_path):
    try:
        with urllib.request.urlopen(url) as response:
            filename = response.info().get_filename()
            if not filename:
                parsed_url = urllib.parse.urlparse(url)
                filename = Path(parsed_url.path).name
            if not filename:
                filename = "archive.mhb"
            file_path = dst_path / filename
            with file_path.open('wb') as f:
                f.write(response.read())
    except Exception as e:
        raise RuntimeError(f"Failed to download file: {e}")
    return file_path

def extract_archive(file_path, dst_path):
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(dst_path)
    except zipfile.BadZipFile:
        raise RuntimeError("The downloaded file is not a valid archive.")
    return

def process_slides(src_path, padding, ratio):
    slides_dir = src_path / "Slides"
    if not slides_dir.exists():
        print(f"Warning: {slides_dir} does not exist. Skipping conversion.")
        return
    xml_files = list(slides_dir.glob("*.xml"))
    if not xml_files:
        print("No XML files found in Slides directory.")
    slides = []
    for xml_path in xml_files:
        output_path = f"{xml_path.stem}.svg"
        convert(str(xml_path), output_path, padding, ratio)
        slides.append(output_path)
    return slides

def parse_args():
    parser = argparse.ArgumentParser()
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('-f', '--file', type=str, help='Path to .mhb file')
    group.add_argument('-l', '--link', type=str, help='MAXHUB URL containing the s_id argument')
    parser.add_argument('-p', '--padding', type=int, default=10, help='Padding size for .svg (integer)')
    parser.add_argument('-r', '--ratio', type=float, default=1.0, help='Stoke width ratio (float)')
    return parser.parse_args()

def parse_metadata(src_path):
    xml_path = src_path / "Document.xml"
    tree = ET.parse(str(xml_path))
    root = tree.getroot()
    metadata = [(child.tag, child.text) for child in root]
    return metadata

if __name__ == "__main__":
    args = parse_args()
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        if args.link:
            sid = extract_sid(args.link)
            print(f"Extracted SID: {sid}")
            file_url = fetch_file_url(sid)
            print(f"File URL: {file_url}")
            file_path = download_archive(file_url, tmp_path)
        else:
            file_path = Path(args.file)
        print("Unzipping archive...")
        extract_archive(file_path, tmp_path)
        metadata = parse_metadata(tmp_path)
        print(f"Metadata of {file_path.name}:")
        for tag, text in metadata:
            print(f"\t{tag}: {text}")
        print("Processing slides...")
        slides = process_slides(tmp_path, args.padding, args.ratio)
        print(f"Generated " + ", ".join(slides))

