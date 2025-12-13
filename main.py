import json
import urllib.parse
import urllib.request
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import argparse
import tempfile

BG_color = '#363b41' # Background color of the slides
API_tpl = 'https://res.maxhub.com/v3/clientairdisk/api/share/v2/{sid}/resources.json' # API template
aspect_ratio = 1.414 # sqrt(2), A series
padding = 10.0
stroke_ratio = 1.0

def write_svg(min_x, min_y, width, height, strokes, output_path, color):
    with open(output_path, 'w') as f:
        f.write(
            f'<svg xmlns="http://www.w3.org/2000/svg" version="1.1" viewBox="0 0 {width + 2 * padding} {height + 2 * padding}">\n'
        )
        if color: # Render fullscreen rectangle with background color
            f.write(
                f'<rect x="0" y="0" width="{width + 2 * padding}" height="{height + 2 * padding}" fill="{BG_color}"/>\n'
            )
        for stroke in strokes:
            if min_y < stroke["points"][0][1] < min_y + height:
                points_str = ' '.join([f'{x - min_x + padding},{y - min_y + padding}' for x, y in stroke["points"]])
                f.write(f'<polyline points="{points_str}" fill="none" stroke="{stroke["color"]}" stroke-width="{stroke["width"]}"/>\n')
        f.write('</svg>\n')

def convert(input_path, output_prefix, color, paging):
    tree = ET.parse(input_path)
    root = tree.getroot()
    lines = []
    min_x, min_y = float('inf'), float('inf')
    max_x, max_y = float('-inf'), float('-inf')
    strokes = []
    for ink in root.findall('.//Ink'):
        stroke_width_str = ink.findtext('Thickness')
        if stroke_width_str:
            stroke_width = float(stroke_width_str)
        else:
            stroke_width = 1.0
        stroke_width *= stroke_ratio
        if color:
            stroke_color = ink.findtext('ForegroundColor')
        else:
            stroke_color = '#000000'
        points = ink.find('Points')
        if points is None:
            continue
        stroke = {"points": [], "color": stroke_color, "width": stroke_width}
        for pt in points.findall('StylusPoint'):
            # Format: 'x,y,pressure'
            parts = pt.text.split(',')
            x = float(parts[0])
            y = float(parts[1])
            if x < min_x: min_x = x
            if x > max_x: max_x = x
            if y < min_y: min_y = y
            if y > max_y: max_y = y
            stroke["points"].append((x, y))
        strokes.append(stroke)
    width = max_x - min_x
    height = max_y - min_y
    pages = int(height / width / aspect_ratio)
    if pages == 0 or not paging:
        write_svg(min_x, min_y, width, height, strokes, output_prefix + '.svg', color)
        return 0
    dh = (height - aspect_ratio * width) / pages
    for i in range(pages + 1):
        write_svg(min_x, min_y + i * dh, width, aspect_ratio * width, strokes, f"{output_prefix}-{i}.svg", color)
    return pages

def extract_sid(url):
    parsed_url = urllib.parse.urlparse(url)
    query_params = urllib.parse.parse_qs(parsed_url.query)
    if 's_id' not in query_params:
        raise ValueError('URL does not contain the "s_id" parameter.')
    return query_params['s_id'][0]

def fetch_file_url(sid):
    api_url = API_tpl.format(sid=sid)
    req = urllib.request.Request(api_url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req) as response:
        data = json.loads(response.read().decode('utf-8'))
        if not data or 'url' not in data[0]:
            raise ValueError('API response did not contain the expected file URL.')
    return data[0]['url']

def download_archive(url, dst_path):
    try:
        with urllib.request.urlopen(url) as response:
            filename = response.info().get_filename()
            if not filename:
                parsed_url = urllib.parse.urlparse(url)
                filename = Path(parsed_url.path).name
            if not filename:
                filename = 'archive.mhb'
            file_path = dst_path / filename
            with file_path.open('wb') as f:
                f.write(response.read())
    except Exception as e:
        raise RuntimeError(f'Failed to download file: {e}')
    return file_path

def extract_archive(file_path, dst_path):
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(dst_path)
    except zipfile.BadZipFile:
        raise RuntimeError('The downloaded file is not a valid archive.')

def process_slides(src_path, color, paging):
    slides_dir = src_path / 'Slides'
    if not slides_dir.exists():
        print(f'Warning: {slides_dir} does not exist. Skipping conversion.')
        return
    xml_files = list(slides_dir.glob('*.xml'))
    if not xml_files:
        print('No XML files found in Slides directory.')
    slides = []
    for xml_path in xml_files:
        output_prefix = xml_path.stem
        pages = convert(str(xml_path), output_prefix, color, paging)
        if pages == 0:
            slides.append(output_prefix)
        else:
            slides.append(f"{output_prefix} ({pages})")
    return slides

def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('-c', '--color',  action='store_true', help='Enable color (default black & white)')
    parser.add_argument('-p', '--paging', action='store_true', help='Enable paging (default none)')
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('-f', '--file', type=str, help='Path to .mhb file')
    group.add_argument('-l', '--link', type=str, help='MAXHUB URL containing the s_id argument')
    return parser.parse_args()

def parse_metadata(src_path):
    xml_path = src_path / 'Document.xml'
    tree = ET.parse(str(xml_path))
    root = tree.getroot()
    metadata = [(child.tag, child.text) for child in root]
    return metadata

if __name__ == '__main__':
    args = parse_args()
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        if args.link:
            sid = extract_sid(args.link)
            print(f'Extracted SID: {sid}')
            file_url = fetch_file_url(sid)
            print(f'File URL: {file_url}')
            file_path = download_archive(file_url, tmp_path)
        else:
            file_path = Path(args.file)
        print('Unzipping archive...')
        extract_archive(file_path, tmp_path)
        metadata = parse_metadata(tmp_path)
        print(f'Metadata of {file_path.name}:')
        for tag, text in metadata:
            print(f'\t{tag}: {text}')
        print('Processing slides...')
        slides = process_slides(tmp_path, args.color, args.paging)
        print(f'Generated ' + ', '.join(slides))

