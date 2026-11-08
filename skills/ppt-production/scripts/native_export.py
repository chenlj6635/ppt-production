import argparse
import hashlib
import json
import math
from pathlib import Path
import re
import sys
import tempfile
import xml.etree.ElementTree as ET

from _native_backend import load_backend


def finite_positive(value, label):
    if isinstance(value, bool) or not isinstance(value, (float, int)) or not math.isfinite(value) or value <= 0:
        raise ValueError(label + ' must be a positive finite number')
    return float(value)


def load_design(project):
    design = json.loads((project / 'design.json').read_text(encoding='utf-8'))
    finite_positive(design['canvas']['width_px'], 'canvas.width_px')
    finite_positive(design['canvas']['height_px'], 'canvas.height_px')
    finite_positive(design['body_min_pt'], 'body_min_pt')
    if design['body_min_pt'] < 20 and not str(design.get('body_min_exception', '')).strip():
        raise ValueError('A floor below 20 pt requires body_min_exception documenting user authorization')
    for field in ['font_latin', 'font_east_asian', 'text_color']:
        if not isinstance(design.get(field), str) or not design[field].strip():
            raise ValueError('Missing explicit design field: ' + field)
    if not design.get('palette'):
        raise ValueError('Provide an explicit project palette; no upstream theme is assumed')
    if not isinstance(design['palette'], list):
        raise ValueError('palette must be a list of explicit project colors')
    design.setdefault('background_color', '#FFFFFF')
    design.setdefault('table_header_fill', design['background_color'])
    for color in design['palette'] + [design['text_color'], design['background_color'], design['table_header_fill']]:
        if not re.fullmatch(r'#[0-9a-fA-F]{6}', color):
            raise ValueError('Use six-digit hex colors in design.json')
    return design


def roster(project):
    files = sorted((project / 'svg_output').glob('*.svg'))
    if not files:
        raise ValueError('No SVG slides found in svg_output')
    return files


def load_objects(project, slide_count, design):
    path = project / 'objects.json'
    objects = json.loads(path.read_text(encoding='utf-8')) if path.exists() else []
    if not isinstance(objects, list):
        raise ValueError('objects.json must contain a list')
    allowed_colors = {color.lower() for color in design['palette'] + [design['text_color'], design['background_color'], design['table_header_fill']]}
    for item in objects:
        if not isinstance(item, dict):
            raise ValueError('Each native data object must be an object')
        if type(item.get('slide')) is not int or not 1 <= item['slide'] <= slide_count:
            raise ValueError('Native object slide must be a valid 1-based integer')
        if item.get('type') not in {'chart', 'table'}:
            raise ValueError('Only native chart/table sidecars are supported')
        colors = [item[field] for field in ['header_fill', 'body_fill'] if field in item]
        if 'series_colors' in item:
            if not isinstance(item['series_colors'], list) or not item['series_colors']:
                raise ValueError('series_colors must be a nonempty list')
            colors += item['series_colors']
        if any(not isinstance(color, str) or color.lower() not in allowed_colors for color in colors):
            raise ValueError('Native object colors must be declared in design.json')
        bounds = item['bounds']
        for field in ['x', 'y', 'width', 'height']:
            value = bounds[field]
            if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
                raise ValueError('Object geometry must be finite numeric pixels')
            if value < 0 or (field in {'width', 'height'} and value == 0):
                raise ValueError('Object geometry must have nonnegative origins and positive extents')
        if bounds['x'] + bounds['width'] > design['canvas']['width_px'] or bounds['y'] + bounds['height'] > design['canvas']['height_px']:
            raise ValueError('Native object exceeds the canvas')
        if item['type'] == 'table':
            rows = item.get('rows')
            if not isinstance(rows, list) or not rows or not all(isinstance(row, list) and row for row in rows):
                raise ValueError('Native tables need a nonempty rows matrix')
            if len({len(row) for row in rows}) != 1:
                raise ValueError('Native table rows must have equal length')
            if any(not isinstance(value, (str, int, float, bool, type(None))) for row in rows for value in row):
                raise ValueError('Native table cells must be scalar values')
        else:
            if item.get('chart_type') not in {'column', 'bar', 'line'}:
                raise ValueError('Supported data chart types: column, bar, line')
            categories, series = item.get('categories'), item.get('series')
            if not isinstance(categories, list) or not categories or not isinstance(series, list) or not series:
                raise ValueError('Charts need categories and series')
            for entry in series:
                if not isinstance(entry.get('name'), str) or len(entry.get('values', [])) != len(categories):
                    raise ValueError('Each series needs a name and one value per category')
                if any(isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) for value in entry['values']):
                    raise ValueError('Chart values must be finite numbers; never invent missing values')
    return objects


def inspect_svg(path, design):
    root = ET.parse(path).getroot()
    expected = [0.0, 0.0, float(design['canvas']['width_px']), float(design['canvas']['height_px'])]
    actual = [float(part) for part in root.get('viewBox', '').replace(',', ' ').split()]
    if actual != expected:
        raise ValueError(path.name + ': viewBox differs from design.json')
    findings = []
    allowed_fonts = {design['font_latin'], design['font_east_asian'], 'Cambria Math'}
    allowed_colors = {color.lower() for color in design['palette'] + [design['text_color'], design['background_color'], design['table_header_fill']]}
    minimum_px = design['body_min_pt'] * 96 / 72
    def visit(element, inherited):
        tag = element.tag.rsplit('}', 1)[-1]
        if tag == 'metadata':
            return
        values = dict(inherited)
        values.update(element.attrib)
        for pair in element.get('style', '').split(';'):
            if ':' in pair:
                name, value = pair.split(':', 1)
                values[name.strip()] = value.strip()
        marker = element.get('data-pptx-replace-with') or element.get('data-pptx-native')
        if marker and marker != 'formula':
            raise ValueError('Use objects.json for chart/table data so project style rules apply')
        if element.get('data-pptx-replace-with') == 'formula':
            metadata = next((child for child in element if child.tag.rsplit('}', 1)[-1] == 'metadata'), None)
            if metadata is None:
                raise ValueError('Formula marker needs explicit JSON metadata')
            payload = json.loads(metadata.text or '{}')
            size = finite_positive(payload.get('font_size'), 'formula.font_size')
            if size < minimum_px - .01:
                findings.append({'severity': 'error', 'kind': 'formula-below-floor', 'size_pt': size * .75})
            if not any(child.tag.rsplit('}', 1)[-1] in {'text', 'path', 'line', 'g'} for child in element):
                findings.append({'severity': 'error', 'kind': 'formula-preview-missing'})
        if tag in {'text', 'tspan'} and ''.join(element.itertext()).strip():
            size = values.get('font-size', '')
            try:
                pixels = float(str(size).removesuffix('px'))
            except ValueError:
                findings.append({'severity': 'error', 'kind': 'font-size-unresolved', 'text': ''.join(element.itertext())[:70]})
            else:
                if pixels < minimum_px - 0.01 and values.get('data-role') != 'page-number':
                    findings.append({'severity': 'error', 'kind': 'font-below-floor', 'size_pt': pixels * .75, 'text': ''.join(element.itertext())[:70]})
            fonts = {part.strip().strip('"\'') for part in values.get('font-family', '').split(',')}
            first_font = values.get('font-family', '').split(',')[0].strip().strip('"\'')
            if not fonts & allowed_fonts or first_font not in allowed_fonts:
                findings.append({'severity': 'error', 'kind': 'font-not-in-design', 'font': values.get('font-family')})
            if 'transform' in element.attrib:
                findings.append({'severity': 'warning', 'kind': 'transformed-text-review'})
        for field in ['fill', 'stroke']:
            color = values.get(field, '').lower()
            if re.fullmatch(r'#[0-9a-f]{6}', color) and color not in allowed_colors:
                findings.append({'severity': 'error', 'kind': 'color-not-in-design', 'color': color})
        for child in element:
            visit(child, values)
    visit(root, {})
    return findings


def check(project):
    design = load_design(project)
    files = roster(project)
    objects = load_objects(project, len(files), design)
    backend = load_backend()
    from svg_to_pptx.drawingml.converter import convert_svg_to_slide_shapes
    results = []
    for number, path in enumerate(files, 1):
        try:
            findings = inspect_svg(path, design)
            convert_svg_to_slide_shapes(path, resource_root=project, slide_num=number,
                                        slide_count=len(files), native_objects=True,
                                        image_optimize=False, primary_language='zh-CN')
        except Exception as error:
            findings = [{'severity': 'error', 'kind': 'native-conversion-rejected', 'message': str(error)}]
        results.append({'slide': number, 'file': path.name, 'findings': findings})
    passed = not any(item['severity'] == 'error' for page in results for item in page['findings'])
    report = {'passed': passed, 'visual_reviewed': False, 'engine_commit': backend['commit'],
              'slides': results, 'native_data_objects': len(objects)}
    return report, design, files, objects


def rgb(color):
    from pptx.dml.color import RGBColor
    return RGBColor.from_string(color.lstrip('#'))


def set_font(font, design):
    from pptx.util import Pt
    font.name = design['font_east_asian']
    font.size = Pt(design['body_min_pt'])
    font.color.rgb = rgb(design['text_color'])


def add_data_objects(deck, objects, design):
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.oxml.xmlchemy import OxmlElement
    from pptx.util import Inches
    for item in objects:
        slide = deck.slides[item['slide'] - 1]
        bounds = item['bounds']
        frame = [Inches(bounds[field] / 96) for field in ['x', 'y', 'width', 'height']]
        if item['type'] == 'table':
            rows = item['rows']
            table = slide.shapes.add_table(len(rows), len(rows[0]), *frame).table
            for row_index, values in enumerate(rows):
                for column_index, value in enumerate(values):
                    cell = table.cell(row_index, column_index)
                    cell.text = '' if value is None else str(value)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = rgb(item.get('header_fill', design['table_header_fill']) if row_index == 0 else item.get('body_fill', design['background_color']))
                    for paragraph in cell.text_frame.paragraphs:
                        for run in paragraph.runs:
                            set_font(run.font, design)
                            run.font.bold = row_index == 0
                            properties = run._r.get_or_add_rPr()
                            from pptx.oxml.ns import qn
                            for tag, face in [('a:latin', design['font_latin']), ('a:ea', design['font_east_asian'])]:
                                child = properties.find(qn(tag))
                                if child is None:
                                    child = OxmlElement(tag)
                                    properties.append(child)
                                child.set('typeface', face)
        else:
            data = CategoryChartData()
            data.categories = item['categories']
            for series in item['series']:
                data.add_series(series['name'], series['values'])
            types = {'column': XL_CHART_TYPE.COLUMN_CLUSTERED, 'bar': XL_CHART_TYPE.BAR_CLUSTERED, 'line': XL_CHART_TYPE.LINE}
            chart = slide.shapes.add_chart(types[item['chart_type']], *frame, data).chart
            chart.has_title = False
            chart.has_legend = bool(item.get('legend', False))
            for axis in [chart.category_axis, chart.value_axis]:
                set_font(axis.tick_labels.font, design)
            if chart.has_legend:
                set_font(chart.legend.font, design)
            colors = item.get('series_colors', [design['palette'][0]])
            for index, series in enumerate(chart.series):
                color = rgb(colors[index % len(colors)])
                if item['chart_type'] == 'line':
                    series.format.line.color.rgb = color
                else:
                    series.format.fill.solid()
                    series.format.fill.fore_color.rgb = color
                    series.format.line.color.rgb = color


def export(project, output, report, design, files, objects):
    from pptx import Presentation
    from svg_to_pptx.pptx_package.builder import create_pptx_with_native_svg
    if output.exists():
        raise FileExistsError('Export requires a new output filename; preserve the existing deck')
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='ppt-production-build-') as directory:
        intermediate = Path(directory) / 'native.pptx'
        success = create_pptx_with_native_svg(
            files, intermediate, resource_root=project, verbose=False,
            transition='none', animation=None, native_objects=True,
            use_native_shapes=True, use_compat_mode=False, image_optimize=False,
            pptx_structure='flat', primary_language='zh-CN', workers=1,
            doc_metadata={'title': design.get('title', ''), 'creator': '', 'last_modified_by': ''})
        if not success:
            raise RuntimeError('Native conversion did not complete every page')
        deck = Presentation(intermediate)
        add_data_objects(deck, objects, design)
        deck.core_properties.author = ''
        deck.core_properties.last_modified_by = ''
        deck.save(intermediate)
        import pptx_audit
        audit = pptx_audit.audit(intermediate, min_pt=design['body_min_pt'])
        report['pptx_audit'] = audit
        if any(item['severity'] == 'error' for item in audit):
            raise ValueError('Generated PPTX failed the font-size audit')
        output.write_bytes(intermediate.read_bytes())
    report['output'] = output.name
    report['output_sha256'] = hashlib.sha256(output.read_bytes()).hexdigest()
    report['visual_reviewed'] = False


def main():
    parser = argparse.ArgumentParser(description='Style-constrained native PPTX export; visual review remains required')
    parser.add_argument('operation', choices=['check', 'export'])
    parser.add_argument('project', type=Path)
    parser.add_argument('--output', type=Path)
    parser.add_argument('--report', type=Path)
    args = parser.parse_args()
    report = None
    try:
        project = args.project.resolve(strict=True)
        report, design, files, objects = check(project)
        if report['passed'] and args.operation == 'export':
            if args.output is None:
                raise ValueError('export requires --output')
            export(project, args.output.resolve(), report, design, files, objects)
        result = 0 if report['passed'] else 1
    except Exception as error:
        report = report or {}
        report.update({'passed': False, 'visual_reviewed': False, 'error': str(error)})
        result = 2
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(text + '\n', encoding='utf-8')
    print(text)
    return result


if __name__ == '__main__':
    raise SystemExit(main())
