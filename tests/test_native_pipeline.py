import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import zipfile

from pptx import Presentation

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / 'skills' / 'ppt-production' / 'scripts'
sys.path.insert(0, str(SCRIPTS))
import native_export
import source_intake


def fixture(project):
    (project / 'svg_output').mkdir(parents=True)
    design = {'canvas': {'width_px': 1280, 'height_px': 720}, 'body_min_pt': 20,
              'font_latin': 'Arial', 'font_east_asian': 'Microsoft YaHei',
              'text_color': '#222222', 'palette': ['#540054', '#FFFFFF', '#EDEDED'],
              'title': 'Native pipeline test'}
    (project / 'design.json').write_text(json.dumps(design), encoding='utf-8')
    (project / 'svg_output' / '01.svg').write_text('''<svg xmlns="http://www.w3.org/2000/svg" width="1280" height="720" viewBox="0 0 1280 720">
<rect x="0" y="0" width="1280" height="720" fill="#FFFFFF"/>
<rect x="60" y="50" width="1160" height="70" fill="#540054"/>
<text x="80" y="98" font-family="Microsoft YaHei" font-size="36" fill="#FFFFFF">原生对象转换验证</text>
<text x="70" y="160" font-family="Microsoft YaHei" font-size="28" fill="#222222">数据与公式由结构化输入生成</text>
<g data-pptx-replace-with="formula"><metadata type="application/json">{"latex":"\\\\alpha=\\\\frac{k}{\\\\rho c_p}","x":90,"y":560,"width":450,"height":90,"font_size":34,"color":"#222222"}</metadata><text x="90" y="610" font-family="Cambria Math" font-size="34" fill="#222222">α = k/(ρ cₚ)</text></g>
</svg>''', encoding='utf-8')
    objects = [
        {'slide': 1, 'type': 'chart', 'chart_type': 'column',
         'bounds': {'x': 70, 'y': 210, 'width': 550, 'height': 300},
         'categories': ['A', 'B'], 'series': [{'name': '示例数据', 'values': [2, 3]}]},
        {'slide': 1, 'type': 'table', 'bounds': {'x': 720, 'y': 240, 'width': 450, 'height': 220},
         'rows': [['参数', '单位'], ['长度', 'mm'], ['温度', 'K']], 'header_fill': '#EDEDED'},
    ]
    (project / 'objects.json').write_text(json.dumps(objects, ensure_ascii=False), encoding='utf-8')


class NativePipelineTests(unittest.TestCase):
    def test_export_contains_shapes_chart_workbook_table_and_office_math(self):
        with tempfile.TemporaryDirectory() as directory:
            project = Path(directory)
            fixture(project)
            output = project / 'native.pptx'
            result = subprocess.run([sys.executable, str(SCRIPTS / 'native_export.py'), 'export', str(project),
                                     '--output', str(output)], capture_output=True, text=True, encoding='utf-8')
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            deck = Presentation(output)
            self.assertEqual(len(deck.slides), 1)
            self.assertTrue(any(shape.has_chart for shape in deck.slides[0].shapes))
            self.assertTrue(any(shape.has_table for shape in deck.slides[0].shapes))
            with zipfile.ZipFile(output) as package:
                xml = package.read('ppt/slides/slide1.xml').decode('utf-8')
                self.assertIn('oMath', xml)
                self.assertIn('原生对象转换验证', xml)
                self.assertTrue(any(name.startswith('ppt/embeddings/') and name.endswith('.xlsx') for name in package.namelist()))

    def test_small_text_blocks_check(self):
        with tempfile.TemporaryDirectory() as directory:
            project = Path(directory)
            fixture(project)
            svg = project / 'svg_output' / '01.svg'
            svg.write_text(svg.read_text(encoding='utf-8').replace('font-size="28"', 'font-size="12"'), encoding='utf-8')
            report, *_ = native_export.check(project)
            self.assertFalse(report['passed'])

    def test_chart_data_mismatch_is_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            project = Path(directory)
            fixture(project)
            objects = json.loads((project / 'objects.json').read_text(encoding='utf-8'))
            objects[0]['series'][0]['values'] = [2]
            (project / 'objects.json').write_text(json.dumps(objects), encoding='utf-8')
            with self.assertRaisesRegex(ValueError, 'one value per category'):
                native_export.check(project)

    def test_foreign_font_and_color_do_not_bypass_project_rules(self):
        with tempfile.TemporaryDirectory() as directory:
            project = Path(directory)
            fixture(project)
            svg = project / 'svg_output' / '01.svg'
            svg.write_text(svg.read_text(encoding='utf-8').replace('font-family="Microsoft YaHei"', 'font-family="Courier New"').replace('#540054', '#FF0000'), encoding='utf-8')
            report, *_ = native_export.check(project)
            self.assertFalse(report['passed'])
            kinds = {item['kind'] for item in report['slides'][0]['findings']}
            self.assertIn('font-not-in-design', kinds)
            self.assertIn('color-not-in-design', kinds)

    def test_small_native_formula_is_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            project = Path(directory)
            fixture(project)
            svg = project / 'svg_output' / '01.svg'
            svg.write_text(svg.read_text(encoding='utf-8').replace('"font_size":34', '"font_size":12'), encoding='utf-8')
            report, *_ = native_export.check(project)
            self.assertFalse(report['passed'])

    def test_intake_keeps_original_and_tracks_provenance(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / '科研资料.txt'
            source.write_text('Greek symbols: α β γ Δ', encoding='utf-8')
            original = source.read_bytes()
            records = source_intake.intake([source], root / 'sources')
            self.assertEqual(source.read_bytes(), original)
            self.assertEqual(len(records[0]['sha256']), 64)
            self.assertIn('α β γ Δ', (root / 'sources' / records[0]['extraction']).read_text(encoding='utf-8'))
            with self.assertRaises(FileExistsError):
                source_intake.intake([source], root / 'sources')


if __name__ == '__main__':
    unittest.main()
