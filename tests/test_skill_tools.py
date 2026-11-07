import copy
import importlib.util
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest

from pptx import Presentation
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / 'skills' / 'ppt-production' / 'scripts'
LIBRARY = ROOT / 'skills' / 'ppt-production' / 'assets' / 'library'
sys.path.insert(0, str(SCRIPTS))
import pptx_audit as audit
import index_library as indexer
import check_library as checker
import render_catalog as gallery


def text(shapes, left, top, content, size=24):
    shape = shapes.add_textbox(Inches(left), Inches(top), Inches(3), Inches(0.6))
    run = shape.text_frame.paragraphs[0].add_run()
    run.text = content
    run.font.size = Pt(size)
    return shape


class AuditTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name) / 'fixture.pptx'
        self.deck = Presentation()

    def tearDown(self):
        self.tmp.cleanup()

    def findings(self):
        self.deck.save(self.path)
        return audit.audit(self.path)

    def test_empty_placeholders_and_background_are_not_errors(self):
        slide = self.deck.slides.add_slide(self.deck.slide_layouts[1])
        slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE,
                               Inches(.5), Inches(1), Inches(4), Inches(2))
        text(slide.shapes, .7, 1.2, 'Readable body')
        self.assertEqual(self.findings(), [])

    def test_group_and_numeric_content_are_checked(self):
        slide = self.deck.slides.add_slide(self.deck.slide_layouts[6])
        text(slide.shapes.add_group_shape().shapes, 1, 1, 'Grouped text', 12)
        text(slide.shapes, 1, 3, '123', 12)
        errors = [f for f in self.findings() if f['severity'] == 'error']
        self.assertEqual(len(errors), 2)
        self.assertEqual({f['text'] for f in errors}, {'Grouped text', '123'})

    def test_paragraph_inheritance_is_detected(self):
        slide = self.deck.slides.add_slide(self.deck.slide_layouts[6])
        shape = text(slide.shapes, 1, 1, 'Inherited size')
        paragraph = shape.text_frame.paragraphs[0]
        paragraph.runs[0].font.size = None
        paragraph.font.size = Pt(18)
        self.assertEqual(self.findings()[0]['kind'], 'font-below-floor')

    def test_unknown_inheritance_is_warning(self):
        slide = self.deck.slides.add_slide(self.deck.slide_layouts[6])
        shape = text(slide.shapes, 1, 1, 'Master size unknown')
        shape.text_frame.paragraphs[0].runs[0].font.size = None
        finding = self.findings()[0]
        self.assertEqual(finding['kind'], 'font-inherited-review')
        self.assertEqual(finding['severity'], 'warning')

    def test_overlap_is_review_not_failure(self):
        slide = self.deck.slides.add_slide(self.deck.slide_layouts[6])
        text(slide.shapes, 1, 1, 'Text A')
        text(slide.shapes, 1.2, 1.1, 'Text B')
        findings = self.findings()
        self.assertEqual(len(findings), 1)
        self.assertEqual(findings[0]['kind'], 'overlap-review')
        result = subprocess.run([sys.executable, str(SCRIPTS / 'pptx_audit.py'), str(self.path), '--json'],
                                capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(len(json.loads(result.stdout)), 1)

    def test_small_text_returns_failure(self):
        slide = self.deck.slides.add_slide(self.deck.slide_layouts[6])
        text(slide.shapes, 1, 1, 'Too small', 12)
        self.findings()
        result = subprocess.run([sys.executable, str(SCRIPTS / 'pptx_audit.py'), str(self.path)],
                                capture_output=True, text=True)
        self.assertEqual(result.returncode, 1, result.stderr)

    def test_unreadable_file_returns_two(self):
        result = subprocess.run([sys.executable, str(SCRIPTS / 'pptx_audit.py'), str(self.path)],
                                capture_output=True, text=True)
        self.assertEqual(result.returncode, 2)


class LibraryTests(unittest.TestCase):
    def test_bundled_library_integrity(self):
        result = checker.validate(LIBRARY)
        self.assertEqual(result['pages'], 13)

    def test_relationship_paths(self):
        self.assertEqual(indexer.resolve_target('ppt/slides/slide1.xml', '../media/image1.png'),
                         'ppt/media/image1.png')
        self.assertEqual(indexer.resolve_target('ppt/presentation.xml', '/ppt/slides/slide1.xml'),
                         'ppt/slides/slide1.xml')

    def test_index_cli_uses_relative_source_path(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            deck = Presentation()
            deck.slides.add_slide(deck.slide_layouts[6])
            deck.save(root / 'deck.pptx')
            result = subprocess.run([sys.executable, str(SCRIPTS / 'index_library.py'), str(root / 'deck.pptx'),
                                     '--out', str(root / 'inventory.json')], capture_output=True, text=True)
            self.assertEqual(result.returncode, 0, result.stderr)
            data = json.loads((root / 'inventory.json').read_text(encoding='utf-8'))
            self.assertEqual(data['source_path'], 'deck.pptx')

    def test_changed_source_is_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / 'library'
            shutil.copytree(LIBRARY, root)
            with (root / 'source.pptx').open('ab') as stream:
                stream.write(b'changed-source')
            with self.assertRaisesRegex(ValueError, 'Source changed'):
                checker.validate(root)

    def test_changed_preview_is_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / 'library'
            shutil.copytree(LIBRARY, root)
            (root / 'previews' / 'page-01.png').write_bytes(b'not-the-expected-preview')
            with self.assertRaisesRegex(ValueError, 'changed preview'):
                checker.validate(root)

    def test_changed_catalog_alignment_is_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / 'library'
            shutil.copytree(LIBRARY, root)
            path = root / 'catalog.json'
            data = json.loads(path.read_text(encoding='utf-8'))
            data['entries'][0]['slide_id'] = 'not-the-slide'
            path.write_text(json.dumps(data), encoding='utf-8')
            with self.assertRaisesRegex(ValueError, 'mismatched slide_id'):
                checker.validate(root)

    def test_gallery_is_reproducible_and_links_every_preview(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / 'library'
            shutil.copytree(LIBRARY, root)
            self.assertEqual(gallery.build(root), 13)
            for name in ('index.md', 'gallery.html'):
                self.assertEqual((root / name).read_bytes(), (LIBRARY / name).read_bytes())
            content = (root / 'gallery.html').read_text(encoding='utf-8')
            for page in range(1, 14):
                self.assertIn(f'previews/page-{page:02d}.png', content)


if __name__ == '__main__':
    unittest.main()