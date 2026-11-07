#!/usr/bin/env python3
"""Audit editable PPTX text; flag geometry and glyphs for visual review.

Exit 1 means a confirmed font-floor violation; warnings do not fail a deck.
Exit 2 means the input could not be inspected. Empty placeholders are ignored.
Bounding boxes cannot prove clipping, glyph support, or rendered overlap.
"""
import argparse
import json
import sys
from collections import Counter
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.oxml.ns import qn

SYMBOLS = set("\u2705\u274c\u26a0\u2605\u2606\u2714\u2716\u2717\u2718")
EMU_IN = 914400


def walk_shapes(shapes):
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from walk_shapes(shape.shapes)
        else:
            yield shape


def paragraphs(shape):
    if getattr(shape, "has_text_frame", False):
        yield from shape.text_frame.paragraphs
    if getattr(shape, "has_table", False):
        for row in shape.table.rows:
            for cell in row.cells:
                if not cell.is_spanned:
                    yield from cell.text_frame.paragraphs


def page_number(shape):
    if shape.is_placeholder and shape.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER:
        return True
    fields = shape._element.xpath('.//a:fld')
    return bool(fields) and all(f.get('type') == 'slidenum' for f in fields) and not shape._element.xpath('.//a:r/a:t')


def font_size(run, paragraph, shape):
    if run.font.size is not None:
        return run.font.size.pt
    if paragraph.font.size is not None:
        return paragraph.font.size.pt
    level = paragraph.level + 1
    for node in shape._element.xpath(f'.//a:lstStyle/a:lvl{level}pPr/a:defRPr'):
        if node.get('sz'):
            return int(node.get('sz')) / 100
    return None  # Master/theme inheritance requires rendering, never assume zero.


def content(shape):
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        return any(content(s) for s in shape.shapes)
    return (shape.shape_type == MSO_SHAPE_TYPE.PICTURE
            or getattr(shape, "has_table", False)
            or getattr(shape, "has_chart", False)
            or (getattr(shape, "has_text_frame", False) and bool(shape.text_frame.text.strip())))


def overlap_ratio(a, b):
    vals = (a.left, a.top, a.width, a.height, b.left, b.top, b.width, b.height)
    if any(v is None for v in vals) or min(a.width, a.height, b.width, b.height) <= 0:
        return 0
    width = min(a.left + a.width, b.left + b.width) - max(a.left, b.left)
    height = min(a.top + a.height, b.top + b.height) - max(a.top, b.top)
    if min(width, height) <= EMU_IN * .02:
        return 0
    return width * height / min(a.width * a.height, b.width * b.height)


def audit(path, min_pt=20.0, overlap_threshold=.30):
    prs = Presentation(path)
    findings = []
    def add(slide, kind, severity, detail, text="", **extra):
        findings.append(dict(slide=slide, kind=kind, severity=severity,
                             detail=detail, text=text[:80], **extra))
    for idx, slide in enumerate(prs.slides, 1):
        shapes = list(slide.shapes)
        for shape in walk_shapes(shapes):
            unresolved = []
            for para in paragraphs(shape):
                for run in para.runs:
                    text = run.text.strip()
                    if not text:
                        continue
                    size = font_size(run, para, shape)
                    if not page_number(shape):
                        if size is None:
                            unresolved.append(text)
                        elif size < min_pt:
                            add(idx, 'font-below-floor', 'error', f'{size:g}pt in {shape.name}', text)
                    bad = SYMBOLS.intersection(text)
                    if bad:
                        add(idx, 'symbol-glyph-review', 'warning', ''.join(sorted(bad)), text)
            if unresolved:
                add(idx, 'font-inherited-review', 'warning', shape.name, ' / '.join(unresolved))
        # Only top-level content shapes: never compare group parents to children.
        # Skip empty/decorative shapes and full-slide background pictures.
        boxes = [s for s in shapes if content(s) and s.width and s.height
                 and not (s.shape_type == MSO_SHAPE_TYPE.PICTURE
                          and s.width * s.height >= .9 * prs.slide_width * prs.slide_height)]
        for i, a in enumerate(boxes):
            for b in boxes[i + 1:]:
                ratio = overlap_ratio(a, b)
                if ratio >= overlap_threshold:
                    add(idx, 'overlap-review', 'warning', f'{a.name} / {b.name}',
                        overlap_fraction=round(ratio, 3))
    return findings


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('deck')
    ap.add_argument('--min-pt', type=float, default=20.0)
    ap.add_argument('--overlap-threshold', type=float, default=.30,
                    help='Intersection / smaller shape area; visual-review warning only')
    ap.add_argument('--json', action='store_true')
    args = ap.parse_args()
    if args.min_pt <= 0 or not 0 < args.overlap_threshold <= 1:
        ap.error('min-pt must be positive; overlap-threshold must be in (0, 1]')
    try:
        findings = audit(args.deck, args.min_pt, args.overlap_threshold)
    except Exception as exc:
        print(f'Cannot inspect deck: {exc}', file=sys.stderr)
        return 2
    if args.json:
        print(json.dumps(findings, ensure_ascii=False, indent=2))
    else:
        counts = Counter(f['severity'] for f in findings)
        print(f"errors={counts['error']}, review_warnings={counts['warning']}")
        for f in findings:
            print(f"slide {f['slide']} [{f['severity']}] {f['kind']}: {f['detail']} | {f['text']}")
        print('Rendered slides still require visual review, including unreported intentional overlays.')
    return int(any(f['severity'] == 'error' for f in findings))


if __name__ == '__main__':
    sys.exit(main())