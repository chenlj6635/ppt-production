#!/usr/bin/env python3
"""Read PPTX into a fingerprinted inventory without modifying the source.

python index_library.py SOURCE.pptx --out LIBRARY/inventory.json
Paths are relative to the output inventory. --source-ref overrides that value.
"""
import argparse
import hashlib
import json
import os
import posixpath
import zipfile
from datetime import datetime, timezone
from pathlib import Path
import xml.etree.ElementTree as ET

NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
SHAPE_TAGS = {'sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp'}


def tag(element):
    return element.tag.rsplit('}', 1)[-1]


def resolve_target(part, target):
    if target.startswith('/'):
        return posixpath.normpath(target).lstrip('/')
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), target))


def shape(element):
    properties = next((x for x in element if tag(x).startswith('nv')), None)
    c = properties.find('p:cNvPr', NS) if properties is not None else None
    transform = None
    for query in ('p:spPr/a:xfrm', 'p:grpSpPr/a:xfrm', 'p:xfrm'):
        transform = element.find(query, NS)
        if transform is not None:
            break
    coordinates = {}
    if transform is not None:
        for child in transform:
            coordinates[tag(child)] = {k: int(v) for k, v in child.attrib.items()}
    return {
        'kind': tag(element),
        'id': c.get('id') if c is not None else None,
        'name': c.get('name') if c is not None else None,
        'text': '\n'.join(''.join(t.text or '' for t in p.findall('.//a:t', NS))
                          for p in element.findall('.//a:p', NS)),
        'transform_emu': coordinates,
        'explicit_font_pts': sorted({int(x.get('sz')) / 100
                                    for x in element.iter() if x.get('sz')}),
        'image_rids': [x.get('{' + NS['r'] + '}embed')
                      for x in element.findall('.//a:blip', NS)],
        'children': [shape(x) for x in element if tag(x) in SHAPE_TAGS],
    }


def inventory(path, source_reference=None):
    source = Path(path).resolve()
    data = {
        'schema_version': 1,
        'source_path': source_reference if source_reference is not None else str(source),
        'source_sha256': hashlib.sha256(source.read_bytes()).hexdigest(),
        'indexed_at': datetime.now(timezone.utc).isoformat(),
        'source_modified_at': datetime.fromtimestamp(source.stat().st_mtime, timezone.utc).isoformat(),
        'coordinate_note': ('Top-level xfrm is slide-local; group children require '
                            'chOff/chExt scaling. Missing geometry may be inherited.'),
        'slides': [],
    }
    with zipfile.ZipFile(source) as package:
        presentation = ET.fromstring(package.read('ppt/presentation.xml'))
        size = presentation.find('p:sldSz', NS)
        data['canvas_inches'] = [int(size.get(k)) / 914400 for k in ('cx', 'cy')]
        relationships = {
            r.get('Id'): resolve_target('ppt/presentation.xml', r.get('Target'))
            for r in ET.fromstring(package.read('ppt/_rels/presentation.xml.rels'))
        }
        for page, slide in enumerate(presentation.findall('p:sldIdLst/p:sldId', NS), 1):
            part = relationships[slide.get('{' + NS['r'] + '}id')]
            raw = package.read(part)
            root = ET.fromstring(raw)
            tree = root.find('p:cSld/p:spTree', NS)
            relpart = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')
            related = []
            if relpart in package.namelist():
                for relation in ET.fromstring(package.read(relpart)):
                    external = relation.get('TargetMode') == 'External'
                    target = relation.get('Target')
                    related.append({
                        'id': relation.get('Id'),
                        'type': relation.get('Type').rsplit('/', 1)[-1],
                        'target': target if external else resolve_target(part, target),
                        'external': external,
                    })
            data['slides'].append({
                'page': page,
                'slide_id': slide.get('id'),
                'slide_part': part,
                'slide_xml_sha256': hashlib.sha256(raw).hexdigest(),
                'text': [''.join(t.itertext()) for t in root.findall('.//a:t', NS)],
                'shapes': [shape(x) for x in tree if tag(x) in SHAPE_TAGS],
                'relationships': related,
                'preview': f'previews/page-{page:02d}.png',
            })
    return data


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source')
    parser.add_argument('--out', required=True)
    parser.add_argument('--source-ref')
    args = parser.parse_args()
    out = Path(args.out).resolve()
    source_reference = args.source_ref or os.path.relpath(Path(args.source).resolve(), out.parent).replace('\\', '/')
    result = inventory(args.source, source_reference)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({'pages': len(result['slides']), 'source_sha256': result['source_sha256'], 'inventory': str(out)}, ensure_ascii=False))


if __name__ == '__main__':
    main()