#!/usr/bin/env python3
"""Validate library fingerprints and page alignment before reusing examples."""
import argparse
import hashlib
import json
from pathlib import Path
from index_library import inventory


def validate(library):
    root = Path(library).resolve()
    catalog = json.loads((root / 'catalog.json').read_text(encoding='utf-8'))
    stored = json.loads((root / 'inventory.json').read_text(encoding='utf-8'))
    source_ref = catalog['source_path']
    source = (root / source_ref).resolve()
    current = inventory(source, source_ref)
    if current['source_sha256'] != catalog['source_sha256'] or current['source_sha256'] != stored['source_sha256']:
        raise ValueError('Source changed: regenerate inventory/previews and review semantic labels.')
    if current['slides'] != stored['slides']:
        raise ValueError('Native object inventory differs from source.')
    entries = catalog['entries']
    if len(entries) != len(current['slides']):
        raise ValueError('Catalog page count differs from source.')
    if len({e['id'] for e in entries}) != len(entries):
        raise ValueError('Duplicate catalog entry IDs.')
    for entry, slide in zip(entries, current['slides']):
        for key in ('page', 'slide_id', 'slide_part', 'slide_xml_sha256', 'preview'):
            if entry[key] != slide[key]:
                raise ValueError(f'Page {slide["page"]}: mismatched {key}')
        preview = (root / entry['preview']).resolve()
        if not preview.is_relative_to(root):
            raise ValueError('Preview path escapes library.')
        if not preview.is_file() or hashlib.sha256(preview.read_bytes()).hexdigest() != entry['preview_sha256']:
            raise ValueError(f'Page {slide["page"]}: missing or changed preview; review required.')
    return {'pages': len(entries), 'source_sha256': current['source_sha256'], 'integrity': 'passed'}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('library')
    args = parser.parse_args()
    try:
        print(json.dumps(validate(args.library), ensure_ascii=False))
    except (OSError, ValueError, KeyError) as error:
        parser.exit(1, str(error) + '\n')


if __name__ == '__main__':
    main()