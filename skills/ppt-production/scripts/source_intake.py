import argparse
import hashlib
import json
from pathlib import Path
import sys


def markdown_table(rows):
    if not rows:
        return ''
    width = max(len(row) for row in rows)
    cleaned = [[str(value if value is not None else '').replace('|', '\\|').replace('\n', '<br>')
                for value in row] + [''] * (width - len(row)) for row in rows]
    lines = ['| ' + ' | '.join(row) + ' |' for row in cleaned]
    lines.insert(1, '| ' + ' | '.join(['---'] * width) + ' |')
    return '\n'.join(lines)


def extract_source(path):
    suffix = path.suffix.lower()
    warnings = []
    blocks = []
    if suffix in {'.md', '.markdown', '.txt'}:
        return path.read_text(encoding='utf-8-sig'), warnings
    if suffix == '.pdf':
        from pypdf import PdfReader
        reader = PdfReader(str(path))
        if reader.is_encrypted and not reader.decrypt(''):
            raise ValueError('Encrypted PDF needs a readable, decrypted copy')
        for number, page in enumerate(reader.pages, 1):
            text = page.extract_text() or ''
            blocks.append(f'## Page {number}\n\n{text}')
            if not text.strip():
                warnings.append(f'Page {number}: no text extracted; inspect the original or use approved OCR')
        warnings.append('PDF figure, table, and formula layout is not preserved; inspect original pages')
    elif suffix == '.docx':
        from docx import Document
        from docx.table import Table
        from docx.text.paragraph import Paragraph
        document = Document(str(path))
        for child in document.element.body:
            tag = child.tag.rsplit('}', 1)[-1]
            if tag == 'p':
                blocks.append(Paragraph(child, document).text)
            elif tag == 'tbl':
                table = Table(child, document)
                blocks.append(markdown_table([[cell.text for cell in row.cells] for row in table.rows]))
        warnings.append('Images, equations, headers, footnotes, and merged-cell semantics require source inspection')
    elif suffix in {'.xlsx', '.xlsm'}:
        from openpyxl import load_workbook
        book = load_workbook(path, read_only=True, data_only=False)
        try:
            for sheet in book:
                rows = list(sheet.iter_rows(values_only=True))
                blocks.append(f'## Sheet: {sheet.title}\n\n' + markdown_table(rows))
        finally:
            book.close()
        warnings.append('Formula text is preserved, not recalculated; charts and cell formatting are not extracted')
    elif suffix == '.pptx':
        from pptx import Presentation
        def shape_text(shapes):
            result = []
            for shape in shapes:
                if getattr(shape, 'has_table', False):
                    result.append(markdown_table([[cell.text for cell in row.cells] for row in shape.table.rows]))
                elif getattr(shape, 'has_text_frame', False):
                    result.append(shape.text)
                if hasattr(shape, 'shapes'):
                    result.extend(shape_text(shape.shapes))
            return result
        for number, slide in enumerate(Presentation(str(path)).slides, 1):
            blocks.append(f'## Slide {number}\n\n' + '\n\n'.join(shape_text(slide.shapes)))
        warnings.append('Reading order, charts, native equations, images, and notes require original-deck inspection')
    else:
        raise ValueError('Unsupported source format: ' + suffix)
    return '\n\n'.join(blocks), warnings


def intake(paths, output):
    output.mkdir(parents=True, exist_ok=True)
    manifest_path = output / 'sources.json'
    if manifest_path.exists():
        raise FileExistsError('Use a new intake output directory; sources.json already exists')
    records = []
    for supplied in paths:
        source = Path(supplied).resolve(strict=True)
        digest = hashlib.sha256(source.read_bytes()).hexdigest()
        target = output / (source.stem + '-' + digest[:10] + '.md')
        if target.resolve() == source or target.exists():
            raise FileExistsError('Intake refuses to overwrite a source or existing extraction: ' + str(target))
        text, warnings = extract_source(source)
        target.write_text(f'# Source: {source.name}\n\n{text}\n', encoding='utf-8')
        records.append({'source_name': source.name, 'sha256': digest,
                        'extraction': target.name, 'warnings': warnings})
    manifest_path.write_text(json.dumps(records, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    return records


def main():
    parser = argparse.ArgumentParser(description='Extract local sources with provenance; no uploads or OCR')
    parser.add_argument('sources', nargs='+')
    parser.add_argument('--out', type=Path, required=True)
    args = parser.parse_args()
    try:
        print(json.dumps(intake(args.sources, args.out), ensure_ascii=False, indent=2))
        return 0
    except (OSError, ValueError, ImportError) as error:
        print(str(error), file=sys.stderr)
        return 2


if __name__ == '__main__':
    raise SystemExit(main())
