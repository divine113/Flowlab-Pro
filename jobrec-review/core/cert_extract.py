"""Deterministic certificate extraction engine — template-driven.

No AI, no LLM, no external API. Every certificate is read against an existing
template to discover which fields are required, then matched dynamically using
exact match, normalized match, synonyms, regular expressions, contextual
sentence patterns, data-type validation, measurement-unit validation, and
proximity analysis.

Public API:
    extract_certificate(path, filename, template) -> dict
        template: EquipmentCertificateTemplate instance
    populate_template_workbook(template, field_matches) -> (Workbook, str)
        Returns an xlsx workbook with values populated into the template layout.
"""
from __future__ import annotations

import re
import csv
import difflib
import io
import time
from datetime import date, datetime
from pathlib import Path

import pdfplumber
from openpyxl import Workbook, load_workbook

from django.conf import settings

from .cert_fields import docx_structure, extract

# ---------------------------------------------------------------------------
# Date parsing
# ---------------------------------------------------------------------------

def _parse_iso_date(value):
    """Parse a date string, returning a date object or None.

    Accepts ordinals (1st, 2nd, 3rd, 27th), full and short month names,
    commas, dots and slashes as separators.
    """
    if not value:
        return None
    value = value.strip()
    # Remove ordinal suffixes (1st, 2nd, 3rd, 27th, etc.)
    value = re.sub(r'(\d+)(st|nd|rd|th)', r'\1', value, flags=re.IGNORECASE)
    formats = (
        '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%d-%m-%Y', '%Y/%m/%d',
        '%d.%m.%Y', '%m.%d.%Y', '%d.%m.%y',
        '%d %b %Y', '%b %d, %Y', '%d %B %Y', '%B %d, %Y',
        '%d %b, %Y', '%d %B, %Y', '%b %d %Y', '%B %d %Y',
    )
    for fmt in formats:
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    return None


# ---------------------------------------------------------------------------
# Document readers
# ---------------------------------------------------------------------------

_MIN_TEXT_CHARS = 20
_MAX_ROWS = 500


def _normalize_whitespace(text):
    text = (text or '').replace('\r\n', '\n').replace('\r', '\n').replace('\x00', ' ')
    text = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', ' ', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text


def _tables_from_text(text):
    tables = []
    current = []
    current_width = None
    for raw in text.split('\n'):
        line = raw.strip()
        if not line:
            if len(current) >= 2:
                tables.append(current)
            current, current_width = [], None
            continue
        if '\t' in line:
            row = [cell.strip() for cell in line.split('\t')]
        elif '|' in line:
            row = [cell.strip() for cell in line.split('|')]
        else:
            if len(current) >= 2:
                tables.append(current)
            current, current_width = [], None
            continue
        non_empty = [cell for cell in row if cell]
        if len(non_empty) < 2:
            if len(current) >= 2:
                tables.append(current)
            current, current_width = [], None
            continue
        if current and len(row) != current_width:
            if len(current) >= 2:
                tables.append(current)
            current, current_width = [], None
            continue
        current.append(row)
        current_width = len(row)
    if len(current) >= 2:
        tables.append(current)
    return tables


def _read_word_binary(path):
    data = Path(path).read_bytes()
    utf16_chunks = re.findall(rb'(?:[\x09\x0A-\x0D\x20-\xFF]\x00){6,}', data)
    utf16_text = '\n'.join(chunk.decode('utf-16-le', 'ignore') for chunk in utf16_chunks)
    ascii_chunks = b'\n'.join(re.findall(rb'[\x20-\x7E]{6,}', data)).decode('latin-1', 'ignore')
    if len(utf16_text.strip()) >= 20:
        text = _normalize_whitespace(utf16_text)
    else:
        text = _normalize_whitespace(ascii_chunks)
    return text, _tables_from_text(text), 'doc'


def _read_docx(path):
    from docx import Document
    document = Document(path)
    paragraphs = [p.text for p in document.paragraphs]
    tables = [[[cell.text.strip() for cell in row.cells] for row in table.rows]
              for table in document.tables]
    text = '\n'.join(paragraphs)
    for table in tables:
        for row in table:
            text += '\n' + ' | '.join(cell for cell in row if cell)
    return _normalize_whitespace(text), tables, 'docx'


def _read_pdf(path):
    with pdfplumber.open(path) as pdf:
        text = '\n'.join((page.extract_text() or '') for page in pdf.pages)
        tables = []
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                tables.append([[('' if cell is None else str(cell)) for cell in row] for row in table])
    if len(text.strip()) >= _MIN_TEXT_CHARS:
        return _normalize_whitespace(text), tables, 'pdf'
    if _ocr_available():
        return _read_pdf_ocr(path)
    return _normalize_whitespace(text), tables, 'pdf-scanned-no-ocr'


def _read_pdf_ocr(path):
    import fitz
    import tempfile
    from .ocr_engine import extract as ocr_extract
    text_parts, tables = [], []
    with fitz.open(path) as doc:
        for page in doc:
            pix = page.get_pixmap(dpi=200)
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as handle:
                temp = handle.name
            pix.save(temp)
            try:
                result = ocr_extract(temp)
            finally:
                Path(temp).unlink(missing_ok=True)
            lines = result.get('lines') or []
            if lines:
                rows = _group_ocr_lines(lines)
                page_tables = _guess_table_from_lines(rows)
                tables.extend(page_tables)
                text_parts.append('\n'.join(' | '.join(row) for row in rows))
            else:
                text_parts.append(result.get('text') or '')
    return _normalize_whitespace('\n'.join(text_parts)), tables, 'pdf-ocr'


def _read_excel(path, suffix):
    text_parts = []
    tables = []
    if suffix == '.xlsx':
        workbook = load_workbook(path, data_only=True, read_only=True)
        try:
            for sheet in workbook.worksheets:
                lines, table = [], []
                for row in sheet.iter_rows(values_only=True):
                    values = [_cell_text(cell) for cell in row]
                    line = ' | '.join(value for value in values if value)
                    if line:
                        lines.append(line)
                    if any(values):
                        table.append(list(values))
                if len(table) >= 2:
                    for group in _group_by_width(table):
                        if len(group) >= 2:
                            tables.append(_trim_table(group))
                if lines:
                    text_parts.append('\n'.join(lines))
        finally:
            workbook.close()
    else:
        import xlrd
        book = xlrd.open_workbook(path, on_demand=True)
        try:
            for sheet in book.sheets():
                lines, table = [], []
                for row_index in range(min(sheet.nrows, _MAX_ROWS)):
                    values = [_cell_text(sheet.cell_value(row_index, col_index))
                              for col_index in range(min(sheet.ncols, 30))]
                    line = ' | '.join(value for value in values if value)
                    if line:
                        lines.append(line)
                    if any(values):
                        table.append(list(values))
                if len(table) >= 2:
                    for group in _group_by_width(table):
                        if len(group) >= 2:
                            tables.append(_trim_table(group))
                if lines:
                    text_parts.append('\n'.join(lines))
        finally:
            book.release_resources()
    return _normalize_whitespace('\n'.join(text_parts)), tables, 'excel'


def _read_csv(path):
    with open(path, newline='', encoding='utf-8-sig', errors='ignore') as handle:
        rows = list(csv.reader(handle))
    table = [list(row) for row in rows if any(value.strip() for value in row)]
    text = '\n'.join(' | '.join(cell.strip() for cell in row) for row in table)
    return _normalize_whitespace(text), [_trim_table(table)] if len(table) >= 2 else [], 'csv'


def _cell_text(value):
    if value is None:
        return ''
    return re.sub(r'\s+', ' ', str(value)).strip()


def _trim_table(rows):
    if not rows:
        return []
    width = max(len(row) for row in rows)
    return [list(row) + [''] * (width - len(row)) for row in rows]


def _ocr_available():
    try:
        from .ocr_engine import is_available
        return is_available()
    except Exception:
        return False


def _group_ocr_lines(lines):
    items = []
    for line in lines:
        box = line.get('bbox')
        if not box:
            if line.get('text'):
                items.append((0, 0, str(line['text'])))
            continue
        ys = [point[1] for point in box]
        xs = [point[0] for point in box]
        items.append((min(ys), min(xs), str(line.get('text') or '')))
    items.sort(key=lambda item: (item[0], item[1]))
    rows, current, last_y = [], [], None
    for y, x, text in items:
        if last_y is None or abs(y - last_y) <= 12:
            current.append((x, text))
            last_y = y if last_y is None else (last_y + y) / 2
        else:
            rows.append([text for _, text in sorted(current, key=lambda item: item[0])])
            current, last_y = [(x, text)], y
    if current:
        rows.append([text for _, text in sorted(current, key=lambda item: item[0])])
    return rows


def _guess_table_from_lines(rows):
    tables, block = [], []
    for row in rows:
        cells = [cell.strip() for cell in row]
        if len(cells) >= 3:
            block.append(cells)
            continue
        if len(block) >= 3:
            tables.append(block)
        block = []
    if len(block) >= 3:
        tables.append(block)
    return tables


def read_document(path, filename=''):
    suffix = Path(filename or path).suffix.lower()
    if suffix == '.pdf':
        return _read_pdf(path)
    if suffix == '.docx':
        return _read_docx(path)
    if suffix == '.doc':
        return _read_word_binary(path)
    if suffix in ('.xlsx', '.xls'):
        return _read_excel(path, suffix)
    if suffix == '.csv':
        return _read_csv(path)
    return '', [], 'unsupported'


# ---------------------------------------------------------------------------
# Normalization
# ---------------------------------------------------------------------------

_UNIT_ALIASES = {
    '°c': '°c', 'c': '°c', 'celcius': '°c', 'celsius': '°c',
    '%rh': '%rh', 'rh': '%rh', 'relative humidity': '%rh',
    'v': 'v', 'volt': 'v', 'volts': 'v',
    'a': 'a', 'amp': 'a', 'amps': 'a', 'ampere': 'a',
    'hz': 'hz', 'hertz': 'hz',
    'rpm': 'rpm', 'r.p.m.': 'rpm',
    'bar': 'bar', 'psi': 'psi', 'pa': 'pa', 'kpa': 'kpa', 'mpa': 'mpa',
    '°f': '°f', 'f': '°f',
    'mm': 'mm', 'cm': 'cm', 'm': 'm', 'inch': 'inch', 'in': 'inch',
}


def _normalize_unit(unit):
    if not unit:
        return ''
    u = unit.strip().lower()
    u = re.sub(r'[^\w°%/]', '', u)
    return _UNIT_ALIASES.get(u, u)


def _normalize_text(text):
    """Normalize text for matching: lowercase, collapse whitespace, strip punctuation variants."""
    if not text:
        return ''
    t = text.lower()
    t = re.sub(r'\s+', ' ', t)
    t = re.sub(r'[._\-–—/\\]', ' ', t)
    t = re.sub(r'[()\[\]{}]', ' ', t)
    t = re.sub(r'[,:;]', ' ', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t


def _split_label_value(text):
    """Try to split a text block into (label, value) using common separators."""
    patterns = [
        r'^(.+?)\s*[:=]\s*(.+)$',
        r'^(.+?)\s*[-–—]\s*(.+)$',
        r'^(.+?)\s*\|\s*(.+)$',
    ]
    for pat in patterns:
        m = re.match(pat, text.strip())
        if m:
            return m.group(1).strip(), m.group(2).strip()
    m = re.match(r'^([A-Za-z][A-Za-z0-9 ]{2,30}?)\s+([0-9].+)$', text.strip())
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return text.strip(), ''


# ---------------------------------------------------------------------------
# Template field discovery
# ---------------------------------------------------------------------------

def discover_template_fields(template):
    """Read an existing EquipmentCertificateTemplate and discover all required
    fields from its stored activities and columns, plus section titles from
    the original worksheet.

    Returns a dict:
        {
            'activities': [{'label': str, 'slug': str, 'section': str|None}],
            'columns': [{'name': str, 'section': str|None}],
            'sections': [{'title': str, 'row_index': int}],
            'all_labels': [str],
        }
    """
    activities = []
    for label in (template.activities or []):
        activities.append({
            'label': label,
            'slug': re.sub(r'[^a-z0-9]+', '_', label.lower()).strip('_') or 'field',
            'section': None,
        })

    columns = []
    for col in (template.columns or []):
        columns.append({
            'name': col,
            'slug': re.sub(r'[^a-z0-9]+', '_', col.lower()).strip('_') or 'field',
            'section': None,
        })

    sections = []
    try:
        if template.original_file:
            with template.original_file.open('rb') as handle:
                ws = load_workbook(io.BytesIO(handle.read()), data_only=True)
                name = ws.sheetnames[0]
                sheet = ws[name]
                for r in range(1, (sheet.max_row or 1) + 1):
                    row = [str(c.value or '').strip() for c in sheet[r]]
                    filled = [c for c in row if c]
                    if len(filled) == 1:
                        text = filled[0]
                        if _looks_like_section_title(text):
                            sections.append({'title': text, 'row_index': r - 1})
    except Exception:
        pass

    all_labels = [a['label'] for a in activities] + [c['name'] for c in columns]

    return {
        'activities': activities,
        'columns': columns,
        'sections': sections,
        'all_labels': all_labels,
    }


def _looks_like_section_title(text):
    if not text or len(text) < 4 or len(text) > 60:
        return False
    letters = [ch for ch in text if ch.isalpha()]
    if len(letters) < 4 or len(text.split()) > 8 or not text[0].isalpha():
        return False
    caps = sum(1 for ch in letters if ch.isupper())
    return caps / len(letters) >= 0.6


# ---------------------------------------------------------------------------
# Content block extraction from certificate text
# ---------------------------------------------------------------------------

def _extract_content_blocks(text, tables):
    """Convert raw document text and tables into normalized content blocks."""
    blocks = []
    lines = text.split('\n') if text else []

    # Also detect pipe-separated tables from plain text
    text_tables = _tables_from_text(text)
    if text_tables:
        tables = list(tables) + text_tables

    for idx, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue
        label, value = _split_label_value(line)
        blocks.append({
            'type': 'line',
            'text': line,
            'norm': _normalize_text(line),
            'label': label,
            'value': value,
            'norm_label': _normalize_text(label),
            'norm_value': _normalize_text(value),
            'row': idx,
            'col': None,
            'source': 'text',
            'confidence': 1.0 if value else 0.5,
        })

    for t_idx, table in enumerate(tables or []):
        for r_idx, row in enumerate(table):
            for c_idx, cell in enumerate(row):
                cell = (cell or '').strip()
                if not cell:
                    continue
                # Table cells are standalone values, not label+value pairs
                blocks.append({
                    'type': 'table_cell',
                    'text': cell,
                    'norm': _normalize_text(cell),
                    'label': '',
                    'value': cell,
                    'norm_label': '',
                    'norm_value': _normalize_text(cell),
                    'row': r_idx,
                    'col': c_idx,
                    'source': f'table_{t_idx}',
                    'confidence': 0.9 if cell else 0.4,
                })

    return blocks


# ---------------------------------------------------------------------------
# Synonym generation from template field names
# ---------------------------------------------------------------------------

def _generate_synonyms(label):
    """Generate alternative spellings and abbreviations for a template label."""
    synonyms = set()
    label_lower = label.lower().strip()
    synonyms.add(label_lower)
    synonyms.add(_normalize_text(label))

    words = label_lower.split()
    if len(words) > 1:
        core = ' '.join(w for w in words if w not in ('of', 'the', 'and', 'for', 'in', 'on', 'at', 'to', 'a', 'an'))
        if core:
            synonyms.add(core)
            synonyms.add(_normalize_text(core))

    abbr_map = {
        'number': ['no', 'no.', '#', 'num'],
        'serial': ['s/n', 'sn', 'serno'],
        'tag': ['tg'],
        'certificate': ['cert', 'cert.'],
        'temperature': ['temp', 'temp.'],
        'pressure': ['press', 'press.'],
        'voltage': ['volt', 'v'],
        'current': ['curr', 'i'],
        'frequency': ['freq', 'f'],
        'power': ['pwr', 'p'],
        'speed': ['spd'],
        'rpm': ['r.p.m.', 'rev/min'],
        'humidity': ['rh', 'hum'],
        'description': ['desc', 'desc.'],
        'manufacturer': ['mfg', 'mfr', 'make'],
        'model': ['mod', 'model no'],
        'owner': ['owner', 'customer', 'client'],
        'tested': ['test', 'calibrated'],
        'date': ['dt'],
        'environment': ['env', 'conditions', 'ambient'],
        'range': ['range', 'span'],
        'size': ['size'],
    }
    for word, abbrs in abbr_map.items():
        if word in label_lower:
            for abbr in abbrs:
                variant = label_lower.replace(word, abbr)
                synonyms.add(variant)
                synonyms.add(_normalize_text(variant))

    return {s for s in synonyms if s}


# ---------------------------------------------------------------------------
# Matching engine
# ---------------------------------------------------------------------------

def _score_exact_match(template_norm, block_norm):
    return 1.0 if template_norm and template_norm == block_norm else 0.0


def _score_contains(template_norm, block_norm):
    if not template_norm or not block_norm:
        return 0.0
    if template_norm in block_norm or block_norm in template_norm:
        return 0.85
    return 0.0


def _score_fuzzy(template_norm, block_norm):
    if not template_norm or not block_norm:
        return 0.0
    return difflib.SequenceMatcher(None, template_norm, block_norm).ratio()


def _score_regex(template_label, block_text):
    if not template_label or not block_text:
        return 0.0
    words = re.findall(r'\w+', template_label.lower())
    if not words:
        return 0.0
    pattern = r'[\s:.\-–/=()|.,;\n]*?'.join(re.escape(w) for w in words)
    if re.search(pattern, block_text, flags=re.IGNORECASE):
        return 0.8
    return 0.0


def _validate_data_type(value, field_hint):
    if not value:
        return 0.0
    hint = field_hint.lower()
    if any(w in hint for w in ('date', 'due', 'test date', 'calibration date')):
        for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%d-%m-%Y', '%Y/%m/%d'):
            try:
                datetime.strptime(value.strip(), fmt)
                return 1.0
            except ValueError:
                continue
        return 0.2
    if any(w in hint for w in ('voltage', 'current', 'frequency', 'power', 'speed', 'rpm',
                                'range', 'size', 'pressure', 'temperature', 'humidity')):
        if re.search(r'[-+]?\d*\.?\d+', value):
            return 1.0
        return 0.2
    return 0.9


def _validate_unit(value, field_hint):
    hint = field_hint.lower()
    unit_hints = {
        'voltage': ['v', 'volt', 'mv', 'kv'],
        'current': ['a', 'amp', 'ma'],
        'frequency': ['hz', 'khz', 'mhz'],
        'power': ['w', 'kw', 'mw'],
        'speed': ['rpm', 'r.p.m.'],
        'pressure': ['bar', 'psi', 'pa', 'kpa', 'mpa'],
        'temperature': ['°c', '°f', 'c', 'f'],
        'humidity': ['%rh', 'rh', '%'],
    }
    for key, units in unit_hints.items():
        if key in hint:
            for unit in units:
                if unit.lower() in value.lower():
                    return 1.0
            if re.search(r'\d', value):
                return 0.6
            return 0.0
    return 1.0


def _proximity_score(template_label, block, all_blocks, block_index):
    if not template_label or block_index < 0:
        return 0.0
    label_words = set(_normalize_text(template_label).split())
    nearby_text = []
    for i in range(max(0, block_index - 2), min(len(all_blocks), block_index + 3)):
        if i != block_index:
            nearby_text.append(all_blocks[i].get('norm', ''))
    nearby = ' '.join(nearby_text)
    if not nearby:
        return 0.5
    matches = sum(1 for w in label_words if w and w in nearby)
    return min(1.0, 0.3 + 0.1 * matches)


def match_template_fields(template_fields, content_blocks):
    """Match each template activity field against content blocks.

    Template table columns are handled separately by match_template_columns().
    Only activity fields (with 'label' key) are matched here.
    """
    results = []

    for field in template_fields['activities']:
        field_label = field['label']
        field_hint = field_label.lower()
        synonyms = _generate_synonyms(field_label)
        template_norm = _normalize_text(field_label)

        candidates = []

        for bi, block in enumerate(content_blocks):
            # Only match text lines as field labels. Table cells are handled
            # separately by match_template_columns().
            if block.get('type') != 'line':
                continue

            block_norm_label = block.get('norm_label', '')
            block_norm = block.get('norm', '')
            block_value = block.get('value', '') or block.get('text', '')

            # --- Strategy scores (label-based only) ---
            s1 = _score_exact_match(template_norm, block_norm_label)
            s2 = _score_contains(template_norm, block_norm_label)
            s3 = _score_fuzzy(template_norm, block_norm_label)
            s4 = _score_regex(field_label, block_norm)
            s5 = 0.0
            for syn in synonyms:
                syn_norm = _normalize_text(syn)
                if syn_norm and (syn_norm == block_norm_label or
                                 syn_norm in block_norm or
                                 block_norm_label in syn_norm):
                    s5 = max(s5, 0.9)

            label_score = max(s1, s2, s3, s4, s5)

            # Only consider blocks with a matching label.
            # Value-only blocks (no label) are not candidates for field matching.
            if label_score == 0.0:
                continue

            # Datatype/unit validation only boosts, never creates, a match
            s6 = _validate_data_type(block_value, field_hint) * 0.15
            s7 = _validate_unit(block_value, field_hint) * 0.1
            s8 = _proximity_score(field_label, block, content_blocks, bi) * 0.1

            total = label_score * 0.75 + s6 + s7 + s8
            total = min(1.0, total)

            if total > 0.5:
                candidates.append({
                    'block_index': bi,
                    'block': block,
                    'value': block_value,
                    'score': total,
                    'strategies': {
                        'exact': s1,
                        'contains': s2,
                        'fuzzy': s3,
                        'regex': s4,
                        'synonym': s5,
                        'datatype': s6,
                        'unit': s7,
                        'proximity': s8,
                    },
                })

        candidates.sort(key=lambda c: c['score'], reverse=True)

        best = candidates[0] if candidates else None
        matched_value = ''
        confidence = 0.0
        source_block = None
        strategy = 'none'

        if best and best['score'] >= 0.5:
            matched_value = best['value']
            confidence = best['score']
            source_block = best['block']
            strat_scores = best['strategies']
            strategy = max(strat_scores, key=strat_scores.get)

        results.append({
            'field': field_label,
            'slug': field['slug'],
            'matched_value': matched_value,
            'confidence': confidence,
            'source_block': source_block,
            'strategy': strategy,
            'all_candidates': candidates[:5],
        })

    return results


def match_template_columns(template_columns, tables):
    """Match template table columns to detected table columns in the certificate."""
    if not template_columns:
        return [], 0.0

    table_results = []
    for t_idx, table in enumerate(tables or []):
        if not table or len(table) < 2:
            continue
        header = table[0]
        data_rows = table[1:]

        col_matches = []
        for t_col in template_columns:
            t_name = t_col['name']
            t_norm = _normalize_text(t_name)
            best_col_match = None
            best_score = 0.0
            for c_idx, cell in enumerate(header):
                cell_norm = _normalize_text(cell)
                s_exact = 1.0 if cell_norm == t_norm else 0.0
                s_contains = 0.85 if (cell_norm in t_norm or t_norm in cell_norm) and cell_norm else 0.0
                s_fuzzy = difflib.SequenceMatcher(None, t_norm, cell_norm).ratio()
                score = max(s_exact, s_contains, s_fuzzy)
                if score > best_score:
                    best_score = score
                    best_col_match = {'cell_index': c_idx, 'cell_value': cell, 'score': score}
            if best_col_match and best_col_match['score'] >= 0.5:
                col_matches.append({
                    'template_column': t_name,
                    'match': best_col_match,
                })

        if col_matches:
            matched_rows = []
            for row in data_rows:
                matched_row = {}
                for cm in col_matches:
                    ci = cm['match']['cell_index']
                    matched_row[cm['template_column']] = str(row[ci]).strip() if ci < len(row) else ''
                matched_rows.append(matched_row)

            table_score = sum(cm['match']['score'] for cm in col_matches) / len(col_matches)
            table_results.append({
                'table_index': t_idx,
                'column_matches': col_matches,
                'rows': matched_rows,
                'score': table_score,
            })

    table_results.sort(key=lambda t: t['score'], reverse=True)
    if table_results:
        return table_results[0]['rows'], table_results[0]['score']
    return [], 0.0


# ---------------------------------------------------------------------------
# Standard field definitions (always extracted, in addition to template labels)
# ---------------------------------------------------------------------------

_STANDARD_FIELDS = [
    ('serial_no', 'Serial/Tag Number'),
    ('model', 'Model Number'),
    ('certificate_no', 'Certificate Number'),
    ('test_date', 'Test Date'),
    ('conclusion', 'Conclusion'),
]


def _match_standard_fields(content_blocks, existing_fields=None):
    """Match standard fields (serial_no, model, certificate_no, test_date, conclusion)
    against content blocks. Only extracts fields not already covered by template matches.

    Returns (fields_dict, confidence_dict, matches_list).
    """
    existing_fields = existing_fields or {}
    fields = {}
    confidence = {}
    matches = []

    for slug, label in _STANDARD_FIELDS:
        # Skip if template already covers this field
        if slug in existing_fields:
            continue

        synonyms = _generate_synonyms(label)
        template_norm = _normalize_text(label)

        candidates = []
        for bi, block in enumerate(content_blocks):
            # Only match text lines as field labels.
            if block.get('type') != 'line':
                continue

            block_norm_label = block.get('norm_label', '')
            block_norm = block.get('norm', '')
            block_value = block.get('value', '') or block.get('text', '')

            s1 = _score_exact_match(template_norm, block_norm_label)
            s2 = _score_contains(template_norm, block_norm_label)
            s3 = _score_fuzzy(template_norm, block_norm_label)
            s4 = _score_regex(label, block_norm)
            s5 = 0.0
            for syn in synonyms:
                syn_norm = _normalize_text(syn)
                if syn_norm and (syn_norm == block_norm_label or
                                 syn_norm in block_norm or
                                 block_norm_label in syn_norm):
                    s5 = max(s5, 0.9)

            label_score = max(s1, s2, s3, s4, s5)
            if label_score == 0.0:
                continue

            s6 = _validate_data_type(block_value, label) * 0.15
            s7 = _validate_unit(block_value, label) * 0.1
            s8 = _proximity_score(label, block, content_blocks, bi) * 0.1

            total = label_score * 0.75 + s6 + s7 + s8
            total = min(1.0, total)

            if total > 0.5:
                candidates.append({
                    'block_index': bi,
                    'block': block,
                    'value': block_value,
                    'score': total,
                })

        candidates.sort(key=lambda c: c['score'], reverse=True)

        best = candidates[0] if candidates else None
        matched_value = ''
        conf = 0.0
        if best and best['score'] >= 0.5:
            matched_value = best['value']
            conf = best['score']

        fields[slug] = matched_value
        confidence[slug] = conf
        matches.append({
            'field': label,
            'slug': slug,
            'matched_value': matched_value,
            'confidence': conf,
            'source_block': best['block'] if best else None,
            'strategy': 'standard',
        })

    return fields, confidence, matches


# ---------------------------------------------------------------------------
# Top-level extraction
# ---------------------------------------------------------------------------

def extract_certificate(path, filename, template):
    """Extract certificate data against a template.

    Args:
        path: Path to the certificate file
        filename: Original filename
        template: EquipmentCertificateTemplate instance

    Returns:
        dict with provider, fields, field_confidence, activities, columns,
        rows, missing, text_length, table_columns_matched, processing_time_ms,
        field_matches
    """
    start_time = time.monotonic()

    template_fields = discover_template_fields(template)
    text, tables, provider = read_document(path, filename)

    content_blocks = _extract_content_blocks(text, tables)

    # Combine tables from the reader with text-detected tables for column matching
    all_tables = list(tables) + _tables_from_text(text)

    # If DOCX, use cert_fields module for extraction
    suffix = Path(filename or path).suffix.lower()
    if suffix == '.docx':
        structure = docx_structure(path)
        result = extract(structure,
                         [f['label'] for f in template_fields['activities']],
                         [c['name'] for c in template_fields['columns']])

        # Build fields dict from template field matches
        fields = {}
        field_confidence = {}
        activities = {}
        all_matches = []

        for label, value in result['fields'].items():
            slug = re.sub(r'[^a-z0-9]+', '_', label.lower()).strip('_') or 'field'
            fields[slug] = value
            field_confidence[slug] = result['confidence'].get(label, 0.0)
            activities[slug] = value
            all_matches.append({
                'field': label,
                'slug': slug,
                'matched_value': value,
                'confidence': result['confidence'].get(label, 0.0),
                'source_block': None,
                'strategy': 'server_extract',
            })

        # Table rows
        columns = [c['name'] for c in template_fields['columns']]
        rows = []
        if result.get('table_rows'):
            for row in result['table_rows']:
                if isinstance(row, dict):
                    rows.append([str(row.get(col, '')) for col in columns])
                else:
                    rows.append([str(cell) for cell in row])

        # Missing fields
        missing = [match['field'] for match in all_matches if not match['matched_value']]
        if not rows and (template_fields['columns'] or tables):
            missing.append('table')

        elapsed_ms = int((time.monotonic() - start_time) * 1000)

        return {
            'provider': provider,
            'fields': fields,
            'field_confidence': field_confidence,
            'activities': activities,
            'columns': columns,
            'rows': rows,
            'missing': missing,
            'text_length': len(text or ''),
            'table_columns_matched': len(columns),
            'processing_time_ms': elapsed_ms,
            'field_matches': all_matches,
            'internal': result.get('internal', {}),
        }

    # Extract template activity fields only — the template defines what is required
    field_matches = match_template_fields(template_fields, content_blocks)

    # Build fields dict from template field matches
    fields = {}
    field_confidence = {}
    activities = {}
    all_matches = []

    for match in field_matches:
        slug = match['slug']
        value = match['matched_value']
        conf = match['confidence']
        fields[slug] = value
        field_confidence[slug] = conf
        all_matches.append(match)
        if value:
            activities[slug] = value

    # Apply fixed values from template (ticked "Keep this value on every certificate")
    fixed_values = template.fixed_values or {}
    for match in field_matches:
        slug = match['slug']
        if slug in fixed_values and fixed_values[slug]:
            fields[slug] = fixed_values[slug]
            field_confidence[slug] = 1.0
            match['matched_value'] = fixed_values[slug]
            match['confidence'] = 1.0
            match['strategy'] = 'fixed'
            activities[slug] = fixed_values[slug]

        # Override with internal values from cert_fields.extract for standard fields
        internal = result.get('internal', {})
        for slug in ['serial_no', 'test_date', 'due_date', 'certificate_no']:
            if internal.get(slug):
                fields[slug] = internal[slug]
                field_confidence[slug] = 1.0
                activities[slug] = internal[slug]

    # Also extract the 5 standard fields if not already covered by template fields
    std_fields, std_confidence, std_matches = _match_standard_fields(content_blocks, fields)
    for slug, value in std_fields.items():
        if slug not in fields:
            fields[slug] = value
            field_confidence[slug] = std_confidence.get(slug, 0.0)
            activities[slug] = value
    all_matches.extend(std_matches)

    matched_rows, table_score = match_template_columns(
        template_fields['columns'], all_tables
    )

    columns = [c['name'] for c in template_fields['columns']]
    rows = matched_rows

    # Merge column data into fields/activities from the first data row
    if rows and columns:
        first_row = rows[0]
        for col in template_fields['columns']:
            col_name = col['name']
            col_slug = col['slug']
            col_value = first_row.get(col_name, '')
            fields[col_slug] = col_value
            field_confidence[col_slug] = 0.9 if col_value else 0.0
            activities[col_slug] = col_value
            all_matches.append({
                'field': col_name,
                'slug': col_slug,
                'matched_value': col_value,
                'confidence': 0.9 if col_value else 0.0,
                'source_block': None,
                'strategy': 'table_column',
            })

    # Only template-defined fields can be missing — no hardcoded requirements
    missing = []
    for match in field_matches:
        if 'label' in match and not match['matched_value']:
            missing.append(match['field'])
    if not rows and (template_fields['columns'] or tables):
        missing.append('table')

    elapsed_ms = int((time.monotonic() - start_time) * 1000)

    return {
        'provider': provider,
        'fields': fields,
        'field_confidence': field_confidence,
        'activities': activities,
        'columns': columns,
        'rows': rows,
        'missing': missing,
        'text_length': len(text or ''),
        'table_columns_matched': len(columns),
        'processing_time_ms': elapsed_ms,
        'field_matches': all_matches,
        'internal': result.get('internal', {}),
    }


# ---------------------------------------------------------------------------
# Template population (generate completed document preserving template layout)
# ---------------------------------------------------------------------------

def _resolve_merge_anchor(ws, row, col):
    """If (row, col) is inside a merged range, return the anchor (top-left)
    cell's (row, col). Otherwise return (row, col) unchanged."""
    try:
        merged = ws.merged_cells.ranges
    except Exception:
        return row, col
    for rng in merged:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return rng.min_row, rng.min_col
    return row, col


def populate_template_workbook(template, field_matches):
    """Generate a completed document based on the existing template structure.

    Preserves the original template layout and fills in matched values into
    adjacent cells. Unmatched fields remain in their original positions.

    Returns (Workbook, output_filename, filled_count).
    """
    import io

    wb = Workbook()
    wb.remove(wb.active)

    try:
        if template.original_file:
            with template.original_file.open('rb') as handle:
                wb = load_workbook(io.BytesIO(handle.read()))
    except Exception:
        pass

    if not wb.sheetnames:
        ws = wb.active
        ws.title = "Certificate"
    else:
        ws = wb[wb.sheetnames[0]]

    # Build a lookup: slug -> matched value
    value_by_slug = {m['slug']: m['matched_value'] for m in field_matches if m['matched_value']}
    confidence_by_slug = {m['slug']: m['confidence'] for m in field_matches}

    filled_count = 0
    # Track anchor cells already written to, so two fields don't overwrite
    # each other when they resolve to the same merged cell.
    written_anchors = set()

    # Scan the template sheet for label cells and fill adjacent value cells
    max_row = ws.max_row or 1
    max_col = ws.max_column or 1

    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell_val = str(cell.value or '').strip()
            if not cell_val:
                continue
            cell_norm = _normalize_text(cell_val)
            if not cell_norm:
                continue

            # Try to match this cell label against template fields
            best_match = None
            best_score = 0.0
            for match in field_matches:
                field_norm = _normalize_text(match['field'])
                if not field_norm:
                    continue
                # Exact or contains match
                if cell_norm == field_norm:
                    score = 1.0
                elif cell_norm in field_norm or field_norm in cell_norm:
                    score = 0.85
                else:
                    score = difflib.SequenceMatcher(None, cell_norm, field_norm).ratio()
                if score > best_score and score >= 0.7:
                    best_score = score
                    best_match = match

            if best_match and best_match['matched_value']:
                # Write the value into the next non-empty cell to the right,
                # or the cell below if no right cell is available
                target_row, target_col = None, None
                for cc in range(c + 1, min(c + 5, max_col + 1)):
                    neighbor = ws.cell(row=r, column=cc)
                    if not neighbor.value or str(neighbor.value or '').strip() == '':
                        target_row, target_col = r, cc
                        break
                if target_col is None:
                    # Try cell below
                    for rr in range(r + 1, min(r + 5, max_row + 1)):
                        neighbor = ws.cell(row=rr, column=c)
                        if not neighbor.value or str(neighbor.value or '').strip() == '':
                            target_row, target_col = rr, c
                            break

                if target_row is not None and target_col is not None:
                    # Resolve to the merged-range anchor cell if needed, since
                    # only the top-left cell of a merged range is writable.
                    anchor_row, anchor_col = _resolve_merge_anchor(ws, target_row, target_col)
                    anchor_key = (anchor_row, anchor_col)
                    if anchor_key in written_anchors:
                        continue  # already filled; skip to avoid overwriting
                    written_anchors.add(anchor_key)
                    ws.cell(row=anchor_row, column=anchor_col).value = best_match['matched_value']
                    filled_count += 1

    output_filename = f"{template.name.replace(' ', '_')}_completed.xlsx"
    return wb, output_filename, filled_count