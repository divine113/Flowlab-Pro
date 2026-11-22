"""Render calibration-certificate sheets from extracted data only.

Every surface that shows a certificate shows the sheet itself — the logo, the
border grid, the LAB GENERAL ACTIVITIES LOG block, the TEST DATA table — never
plain form cards. Values from the certificate (or a saved record) are written
into the same cells the sheet prints.

Public API:
    render_sheet(ws=None, editable=False, values=None, confidence=None) -> html string
    values = {
        'activities': {slug: value},   # prefill field inputs (act_<slug>)
        'rows': [[cell, …], …],        # prefill TEST DATA rows, column order
        'columns': [col1, col2, …],     # test-data table column names
    }
"""
from __future__ import annotations

import base64
import html
import re


def _mark_safe(value):
    try:
        from django.utils.safestring import mark_safe
        return mark_safe(value)
    except Exception:
        return value


def label_slug(label):
    slug = re.sub(r'[^a-z0-9]+', '_', str(label).lower()).strip('_')
    return slug or 'field'


def load_sheet(file_or_path, sheet_name=None):
    from openpyxl import load_workbook
    workbook = load_workbook(file_or_path, data_only=True)
    name = sheet_name if sheet_name in workbook.sheetnames else workbook.sheetnames[0]
    return workbook[name]


def _text(value):
    return re.sub(r'\s+', ' ', str(value if value is not None else '')).strip()


# ----------------------------------------------------------------- image

def _image_uri(ws):
    try:
        images = list(getattr(ws, '_images', []) or [])
    except Exception:
        images = []
    if not images:
        return None
    try:
        payload = images[0]._data()
    except Exception:
        return None
    return 'data:image/png;base64,%s' % base64.b64encode(payload).decode('ascii')


# ----------------------------------------------------------------- render

def render_sheet(ws=None, editable=False, values=None, confidence=None):
    """Render a calibration-certificate sheet from extracted data only.

    Builds an HTML table from template structure (activities labels, columns)
    and fills values from field_values and table rows. Does NOT read cell
    values from the template XLSX — the template XLSX is used only for
    logo/image reference.

    Args:
        ws: Worksheet object (optional, for logo image only)
        editable: Whether inputs should be editable
        values: dict with 'activities', 'rows', 'columns' keys
        confidence: dict with slug -> confidence score

    Returns HTML string.
    """
    values = values or {}
    activity_values = values.get('activities') or {}
    data_rows = values.get('rows') or []
    headers = values.get('columns') or []
    confidence = confidence or {}

    parts = ['<div class="cert-sheet-wrap">']

    # Logo (from template, if available)
    image_uri = _image_uri(ws) if ws else None
    if image_uri:
        parts.append('<div class="cert-logo"><img src="%s" alt="Company logo"></div>' % image_uri)

    parts.append('<table class="cert-sheet">')
    parts.append('<tbody>')

    # Section title: LAB GENERAL ACTIVITIES LOG
    parts.append('<tr><th class="sv-title" colspan="2">LAB GENERAL ACTIVITIES LOG</th></tr>')

    # Activity fields: label + input
    for slug, value in activity_values.items():
        label = slug.replace('_', ' ').title()
        conf_badge = _conf_badge(confidence.get(slug, 1.0))
        cells_html = (
            '<tr>'
            '<th class="sv-label">%s%s</th>'
            '<td class="sv-input"><input class="qinput" name="act_%s" value="%s"></td>'
            '</tr>'
        ) % (
            html.escape(label),
            conf_badge,
            html.escape(slug),
            html.escape(str(value or ''))
        )
        parts.append(cells_html)

    # TEST DATA section
    if headers:
        parts.append('<tr><th class="sv-head" colspan="%d">TEST DATA</th></tr>' % len(headers))
        # Header row
        header_cells = ''.join(
            '<th class="sv-head">%s</th>' % html.escape(h) for h in headers
        )
        parts.append('<tr>%s</tr>' % header_cells)
        # Data rows
        for row_index, row in enumerate(data_rows):
            cells_html = []
            for c_idx, col_name in enumerate(headers):
                cell_value = ''
                if isinstance(row, dict):
                    cell_value = str(row.get(col_name, ''))
                elif isinstance(row, (list, tuple)) and c_idx < len(row):
                    cell_value = str(row[c_idx])
                cells_html.append(
                    '<td class="sv-input sv-td"><input class="qinput cell" name="cell_%d_%d" value="%s"></td>'
                    % (row_index, c_idx, html.escape(cell_value))
                )
            parts.append('<tr>%s</tr>' % ''.join(cells_html))

    parts.append('</tbody>')
    parts.append('</table></div>')
    return _mark_safe(''.join(parts))


def _conf_badge(score):
    """Small confidence badge for extracted fields. Low scores flag values the
    operator should double-check in Review."""
    if score is None or score >= 0.8:
        return ''
    if score >= 0.5:
        return ' <span class="el-conf el-conf-med" title="Extracted with moderate confidence — please verify">⚠</span>'
    return ' <span class="el-conf el-conf-low" title="Extracted with low confidence — please verify">⚠</span>'