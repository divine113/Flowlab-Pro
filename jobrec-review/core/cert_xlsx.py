"""
cert_xlsx.py - draws certificate-log records as Excel sheets in the
Engineering Automation house style (Arial 11, medium black borders, logo top left).

    records = [{
        "name":    "DM-003",                              # sheet name (file name without extension)
        "fields":  [("Manufacturer", "Dewalt"), ...],     # label, value  (template order)
        "columns": ["Applied voltage (V)", ...],          # test-data headers
        "rows":    [["220-240", "50/60", "3.5", ...]],    # test-data rows
    }]
    wb = build_workbook(records, logo_path)               # openpyxl Workbook, one sheet per record
    data = workbook_bytes(records, logo_path)             # same, as bytes for an HTTP response

Nothing is read from the template .xlsx, so every equipment looks the same.
"""
import io
import math
import re

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

_SIDE = Side(style="medium", color="000000")
_BORDER = Border(left=_SIDE, right=_SIDE, top=_SIDE, bottom=_SIDE)
_BOLD = Font(name="Arial", size=11, bold=True)
_NORMAL = Font(name="Arial", size=11)
_LINE = 15.0          # row height of one line, in points
_LABEL_MIN = 24.6     # width of the label column (B)
_COL_MIN = 12.0
_COL_MAX = 30.0


def _txt(v):
    return "" if v is None else str(v).strip()


def _lines(text, width, bold=False):
    per_line = max(1, int(width * (0.9 if bold else 1.0)))
    return max(1, sum(max(1, math.ceil(len(p) / per_line)) for p in _txt(text).split("\n")))


def _put(ws, row, col, value, font, horizontal, wrap=True):
    c = ws.cell(row=row, column=col)
    text = _txt(value)
    c.value = text
    if text.startswith("="):          # never let certificate text become a formula
        c.data_type = "s"
    c.font = font
    c.alignment = Alignment(horizontal=horizontal, vertical="center", wrap_text=wrap)
    c.border = _BORDER
    return c


def _box(ws, row, c1, c2):
    """Borders on every cell of a merged range, so the outline is complete."""
    for col in range(c1, c2 + 1):
        ws.cell(row=row, column=col).border = _BORDER


def _sheet_name(name, used):
    base = re.sub(r"[\[\]:*?/\\]", "-", _txt(name) or "Record")[:31].strip("'") or "Record"
    cand, n = base, 2
    while cand.lower() in used:
        suffix = f" ({n})"
        cand = base[: 31 - len(suffix)] + suffix
        n += 1
    used.add(cand.lower())
    return cand


def write_record_sheet(ws, fields, columns, rows, logo_path=None,
                       title="LAB GENERAL ACTIVITIES LOG", test_title="TEST DATA"):
    fields = [(_txt(a), _txt(b)) for a, b in fields]
    columns = [_txt(c) for c in columns]
    rows = [[_txt(x) for x in r] for r in rows]
    has_table = bool(columns)
    ncols = max(len(columns), 4)            # block spans B .. last
    first, last = 2, 1 + ncols

    # ---- column widths: label column fits labels and the first table column
    widths = {}
    for i in range(ncols):
        col = first + i
        w = _LABEL_MIN if i == 0 else _COL_MIN
        if i < len(columns):
            w = max(w, len(columns[i]) * 1.15 + 3)
            for r in rows:
                if i < len(r):
                    w = max(w, len(r[i]) * 1.05 + 2)
        if i == 0:
            for lab, _v in fields:
                w = max(w, len(lab) * 1.15 + 3)
        widths[col] = min(max(w, _LABEL_MIN if i == 0 else _COL_MIN), _COL_MAX + (10 if i == 0 else 0))
        ws.column_dimensions[get_column_letter(col)].width = round(widths[col], 1)
    value_width = sum(widths[c] for c in range(first + 1, last + 1))

    # ---- logo (rows 1-3)
    if logo_path:
        try:
            img = XLImage(logo_path)
            img.width, img.height = 262, 56
            ws.add_image(img, f"{get_column_letter(first)}1")
        except Exception:
            pass

    # ---- title bar (row 5)
    r = 5
    ws.merge_cells(start_row=r, start_column=first, end_row=r, end_column=last)
    _put(ws, r, first, title, _BOLD, "center")
    _box(ws, r, first, last)
    ws.row_dimensions[r].height = _LINE

    # ---- label / value rows
    for lab, val in fields:
        r += 1
        _put(ws, r, first, lab, _BOLD, None)
        ws.merge_cells(start_row=r, start_column=first + 1, end_row=r, end_column=last)
        _put(ws, r, first + 1, val, _NORMAL, "left")
        _box(ws, r, first + 1, last)
        n = max(_lines(lab, widths[first], True), _lines(val, value_width))
        ws.row_dimensions[r].height = _LINE * n

    # ---- test data
    if has_table:
        r += 4                                   # three blank rows
        ws.merge_cells(start_row=r, start_column=first, end_row=r, end_column=last)
        _put(ws, r, first, test_title, _BOLD, "center")
        _box(ws, r, first, last)
        ws.row_dimensions[r].height = _LINE

        r += 1
        n = 1
        for i, h in enumerate(columns):
            _put(ws, r, first + i, h, _BOLD, None if i == 0 else "center")
            n = max(n, _lines(h, widths[first + i], True))
        ws.row_dimensions[r].height = _LINE * n

        for data in (rows or [[""] * len(columns)]):
            r += 1
            n = 1
            for i in range(len(columns)):
                v = data[i] if i < len(data) else ""
                _put(ws, r, first + i, v, _NORMAL, None if i == 0 else "center")
                n = max(n, _lines(v, widths[first + i]))
            ws.row_dimensions[r].height = _LINE * n

    # ---- print: one page wide, centred
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.print_area = f"A1:{get_column_letter(last + 1)}{r + 1}"


def build_workbook(records, logo_path=None):
    wb = Workbook()
    wb.remove(wb.active)
    used = set()
    for rec in records or [{"name": "Record", "fields": [], "columns": [], "rows": []}]:
        ws = wb.create_sheet(_sheet_name(rec.get("name"), used))
        write_record_sheet(ws, rec.get("fields", []), rec.get("columns", []),
                           rec.get("rows", []), logo_path)
    return wb


def workbook_bytes(records, logo_path=None):
    buf = io.BytesIO()
    build_workbook(records, logo_path).save(buf)
    return buf.getvalue()
