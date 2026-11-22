"""
cert_fields.py - template-driven extraction from a certificate.

The template supplies two lists:  labels  (e.g. "Manufacturer") and  columns
(e.g. "Applied voltage (V)").  Nothing else is hard-coded.

  structure = docx_structure("DM-003.docx")        # DOCX reader
  result    = extract(structure, labels, columns)  # works on any structure

A "structure" is just:  {"tables": [[[cell, cell, ...], ...], ...], "lines": [str, ...]}
so PDF / XLSX / CSV readers only need to produce the same shape.
"""
import re
import difflib
from docx import Document
from docx.oxml.ns import qn

_DEG = {"0", "o", "O", "\u00ba", "\u00b0"}


# ------------------------------------------------------------------ DOCX reader
def _para_text(p):
    """Text of one w:p. Superscript 0/o becomes the degree sign (31<sup>0</sup>C -> 31°C)."""
    out = []
    for r in p.xpath("./w:r | ./w:hyperlink/w:r"):
        sup = False
        rpr = r.find(qn("w:rPr"))
        if rpr is not None:
            va = rpr.find(qn("w:vertAlign"))
            sup = va is not None and va.get(qn("w:val")) == "superscript"
        txt = ""
        for ch in r:
            if ch.tag == qn("w:t"):
                txt += ch.text or ""
            elif ch.tag == qn("w:tab"):
                txt += "\t"
            elif ch.tag == qn("w:br"):
                txt += "\n"
        if sup and txt.strip() in _DEG:
            txt = "\u00b0"
        out.append(txt)
    return "".join(out)


def _collect(root, tables, lines, seen):
    """Text boxes first, then body paragraphs (lines) and tables, in document order."""
    for p in root.xpath(".//w:txbxContent//w:p"):
        t = _para_text(p).strip()
        if t and t not in seen:
            seen.add(t)
            lines.extend(x for x in t.split("\n") if x.strip())
    for el in root:
        if el.tag == qn("w:p"):
            lines.extend(x for x in _para_text(el).split("\n") if x.strip())
        elif el.tag == qn("w:tbl"):
            rows = []
            for tr in el.xpath("./w:tr"):
                rows.append(["\n".join(_para_text(p) for p in tc.xpath("./w:p")).strip()
                             for tc in tr.xpath("./w:tc")])
            tables.append(rows)


def docx_structure(path):
    doc = Document(path)
    tables, lines, seen = [], [], set()
    # headers first: certificate number / date of issue often sit in a header text box
    done = set()
    for sec in doc.sections:
        for h in (sec.header, sec.first_page_header, sec.even_page_header):
            el = h._element
            if id(el) not in done:
                done.add(id(el))
                _collect(el, tables, lines, seen)
    _collect(doc.element.body, tables, lines, seen)
    return {"tables": tables, "lines": lines}


# ------------------------------------------------------------------ matching
_SAME = {"no": "number", "num": "number", "nr": "number", "nos": "number"}


def _norm(s):
    toks = re.sub(r"[^a-z0-9]+", " ", (s or "").lower()).split()
    return " ".join(_SAME.get(t, t) for t in toks)


def _score(a, b):
    na, nb = _norm(a), _norm(b)
    if not na or not nb:
        return 0.0
    if na == nb:
        return 1.0
    ta, tb = set(na.split()), set(nb.split())
    contain = 0.88 if (ta <= tb or tb <= ta) else 0.0
    ratio = difflib.SequenceMatcher(None, na, nb).ratio()
    return max(contain, ratio if ratio >= 0.8 else 0.0)


def _clean(v):
    return re.sub(r"[ \t]+", " ", (v or "")).strip()


def _col_name(col):  # "Applied voltage (V)" -> "Applied voltage"
    return re.sub(r"\(.*?\)", "", col).strip()


def _col_unit(col):
    m = re.search(r"\(([^)]*)\)", col)
    return m.group(1).strip() if m else ""


# ------------------------------------------------------------------ label/value pairs
def _label_lines(cell, labels):
    """If the cell is a label cell return its label(s), else None."""
    ls = [l.strip() for l in cell.split("\n") if l.strip()]
    if not ls:
        return None
    if all(l.endswith(":") for l in ls):
        return [l.rstrip(":").strip() for l in ls]
    if len(ls) == 1 and any(_score(ls[0], t) >= 0.85 for t in labels):
        return [ls[0]]
    return None


def _spread(labels_, value_cell):
    """Give each stacked label its own line of the value cell (last one gets the rest)."""
    vl = [l.strip() for l in value_cell.split("\n") if l.strip()]
    if len(labels_) == 1:
        return [_clean(" ".join(vl))]
    if len(vl) >= len(labels_):
        head = vl[: len(labels_) - 1]
        tail = " ".join(vl[len(labels_) - 1:])
        return [_clean(x) for x in head + [tail]]
    return [_clean(x) for x in vl] + [""] * (len(labels_) - len(vl))


def _table_pairs(tables, labels):
    pairs = []
    for rows in tables:
        for ri, row in enumerate(rows):
            i = 0
            while i < len(row):
                c = row[i]
                labs = _label_lines(c, labels) if c.strip() else None
                if labs is None:
                    i += 1
                    continue
                j = i + 1
                while j < len(row) and not row[j].strip():
                    j += 1
                vcell = row[j] if j < len(row) else ""
                has_colon = c.strip().endswith(":")
                below = rows[ri + 1][i] if ri + 1 < len(rows) and i < len(rows[ri + 1]) else ""
                # header-style block:  "Tested By" | "Authorized By"  over  "Name:" | "Name:"
                if (not has_colon and vcell and not vcell.strip().endswith(":")
                        and below.strip().endswith(":")):
                    v = re.sub(r"^[^:]{1,30}:\s*", "", below.strip())
                    pairs.append((labs[0], _clean("" if v.strip("_ ") == "" else v)))
                    i += 1
                    continue
                if vcell and _label_lines(vcell, labels) is not None:
                    vals = [""] * len(labs)          # next cell is another label -> no value here
                    j = i
                else:
                    vals = _spread(labs, vcell)
                for lab, v in zip(labs, vals):
                    pairs.append((lab, v))
                i = j + 1
    return pairs


def _line_pairs(lines):
    pairs = []
    for ln in lines:
        m = re.match(r"^\s*([^:=\t]{1,40}?)\s*[:=]\s*(.+)$", ln)
        if m:
            pairs.append((m.group(1).strip(), _clean(m.group(2))))
    return pairs


# ------------------------------------------------------------------ test data
def _table_rows(tables, columns):
    names = [_col_name(c) for c in columns]
    need = max(2, (len(columns) + 1) // 2)
    for rows in tables:
        for hi, hrow in enumerate(rows):
            m = {}
            for ci, h in enumerate(hrow):
                for k, n in enumerate(names):
                    if k not in m.values() and _score(h.replace("\n", " "), n) >= 0.8:
                        m[ci] = k
                        break
            if len(m) >= need:
                out = []
                for r in rows[hi + 1:]:
                    if not any(x.strip() for x in r):
                        break
                    row = [""] * len(columns)
                    for ci, k in m.items():
                        if ci < len(r):
                            row[k] = _clean(r[ci])
                    out.append(row)
                return out
    return []


def _strip_unit(v, unit):
    if unit:
        v = re.sub(r"(?<=\d)\s*%s\b" % re.escape(unit), "", v, flags=re.I)
    return _clean(v)


def _line_rows(lines, columns):
    """'Applied voltage = 220V-240V' style lines -> rows."""
    names = [_col_name(c) for c in columns]
    rows, cur = [], {}
    for ln in lines:
        m = re.match(r"^\s*(.{1,40}?)\s*=\s*(.+)$", ln)
        if not m:
            continue
        key = max(range(len(names)), key=lambda k: _score(m.group(1), names[k]))
        if _score(m.group(1), names[key]) < 0.8:
            continue
        if key in cur:
            rows.append(cur)
            cur = {}
        cur[key] = _strip_unit(m.group(2), _col_unit(columns[key]))
    if cur:
        rows.append(cur)
    return [[r.get(k, "") for k in range(len(columns))] for r in rows]


# ------------------------------------------------------------------ main
def extract(structure, labels, columns):
    pairs = _table_pairs(structure["tables"], labels) + _line_pairs(structure["lines"])
    cands = []
    for li, L in enumerate(labels):
        for pi, (pl, _v) in enumerate(pairs):
            s = _score(L, pl)
            if s >= 0.8:
                cands.append((-s, pi, li))
    cands.sort()
    used_l, used_p, fields, conf = set(), set(), {}, {}
    for ns, pi, li in cands:
        if li in used_l or pi in used_p:
            continue
        used_l.add(li)
        used_p.add(pi)
        v = pairs[pi][1]
        if v.endswith(":") or any(_score(v, t) >= 0.95 for t in labels):
            v = ""                                   # a label is never a value
        fields[labels[li]] = v
        conf[labels[li]] = round(-ns, 2)
    for L in labels:
        fields.setdefault(L, "")
        conf.setdefault(L, 0.0)

    rows = _table_rows(structure["tables"], columns) or _line_rows(structure["lines"], columns)

    def find(*keys):
        for pl, v in pairs:
            n = _norm(pl)
            if v and all(k in n for k in keys):
                return v
        return ""

    internal = {
        "serial_no": find("serial") or find("tag"),
        "test_date": find("test", "date") or find("date", "calibration"),
        "due_date": find("due", "date"),
        "certificate_no": find("certificate", "number"),
    }
    missing = [L for L in labels if not fields[L]]
    if not rows:
        missing.append("table")
    return {"fields": fields, "confidence": conf, "table_rows": rows,
            "internal": internal, "missing": missing}
