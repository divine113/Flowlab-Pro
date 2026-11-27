"""Equipment Certificate Logs — Administrator-only certificate log.

RECORDS → Equipment Logs

No AI, no external service: text PDFs/DOCX/Excel are parsed locally.
Certificates are processed automatically against an existing template:
the template determines what fields are required, values are extracted
from the certificate using deterministic matching, and the completed
template is saved to the log without intermediate extraction screens.
"""
import functools
import json
import re
import uuid
from datetime import datetime
from pathlib import Path

from django.conf import settings
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.core.cache import cache
from django.core.exceptions import PermissionDenied
from django.http import FileResponse, Http404, JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.utils import timezone
from django.db.models import Count

from . import ei_workflow as service
from .auth_utils import is_administrator
from .cert_extract import extract_certificate, populate_template_workbook, _parse_iso_date
from .cert_sheet import label_slug
from .cert_xlsx import workbook_bytes
from .models import CompanyProfile, EquipmentCertificateRecord, EquipmentCertificateTemplate, Primary

TMP_DIR = Path(settings.EI_PRIVATE_STORAGE_ROOT) / 'tmp' / 'cert_upload'
SESSION_KEY = 'equipment_log_upload'
MAX_ROWS = 500
MAX_COLS = 60
ALLOWED_SUFFIXES = {'.pdf', '.doc', '.docx', '.xls', '.xlsx', '.csv'}
MAX_UPLOAD_BYTES = 20 * 1024 * 1024
PROCESSING_TTL = 600  # seconds


# ----------------------------------------------------------------- helpers

def _admin_required(view):
    @login_required
    @functools.wraps(view)
    def wrapper(request, *args, **kwargs):
        if not is_administrator(request.user):
            raise PermissionDenied('Equipment Logs is restricted to administrators.')
        return view(request, *args, **kwargs)
    return wrapper


def _cache_key(token):
    return f'equip_log_proc:{token}'


def _store_result(token, data, ttl=PROCESSING_TTL):
    cache.set(_cache_key(token), data, ttl)


def _get_result(token):
    return cache.get(_cache_key(token))


def _session_upload(request, key=None):
    data = request.session.get(SESSION_KEY, None) or {}
    if key is None:
        return data
    return data.get(key)


def _stash_upload(request, **payload):
    data = request.session.get(SESSION_KEY, None) or {}
    data.update(payload)
    request.session[SESSION_KEY] = data
    request.session.modified = True
    return data


def _clear_upload(request):
    data = request.session.pop(SESSION_KEY, None) or {}
    token = data.get('token')
    template_token = data.get('template_token')
    for candidate in (token, template_token):
        if not candidate:
            continue
        safe = re.sub(r'[^a-f0-9]', '', str(candidate).lower())
        for suffix in ('.pdf', '.doc', '.docx', '.xls', '.xlsx', '.csv', '.html', '.json', '.xlsx'):
            try:
                (TMP_DIR / (safe + suffix)).unlink(missing_ok=True)
            except OSError:
                pass


_CONTENT_TYPES = {
    '.pdf': 'application/pdf',
    '.doc': 'application/msword',
    '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    '.xls': 'application/vnd.ms-excel',
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    '.csv': 'text/csv',
}


def _content_type(filename):
    suffix = Path(filename).suffix.lower()
    return _CONTENT_TYPES.get(suffix, 'application/octet-stream')


def _logo_path():
    """Filesystem path to the Engineering Automation logo shown in the top bar."""
    try:
        company = CompanyProfile.objects.first()
        if company and company.logo:
            return company.logo.path
    except Exception:
        pass
    static_dirs = settings.STATICFILES_DIRS or []
    for static_dir in static_dirs:
        path = static_dir / 'core' / 'images' / 'image2.jpeg'
        if path.is_file():
            return str(path)
    return None


# ----------------------------------------------------------------- home

@_admin_required
def home(request):
    records = EquipmentCertificateRecord.objects.all()
    years = list(records.values_list('year', flat=True).distinct().order_by('-year'))
    current_year = timezone.localdate().year
    if current_year not in years:
        years = sorted({*years, current_year}, reverse=True)
    year_cards = []
    for year in years:
        year_records = records.filter(year=year)
        year_cards.append({
            'year': year,
            'types': year_records.values('template').distinct().count(),
            'certificates': year_records.count(),
        })
    return render(request, 'core/equipment_logs/home.html', {
        'year_cards': year_cards,
        'template_count': EquipmentCertificateTemplate.objects.count(),
        'page_title': 'Equipment Logs',
    })


# ----------------------------------------------------------------- step 1: upload

@_admin_required
def upload(request):
    if request.method == 'POST':
        template_id = request.POST.get('template')
        file = request.FILES.get('certificate')
        if not file:
            messages.error(request, 'Choose a certificate file (PDF, Word or Excel).')
            return redirect('equipment_log_upload')
        suffix = Path(file.name or '').suffix.lower()
        if suffix not in ALLOWED_SUFFIXES:
            messages.error(request, 'Supported file types: PDF, DOC, DOCX, XLS, XLSX, CSV.')
            return redirect('equipment_log_upload')
        if file.size > MAX_UPLOAD_BYTES:
            messages.error(request, 'File is larger than 20 MB.')
            return redirect('equipment_log_upload')
        if not template_id:
            messages.error(request, 'Select a certificate template before uploading.')
            return redirect('equipment_log_upload')
        template = get_object_or_404(EquipmentCertificateTemplate, pk=template_id)

        _clear_upload(request)
        token = uuid.uuid4().hex
        TMP_DIR.mkdir(parents=True, exist_ok=True)
        staged = TMP_DIR / (token + suffix)
        with staged.open('wb') as handle:
            for chunk in file.chunks():
                handle.write(chunk)

        # Run the full extraction pipeline automatically against the template.
        extraction = extract_certificate(str(staged), file.name or f'certificate{suffix}', template)

        # Populate the template workbook with matched values.
        wb, out_filename, filled_count = populate_template_workbook(template, extraction['field_matches'])

        # Save the completed workbook to a temp file for download.
        out_staged = TMP_DIR / (token + '_completed.xlsx')
        wb.save(str(out_staged))

        # Store extraction result for the review page. Do NOT save to DB yet.
        _store_result(token, {
            'template_id': template.pk,
            'template_name': template.name,
            'filename': file.name or f'certificate{suffix}',
            'output_filename': out_filename,
            'provider': extraction['provider'],
            'processing_time_ms': extraction['processing_time_ms'],
            'fields': extraction['fields'],
            'field_confidence': extraction['field_confidence'],
            'activities': extraction['activities'],
            'columns': extraction['columns'],
            'rows': extraction['rows'],
            'missing': extraction['missing'],
            'field_matches': extraction['field_matches'],
            'filled_count': filled_count,
            'staged_path': str(staged),
            'out_staged_path': str(out_staged),
        })
        _stash_upload(request, token=token, filename=file.name or f'certificate{suffix}',
                      template_id=template.pk)

        return redirect('equipment_log_review', token=token)

    templates = EquipmentCertificateTemplate.objects.all().order_by('name')
    return render(request, 'core/equipment_logs/upload_step1.html', {
        'templates': templates,
        'page_title': 'Upload Certificate',
    })


# ----------------------------------------------------------------- step 2: review

@_admin_required
def review(request, token):
    """Review extracted certificate data before saving to the log."""
    result = _get_result(token)
    if not result:
        messages.error(request, 'Processing session not found. Please upload again.')
        return redirect('equipment_log_upload')

    # Render the sheet preview from extracted data only (not template file cells).
    sheet_html = ''
    try:
        from .cert_sheet import render_sheet, load_sheet
        template = get_object_or_404(EquipmentCertificateTemplate, pk=result['template_id'])
        with template.original_file.open('rb') as handle:
            ws = load_sheet(handle, None)
            sheet_html = render_sheet(
                ws, editable=True,
                values={
                    'activities': result['activities'] or {},
                    'rows': result['rows'] or [],
                    'columns': result['columns'] or [],
                },
                confidence=result['field_confidence'] or {},
            )
    except Exception:
        sheet_html = ''

    # Determine how many fields need review (confidence < 0.8)
    needs_review = [
        m for m in result['field_matches']
        if m['confidence'] < 0.8 and m['matched_value']
    ]

    if request.method == 'POST':
        # Save the record from edited values.
        from .ei_workflow import actor_fields

        values = {
            'columns': result['columns'],
            'rows': result['rows'],
            'field_values': result['activities'],
        }
        # Copy any template-defined fields into the values dict
        for slug, val in result['fields'].items():
            values[slug] = val

        # Override with POST data (edited values from review page)
        for slug in values.get('field_values', {}):
            posted = request.POST.get(f'act_{slug}', '')
            if posted != '':
                values['field_values'][slug] = posted

        # Read test-data cell inputs from the review page (cell_<row>_<col>).
        # These take priority over rows_json, which is only a fallback.
        cells = {}
        for key, value in request.POST.items():
            if key.startswith('cell_'):
                parts = key.split('_')
                if len(parts) == 3:
                    try:
                        row_idx = int(parts[1])
                        col_idx = int(parts[2])
                        cells[(row_idx, col_idx)] = value
                    except ValueError:
                        pass
        if cells:
            max_row = max(r for r, _ in cells.keys())
            max_col = max(c for _, c in cells.keys())
            new_rows = []
            for r in range(max_row + 1):
                row = []
                for c in range(max_col + 1):
                    row.append(cells.get((r, c), ''))
                new_rows.append(row)
            values['rows'] = new_rows
        else:
            # Use edited cell values from the review page (rows_json) when present
            rows_raw = request.POST.get('rows_json') or ''
            if rows_raw:
                try:
                    values['rows'] = json.loads(rows_raw)
                except (ValueError, TypeError):
                    pass

        # Take year from posted test date; fall back to current year if unparseable
        test_date_slug = ''
        for lbl in (template.activities or []):
            if 'test date' in lbl.lower() or lbl.lower() == 'test_date':
                test_date_slug = label_slug(lbl)
                break
        posted_test_date = (request.POST.get(f'act_{test_date_slug}', '') or '').strip() if test_date_slug else ''
        test_date = _parse_iso_date(posted_test_date) if posted_test_date else None
        if test_date is None:
            test_date_str = (values.get('test_date') or '').strip()
            test_date = _parse_iso_date(test_date_str)
        if test_date is None:
            test_date = timezone.localdate()
        year = test_date.year

        file_name = result['filename']

        # Check for duplicate (template + file_name + year, case-insensitive)
        existing = EquipmentCertificateRecord.objects.filter(
            template=template, file_name__iexact=file_name, year=year
        ).first()

        if existing and not request.POST.get('replace'):
            # Show replace confirmation dialog
            sheet_html = ''
            try:
                from .cert_sheet import render_sheet, load_sheet
                tpl = get_object_or_404(EquipmentCertificateTemplate, pk=result['template_id'])
                with tpl.original_file.open('rb') as handle:
                    ws = load_sheet(handle, None)
                    sheet_html = render_sheet(
                        ws, editable=True,
                        values={
                            'activities': result['activities'] or {},
                            'rows': result['rows'] or [],
                            'columns': result['columns'] or [],
                        },
                        confidence=result['field_confidence'] or {},
                    )
            except Exception:
                sheet_html = ''
            needs_review = [
                m for m in result['field_matches']
                if m['confidence'] < 0.8 and m['matched_value']
            ]
            return render(request, 'core/equipment_logs/review.html', {
                'token': token,
                'template_name': result['template_name'],
                'filename': result['filename'],
                'provider': result['provider'],
                'processing_time_ms': result['processing_time_ms'],
                'processing_time_s': round(result['processing_time_ms'] / 1000.0, 2),
                'fields': result['fields'],
                'field_confidence': result['field_confidence'],
                'activities': result['activities'],
                'columns': result['columns'],
                'rows': result['rows'],
                'missing': result['missing'],
                'field_matches': result['field_matches'],
                'sheet_html': sheet_html,
                'filled_count': result['filled_count'],
                'needs_review': needs_review,
                'total_fields': len(result['field_matches']),
                'page_title': 'Review & fix',
                'replace_confirm': True,
                'existing_file_name': existing.file_name,
            })

        # Save the record (delete existing if replacing)
        if existing:
            existing.delete()

        record = EquipmentCertificateRecord(
            template=template,
            template_name=template.name,
            file_name=file_name,
            year=year,
            test_date=test_date,
            **actor_fields(request.user),
        )
        record.due_date = _parse_iso_date(values.get('due_date'))
        record.model = values.get('model', '')
        record.certificate_no = values.get('certificate_no', '')
        record.environment = values.get('environment', '')
        record.conclusion = values.get('conclusion', '') or ''
        record.columns = values.get('columns', [])
        record.rows = values.get('rows', [])
        record.field_values = values.get('field_values') or {}
        history_entry = {
            'when': timezone.now().isoformat(timespec='seconds'),
            'who': request.user.username,
            'reason': 'Certificate uploaded',
        }
        record.change_history = [history_entry]
        record.save()
        year = record.year

        # Clean up staged upload
        _clear_upload(request)

        # Redirect using record id + file query param
        url = reverse('equipment_log_type', kwargs={'year': record.year, 'template_pk': template.pk})
        return redirect('%s?file=%s' % (url, record.file_name))

    # Render review page
    sheet_html = ''
    try:
        from .cert_sheet import render_sheet, load_sheet
        template = get_object_or_404(EquipmentCertificateTemplate, pk=result['template_id'])
        with template.original_file.open('rb') as handle:
            ws = load_sheet(handle, None)
            sheet_html = render_sheet(
                ws, editable=True,
                values={
                    'activities': result['activities'] or {},
                    'rows': result['rows'] or [],
                    'columns': result['columns'] or [],
                },
                confidence=result['field_confidence'] or {},
            )
    except Exception:
        sheet_html = ''

    return render(request, 'core/equipment_logs/review.html', {
        'token': token,
        'template_name': result['template_name'],
        'filename': result['filename'],
        'provider': result['provider'],
        'processing_time_ms': result['processing_time_ms'],
        'processing_time_s': round(result['processing_time_ms'] / 1000.0, 2),
        'fields': result['fields'],
        'field_confidence': result['field_confidence'],
        'activities': result['activities'],
        'columns': result['columns'],
        'rows': result['rows'],
        'missing': result['missing'],
        'field_matches': result['field_matches'],
        'sheet_html': sheet_html,
        'filled_count': result['filled_count'],
        'needs_review': needs_review,
        'total_fields': len(result['field_matches']),
        'page_title': 'Review & fix',
    })


# ----------------------------------------------------------------- existing log views (unchanged except minor updates)

@_admin_required
def year_view(request, year):
    year = int(year)
    records = EquipmentCertificateRecord.objects.filter(year=year)
    types = list(records.values('template', 'template_name')
                 .annotate(certificates=Count('pk'))
                 .order_by('template_name'))
    return render(request, 'core/equipment_logs/year.html', {
        'year': year,
        'types': types,
        'page_title': f'Equipment Logs {year}',
    })


@_admin_required
def type_view(request, year, template_pk):
    year = int(year)
    template = get_object_or_404(EquipmentCertificateTemplate, pk=template_pk)
    records = list(EquipmentCertificateRecord.objects.filter(year=year, template=template).order_by('file_name'))
    selected_file = request.GET.get('file')
    selected = None
    for record in records:
        if record.file_name == (selected_file or '') or (selected is None and not selected_file):
            selected = record
            break
    unlock_flag = f'equipment_log_unlock_{template.pk}_{year}_{selected.pk}' if selected else None
    unlock = bool(request.session.get(unlock_flag)) if selected else False
    sheet_html = ''
    if selected:
        from .cert_sheet import render_sheet, load_sheet
        with template.original_file.open('rb') as handle:
            ws = load_sheet(handle, None)
            sheet_html = render_sheet(
                ws, editable=unlock,
                values={
                    'activities': selected.field_values or {},
                    'rows': selected.rows or [],
                    'columns': list(template.columns or []),
                },
            )
    return render(request, 'core/equipment_logs/type.html', {
        'year': year,
        'template': template,
        'records': records,
        'selected': selected,
        'unlock': unlock,
        'activity_pairs': _activity_pairs(list(template.activities or []), selected.field_values) if selected else [],
        'sheet_html': sheet_html,
        'page_title': f'{template.name} · {year}',
    })


def _activity_pairs(labels, values):
    values = values or {}
    return [(label, label_slug(label), values.get(label_slug(label), '')) for label in labels]


@_admin_required
def unlock_record(request, record_pk):
    if request.method != 'POST':
        raise Http404()
    record = get_object_or_404(EquipmentCertificateRecord, pk=record_pk)
    reason = (request.POST.get('reason') or '').strip()
    if not reason:
        messages.error(request, 'A reason is required to edit a saved record.')
    else:
        request.session[f'equipment_log_unlock_{record.template_id}_{record.year}_{record.pk}'] = True
        request.session[f'equipment_log_reason_{record.pk}'] = reason
        request.session.modified = True
        messages.info(request, 'Editing unlocked. Remember: every change is recorded.')
    return redirect('%s?file=%s' % (
        reverse('equipment_log_type', kwargs={'year': record.year, 'template_pk': record.template_id}),
        record.file_name))


@_admin_required
def update_record(request, record_pk):
    if request.method != 'POST':
        raise Http404()
    record = get_object_or_404(EquipmentCertificateRecord, pk=record_pk)
    flag = f'equipment_log_unlock_{record.template_id}_{record.year}_{record.pk}'
    if not request.session.get(flag):
        messages.error(request, 'The record is locked. State a reason to edit it.')
        return redirect('equipment_log_type', year=record.year, template_pk=record.template_id)
    template = record.template
    from .cert_extract import _parse_iso as _parse_iso
    values = _values_from_post_auto(request, template)
    reason = request.session.get(f'equipment_log_reason_{record.pk}', 'Edited')
    record.model = values.get('model', '')
    record.certificate_no = values.get('certificate_no', '')
    record.environment = values.get('environment', '')
    test_date = _parse_iso(values.get('test_date'))
    if test_date:
        record.test_date = test_date
    record.due_date = _parse_iso(values.get('due_date'))
    if values.get('conclusion') in ('pass', 'fail', ''):
        record.conclusion = values['conclusion']
    record.columns = values.get('columns', record.columns)
    record.rows = values.get('rows', record.rows)
    record.field_values = values.get('field_values', record.field_values)
    record.change_history = list(record.change_history or []) + [{
        'when': timezone.now().isoformat(timespec='seconds'),
        'who': request.user.username,
        'reason': reason,
    }]
    record.save()
    request.session.pop(flag, None)
    request.session.pop(f'equipment_log_reason_{record.pk}', None)
    service.audit(request.user, None, 'equipment_certificate_edited',
                  after={'record': record.pk, 'reason': reason})
    messages.success(request, 'Record updated and history logged.')
    return redirect('%s?file=%s' % (
        reverse('equipment_log_type', kwargs={'year': record.year, 'template_pk': record.template_id}),
        record.file_name))


def _values_from_post_auto(request, template):
    activities = list(template.activities or []) if template else []
    rows_raw = request.POST.get('rows_json') or '[]'
    try:
        rows = [[str(cell)[:200] for cell in row][:MAX_COLS] for row in json.loads(rows_raw)][:MAX_ROWS]
    except (ValueError, TypeError):
        rows = []
    conclusion = request.POST.get('conclusion', '').strip().lower()
    conclusion = conclusion if conclusion in ('pass', 'fail') else ''
    field_values = {}
    for label in activities:
        slug = label_slug(label)
        field_values[slug] = request.POST.get(f'act_{slug}', '').strip()
    flat = {
        'serial_no': request.POST.get('serial_no', '').strip(),
        'model': request.POST.get('model', '').strip(),
        'certificate_no': request.POST.get('certificate_no', '').strip(),
        'test_date': request.POST.get('test_date', '').strip(),
        'due_date': request.POST.get('due_date', '').strip(),
        'environment': request.POST.get('environment', '').strip(),
    }
    flat['conclusion'] = conclusion
    flat['columns'] = list(template.columns or []) if template else []
    flat['rows'] = rows
    flat['field_values'] = field_values
    return flat


# ----------------------------------------------------------------- downloads

def _safe_filename(name):
    return re.sub(r'[\\/:*?"<>|]+', ' ', str(name)).strip() or 'download'


def _sheet_title(workbook, name):
    clean = re.sub(r'[\[\]:*?/\\]', ' ', str(name))[:31].strip() or 'Sheet'
    title, index = clean, 2
    while title in workbook.sheetnames:
        suffix = f' ({index})'
        title = (clean[:31 - len(suffix)] + suffix).strip()
        index += 1
    return title


def _record_sheet(workbook, record):
    template = record.template
    sheet = workbook.create_sheet(_sheet_title(workbook, record.file_name or 'serial'))
    sheet.append(['Equipment Log — certificate record'])
    sheet.append([])
    if record.field_values and (template and template.activities):
        sheet.append(['LAB GENERAL ACTIVITIES LOG'])
        for label in template.activities:
            sheet.append([label, record.field_values.get(label_slug(label), '')])
        sheet.append([])
    sheet.append(['TEST DATA'])
    if record.columns:
        sheet.append(list(record.columns))
        for row in record.rows or []:
            sheet.append(list(row))
    sheet.append([])
    sheet.append(['Conclusion', record.get_conclusion_display() if record.conclusion else ''])
    for col in sheet.columns:
        width = max(len(str(cell.value or '')) for cell in col) + 2
        sheet.column_dimensions[col[0].column_letter].width = min(width, 60)
    return sheet


def _xlsx_response(workbook, filename):
    from django.http import HttpResponse
    import io
    buffer = io.BytesIO()
    workbook.save(buffer)
    response = HttpResponse(buffer.getvalue(),
                            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml')
    response['Content-Disposition'] = 'attachment; filename="%s"' % _safe_filename(filename)
    return response


@_admin_required
def download_record(request, record_pk):
    record = get_object_or_404(EquipmentCertificateRecord, pk=record_pk)
    template = record.template
    logo_path = _logo_path()
    stem = Path(record.file_name or '').stem or 'Record'
    fields = [(label, record.field_values.get(label_slug(label), '')) for label in (template.activities or [])]
    records = [{"name": stem, "fields": fields, "columns": list(record.columns or []), "rows": list(record.rows or [])}]
    data = workbook_bytes(records, logo_path)
    from django.http import HttpResponse
    response = HttpResponse(data, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml')
    response['Content-Disposition'] = 'attachment; filename="%s"' % _safe_filename(f'{record.template_name} {record.year} {record.file_name}.xlsx')
    return response


@_admin_required
def download_type(request, year, template_pk):
    year = int(year)
    template = get_object_or_404(EquipmentCertificateTemplate, pk=template_pk)
    logo_path = _logo_path()
    records_qs = EquipmentCertificateRecord.objects.filter(year=year, template=template).order_by('file_name')
    records = []
    for record in records_qs:
        stem = Path(record.file_name or '').stem or 'Record'
        fields = [(label, record.field_values.get(label_slug(label), '')) for label in (template.activities or [])]
        records.append({"name": stem, "fields": fields, "columns": list(record.columns or []), "rows": list(record.rows or [])})
    data = workbook_bytes(records, logo_path)
    from django.http import HttpResponse
    response = HttpResponse(data, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml')
    response['Content-Disposition'] = 'attachment; filename="%s"' % _safe_filename(f'{template.name} Certificates Log {year}.xlsx')
    return response


@_admin_required
def download_year(request, year):
    year = int(year)
    logo_path = _logo_path()
    records_qs = EquipmentCertificateRecord.objects.filter(year=year).order_by('template_name', 'file_name')
    records = []
    for record in records_qs:
        stem = Path(record.file_name or '').stem or 'Record'
        template = record.template
        fields = [(label, record.field_values.get(label_slug(label), '')) for label in (template.activities or [])]
        records.append({"name": stem, "fields": fields, "columns": list(record.columns or []), "rows": list(record.rows or [])})
    data = workbook_bytes(records, logo_path)
    from django.http import HttpResponse
    response = HttpResponse(data, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml')
    response['Content-Disposition'] = 'attachment; filename="%s"' % _safe_filename(f'Equipment Certificates Log {year}.xlsx')
    return response


# ----------------------------------------------------------------- templates folder

@_admin_required
def templates_folder(request):
    templates = EquipmentCertificateTemplate.objects.annotate(
        records_count=Count('records')).order_by('name')
    return render(request, 'core/equipment_logs/templates.html', {
        'templates': templates,
        'page_title': 'Certificate Templates',
    })


@_admin_required
def upload_template(request):
    if request.method == 'POST':
        file = request.FILES.get('template')
        name = (request.POST.get('name') or '').strip()
        if not file or not (file.name or '').lower().endswith('.xlsx'):
            messages.error(request, 'Choose an .xlsx template file.')
            return redirect('equipment_log_template_upload')
        if file.size > MAX_UPLOAD_BYTES:
            messages.error(request, 'File is larger than 20 MB.')
            return redirect('equipment_log_template_upload')
        if not name:
            name = Path(file.name).stem
        _clear_upload(request)
        token = uuid.uuid4().hex
        TMP_DIR.mkdir(parents=True, exist_ok=True)
        staged_path = TMP_DIR / (token + '.xlsx')
        with staged_path.open('wb') as handle:
            for chunk in file.chunks():
                handle.write(chunk)
        _stash_upload(request, template_token=token, template_name=name, filename=file.name)
        return redirect('equipment_log_template_confirm')
    return render(request, 'core/equipment_logs/template_upload.html', {
        'page_title': 'Upload template',
    })


def _template_staged(request):
    data = request.session.get(SESSION_KEY) or {}
    token = data.get('template_token')
    if not token:
        return None, data
    staged = TMP_DIR / (re.sub(r'[^a-f0-9]', '', str(token).lower()) + '.xlsx')
    if not staged.is_file():
        return None, data
    return staged, data


@_admin_required
def template_confirm(request):
    staged, data = _template_staged(request)
    if staged is None or not data.get('template_token'):
        messages.error(request, 'Upload a template first.')
        return redirect('equipment_log_template_upload')

    if request.method == 'POST':
        name = (request.POST.get('name') or data.get('template_name') or '').strip()
        selected = set(request.POST.getlist('include'))
        from .template_parser import parse_sheet
        parse = parse_sheet(str(staged), request.POST.get('sheet') or None)
        activities, columns = [], []
        for index, candidate in enumerate(parse['candidates']):
            if str(index) not in selected:
                continue
            if candidate['type'] == 'fields':
                activities.extend(candidate['labels'])
            elif candidate['type'] == 'table':
                columns = candidate['columns']
        if not activities and not columns:
            messages.error(request, 'Include at least one activities block or the test-data table.')
            return redirect('equipment_log_template_confirm')
        if EquipmentCertificateTemplate.objects.filter(name__iexact=name).exists():
            messages.error(request, f'A template named "{name}" already exists.')
            return redirect('equipment_log_template_confirm')
        template = EquipmentCertificateTemplate(
            name=name,
            activities=activities,
            columns=columns,
            original_name=data.get('filename') or 'template.xlsx',
            **service.actor_fields(request.user),
        )
        # Collect fixed values from "Keep this value on every certificate" checkboxes
        fixed_values = {}
        for key, value in request.POST.items():
            if key.startswith('fixed_'):
                parts = key.split('_')
                if len(parts) == 3:
                    try:
                        ci = int(parts[1])
                        li = int(parts[2])
                        candidate = parse['candidates'][ci]
                        if candidate['type'] == 'fields':
                            # Checkbox value is f.key (the label_slug key); only sent when checked
                            slug = value
                            if slug:
                                fixed_value = request.POST.get(f'fixed_value_{ci}_{li}', '')
                                samples = candidate.get('samples') or []
                                sample = samples[li] if li < len(samples) else ''
                                if fixed_value or sample:
                                    fixed_values[slug] = fixed_value or sample
                    except (IndexError, ValueError):
                        pass
        template.fixed_values = fixed_values
        with staged.open('rb') as handle:
            from django.core.files import File
            template.original_file.save(data.get('filename') or 'template.xlsx', File(handle), save=True)
        service.audit(request.user, None, 'equipment_template_uploaded',
                      after={'template': template.pk, 'name': template.name,
                             'activities': activities, 'columns': columns})
        _clear_upload(request)
        messages.success(request, f'Template "{name}" saved: {len(activities)} activity fields, {len(columns)} test-data columns.')
        return redirect('equipment_log_templates')

    sheet_name = request.GET.get('sheet')
    from .template_parser import parse_sheet
    parse = parse_sheet(str(staged), sheet_name)
    from .cert_sheet import render_sheet, load_sheet
    with staged.open('rb') as handle:
        ws = load_sheet(handle, None)
        # Build activities dict from parsed template candidates
        activities = {}
        columns = []
        for c in parse['candidates']:
            if c['type'] == 'fields':
                c['fields'] = []
                samples = c.get('samples') or []
                for i, label in enumerate(c['labels']):
                    c['fields'].append({
                        'label': label,
                        'key': label_slug(label),
                        'sample': samples[i] if i < len(samples) else '',
                    })
                for label in c['labels']:
                    activities[label_slug(label)] = ''
            elif c['type'] == 'table':
                columns = c['columns']
        sheet_html = render_sheet(ws, editable=False, values={
            'activities': activities,
            'columns': columns,
            'rows': [],
        })
    return render(request, 'core/equipment_logs/template_confirm.html', {
        'sheets': parse['sheets'],
        'sheet': parse['sheet'],
        'candidates': parse['candidates'],
        'preview_rows': parse['preview_rows'],
        'sheet_html': sheet_html,
        'default_name': data.get('template_name') or Path(data.get('filename') or 'template').stem,
        'page_title': 'Review what was detected',
    })


@_admin_required
def template_delete(request, template_pk):
    if request.method != 'POST':
        raise Http404()
    template = get_object_or_404(EquipmentCertificateTemplate, pk=template_pk)
    record_count = template.records.count()
    if template.original_file:
        template.original_file.delete(save=False)
    template.records.all().delete()
    name = template.name
    template.delete()
    service.audit(request.user, None, 'equipment_template_deleted',
                  after={'name': name, 'records_deleted': record_count})
    messages.info(request, f'Deleted template "{name}" with {record_count} log record(s).')
    return redirect('equipment_log_templates')


@_admin_required
def delete_record(request, record_pk):
    if request.method != 'POST':
        raise Http404()
    record = get_object_or_404(EquipmentCertificateRecord, pk=record_pk)
    year = record.year
    template_pk = record.template_id
    serial = record.serial_no
    file_name = record.file_name
    name = record.template_name
    record.delete()
    service.audit(request.user, None, 'equipment_certificate_deleted',
                  after={'record': record_pk, 'serial': serial, 'year': year,
                         'template': name})
    messages.info(request, f'Deleted log record for file "{file_name}" ({name}, {year}).')
    return redirect('equipment_log_type', year=year, template_pk=template_pk)