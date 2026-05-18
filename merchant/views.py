import io
import os
import pickle
import shutil
from pathlib import Path
from datetime import datetime, timedelta

import pandas as pd
from django.conf import settings
from django.http import HttpResponse, HttpResponseRedirect
from django.shortcuts import render, redirect
from django.views.decorators.http import require_POST

from .forms import UploadForm, SheetSelectForm, CIFLookupForm, DateRangeForm
from .utils import (
    normalize_columns,
    extract_date_columns,
    _to_numeric,
    _parse_col_date,
    resolve_dates,
    build_top15_data,
    build_plotly_bar_div,
    build_performance_report_bytes,
    _build_gainers_shakers,
    _write_gainers_excel,
    build_top25_data,
    build_top25_excel_bytes,
    find_idle_terminals,
    idle_excel_bytes,
    pos_statement_bytes,
    champions_zinara_bytes,
    champions_insurance_bytes,
    compute_mtd_from_total_row,
    compute_daily_from_total_row,
    build_merchant_revenue_chart,
    merchant_period_excel_bytes,
    validate_b02_file,
    parse_b02_files,
    update_merchant_report,
)


# =============================================================================
# SESSION / FILE HELPERS
# =============================================================================

def _session_dir(request):
    key = request.session.session_key
    if not key:
        request.session.create()
        key = request.session.session_key
    d = Path(settings.TEMP_DATA_DIR) / 'uploads' / key
    d.mkdir(parents=True, exist_ok=True)
    return d


def _pickle_dir(request):
    key = request.session.session_key
    if not key:
        request.session.create()
        key = request.session.session_key
    d = Path(settings.TEMP_DATA_DIR) / 'pickles' / key
    d.mkdir(parents=True, exist_ok=True)
    return d


def _save_df(request, name, df):
    path = _pickle_dir(request) / f'{name}.pkl'
    with open(path, 'wb') as fh:
        pickle.dump(df, fh)


def _load_df(request, name):
    path = _pickle_dir(request) / f'{name}.pkl'
    if not path.exists():
        return None
    with open(path, 'rb') as fh:
        return pickle.load(fh)


def _require_data(request):
    """Redirect to upload if data not loaded yet."""
    if not request.session.get('data_loaded'):
        return redirect('upload')
    return None


def _require_files(request):
    """Redirect to upload if files not validated yet."""
    if not request.session.get('files_validated'):
        return redirect('upload')
    return None


def _load_all_sheets(request):
    """Load ALL sheets from both uploaded files, concat per currency, save as pickles.
    Returns an error string on failure, or None on success."""
    d = _session_dir(request)
    all_zwg, all_usd = [], []
    required = {'TID', 'CIF', 'MERCHANT NAME'}
    try:
        for sheet in pd.ExcelFile(d / 'zwg.xlsx').sheet_names:
            try:
                df = normalize_columns(pd.read_excel(d / 'zwg.xlsx', sheet_name=sheet))
                if not required.issubset(set(df.columns)):
                    continue
                df = df[~df['MERCHANT NAME'].astype(str).str.upper().str.strip().isin(['TOTAL', 'GRAND TOTAL'])]
                all_zwg.append(df)
            except Exception:
                pass
        for sheet in pd.ExcelFile(d / 'usd.xlsx').sheet_names:
            try:
                df = normalize_columns(pd.read_excel(d / 'usd.xlsx', sheet_name=sheet))
                if not required.issubset(set(df.columns)):
                    continue
                df = df[~df['MERCHANT NAME'].astype(str).str.upper().str.strip().isin(['TOTAL', 'GRAND TOTAL'])]
                all_usd.append(df)
            except Exception:
                pass
    except Exception as e:
        return str(e)

    if all_zwg:
        _save_df(request, 'zwg_all_df', pd.concat(all_zwg, ignore_index=True))
    if all_usd:
        _save_df(request, 'usd_all_df', pd.concat(all_usd, ignore_index=True))

    request.session['periodic_data_loaded'] = True

    try:
        combined = _load_df(request, 'zwg_all_df')
        if combined is not None:
            date_cols = extract_date_columns(combined)
            parsed = sorted([_parse_col_date(c) for c in date_cols if pd.notnull(_parse_col_date(c))])
            if parsed:
                request.session['data_period_all'] = (
                    f"{parsed[0].strftime('%d %b %Y')} → {parsed[-1].strftime('%d %b %Y')}"
                )
    except Exception:
        pass

    return None


def _aggregate_by_tid(df):
    """Collapse multi-month concatenated rows per TID into one row by summing date columns.
    Needed so idle-terminal detection works correctly across months."""
    if df is None or df.empty:
        return df
    date_cols = extract_date_columns(df)
    df = df.copy()
    for col in date_cols:
        df[col] = _to_numeric(df[col])
    fixed = [c for c in ['TID', 'CIF', 'MERCHANT NAME', 'BUSINESS UNIT'] if c in df.columns]
    agg = {c: 'first' for c in fixed if c != 'TID'}
    agg.update({c: 'sum' for c in date_cols if c in df.columns})
    return df.groupby('TID', as_index=False).agg(agg)


def _tids_in_month(df, all_date_cols, year, month):
    """Return set of TIDs whose rows come from the given month's sheet.
    Rows from month X have real values (0 or positive) for that month's date columns
    and NaN for all other months — notna() therefore reliably detects presence."""
    month_cols = [
        c for c in all_date_cols
        if pd.notnull(_parse_col_date(c))
        and _parse_col_date(c).year  == year
        and _parse_col_date(c).month == month
    ]
    if not month_cols:
        return None
    mask = df[month_cols].notna().any(axis=1)
    return set(df.loc[mask, 'TID'].unique())


# =============================================================================
# UPLOAD
# =============================================================================

def upload_view(request):
    form = UploadForm()
    errors = []

    if request.method == 'POST':
        form = UploadForm(request.POST, request.FILES)
        if form.is_valid():
            request.session.create()
            d = _session_dir(request)

            for field, name in [('zwg_file', 'zwg.xlsx'), ('usd_file', 'usd.xlsx')]:
                uploaded = form.cleaned_data[field]
                dest = d / name
                with open(dest, 'wb') as fh:
                    for chunk in uploaded.chunks():
                        fh.write(chunk)

            try:
                zwg_xls = pd.ExcelFile(d / 'zwg.xlsx')
                usd_xls = pd.ExcelFile(d / 'usd.xlsx')
            except Exception as e:
                errors.append(f'Could not read one of the uploaded files: {e}')
            else:
                request.session['zwg_sheet_names'] = zwg_xls.sheet_names
                request.session['usd_sheet_names'] = usd_xls.sheet_names
                request.session['files_validated'] = True
                request.session['data_loaded'] = False
                request.session['periodic_data_loaded'] = False
                return redirect('choose_analytics')

    return render(request, 'merchant/upload.html', {'form': form, 'errors': errors})


# =============================================================================
# SHEET SELECTION
# =============================================================================

def select_sheet_view(request):
    if not request.session.get('files_validated'):
        return redirect('upload')

    zwg_sheets = request.session.get('zwg_sheet_names', [])
    usd_sheets = request.session.get('usd_sheet_names', [])

    form = SheetSelectForm(zwg_sheets=zwg_sheets, usd_sheets=usd_sheets)

    if request.method == 'POST':
        form = SheetSelectForm(request.POST, zwg_sheets=zwg_sheets, usd_sheets=usd_sheets)
        if form.is_valid():
            zwg_sheet = form.cleaned_data['zwg_sheet']
            usd_sheet = form.cleaned_data['usd_sheet']

            d = _session_dir(request)
            try:
                zwg_df = normalize_columns(pd.read_excel(d / 'zwg.xlsx', sheet_name=zwg_sheet))
                usd_df = normalize_columns(pd.read_excel(d / 'usd.xlsx', sheet_name=usd_sheet))
            except Exception as e:
                form.add_error(None, f'Error loading data: {e}')
            else:
                _save_df(request, 'zwg_df', zwg_df)
                _save_df(request, 'usd_df', usd_df)
                request.session['data_loaded'] = True

                # Pre-compute data period label for sidebar
                try:
                    date_cols = extract_date_columns(zwg_df)
                    parsed = sorted([_parse_col_date(c) for c in date_cols if pd.notnull(_parse_col_date(c))])
                    if parsed:
                        request.session['data_period'] = (
                            f"{parsed[0].strftime('%d %b %Y')} → {parsed[-1].strftime('%d %b %Y')}"
                        )
                except Exception:
                    pass

                return redirect('zwg_performance')

    return render(request, 'merchant/select_sheet.html', {'form': form})


# =============================================================================
# ANALYTICS CHOICE
# =============================================================================

def choose_analytics_view(request):
    if not request.session.get('files_validated'):
        return redirect('upload')
    return render(request, 'merchant/choose_analytics.html', {
        'zwg_sheet_count': len(request.session.get('zwg_sheet_names', [])),
        'usd_sheet_count': len(request.session.get('usd_sheet_names', [])),
    })


# =============================================================================
# PERIODIC ANALYTICS HUB
# =============================================================================

def periodic_analytics_view(request):
    redir = _require_files(request)
    if redir:
        return redir

    if not request.session.get('periodic_data_loaded'):
        err = _load_all_sheets(request)
        if err:
            return render(request, 'merchant/periodic_analytics.html', {
                'error': f'Could not load sheets: {err}',
                'active_page': 'periodic_hub',
                'section': 'periodic',
                'data_period': '',
            })

    return render(request, 'merchant/periodic_analytics.html', {
        'active_page': 'periodic_hub',
        'section': 'periodic',
        'data_period': request.session.get('data_period_all', request.session.get('data_period', '')),
        'zwg_sheet_count': len(request.session.get('zwg_sheet_names', [])),
        'usd_sheet_count': len(request.session.get('usd_sheet_names', [])),
    })


# =============================================================================
# ZWG PERFORMANCE
# =============================================================================

def zwg_performance_view(request):
    redir = _require_data(request)
    if redir:
        return redir

    df_raw = _load_df(request, 'zwg_df')
    if df_raw is None:
        return redirect('upload')

    ctx = {'active_page': 'zwg', 'data_period': request.session.get('data_period', '')}
    fixed_cols = ['TID', 'CIF', 'MERCHANT NAME', 'BUSINESS UNIT']

    # Validate columns
    missing = [c for c in fixed_cols if c not in df_raw.columns]
    if missing:
        ctx['error'] = f'Missing required columns: {", ".join(missing)}'
        return render(request, 'merchant/zwg_performance.html', ctx)

    df = df_raw.copy()
    df = df[~df['MERCHANT NAME'].astype(str).str.upper().str.strip().isin(['TOTAL', 'GRAND TOTAL'])]
    df['CIF'] = df['CIF'].astype(str).str.strip().str.split('.').str[0].str.zfill(6)
    df = df[~df['CIF'].str.lower().str.contains('nan')]

    date_columns = extract_date_columns(df)
    if not date_columns:
        ctx['error'] = 'No transaction date columns detected.'
        return render(request, 'merchant/zwg_performance.html', ctx)

    df[date_columns] = df[date_columns].apply(_to_numeric)

    date_map, _, latest_date_obj, latest_col_name, prev_date_obj = resolve_dates(df, date_columns)
    if not date_map:
        ctx['error'] = 'Date parsing failed. Ensure headers are in a recognised date format.'
        return render(request, 'merchant/zwg_performance.html', ctx)

    # Top 15
    metrics, top15_rows = build_top15_data(df, date_map, latest_date_obj, latest_col_name, prev_date_obj)
    metrics['mtd_revenue'] = f"{compute_mtd_from_total_row(df_raw, date_columns):,.0f}"
    metrics['total_revenue'] = f"{compute_daily_from_total_row(df_raw, latest_col_name):,.0f}"
    ctx['metrics']     = metrics
    ctx['top15_rows']  = top15_rows
    ctx['display_date'] = latest_date_obj.strftime('%d %b %Y')

    # Bar chart
    if top15_rows:
        labels = [r['merchant'] for r in top15_rows]
        values = [float(r['revenue'].replace(',', '')) for r in top15_rows]
        ctx['chart_div'] = build_plotly_bar_div(
            labels, values,
            title='Top 15 Merchants by Revenue (ZWG)',
            x_label='Revenue (ZWG)',
        )

    # Build report + gainers
    report_bytes, report_sheets = build_performance_report_bytes(df_raw)
    gainers_df = _build_gainers_shakers(report_sheets, date_map, latest_col_name, threshold=50000)

    # Store in session temp dir for download
    pkl_d = _pickle_dir(request)
    (pkl_d / 'zwg_report.xlsx').write_bytes(report_bytes)
    ctx['has_zwg_report'] = True

    if not gainers_df.empty:
        gainers_bytes = _write_gainers_excel(gainers_df).getvalue()
        (pkl_d / 'zwg_gainers.xlsx').write_bytes(gainers_bytes)
        ctx['has_zwg_gainers'] = True
    else:
        ctx['has_zwg_gainers'] = False

    # Top 25 form
    top25_form = DateRangeForm(prefix='top25')
    top25_results = None
    top25_error   = None

    if request.method == 'POST' and 'top25' in request.POST.get('action', ''):
        top25_form = DateRangeForm(request.POST, prefix='top25')
        if top25_form.is_valid():
            start_d = top25_form.cleaned_data['from_date']
            end_d   = top25_form.cleaned_data['to_date']
            top_25, total_or_err = build_top25_data(df, fixed_cols, date_columns, start_d, end_d)
            if top_25 is None:
                top25_error = total_or_err
            else:
                top25_excel = build_top25_excel_bytes(top_25)
                (pkl_d / 'zwg_top25.xlsx').write_bytes(top25_excel)
                display = top_25[['CIF', 'MERCHANT NAME', 'Revenue', 'BUSINESS UNIT', 'Revenue Percentage']].copy()
                display.columns = ['CIF', 'MERCHANT_NAME', 'Revenue', 'BUSINESS_UNIT', 'Revenue_Percentage']
                top25_results = {
                    'rows': display.to_dict('records'),
                    'from_date': start_d,
                    'to_date':   end_d,
                    'has_download': True,
                }

    ctx['top25_form']    = top25_form
    ctx['top25_results'] = top25_results
    ctx['top25_error']   = top25_error
    return render(request, 'merchant/zwg_performance.html', ctx)


# =============================================================================
# USD PERFORMANCE
# =============================================================================

def usd_performance_view(request):
    redir = _require_data(request)
    if redir:
        return redir

    df_raw = _load_df(request, 'usd_df')
    if df_raw is None:
        return redirect('upload')

    ctx = {'active_page': 'usd', 'data_period': request.session.get('data_period', '')}
    fixed_cols = ['TID', 'CIF', 'MERCHANT NAME', 'BUSINESS UNIT']

    missing = [c for c in fixed_cols if c not in df_raw.columns]
    if missing:
        ctx['error'] = f'Missing required columns: {", ".join(missing)}'
        return render(request, 'merchant/usd_performance.html', ctx)

    df = df_raw.copy()
    df = df[~df['MERCHANT NAME'].astype(str).str.upper().str.strip().isin(['TOTAL', 'GRAND TOTAL'])]
    df['CIF'] = df['CIF'].astype(str).str.strip().str.split('.').str[0].str.zfill(6)
    df = df[~df['CIF'].str.lower().str.contains('nan')]

    date_columns = extract_date_columns(df)
    if not date_columns:
        ctx['error'] = 'No transaction date columns detected.'
        return render(request, 'merchant/usd_performance.html', ctx)

    df[date_columns] = df[date_columns].apply(_to_numeric)

    date_map, _, latest_date_obj, latest_col_name, prev_date_obj = resolve_dates(df, date_columns)
    if not date_map:
        ctx['error'] = 'Date parsing failed.'
        return render(request, 'merchant/usd_performance.html', ctx)

    metrics, top15_rows = build_top15_data(df, date_map, latest_date_obj, latest_col_name, prev_date_obj)
    metrics['mtd_revenue'] = f"{compute_mtd_from_total_row(df_raw, date_columns):,.0f}"
    metrics['total_revenue'] = f"{compute_daily_from_total_row(df_raw, latest_col_name):,.0f}"
    ctx['metrics']      = metrics
    ctx['top15_rows']   = top15_rows
    ctx['display_date'] = latest_date_obj.strftime('%d %b %Y')

    if top15_rows:
        labels = [r['merchant'] for r in top15_rows]
        values = [float(r['revenue'].replace(',', '')) for r in top15_rows]
        ctx['chart_div'] = build_plotly_bar_div(
            labels, values,
            title='Top 15 Merchants by Revenue (USD)',
            x_label='Revenue (USD)',
        )

    report_bytes, report_sheets = build_performance_report_bytes(df_raw)
    gainers_df = _build_gainers_shakers(report_sheets, date_map, latest_col_name, threshold=1000)

    pkl_d = _pickle_dir(request)
    (pkl_d / 'usd_report.xlsx').write_bytes(report_bytes)
    ctx['has_usd_report'] = True

    if not gainers_df.empty:
        gainers_bytes = _write_gainers_excel(gainers_df).getvalue()
        (pkl_d / 'usd_gainers.xlsx').write_bytes(gainers_bytes)
        ctx['has_usd_gainers'] = True
    else:
        ctx['has_usd_gainers'] = False

    top25_form    = DateRangeForm(prefix='top25')
    top25_results = None
    top25_error   = None

    if request.method == 'POST' and 'top25' in request.POST.get('action', ''):
        top25_form = DateRangeForm(request.POST, prefix='top25')
        if top25_form.is_valid():
            start_d = top25_form.cleaned_data['from_date']
            end_d   = top25_form.cleaned_data['to_date']
            top_25, total_or_err = build_top25_data(df, fixed_cols, date_columns, start_d, end_d)
            if top_25 is None:
                top25_error = total_or_err
            else:
                top25_excel = build_top25_excel_bytes(top_25)
                (pkl_d / 'usd_top25.xlsx').write_bytes(top25_excel)
                display = top_25[['CIF', 'MERCHANT NAME', 'Revenue', 'BUSINESS UNIT', 'Revenue Percentage']].copy()
                display.columns = ['CIF', 'MERCHANT_NAME', 'Revenue', 'BUSINESS_UNIT', 'Revenue_Percentage']
                top25_results = {
                    'rows': display.to_dict('records'),
                    'from_date': start_d,
                    'to_date':   end_d,
                    'has_download': True,
                }

    ctx['top25_form']    = top25_form
    ctx['top25_results'] = top25_results
    ctx['top25_error']   = top25_error
    return render(request, 'merchant/usd_performance.html', ctx)


# =============================================================================
# PoS STATEMENTS
# =============================================================================

def pos_statements_view(request):
    redir = _require_data(request)
    if redir:
        return redir

    ctx = {
        'active_page': 'pos',
        'data_period': request.session.get('data_period', ''),
        'form': CIFLookupForm(),
    }

    if request.method == 'POST':
        form = CIFLookupForm(request.POST)
        ctx['form'] = form
        if form.is_valid():
            cif = form.cleaned_data['cif']
            df_zwg = _load_df(request, 'zwg_df')
            df_usd = _load_df(request, 'usd_df')

            for _df in [df_zwg, df_usd]:
                _df['CIF'] = _df['CIF'].astype(str).str.strip().str.split('.').str[0].str.zfill(6)
                _df['TID'] = _df['TID'].astype(str).str.strip()

            filt_zwg = df_zwg[df_zwg['CIF'] == cif]
            filt_usd = df_usd[df_usd['CIF'] == cif]

            if filt_zwg.empty and filt_usd.empty:
                ctx['warning'] = f'No records found for Customer ID {cif}.'
                return render(request, 'merchant/pos_statements.html', ctx)

            source     = filt_zwg if not filt_zwg.empty else filt_usd
            short_name = str(source['MERCHANT NAME'].iloc[0]).split()[0]
            request.session['pos_short_name'] = short_name

            pkl_d = _pickle_dir(request)
            stmt_bytes = pos_statement_bytes(filt_zwg, filt_usd)
            (pkl_d / 'pos_statement.xlsx').write_bytes(stmt_bytes)

            ctx['cif']           = cif
            ctx['short_name']    = short_name
            ctx['zwg_records']   = len(filt_zwg)
            ctx['usd_records']   = len(filt_usd)
            ctx['file_label']    = f'{short_name} Daily Pos Statements'
            ctx['has_statement'] = True

            if cif == '012499':
                (pkl_d / 'champions_zinara.xlsx').write_bytes(champions_zinara_bytes(filt_zwg, filt_usd))
                (pkl_d / 'champions_insurance.xlsx').write_bytes(champions_insurance_bytes(filt_zwg, filt_usd))
                ctx['is_champions'] = True

    return render(request, 'merchant/pos_statements.html', ctx)


# =============================================================================
# IDLE TERMINALS  (Periodic Analytics — scans ALL sheets)
# =============================================================================

def idle_terminals_view(request):
    redir = _require_files(request)
    if redir:
        return redir

    # Lazy-load all sheets the first time this page is visited
    if not request.session.get('periodic_data_loaded'):
        err = _load_all_sheets(request)
        if err:
            return render(request, 'merchant/idle_terminals.html', {
                'error': f'Could not load data: {err}',
                'active_page': 'idle',
                'section': 'periodic',
                'form': DateRangeForm(),
                'data_period': '',
            })

    ctx = {
        'active_page': 'idle',
        'section': 'periodic',
        'data_period': request.session.get('data_period_all', request.session.get('data_period', '')),
        'form': DateRangeForm(),
    }

    if request.method == 'POST':
        form = DateRangeForm(request.POST)
        ctx['form'] = form
        if form.is_valid():
            from_date = form.cleaned_data['from_date']
            to_date   = form.cleaned_data['to_date']

            df_zwg_raw = _load_df(request, 'zwg_all_df')
            df_usd_raw = _load_df(request, 'usd_all_df')

            if df_zwg_raw is None or df_usd_raw is None:
                ctx['error'] = 'Data unavailable. Please re-upload files and try again.'
                return render(request, 'merchant/idle_terminals.html', ctx)

            for _df in [df_zwg_raw, df_usd_raw]:
                _df['CIF']           = _df['CIF'].astype(str).str.strip().str.split('.').str[0].str.zfill(6)
                _df['TID']           = _df['TID'].astype(str).str.strip()
                _df['MERCHANT NAME'] = _df['MERCHANT NAME'].astype(str).str.strip()
                if 'BUSINESS UNIT' in _df.columns:
                    _df['BUSINESS UNIT'] = _df['BUSINESS UNIT'].astype(str).str.strip()

            zwg_all_dcols = extract_date_columns(df_zwg_raw)
            usd_all_dcols = extract_date_columns(df_usd_raw)

            zwg_start = _tids_in_month(df_zwg_raw, zwg_all_dcols, from_date.year, from_date.month)
            zwg_end   = _tids_in_month(df_zwg_raw, zwg_all_dcols, to_date.year,   to_date.month)
            usd_start = _tids_in_month(df_usd_raw, usd_all_dcols, from_date.year, from_date.month)
            usd_end   = _tids_in_month(df_usd_raw, usd_all_dcols, to_date.year,   to_date.month)

            if zwg_start is None or zwg_end is None:
                ctx['error'] = (
                    'ZWG data does not cover both the start and end months of the selected period. '
                    'Please upload sheets for all months in the range.'
                )
                return render(request, 'merchant/idle_terminals.html', ctx)
            if usd_start is None or usd_end is None:
                ctx['error'] = (
                    'USD data does not cover both the start and end months of the selected period. '
                    'Please upload sheets for all months in the range.'
                )
                return render(request, 'merchant/idle_terminals.html', ctx)

            # Keep only TIDs present in BOTH the start month and the end month.
            # This excludes terminals that were removed before the end, and terminals
            # that were added after the period began.
            zwg_throughout = zwg_start & zwg_end
            usd_throughout = usd_start & usd_end

            # Aggregate by TID (collapses multi-month rows, NaN → 0 for missing dates)
            df_zwg = _aggregate_by_tid(df_zwg_raw)
            df_usd = _aggregate_by_tid(df_usd_raw)

            # Restrict to TIDs present throughout the period
            df_zwg = df_zwg[df_zwg['TID'].isin(zwg_throughout)].copy()
            df_usd = df_usd[df_usd['TID'].isin(usd_throughout)].copy()

            if df_zwg.empty or df_usd.empty:
                ctx['error'] = 'No terminals were present in both the start and end months of the selected period.'
                return render(request, 'merchant/idle_terminals.html', ctx)

            zwg_date_cols = extract_date_columns(df_zwg)
            usd_date_cols = extract_date_columns(df_usd)

            try:
                total_idle, idle_zwg, idle_usd, common_sfx = find_idle_terminals(
                    df_zwg, df_usd, zwg_date_cols, usd_date_cols, from_date, to_date
                )
            except ValueError as e:
                ctx['error'] = str(e)
                return render(request, 'merchant/idle_terminals.html', ctx)

            date_label = f"{from_date.strftime('%d %b %Y')} to {to_date.strftime('%d %b %Y')}"

            pkl_d = _pickle_dir(request)
            (pkl_d / 'idle_total.xlsx').write_bytes(idle_excel_bytes(total_idle))

            ctx['date_label']   = date_label
            ctx['total_idle']   = len(total_idle)
            ctx['has_download'] = True

            if 'BUSINESS UNIT' in total_idle.columns:
                unit_counts = (
                    total_idle.groupby('BUSINESS UNIT')['TID']
                    .count()
                    .reset_index()
                    .rename(columns={'TID': 'Idle Terminals'})
                    .sort_values('Idle Terminals', ascending=False)
                )
                ctx['chart_div'] = build_plotly_bar_div(
                    labels=unit_counts['BUSINESS UNIT'].tolist(),
                    values=unit_counts['Idle Terminals'].tolist(),
                    title=f'Idle Terminals by Business Unit — {date_label}',
                    x_label='Number of Idle Terminals',
                )
                bu_list = sorted(total_idle['BUSINESS UNIT'].dropna().astype(str).str.strip().unique().tolist())
                ctx['business_units'] = bu_list

                selected_bu = request.POST.get('selected_bu', '')
                if selected_bu and selected_bu in bu_list:
                    unit_df = total_idle[
                        total_idle['BUSINESS UNIT'].astype(str).str.strip() == selected_bu
                    ].copy()
                    (pkl_d / 'idle_unit.xlsx').write_bytes(idle_excel_bytes(unit_df))
                    ctx['selected_bu']       = selected_bu
                    ctx['has_unit_download'] = True

    return render(request, 'merchant/idle_terminals.html', ctx)


# =============================================================================
# ZWG WEEKLY TOP 25  (Periodic Analytics — scans ALL sheets)
# =============================================================================

def zwg_top25_view(request):
    redir = _require_files(request)
    if redir:
        return redir

    if not request.session.get('periodic_data_loaded'):
        err = _load_all_sheets(request)
        if err:
            return render(request, 'merchant/zwg_top25.html', {
                'error': f'Could not load data: {err}',
                'active_page': 'zwg_top25',
                'section': 'periodic',
                'form': DateRangeForm(),
            })

    ctx = {
        'active_page':   'zwg_top25',
        'section':       'periodic',
        'data_period':   request.session.get('data_period_all', request.session.get('data_period', '')),
        'form':          DateRangeForm(),
        'top25_results': None,
        'top25_error':   None,
    }

    if request.method == 'POST':
        form = DateRangeForm(request.POST)
        ctx['form'] = form
        if form.is_valid():
            start_d = form.cleaned_data['from_date']
            end_d   = form.cleaned_data['to_date']
            df_raw  = _load_df(request, 'zwg_all_df')
            if df_raw is None:
                ctx['top25_error'] = 'Data unavailable. Please re-upload files and try again.'
            else:
                df = df_raw.copy()
                df = df[~df['MERCHANT NAME'].astype(str).str.upper().str.strip().isin(['TOTAL', 'GRAND TOTAL'])]
                df['CIF'] = df['CIF'].astype(str).str.strip().str.split('.').str[0].str.zfill(6)
                fixed_cols   = ['TID', 'CIF', 'MERCHANT NAME', 'BUSINESS UNIT']
                date_columns = extract_date_columns(df)
                df[date_columns] = df[date_columns].apply(_to_numeric)

                top_25, total_or_err = build_top25_data(df, fixed_cols, date_columns, start_d, end_d)
                if top_25 is None:
                    ctx['top25_error'] = total_or_err
                else:
                    pkl_d = _pickle_dir(request)
                    (pkl_d / 'zwg_top25_weekly.xlsx').write_bytes(build_top25_excel_bytes(top_25))
                    display = top_25[['CIF', 'MERCHANT NAME', 'Revenue', 'BUSINESS UNIT', 'Revenue Percentage']].copy()
                    display.columns = ['CIF', 'MERCHANT_NAME', 'Revenue', 'BUSINESS_UNIT', 'Revenue_Percentage']
                    ctx['top25_results'] = {
                        'rows':        display.to_dict('records'),
                        'from_date':   start_d,
                        'to_date':     end_d,
                        'has_download': True,
                    }

    return render(request, 'merchant/zwg_top25.html', ctx)


# =============================================================================
# USD WEEKLY TOP 25  (Periodic Analytics — scans ALL sheets)
# =============================================================================

def usd_top25_view(request):
    redir = _require_files(request)
    if redir:
        return redir

    if not request.session.get('periodic_data_loaded'):
        err = _load_all_sheets(request)
        if err:
            return render(request, 'merchant/usd_top25.html', {
                'error': f'Could not load data: {err}',
                'active_page': 'usd_top25',
                'section': 'periodic',
                'form': DateRangeForm(),
            })

    ctx = {
        'active_page':   'usd_top25',
        'section':       'periodic',
        'data_period':   request.session.get('data_period_all', request.session.get('data_period', '')),
        'form':          DateRangeForm(),
        'top25_results': None,
        'top25_error':   None,
    }

    if request.method == 'POST':
        form = DateRangeForm(request.POST)
        ctx['form'] = form
        if form.is_valid():
            start_d = form.cleaned_data['from_date']
            end_d   = form.cleaned_data['to_date']
            df_raw  = _load_df(request, 'usd_all_df')
            if df_raw is None:
                ctx['top25_error'] = 'Data unavailable. Please re-upload files and try again.'
            else:
                df = df_raw.copy()
                df = df[~df['MERCHANT NAME'].astype(str).str.upper().str.strip().isin(['TOTAL', 'GRAND TOTAL'])]
                df['CIF'] = df['CIF'].astype(str).str.strip().str.split('.').str[0].str.zfill(6)
                fixed_cols   = ['TID', 'CIF', 'MERCHANT NAME', 'BUSINESS UNIT']
                date_columns = extract_date_columns(df)
                df[date_columns] = df[date_columns].apply(_to_numeric)

                top_25, total_or_err = build_top25_data(df, fixed_cols, date_columns, start_d, end_d)
                if top_25 is None:
                    ctx['top25_error'] = total_or_err
                else:
                    pkl_d = _pickle_dir(request)
                    (pkl_d / 'usd_top25_weekly.xlsx').write_bytes(build_top25_excel_bytes(top_25))
                    display = top_25[['CIF', 'MERCHANT NAME', 'Revenue', 'BUSINESS UNIT', 'Revenue Percentage']].copy()
                    display.columns = ['CIF', 'MERCHANT_NAME', 'Revenue', 'BUSINESS_UNIT', 'Revenue_Percentage']
                    ctx['top25_results'] = {
                        'rows':        display.to_dict('records'),
                        'from_date':   start_d,
                        'to_date':     end_d,
                        'has_download': True,
                    }

    return render(request, 'merchant/usd_top25.html', ctx)


# =============================================================================
# DETAILED MERCHANT PERFORMANCE  (Periodic Analytics — scans ALL sheets)
# =============================================================================

def merchant_performance_view(request):
    redir = _require_files(request)
    if redir:
        return redir

    if not request.session.get('periodic_data_loaded'):
        err = _load_all_sheets(request)
        if err:
            return render(request, 'merchant/merchant_performance.html', {
                'error': f'Could not load data: {err}',
                'active_page': 'merchant_perf',
                'section': 'periodic',
                'cif_form': CIFLookupForm(),
                'data_period': '',
            })

    ctx = {
        'active_page': 'merchant_perf',
        'section': 'periodic',
        'data_period': request.session.get('data_period_all', request.session.get('data_period', '')),
        'cif_form': CIFLookupForm(),
    }

    if request.method != 'POST':
        return render(request, 'merchant/merchant_performance.html', ctx)

    step = request.POST.get('step', '')

    # ── STEP 1: Validate CIF ──────────────────────────────────────────────────
    if step == 'cif':
        cif_form = CIFLookupForm(request.POST)
        ctx['cif_form'] = cif_form
        if not cif_form.is_valid():
            return render(request, 'merchant/merchant_performance.html', ctx)

        cif = cif_form.cleaned_data['cif'].strip().zfill(6)
        df_zwg = _load_df(request, 'zwg_all_df')
        df_usd = _load_df(request, 'usd_all_df')
        if df_zwg is None or df_usd is None:
            cif_form.add_error(None, 'Data unavailable. Please re-upload files.')
            return render(request, 'merchant/merchant_performance.html', ctx)

        df_zwg = df_zwg.copy()
        df_usd = df_usd.copy()
        df_zwg['CIF'] = df_zwg['CIF'].astype(str).str.strip().str.split('.').str[0].str.zfill(6)
        df_usd['CIF'] = df_usd['CIF'].astype(str).str.strip().str.split('.').str[0].str.zfill(6)

        fz = df_zwg[df_zwg['CIF'] == cif]
        fu = df_usd[df_usd['CIF'] == cif]
        if fz.empty and fu.empty:
            cif_form.add_error('cif', f'No records found for Customer ID {cif}.')
            return render(request, 'merchant/merchant_performance.html', ctx)

        src = fz if not fz.empty else fu
        merchant_name = str(src['MERCHANT NAME'].iloc[0]).strip()
        business_unit = str(src['BUSINESS UNIT'].iloc[0]).strip() if 'BUSINESS UNIT' in src.columns else '—'
        ctx.update({
            'confirmed_cif':      cif,
            'merchant_name':      merchant_name,
            'merchant_short_name': merchant_name.split()[0] if merchant_name else cif,
            'business_unit':      business_unit,
            'date_form':          DateRangeForm(),
        })
        return render(request, 'merchant/merchant_performance.html', ctx)

    # ── STEP 2: Generate Dashboard ────────────────────────────────────────────
    if step == 'report':
        cif = request.POST.get('cif_value', '').strip().zfill(6)
        date_form = DateRangeForm(request.POST)
        ctx['date_form']     = date_form
        ctx['confirmed_cif'] = cif

        df_zwg_raw = _load_df(request, 'zwg_all_df')
        df_usd_raw = _load_df(request, 'usd_all_df')
        if df_zwg_raw is None or df_usd_raw is None:
            ctx['error'] = 'Data unavailable. Please re-upload files and try again.'
            return render(request, 'merchant/merchant_performance.html', ctx)

        df_zwg_raw = df_zwg_raw.copy()
        df_usd_raw = df_usd_raw.copy()
        for _df in [df_zwg_raw, df_usd_raw]:
            _df['CIF']           = _df['CIF'].astype(str).str.strip().str.split('.').str[0].str.zfill(6)
            _df['TID']           = _df['TID'].astype(str).str.strip()
            _df['MERCHANT NAME'] = _df['MERCHANT NAME'].astype(str).str.strip()
            if 'BUSINESS UNIT' in _df.columns:
                _df['BUSINESS UNIT'] = _df['BUSINESS UNIT'].astype(str).str.strip()

        fz_all = df_zwg_raw[df_zwg_raw['CIF'] == cif]
        fu_all = df_usd_raw[df_usd_raw['CIF'] == cif]
        if fz_all.empty and fu_all.empty:
            ctx['error'] = f'No records found for CIF {cif}. Please go back and verify.'
            return render(request, 'merchant/merchant_performance.html', ctx)

        src = fz_all if not fz_all.empty else fu_all
        merchant_name = str(src['MERCHANT NAME'].iloc[0]).strip()
        business_unit = str(src['BUSINESS UNIT'].iloc[0]).strip() if 'BUSINESS UNIT' in src.columns else '—'
        ctx['merchant_name']       = merchant_name
        ctx['merchant_short_name'] = merchant_name.split()[0] if merchant_name else cif
        ctx['business_unit']       = business_unit

        if not date_form.is_valid():
            return render(request, 'merchant/merchant_performance.html', ctx)

        from_date  = date_form.cleaned_data['from_date']
        to_date    = date_form.cleaned_data['to_date']
        date_label = f"{from_date.strftime('%d %b %Y')} to {to_date.strftime('%d %b %Y')}"
        ctx.update({'date_label': date_label, 'from_date': from_date, 'to_date': to_date})

        # ── Date columns ─────────────────────────────────────────────────────
        zwg_all_dcols = extract_date_columns(df_zwg_raw)
        usd_all_dcols = extract_date_columns(df_usd_raw)

        def _period_cols(all_cols, from_d, to_d):
            return [
                c for c in all_cols
                if pd.notnull(_parse_col_date(c)) and from_d <= _parse_col_date(c).date() <= to_d
            ]

        zwg_pcols = _period_cols(zwg_all_dcols, from_date, to_date)
        usd_pcols = _period_cols(usd_all_dcols, from_date, to_date)

        if not zwg_pcols and not usd_pcols:
            ctx['error'] = f'No transaction date columns found in the selected period ({date_label}).'
            return render(request, 'merchant/merchant_performance.html', ctx)

        # ── Revenue for the period ────────────────────────────────────────────
        def _rev(frame, pcols):
            if frame.empty or not pcols:
                return 0.0
            avail = [c for c in pcols if c in frame.columns]
            return float(frame[avail].apply(_to_numeric).sum().sum()) if avail else 0.0

        zwg_rev = _rev(fz_all, zwg_pcols)
        usd_rev = _rev(fu_all, usd_pcols)

        # ── Total terminals (all-time unique TID suffixes for this CIF) ──────
        zwg_tids = set(fz_all['TID'].dropna().astype(str).str.strip().unique())
        usd_tids = set(fu_all['TID'].dropna().astype(str).str.strip().unique())
        all_sfx  = {t[-4:] for t in (zwg_tids | usd_tids) if len(t) >= 4}
        total_terminals = len(all_sfx)

        # ── Inactive terminals (same mechanism as Idle Terminals) ─────────────
        inactive_count = 0
        idle_df = pd.DataFrame()
        if not fz_all.empty and not fu_all.empty:
            fz_dcols = extract_date_columns(fz_all)
            fu_dcols = extract_date_columns(fu_all)
            zs = _tids_in_month(fz_all, fz_dcols, from_date.year, from_date.month)
            ze = _tids_in_month(fz_all, fz_dcols, to_date.year,   to_date.month)
            us = _tids_in_month(fu_all, fu_dcols, from_date.year, from_date.month)
            ue = _tids_in_month(fu_all, fu_dcols, to_date.year,   to_date.month)
            if zs and ze and us and ue:
                fz_scope = _aggregate_by_tid(fz_all[fz_all['TID'].isin(zs & ze)].copy())
                fu_scope = _aggregate_by_tid(fu_all[fu_all['TID'].isin(us & ue)].copy())
                if not fz_scope.empty and not fu_scope.empty:
                    try:
                        idle_total, _, _, _ = find_idle_terminals(
                            fz_scope, fu_scope,
                            extract_date_columns(fz_scope), extract_date_columns(fu_scope),
                            from_date, to_date,
                        )
                        idle_df       = idle_total
                        inactive_count = len(idle_total)
                    except ValueError:
                        pass

        active_terminals = total_terminals - inactive_count
        activity_ratio   = round(active_terminals / total_terminals * 100, 1) if total_terminals > 0 else 0.0

        ctx.update({
            'zwg_rev':          f'{zwg_rev:,.0f}',
            'usd_rev':          f'{usd_rev:,.2f}',
            'total_terminals':  total_terminals,
            'inactive_count':   inactive_count,
            'active_terminals': active_terminals,
            'activity_ratio':   f'{activity_ratio:.1f}%',
            'activity_ratio_ok': activity_ratio >= 70,
        })

        # ── Daily revenue data for chart ──────────────────────────────────────
        daily_zwg, daily_usd = {}, {}
        for col in zwg_pcols:
            d = _parse_col_date(col)
            if pd.notnull(d) and col in fz_all.columns:
                daily_zwg[d.date()] = float(_to_numeric(fz_all[col]).sum())
        for col in usd_pcols:
            d = _parse_col_date(col)
            if pd.notnull(d) and col in fu_all.columns:
                daily_usd[d.date()] = float(_to_numeric(fu_all[col]).sum())

        if daily_zwg or daily_usd:
            ctx['daily_chart_div'] = build_merchant_revenue_chart(
                daily_zwg, daily_usd, merchant_name, date_label
            )

        # Peak days
        if daily_zwg:
            pk = max(daily_zwg, key=daily_zwg.get)
            ctx['peak_zwg'] = {'date': pk.strftime('%d %b %Y'), 'amount': f"{daily_zwg[pk]:,.0f}"}
        if daily_usd:
            pk = max(daily_usd, key=daily_usd.get)
            ctx['peak_usd'] = {'date': pk.strftime('%d %b %Y'), 'amount': f"{daily_usd[pk]:,.2f}"}

        # ── Top 10 terminals by ZWG revenue ───────────────────────────────────
        if zwg_pcols and not fz_all.empty:
            fz_agg = _aggregate_by_tid(fz_all.copy())
            p_avail = [c for c in zwg_pcols if c in fz_agg.columns]
            if p_avail:
                fz_agg['_rev'] = fz_agg[p_avail].apply(_to_numeric).sum(axis=1)
                ctx['top_terminals'] = [
                    {'tid': row['TID'], 'zwg_rev': f"{row['_rev']:,.0f}"}
                    for _, row in fz_agg.nlargest(10, '_rev').iterrows()
                    if row['_rev'] > 0
                ]

        # ── Save downloadable files ───────────────────────────────────────────
        pkl_d = _pickle_dir(request)
        perf_bytes = merchant_period_excel_bytes(fz_all, fu_all, zwg_pcols, usd_pcols)
        (pkl_d / 'merchant_perf.xlsx').write_bytes(perf_bytes)
        ctx['has_perf_report'] = True

        if not idle_df.empty:
            (pkl_d / 'merchant_idle.xlsx').write_bytes(idle_excel_bytes(idle_df))
            ctx['has_idle_report'] = True

        short_name = merchant_name.split()[0]
        request.session['merch_perf_label'] = f"{merchant_name} Performance Report for {date_label}"
        request.session['merch_idle_label'] = f"Total Idle Terminals for {short_name} from {date_label}"
        ctx['has_dashboard'] = True

    return render(request, 'merchant/merchant_performance.html', ctx)


def download_merchant_perf_report(request):
    label = request.session.get('merch_perf_label', 'Merchant Performance Report')
    return _serve_file(request, 'merchant_perf.xlsx', f'{label}.xlsx')


def download_merchant_idle(request):
    label = request.session.get('merch_idle_label', 'Merchant Idle Terminals')
    return _serve_file(request, 'merchant_idle.xlsx', f'{label}.xlsx')


# =============================================================================
# BUSINESS UNIT PERFORMANCE  (Periodic Analytics — scans ALL sheets)
# =============================================================================

def bu_performance_view(request):
    redir = _require_files(request)
    if redir:
        return redir

    if not request.session.get('periodic_data_loaded'):
        err = _load_all_sheets(request)
        if err:
            return render(request, 'merchant/bu_performance.html', {
                'error': f'Could not load data: {err}',
                'active_page': 'bu_perf',
                'section': 'periodic',
                'data_period': '',
                'bu_list': [],
            })

    df_zwg_raw = _load_df(request, 'zwg_all_df')
    df_usd_raw = _load_df(request, 'usd_all_df')

    # Collect unique Business Units from both files
    bu_set = set()
    for _df in [df_zwg_raw, df_usd_raw]:
        if _df is not None and 'BUSINESS UNIT' in _df.columns:
            bu_set.update(
                _df['BUSINESS UNIT'].dropna().astype(str).str.strip().unique().tolist()
            )
    bu_list = sorted(b for b in bu_set if b and b.lower() not in ('nan', ''))

    ctx = {
        'active_page': 'bu_perf',
        'section': 'periodic',
        'data_period': request.session.get('data_period_all', request.session.get('data_period', '')),
        'bu_list': bu_list,
    }

    if request.method != 'POST':
        return render(request, 'merchant/bu_performance.html', ctx)

    step = request.POST.get('step', '')

    # ── STEP 1: Confirm Business Unit ─────────────────────────────────────────
    if step == 'bu':
        selected_bu = request.POST.get('business_unit', '').strip()
        if not selected_bu or selected_bu not in bu_list:
            ctx['bu_error'] = 'Please select a valid Business Unit from the list.'
            return render(request, 'merchant/bu_performance.html', ctx)
        ctx.update({
            'confirmed_bu': selected_bu,
            'date_form': DateRangeForm(),
        })
        return render(request, 'merchant/bu_performance.html', ctx)

    # ── STEP 2: Generate Dashboard ────────────────────────────────────────────
    if step == 'report':
        selected_bu = request.POST.get('bu_value', '').strip()
        date_form = DateRangeForm(request.POST)
        ctx['date_form'] = date_form
        ctx['confirmed_bu'] = selected_bu

        if df_zwg_raw is None or df_usd_raw is None:
            ctx['error'] = 'Data unavailable. Please re-upload files and try again.'
            return render(request, 'merchant/bu_performance.html', ctx)

        df_zwg_raw = df_zwg_raw.copy()
        df_usd_raw = df_usd_raw.copy()
        for _df in [df_zwg_raw, df_usd_raw]:
            _df['CIF']           = _df['CIF'].astype(str).str.strip().str.split('.').str[0].str.zfill(6)
            _df['TID']           = _df['TID'].astype(str).str.strip()
            _df['MERCHANT NAME'] = _df['MERCHANT NAME'].astype(str).str.strip()
            if 'BUSINESS UNIT' in _df.columns:
                _df['BUSINESS UNIT'] = _df['BUSINESS UNIT'].astype(str).str.strip()

        # Filter by Business Unit
        if 'BUSINESS UNIT' in df_zwg_raw.columns:
            fz_all = df_zwg_raw[df_zwg_raw['BUSINESS UNIT'] == selected_bu].copy()
        else:
            fz_all = pd.DataFrame()
        if 'BUSINESS UNIT' in df_usd_raw.columns:
            fu_all = df_usd_raw[df_usd_raw['BUSINESS UNIT'] == selected_bu].copy()
        else:
            fu_all = pd.DataFrame()

        if fz_all.empty and fu_all.empty:
            ctx['error'] = f'No records found for Business Unit "{selected_bu}". Please go back and verify.'
            return render(request, 'merchant/bu_performance.html', ctx)

        if not date_form.is_valid():
            return render(request, 'merchant/bu_performance.html', ctx)

        from_date  = date_form.cleaned_data['from_date']
        to_date    = date_form.cleaned_data['to_date']
        date_label = f"{from_date.strftime('%d %b %Y')} to {to_date.strftime('%d %b %Y')}"
        ctx.update({'date_label': date_label, 'from_date': from_date, 'to_date': to_date})

        # ── Date columns ──────────────────────────────────────────────────────
        zwg_all_dcols = extract_date_columns(df_zwg_raw)
        usd_all_dcols = extract_date_columns(df_usd_raw)

        def _period_cols(all_cols, from_d, to_d):
            return [
                c for c in all_cols
                if pd.notnull(_parse_col_date(c)) and from_d <= _parse_col_date(c).date() <= to_d
            ]

        zwg_pcols = _period_cols(zwg_all_dcols, from_date, to_date)
        usd_pcols = _period_cols(usd_all_dcols, from_date, to_date)

        if not zwg_pcols and not usd_pcols:
            ctx['error'] = f'No transaction date columns found in the selected period ({date_label}).'
            return render(request, 'merchant/bu_performance.html', ctx)

        # ── Revenue for the period ────────────────────────────────────────────
        def _rev(frame, pcols):
            if frame.empty or not pcols:
                return 0.0
            avail = [c for c in pcols if c in frame.columns]
            return float(frame[avail].apply(_to_numeric).sum().sum()) if avail else 0.0

        zwg_rev = _rev(fz_all, zwg_pcols)
        usd_rev = _rev(fu_all, usd_pcols)

        # ── Total terminals (unique TIDs in this BU across both files) ────────
        zwg_tids = set(fz_all['TID'].dropna().astype(str).str.strip().unique()) if not fz_all.empty else set()
        usd_tids = set(fu_all['TID'].dropna().astype(str).str.strip().unique()) if not fu_all.empty else set()
        total_terminals = len(zwg_tids | usd_tids)

        # ── Inactive terminals (same mechanism as Idle Terminals) ─────────────
        inactive_count = 0
        idle_df = pd.DataFrame()
        if not fz_all.empty and not fu_all.empty:
            fz_dcols = extract_date_columns(fz_all)
            fu_dcols = extract_date_columns(fu_all)
            zs = _tids_in_month(fz_all, fz_dcols, from_date.year, from_date.month)
            ze = _tids_in_month(fz_all, fz_dcols, to_date.year,   to_date.month)
            us = _tids_in_month(fu_all, fu_dcols, from_date.year, from_date.month)
            ue = _tids_in_month(fu_all, fu_dcols, to_date.year,   to_date.month)
            if zs and ze and us and ue:
                fz_scope = _aggregate_by_tid(fz_all[fz_all['TID'].isin(zs & ze)].copy())
                fu_scope = _aggregate_by_tid(fu_all[fu_all['TID'].isin(us & ue)].copy())
                if not fz_scope.empty and not fu_scope.empty:
                    try:
                        idle_total, _, _, _ = find_idle_terminals(
                            fz_scope, fu_scope,
                            extract_date_columns(fz_scope), extract_date_columns(fu_scope),
                            from_date, to_date,
                        )
                        idle_df       = idle_total
                        inactive_count = len(idle_total)
                    except ValueError:
                        pass

        active_terminals = total_terminals - inactive_count
        activity_ratio   = round(active_terminals / total_terminals * 100, 1) if total_terminals > 0 else 0.0

        ctx.update({
            'zwg_rev':          f'{zwg_rev:,.0f}',
            'usd_rev':          f'{usd_rev:,.2f}',
            'total_terminals':  total_terminals,
            'inactive_count':   inactive_count,
            'active_terminals': active_terminals,
            'activity_ratio':   f'{activity_ratio:.1f}%',
            'activity_ratio_ok': activity_ratio >= 70,
        })

        # ── Daily revenue data for chart ──────────────────────────────────────
        daily_zwg, daily_usd = {}, {}
        for col in zwg_pcols:
            d = _parse_col_date(col)
            if pd.notnull(d) and col in fz_all.columns:
                daily_zwg[d.date()] = float(_to_numeric(fz_all[col]).sum())
        for col in usd_pcols:
            d = _parse_col_date(col)
            if pd.notnull(d) and col in fu_all.columns:
                daily_usd[d.date()] = float(_to_numeric(fu_all[col]).sum())

        if daily_zwg or daily_usd:
            ctx['daily_chart_div'] = build_merchant_revenue_chart(
                daily_zwg, daily_usd, selected_bu, date_label
            )

        # Peak days
        if daily_zwg:
            pk = max(daily_zwg, key=daily_zwg.get)
            ctx['peak_zwg'] = {'date': pk.strftime('%d %b %Y'), 'amount': f"{daily_zwg[pk]:,.0f}"}
        if daily_usd:
            pk = max(daily_usd, key=daily_usd.get)
            ctx['peak_usd'] = {'date': pk.strftime('%d %b %Y'), 'amount': f"{daily_usd[pk]:,.2f}"}

        # ── Top 10 merchants by ZWG revenue in this BU (grouped by CIF) ─────
        if zwg_pcols and not fz_all.empty:
            p_avail = [c for c in zwg_pcols if c in fz_all.columns]
            if p_avail:
                fz_cif = fz_all.copy()
                for col in p_avail:
                    fz_cif[col] = _to_numeric(fz_cif[col])
                fz_cif['_rev'] = fz_cif[p_avail].sum(axis=1)
                cif_grp = fz_cif.groupby('CIF', as_index=False).agg(
                    _rev=('_rev', 'sum'),
                    merchant_name=('MERCHANT NAME', 'first'),
                )
                ctx['top_merchants'] = [
                    {
                        'cif':        row['CIF'],
                        'short_name': str(row['merchant_name']).strip().split()[0]
                                      if str(row['merchant_name']).strip() else row['CIF'],
                        'zwg_rev':    f"{row['_rev']:,.0f}",
                    }
                    for _, row in cif_grp.nlargest(10, '_rev').iterrows()
                    if row['_rev'] > 0
                ]

        # ── Top 10 merchants by USD revenue in this BU (grouped by CIF) ─────
        if usd_pcols and not fu_all.empty:
            p_avail_u = [c for c in usd_pcols if c in fu_all.columns]
            if p_avail_u:
                fu_cif = fu_all.copy()
                for col in p_avail_u:
                    fu_cif[col] = _to_numeric(fu_cif[col])
                fu_cif['_rev'] = fu_cif[p_avail_u].sum(axis=1)
                cif_grp_u = fu_cif.groupby('CIF', as_index=False).agg(
                    _rev=('_rev', 'sum'),
                    merchant_name=('MERCHANT NAME', 'first'),
                )
                ctx['top_merchants_usd'] = [
                    {
                        'cif':        row['CIF'],
                        'short_name': str(row['merchant_name']).strip().split()[0]
                                      if str(row['merchant_name']).strip() else row['CIF'],
                        'usd_rev':    f"{row['_rev']:,.2f}",
                    }
                    for _, row in cif_grp_u.nlargest(10, '_rev').iterrows()
                    if row['_rev'] > 0
                ]

        # ── Save downloadable files ───────────────────────────────────────────
        pkl_d = _pickle_dir(request)
        perf_bytes = merchant_period_excel_bytes(fz_all, fu_all, zwg_pcols, usd_pcols)
        (pkl_d / 'bu_perf.xlsx').write_bytes(perf_bytes)
        ctx['has_perf_report'] = True

        if not idle_df.empty:
            (pkl_d / 'bu_idle.xlsx').write_bytes(idle_excel_bytes(idle_df))
            ctx['has_idle_report'] = True

        request.session['bu_perf_label'] = f"{selected_bu} Business Unit Performance from {date_label}"
        request.session['bu_idle_label'] = f"Idle Terminals for {selected_bu} from {date_label}"
        ctx['has_dashboard'] = True

    return render(request, 'merchant/bu_performance.html', ctx)


def download_bu_perf_report(request):
    label = request.session.get('bu_perf_label', 'Business Unit Performance Report')
    return _serve_file(request, 'bu_perf.xlsx', f'{label}.xlsx')


def download_bu_idle(request):
    label = request.session.get('bu_idle_label', 'Business Unit Idle Terminals')
    return _serve_file(request, 'bu_idle.xlsx', f'{label}.xlsx')


# =============================================================================
# PERFORMANCE BY SECTOR  (Periodic Analytics — scans ALL sheets)
# =============================================================================

def _sector_combined_excel_bytes(top10_zwg, top10_usd):
    """Two-sheet Excel: 'zwg' for ZWG top-10 sectors, 'usd' for USD top-10 sectors."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        if top10_zwg is not None and not top10_zwg.empty:
            export = top10_zwg[['SECTOR', '_rev']].copy()
            export.columns = ['SECTOR', 'ZWG REVENUE']
            export.to_excel(writer, index=False, sheet_name='zwg')
        if top10_usd is not None and not top10_usd.empty:
            export = top10_usd[['SECTOR', '_rev']].copy()
            export.columns = ['SECTOR', 'USD REVENUE']
            export.to_excel(writer, index=False, sheet_name='usd')
    return buf.getvalue()


def _sector_gainers_shakers(curr_grp, prev_grp, threshold):
    """Return gainers/shakers for sectors: rows where |variance| >= threshold."""
    if curr_grp is None or curr_grp.empty:
        return pd.DataFrame(columns=['SECTOR', 'PREVIOUSLY', 'CURRENTLY', 'VARIANCE'])
    curr = curr_grp[['SECTOR', '_rev']].rename(columns={'_rev': 'CURRENTLY'}).copy()
    if prev_grp is not None and not prev_grp.empty:
        prev = prev_grp[['SECTOR', '_rev']].rename(columns={'_rev': 'PREVIOUSLY'})
        merged = curr.merge(prev, on='SECTOR', how='outer')
    else:
        merged = curr.copy()
        merged['PREVIOUSLY'] = 0.0
    merged['CURRENTLY'] = merged['CURRENTLY'].fillna(0.0)
    merged['PREVIOUSLY'] = merged['PREVIOUSLY'].fillna(0.0)
    merged['VARIANCE'] = merged['CURRENTLY'] - merged['PREVIOUSLY']
    result = merged[merged['VARIANCE'].abs() >= threshold].copy()
    return result.sort_values('VARIANCE', ascending=False).reset_index(drop=True)[
        ['SECTOR', 'PREVIOUSLY', 'CURRENTLY', 'VARIANCE']
    ]


def _idle_sectors(df, pcols):
    """Return sectors that exist in the data but earned zero revenue in pcols."""
    if 'SECTOR' not in df.columns:
        return pd.DataFrame(columns=['SECTOR'])
    all_secs = {
        s for s in df['SECTOR'].dropna().astype(str).str.strip().unique()
        if s.upper() not in ('NAN', '')
    }
    p_avail = [c for c in pcols if c in df.columns]
    if not p_avail:
        return pd.DataFrame({'SECTOR': sorted(all_secs)})
    d = df.copy()
    for col in p_avail:
        d[col] = _to_numeric(d[col])
    d['_rev'] = d[p_avail].sum(axis=1)
    grp = d.groupby('SECTOR', as_index=False)['_rev'].sum()
    active = {
        str(s).strip() for s in grp.loc[grp['_rev'] > 0, 'SECTOR']
    }
    idle = all_secs - active
    return pd.DataFrame({'SECTOR': sorted(idle)})


def _sector_detailed_excel_bytes(
    zwg_gs, usd_gs, zwg_idle, usd_idle,
    zwg_full, usd_full, zwg_prev, usd_prev,
    from_date, to_date, prior_from, prior_to,
):
    """4-sheet Excel: Gainers and Shakers, Idle Sectors, All Sectors Comparison, Summary."""
    buf = io.BytesIO()
    acct_fmt = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'

    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        # ── Sheet 1: Gainers and Shakers ──────────────────────────────────────
        gs_rows = (
            [{'SECTOR': r['SECTOR'], 'CURRENCY': 'ZWG', 'PREVIOUSLY': r['PREVIOUSLY'],
              'CURRENTLY': r['CURRENTLY'], 'VARIANCE': r['VARIANCE']} for _, r in zwg_gs.iterrows()]
            + [{'SECTOR': r['SECTOR'], 'CURRENCY': 'USD', 'PREVIOUSLY': r['PREVIOUSLY'],
                'CURRENTLY': r['CURRENTLY'], 'VARIANCE': r['VARIANCE']} for _, r in usd_gs.iterrows()]
        )
        gs_df = (pd.DataFrame(gs_rows) if gs_rows
                 else pd.DataFrame(columns=['SECTOR', 'CURRENCY', 'PREVIOUSLY', 'CURRENTLY', 'VARIANCE']))
        gs_df.to_excel(writer, index=False, sheet_name='Gainers and Shakers')
        ws_gs = writer.sheets['Gainers and Shakers']
        for row in ws_gs.iter_rows(min_row=2, min_col=3, max_col=5):
            for cell in row:
                cell.number_format = acct_fmt

        # ── Sheet 2: Idle Sectors ──────────────────────────────────────────────
        idle_rows = (
            [{'SECTOR': r['SECTOR'], 'CURRENCY': 'ZWG'} for _, r in zwg_idle.iterrows()]
            + [{'SECTOR': r['SECTOR'], 'CURRENCY': 'USD'} for _, r in usd_idle.iterrows()]
        )
        idle_df = (pd.DataFrame(idle_rows) if idle_rows
                   else pd.DataFrame(columns=['SECTOR', 'CURRENCY']))
        idle_df.to_excel(writer, index=False, sheet_name='Idle Sectors')

        # ── Sheet 3: All Sectors Comparison (current vs prior period) ─────────
        comp_parts = []
        for grp_curr, grp_prev_data, ccy in (
            (zwg_full, zwg_prev, 'ZWG'),
            (usd_full, usd_prev, 'USD'),
        ):
            if grp_curr is None or grp_curr.empty:
                continue
            comp = grp_curr[['SECTOR', '_rev']].rename(columns={'_rev': 'CURRENTLY'}).copy()
            if grp_prev_data is not None and not grp_prev_data.empty:
                comp = comp.merge(
                    grp_prev_data[['SECTOR', '_rev']].rename(columns={'_rev': 'PREVIOUSLY'}),
                    on='SECTOR', how='outer',
                )
            else:
                comp['PREVIOUSLY'] = 0.0
            comp['CURRENTLY'] = comp['CURRENTLY'].fillna(0.0)
            comp['PREVIOUSLY'] = comp['PREVIOUSLY'].fillna(0.0)
            comp['VARIANCE'] = comp['CURRENTLY'] - comp['PREVIOUSLY']
            comp['CURRENCY'] = ccy
            comp_parts.append(comp[['SECTOR', 'CURRENCY', 'PREVIOUSLY', 'CURRENTLY', 'VARIANCE']])
        if comp_parts:
            all_comp = pd.concat(comp_parts, ignore_index=True)
            all_comp.sort_values(['CURRENCY', 'CURRENTLY'], ascending=[True, False], inplace=True)
            all_comp.to_excel(writer, index=False, sheet_name='All Sectors Comparison')
            ws_comp = writer.sheets['All Sectors Comparison']
            for row in ws_comp.iter_rows(min_row=2, min_col=3, max_col=5):
                for cell in row:
                    cell.number_format = acct_fmt

        # ── Sheet 4: Summary ───────────────────────────────────────────────────
        zwg_ct = zwg_full['_rev'].sum() if zwg_full is not None and not zwg_full.empty else 0
        usd_ct = usd_full['_rev'].sum() if usd_full is not None and not usd_full.empty else 0
        zwg_pt = zwg_prev['_rev'].sum() if zwg_prev is not None and not zwg_prev.empty else 0
        usd_pt = usd_prev['_rev'].sum() if usd_prev is not None and not usd_prev.empty else 0

        def _cnt(df, cond):
            return len(df[df['VARIANCE'].apply(cond)]) if df is not None and not df.empty else 0

        summary = pd.DataFrame({
            'Metric': [
                'Report Period',
                'Comparison (Prior) Period',
                'ZWG Revenue — Current Period',
                'ZWG Revenue — Prior Period',
                'ZWG Net Variance',
                'USD Revenue — Current Period',
                'USD Revenue — Prior Period',
                'USD Net Variance',
                'ZWG Active Sectors (Current)',
                'ZWG Idle Sectors',
                'USD Active Sectors (Current)',
                'USD Idle Sectors',
                'ZWG Gainers (variance >= +1,100,000)',
                'ZWG Shakers (variance <= -1,100,000)',
                'USD Gainers (variance >= +10,000)',
                'USD Shakers (variance <= -10,000)',
            ],
            'Value': [
                f"{from_date.strftime('%d %b %Y')} to {to_date.strftime('%d %b %Y')}",
                f"{prior_from.strftime('%d %b %Y')} to {prior_to.strftime('%d %b %Y')}",
                f"{zwg_ct:,.0f}",
                f"{zwg_pt:,.0f}",
                f"{zwg_ct - zwg_pt:,.0f}",
                f"{usd_ct:,.2f}",
                f"{usd_pt:,.2f}",
                f"{usd_ct - usd_pt:,.2f}",
                len(zwg_full) if zwg_full is not None and not zwg_full.empty else 0,
                len(zwg_idle),
                len(usd_full) if usd_full is not None and not usd_full.empty else 0,
                len(usd_idle),
                _cnt(zwg_gs, lambda v: v > 0),
                _cnt(zwg_gs, lambda v: v < 0),
                _cnt(usd_gs, lambda v: v > 0),
                _cnt(usd_gs, lambda v: v < 0),
            ],
        })
        summary.to_excel(writer, index=False, sheet_name='Summary')

    return buf.getvalue()


def sector_performance_view(request):
    redir = _require_files(request)
    if redir:
        return redir

    if not request.session.get('periodic_data_loaded'):
        err = _load_all_sheets(request)
        if err:
            return render(request, 'merchant/sector_performance.html', {
                'error': f'Could not load data: {err}',
                'active_page': 'sector_perf',
                'section': 'periodic',
                'form': DateRangeForm(),
                'data_period': '',
            })

    ctx = {
        'active_page': 'sector_perf',
        'section': 'periodic',
        'data_period': request.session.get('data_period_all', request.session.get('data_period', '')),
        'form': DateRangeForm(),
    }

    if request.method != 'POST':
        return render(request, 'merchant/sector_performance.html', ctx)

    form = DateRangeForm(request.POST)
    ctx['form'] = form
    if not form.is_valid():
        return render(request, 'merchant/sector_performance.html', ctx)

    from_date  = form.cleaned_data['from_date']
    to_date    = form.cleaned_data['to_date']
    date_label = f"{from_date.strftime('%d %b %Y')} to {to_date.strftime('%d %b %Y')}"
    ctx.update({'date_label': date_label, 'from_date': from_date, 'to_date': to_date})

    df_zwg_raw = _load_df(request, 'zwg_all_df')
    df_usd_raw = _load_df(request, 'usd_all_df')

    if df_zwg_raw is None or df_usd_raw is None:
        ctx['error'] = 'Data unavailable. Please re-upload files and try again.'
        return render(request, 'merchant/sector_performance.html', ctx)

    has_sector_zwg = 'SECTOR' in df_zwg_raw.columns
    has_sector_usd = 'SECTOR' in df_usd_raw.columns
    if not has_sector_zwg and not has_sector_usd:
        ctx['error'] = (
            'No SECTOR column was found in the uploaded files. '
            'Please ensure your Excel files include a SECTOR column.'
        )
        return render(request, 'merchant/sector_performance.html', ctx)

    # Normalise
    df_zwg = df_zwg_raw.copy()
    df_usd = df_usd_raw.copy()
    for _df in [df_zwg, df_usd]:
        _df['MERCHANT NAME'] = _df['MERCHANT NAME'].astype(str).str.strip()
        if 'SECTOR' in _df.columns:
            _df['SECTOR'] = _df['SECTOR'].astype(str).str.strip()

    def _period_cols(all_cols, from_d, to_d):
        return [
            c for c in all_cols
            if pd.notnull(_parse_col_date(c)) and from_d <= _parse_col_date(c).date() <= to_d
        ]

    zwg_pcols = _period_cols(extract_date_columns(df_zwg), from_date, to_date)
    usd_pcols = _period_cols(extract_date_columns(df_usd), from_date, to_date)

    if not zwg_pcols and not usd_pcols:
        ctx['error'] = f'No transaction date columns found in the selected period ({date_label}).'
        return render(request, 'merchant/sector_performance.html', ctx)

    pkl_d = _pickle_dir(request)

    def _sector_group(df, pcols):
        """Return a DataFrame sorted by revenue desc, with 'share' column added."""
        p_avail = [c for c in pcols if c in df.columns]
        if not p_avail:
            return pd.DataFrame()
        d = df.copy()
        for col in p_avail:
            d[col] = _to_numeric(d[col])
        d['_rev'] = d[p_avail].sum(axis=1)
        grp = d.groupby('SECTOR', as_index=False)['_rev'].sum()
        grp = grp[~grp['SECTOR'].str.upper().isin(['NAN', ''])].copy()
        grp = grp.sort_values('_rev', ascending=False).reset_index(drop=True)
        total = grp['_rev'].sum()
        grp['share'] = (grp['_rev'] / total * 100).round(1) if total > 0 else 0.0
        return grp

    # ── ZWG ──────────────────────────────────────────────────────────────────
    zwg_grp  = pd.DataFrame()
    top10_zwg = pd.DataFrame()
    if has_sector_zwg and zwg_pcols:
        zwg_grp = _sector_group(df_zwg, zwg_pcols)
        if not zwg_grp.empty:
            total_zwg = zwg_grp['_rev'].sum()
            top10_zwg = zwg_grp.head(10)

            ctx['total_zwg_rev']  = f'{total_zwg:,.0f}'
            ctx['top_zwg_sector'] = {
                'name':    zwg_grp.iloc[0]['SECTOR'],
                'revenue': f"{zwg_grp.iloc[0]['_rev']:,.0f}",
                'share':   f"{zwg_grp.iloc[0]['share']:.1f}%",
            }
            ctx['zwg_chart_div'] = build_plotly_bar_div(
                labels=top10_zwg['SECTOR'].tolist(),
                values=top10_zwg['_rev'].tolist(),
                title=f'Top 10 Sectors — ZWG Revenue  |  {date_label}',
                x_label='ZWG Revenue',
            )
            ctx['zwg_sector_rows'] = [
                {
                    'rank':    i + 1,
                    'sector':  r['SECTOR'],
                    'revenue': f"{r['_rev']:,.0f}",
                    'share':   f"{r['share']:.1f}%",
                }
                for i, r in top10_zwg.iterrows()
            ]

    # ── USD ──────────────────────────────────────────────────────────────────
    usd_grp  = pd.DataFrame()
    top10_usd = pd.DataFrame()
    if has_sector_usd and usd_pcols:
        usd_grp = _sector_group(df_usd, usd_pcols)
        if not usd_grp.empty:
            total_usd = usd_grp['_rev'].sum()
            top10_usd = usd_grp.head(10)

            ctx['total_usd_rev']  = f'{total_usd:,.2f}'
            ctx['top_usd_sector'] = {
                'name':    usd_grp.iloc[0]['SECTOR'],
                'revenue': f"{usd_grp.iloc[0]['_rev']:,.2f}",
                'share':   f"{usd_grp.iloc[0]['share']:.1f}%",
            }
            ctx['usd_chart_div'] = build_plotly_bar_div(
                labels=top10_usd['SECTOR'].tolist(),
                values=top10_usd['_rev'].tolist(),
                title=f'Top 10 Sectors — USD Revenue  |  {date_label}',
                x_label='USD Revenue',
            )
            ctx['usd_sector_rows'] = [
                {
                    'rank':    i + 1,
                    'sector':  r['SECTOR'],
                    'revenue': f"{r['_rev']:,.2f}",
                    'share':   f"{r['share']:.1f}%",
                }
                for i, r in top10_usd.iterrows()
            ]

    # ── Total / Active sectors (mirrors Merchant Performance terminal counts) ─
    all_sec_set = set()
    if has_sector_zwg:
        all_sec_set |= {
            s for s in df_zwg['SECTOR'].dropna().astype(str).str.strip().unique()
            if s.upper() not in ('NAN', '')
        }
    if has_sector_usd:
        all_sec_set |= {
            s for s in df_usd['SECTOR'].dropna().astype(str).str.strip().unique()
            if s.upper() not in ('NAN', '')
        }
    total_sectors = len(all_sec_set)

    active_sec_set = (
        (set(zwg_grp.loc[zwg_grp['_rev'] > 0, 'SECTOR'].astype(str).str.strip()) if not zwg_grp.empty else set())
        | (set(usd_grp.loc[usd_grp['_rev'] > 0, 'SECTOR'].astype(str).str.strip()) if not usd_grp.empty else set())
    )
    total_active_sectors = len(active_sec_set)

    ctx['total_sectors']        = total_sectors
    ctx['total_active_sectors'] = total_active_sectors

    # ── Combined Top 10 Sectors Excel ─────────────────────────────────────────
    if not top10_zwg.empty or not top10_usd.empty:
        (pkl_d / 'sector_combined.xlsx').write_bytes(
            _sector_combined_excel_bytes(
                top10_zwg if not top10_zwg.empty else None,
                top10_usd if not top10_usd.empty else None,
            )
        )
        ctx['has_combined_download'] = True
        request.session['sector_combined_label'] = (
            f"Sector Revenues for {from_date.strftime('%d %b %Y')} to {to_date.strftime('%d %b %Y')}"
        )

    # ── Prior period for Detailed Sector Performance Report ───────────────────
    period_days    = (to_date - from_date).days + 1
    prior_to_date  = from_date - timedelta(days=1)
    prior_from_date = prior_to_date - timedelta(days=period_days - 1)

    zwg_prior_pcols = _period_cols(extract_date_columns(df_zwg), prior_from_date, prior_to_date) if has_sector_zwg else []
    usd_prior_pcols = _period_cols(extract_date_columns(df_usd), prior_from_date, prior_to_date) if has_sector_usd else []

    zwg_prev = _sector_group(df_zwg, zwg_prior_pcols) if has_sector_zwg and zwg_prior_pcols else pd.DataFrame()
    usd_prev = _sector_group(df_usd, usd_prior_pcols) if has_sector_usd and usd_prior_pcols else pd.DataFrame()

    zwg_gs   = _sector_gainers_shakers(zwg_grp, zwg_prev, 1_100_000)
    usd_gs   = _sector_gainers_shakers(usd_grp, usd_prev, 10_000)

    zwg_idle = _idle_sectors(df_zwg, zwg_pcols) if has_sector_zwg else pd.DataFrame(columns=['SECTOR'])
    usd_idle = _idle_sectors(df_usd, usd_pcols) if has_sector_usd else pd.DataFrame(columns=['SECTOR'])

    if not zwg_grp.empty or not usd_grp.empty:
        (pkl_d / 'sector_detailed.xlsx').write_bytes(
            _sector_detailed_excel_bytes(
                zwg_gs, usd_gs, zwg_idle, usd_idle,
                zwg_grp, usd_grp, zwg_prev, usd_prev,
                from_date, to_date, prior_from_date, prior_to_date,
            )
        )
        ctx['has_detailed_download'] = True
        request.session['sector_detailed_label'] = (
            f"Detailed Sector Performance Report for {from_date.strftime('%d %b %Y')} to {to_date.strftime('%d %b %Y')}"
        )

    ctx['has_dashboard'] = True
    return render(request, 'merchant/sector_performance.html', ctx)


def download_sector_combined(request):
    label = request.session.get('sector_combined_label', 'Top 10 Sectors')
    return _serve_file(request, 'sector_combined.xlsx', f'{label}.xlsx')


def download_sector_detailed(request):
    label = request.session.get('sector_detailed_label', 'Detailed Sector Performance Report')
    return _serve_file(request, 'sector_detailed.xlsx', f'{label}.xlsx')


# =============================================================================
# DOWNLOAD VIEWS
# =============================================================================

def _serve_file(request, filename, download_name, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'):
    pkl_d = _pickle_dir(request)
    path  = pkl_d / filename
    if not path.exists():
        return HttpResponse('File not found. Please regenerate the report.', status=404)
    data = path.read_bytes()
    resp = HttpResponse(data, content_type=mime)
    resp['Content-Disposition'] = f'attachment; filename="{download_name}"'
    return resp


def download_zwg_report(request):
    return _serve_file(request, 'zwg_report.xlsx', 'ZWG MERCHANTS DAILY PERFORMANCE REPORT.xlsx')


def download_usd_report(request):
    return _serve_file(request, 'usd_report.xlsx', 'USD MERCHANTS DAILY PERFORMANCE REPORT.xlsx')


def download_zwg_gainers(request):
    return _serve_file(request, 'zwg_gainers.xlsx', 'ZWG Gainers and Shakers Report.xlsx')


def download_usd_gainers(request):
    return _serve_file(request, 'usd_gainers.xlsx', 'USD Gainers and Shakers Report.xlsx')


def download_zwg_top25(request):
    return _serve_file(request, 'zwg_top25.xlsx', 'ZWG Top 25 Report.xlsx')


def download_usd_top25(request):
    return _serve_file(request, 'usd_top25.xlsx', 'USD Top 25 Report.xlsx')


def download_pos_statement(request):
    pkl_d     = _pickle_dir(request)
    short_name = request.session.get('pos_short_name', 'Merchant')
    return _serve_file(request, 'pos_statement.xlsx', f'{short_name} Daily Pos Statements.xlsx')


def download_champions_zinara(request):
    return _serve_file(request, 'champions_zinara.xlsx', 'CHAMPIONS ZINARA DAILY POS STATEMENT.xlsx')


def download_champions_insurance(request):
    return _serve_file(request, 'champions_insurance.xlsx', 'CHAMPIONS INSURANCE DAILY POS STATEMENT.xlsx')


def download_idle_total(request):
    date_label = request.GET.get('date_label', 'period')
    return _serve_file(request, 'idle_total.xlsx', f'Total Idle Terminals for {date_label}.xlsx')


def download_idle_unit(request):
    bu         = request.GET.get('bu', 'Unit')
    date_label = request.GET.get('date_label', 'period')
    return _serve_file(request, 'idle_unit.xlsx', f'{bu} Idle Terminals for the period {date_label}.xlsx')


def download_zwg_top25_weekly(request):
    return _serve_file(request, 'zwg_top25_weekly.xlsx', 'ZWG Weekly Top 25 Report.xlsx')


def download_usd_top25_weekly(request):
    return _serve_file(request, 'usd_top25_weekly.xlsx', 'USD Weekly Top 25 Report.xlsx')


# =============================================================================
# AUTOMATIC REPORTS GENERATION
# =============================================================================

def auto_reports_view(request):
    ctx = {'has_results': False, 'errors': []}

    if request.method == 'POST':
        # Force the session to fully load NOW, before _pickle_dir is called.
        # Django lazily loads session data on first access.  If the browser
        # sends a stale session cookie whose file no longer exists, the load
        # triggers an internal create() that silently replaces the session key.
        # If that happens after _pickle_dir has already used the old key, the
        # generated file ends up in the wrong directory and the download fails.
        # Accessing the session here ensures any key rotation happens before
        # we compute the pickle path, so both the file write and the response
        # cookie always reference the same key.
        request.session.get('_ready')
        if not request.session.session_key:
            request.session.create()

        b02_files = request.FILES.getlist('b02_files')
        zwg_file  = request.FILES.get('zwg_report')
        usd_file  = request.FILES.get('usd_report')

        if not b02_files:
            ctx['errors'].append('Please upload at least one B02 file.')
            return render(request, 'merchant/auto_reports.html', ctx)
        if not zwg_file and not usd_file:
            ctx['errors'].append('Please upload at least one merchant report file (ZWG or USD).')
            return render(request, 'merchant/auto_reports.html', ctx)

        # Validate each B02 file before processing
        validated_b02 = []
        for f in b02_files:
            fname = getattr(f, 'name', 'uploaded file')
            try:
                f.seek(0)
                raw = f.read()
            except Exception:
                ctx['errors'].append(f'Could not read "{fname}". Please try uploading it again.')
                continue
            err = validate_b02_file(raw, fname)
            if err:
                ctx['errors'].append(err)
            else:
                # Re-wrap as a readable file-like for parse_b02_files
                wrapped = io.BytesIO(raw)
                wrapped.name = fname
                validated_b02.append(wrapped)

        if ctx['errors']:
            return render(request, 'merchant/auto_reports.html', ctx)

        if not validated_b02:
            ctx['errors'].append('No valid B02 files were found after validation.')
            return render(request, 'merchant/auto_reports.html', ctx)

        # Parse all validated B02 files
        try:
            b02_totals = parse_b02_files(validated_b02)
        except Exception as e:
            ctx['errors'].append(f'Error parsing B02 files: {e}')
            return render(request, 'merchant/auto_reports.html', ctx)

        if not b02_totals:
            ctx['errors'].append('No matching transactions found in the uploaded B02 files. '
                                  'Check that the files contain "Goods and Services" / '
                                  '"Approved or Completed" transactions.')
            return render(request, 'merchant/auto_reports.html', ctx)

        pkl_d = _pickle_dir(request)
        zwg_stats = usd_stats = None

        # Process ZWG report
        if zwg_file:
            try:
                zwg_bytes = zwg_file.read()
                updated_zwg, zwg_stats = update_merchant_report(zwg_bytes, b02_totals, 'ZWG')
                (pkl_d / 'auto_zwg.xlsx').write_bytes(updated_zwg)
                request.session['auto_zwg_filename'] = zwg_file.name
            except Exception as e:
                ctx['errors'].append(f'Error processing ZWG report: {e}')

        # Process USD report
        if usd_file:
            try:
                usd_bytes = usd_file.read()
                updated_usd, usd_stats = update_merchant_report(usd_bytes, b02_totals, 'USD')
                (pkl_d / 'auto_usd.xlsx').write_bytes(updated_usd)
                request.session['auto_usd_filename'] = usd_file.name
            except Exception as e:
                ctx['errors'].append(f'Error processing USD report: {e}')

        if not ctx['errors'] or zwg_stats or usd_stats:
            dates_processed = sorted({d.strftime('%d %b %Y') for (_, d, _) in b02_totals})
            b02_tids = len({tid for (tid, _, _) in b02_totals})
            ctx.update({
                'has_results':      True,
                'b02_files_count':  len(b02_files),
                'b02_tids':         b02_tids,
                'dates_processed':  ', '.join(dates_processed) if dates_processed else '—',
                'zwg_stats':        zwg_stats,
                'usd_stats':        usd_stats,
                'has_zwg_download': zwg_stats is not None,
                'has_usd_download': usd_stats is not None,
            })

    return render(request, 'merchant/auto_reports.html', ctx)


def auto_reports_proceed(request):
    """
    Copy auto-generated reports into the session upload dir then fully initialise
    the analytics session — exactly what the regular upload + sheet-select flow
    does — so every analytics page works immediately.
    """
    pkl_d    = _pickle_dir(request)
    upload_d = _session_dir(request)

    # ── 1. Copy generated files into the session upload dir ───────────────────
    zwg_src = pkl_d / 'auto_zwg.xlsx'
    usd_src = pkl_d / 'auto_usd.xlsx'

    if not zwg_src.exists() and not usd_src.exists():
        return redirect('auto_reports')

    if zwg_src.exists():
        shutil.copy2(zwg_src, upload_d / 'zwg.xlsx')
    if usd_src.exists():
        shutil.copy2(usd_src, upload_d / 'usd.xlsx')

    # ── 2. Read sheet names (needed by choose_analytics and other views) ───────
    required = {'TID', 'CIF', 'MERCHANT NAME'}

    def _valid_sheets(path):
        """Return list of sheet names and the DataFrame from the last valid sheet."""
        try:
            xls = pd.ExcelFile(path)
        except Exception:
            return [], None
        sheet_names = xls.sheet_names
        last_df = None
        for sname in reversed(sheet_names):
            try:
                df = normalize_columns(pd.read_excel(path, sheet_name=sname))
                if required.issubset(set(df.columns)):
                    df = df[~df['MERCHANT NAME'].astype(str).str.upper().str.strip()
                            .isin(['TOTAL', 'GRAND TOTAL'])]
                    last_df = df
                    break
            except Exception:
                continue
        return sheet_names, last_df

    zwg_sheet_names, zwg_df = _valid_sheets(upload_d / 'zwg.xlsx') if (upload_d / 'zwg.xlsx').exists() else ([], None)
    usd_sheet_names, usd_df = _valid_sheets(upload_d / 'usd.xlsx') if (upload_d / 'usd.xlsx').exists() else ([], None)

    # ── 3. Save daily-analytics DataFrames (zwg_df / usd_df) ──────────────────
    if zwg_df is not None:
        _save_df(request, 'zwg_df', zwg_df)
    if usd_df is not None:
        _save_df(request, 'usd_df', usd_df)

    # data_period label for sidebar (derived from the auto-selected sheet)
    try:
        ref_df = zwg_df if zwg_df is not None else usd_df
        if ref_df is not None:
            date_cols = extract_date_columns(ref_df)
            parsed = sorted([_parse_col_date(c) for c in date_cols if pd.notnull(_parse_col_date(c))])
            if parsed:
                request.session['data_period'] = (
                    f"{parsed[0].strftime('%d %b %Y')} → {parsed[-1].strftime('%d %b %Y')}"
                )
    except Exception:
        pass

    # ── 4. Load all sheets for periodic analytics ──────────────────────────────
    # Inline version of _load_all_sheets that handles missing files gracefully.
    all_zwg, all_usd = [], []
    for path, bucket in ((upload_d / 'zwg.xlsx', all_zwg), (upload_d / 'usd.xlsx', all_usd)):
        if not path.exists():
            continue
        try:
            for sname in pd.ExcelFile(path).sheet_names:
                try:
                    df = normalize_columns(pd.read_excel(path, sheet_name=sname))
                    if not required.issubset(set(df.columns)):
                        continue
                    df = df[~df['MERCHANT NAME'].astype(str).str.upper().str.strip()
                            .isin(['TOTAL', 'GRAND TOTAL'])]
                    bucket.append(df)
                except Exception:
                    pass
        except Exception:
            pass

    if all_zwg:
        _save_df(request, 'zwg_all_df', pd.concat(all_zwg, ignore_index=True))
    if all_usd:
        _save_df(request, 'usd_all_df', pd.concat(all_usd, ignore_index=True))

    # data_period_all label (across all sheets, for periodic analytics sidebar)
    try:
        combined = _load_df(request, 'zwg_all_df')
        if combined is not None:
            date_cols = extract_date_columns(combined)
            parsed = sorted([_parse_col_date(c) for c in date_cols if pd.notnull(_parse_col_date(c))])
            if parsed:
                request.session['data_period_all'] = (
                    f"{parsed[0].strftime('%d %b %Y')} → {parsed[-1].strftime('%d %b %Y')}"
                )
    except Exception:
        pass

    # ── 5. Set all session flags ───────────────────────────────────────────────
    request.session['zwg_sheet_names']      = zwg_sheet_names
    request.session['usd_sheet_names']      = usd_sheet_names
    request.session['files_validated']      = True
    request.session['data_loaded']          = (zwg_df is not None or usd_df is not None)
    request.session['periodic_data_loaded'] = True

    return redirect('choose_analytics')


def download_auto_zwg(request):
    name = request.session.get('auto_zwg_filename', 'ZWG All Merchants Report.xlsx')
    return _serve_file(request, 'auto_zwg.xlsx', name)


def download_auto_usd(request):
    name = request.session.get('auto_usd_filename', 'USD All Merchants Report.xlsx')
    return _serve_file(request, 'auto_usd.xlsx', name)


# =============================================================================
# LOGOUT
# =============================================================================

def logout_view(request):
    key = request.session.session_key
    if key:
        for sub in ('uploads', 'pickles'):
            d = Path(settings.TEMP_DATA_DIR) / sub / key
            if d.exists():
                shutil.rmtree(d, ignore_errors=True)
    request.session.flush()
    return redirect('upload')
