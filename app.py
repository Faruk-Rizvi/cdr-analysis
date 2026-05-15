"""
==============================================================================
CDR Analysis Web App — Streamlit
==============================================================================
Upload any CDR Excel file → Get HTML + Word Report instantly.

Run locally:
    streamlit run app.py

Deploy free:
    https://streamlit.io/cloud
==============================================================================
"""

import streamlit as st
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import re
import os
import io
import base64
import tempfile
import warnings
from datetime import datetime

warnings.filterwarnings('ignore')

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="CDR Analysis Tool",
    page_icon="📞",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ─────────────────────────────────────────────
# CUSTOM CSS — Professional Design
# ─────────────────────────────────────────────
st.markdown("""
<style>
    /* Global */
    .stApp { background: #f1f5f9; }
    .main .block-container { padding-top: 1rem; padding-bottom: 2rem; max-width: 1400px; }
    #MainMenu, footer, header[data-testid="stHeader"] { visibility: hidden; }

    /* Top App Header */
    .app-header {
        background: white;
        padding: 1rem 1.5rem;
        border-radius: 12px;
        margin-bottom: 1.25rem;
        display: flex;
        align-items: center;
        justify-content: space-between;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }
    .app-header-left { display: flex; align-items: center; gap: 0.75rem; }
    .app-header-logo {
        width: 42px; height: 42px;
        background: linear-gradient(135deg, #2563eb 0%, #1e40af 100%);
        border-radius: 10px;
        display: flex; align-items: center; justify-content: center;
        color: white; font-size: 1.3rem;
    }
    .app-header-title {
        font-size: 1.15rem; font-weight: 700; color: #0f172a;
    }
    .app-header-nav { display: flex; gap: 1.75rem; color: #475569; font-size: 0.95rem; font-weight: 500; }
    .app-header-nav span { display: flex; align-items: center; gap: 0.4rem; cursor: default; }

    /* Hero Section */
    .hero {
        background: linear-gradient(135deg, #1e3a8a 0%, #1e40af 50%, #2563eb 100%);
        border-radius: 16px;
        padding: 2.75rem 2.5rem;
        color: white;
        margin-bottom: 1.5rem;
        position: relative;
        overflow: hidden;
        box-shadow: 0 10px 30px rgba(30, 58, 138, 0.25);
    }
    .hero::before {
        content: '';
        position: absolute;
        top: -50%; right: -10%;
        width: 500px; height: 500px;
        background: radial-gradient(circle, rgba(96, 165, 250, 0.15) 0%, transparent 70%);
        border-radius: 50%;
    }
    .hero-icon-area {
        position: absolute;
        right: 2.5rem; top: 50%;
        transform: translateY(-50%);
        font-size: 5rem;
        opacity: 0.15;
        z-index: 0;
    }
    .hero h1 {
        color: white !important;
        font-size: 2.8rem !important;
        font-weight: 800;
        margin: 0 0 0.5rem 0 !important;
        line-height: 1.1;
        position: relative; z-index: 1;
    }
    .hero p {
        color: #cfe0ff !important;
        font-size: 1.05rem;
        margin: 0 0 1.5rem 0;
        max-width: 600px;
        line-height: 1.5;
        position: relative; z-index: 1;
    }
    .hero-supports {
        display: flex; align-items: center; gap: 0.5rem;
        color: #93c5fd; font-size: 0.9rem; font-weight: 600;
        margin-bottom: 1rem;
        position: relative; z-index: 1;
    }
    .operator-row {
        display: flex; gap: 0.6rem; flex-wrap: wrap;
        position: relative; z-index: 1;
    }
    .op-badge {
        background: white;
        padding: 0.6rem 1rem;
        border-radius: 10px;
        display: flex; align-items: center; gap: 0.5rem;
        color: #0f172a; font-weight: 600; font-size: 0.9rem;
        box-shadow: 0 2px 6px rgba(0,0,0,0.08);
    }
    .op-dot { width: 14px; height: 14px; border-radius: 50%; }

    /* Card Container */
    .card {
        background: white;
        border-radius: 14px;
        padding: 1.75rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        margin-bottom: 1.25rem;
    }
    .card-title {
        display: flex; align-items: center; gap: 0.7rem;
        font-size: 1.15rem; font-weight: 700; color: #0f172a;
        margin-bottom: 0.25rem;
    }
    .card-subtitle { color: #64748b; font-size: 0.88rem; margin-bottom: 1rem; }
    .card-icon-blue {
        width: 36px; height: 36px;
        background: #dbeafe;
        border-radius: 9px;
        display: flex; align-items: center; justify-content: center;
        color: #2563eb; font-size: 1.1rem;
    }

    /* Info Banner */
    .info-banner {
        background: #eff6ff;
        border-left: 4px solid #2563eb;
        border-radius: 12px;
        padding: 1rem 1.5rem;
        display: flex; align-items: center; gap: 1rem;
        margin-bottom: 1.5rem;
    }
    .info-banner-icon {
        width: 40px; height: 40px;
        background: #2563eb;
        border-radius: 50%;
        color: white;
        display: flex; align-items: center; justify-content: center;
        font-size: 1.2rem;
        flex-shrink: 0;
    }
    .info-banner-title { font-weight: 700; color: #1e3a8a; font-size: 1rem; }
    .info-banner-text { color: #475569; font-size: 0.9rem; margin-top: 0.15rem; }

    /* Capabilities Grid */
    .cap-card {
        background: white;
        border-radius: 14px;
        padding: 1.5rem 1.75rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        margin-bottom: 1.25rem;
    }
    .cap-title {
        display: flex; align-items: center; gap: 0.6rem;
        font-size: 1.1rem; font-weight: 700; color: #0f172a;
        margin-bottom: 1.25rem;
    }
    .cap-grid {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 0.75rem 2rem;
    }
    .cap-item {
        display: flex; align-items: center; gap: 0.85rem;
        padding: 0.5rem 0;
    }
    .cap-icon {
        width: 38px; height: 38px;
        border-radius: 10px;
        display: flex; align-items: center; justify-content: center;
        font-size: 1.05rem;
        flex-shrink: 0;
    }
    .cap-text { color: #334155; font-size: 0.92rem; font-weight: 500; }

    /* Trust Badges */
    .trust-row {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 1rem;
        background: white;
        border-radius: 14px;
        padding: 1.5rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .trust-item { display: flex; align-items: center; gap: 1rem; }
    .trust-icon {
        width: 48px; height: 48px;
        border-radius: 12px;
        display: flex; align-items: center; justify-content: center;
        font-size: 1.4rem;
    }
    .trust-title { font-weight: 700; color: #0f172a; font-size: 0.98rem; }
    .trust-sub { color: #64748b; font-size: 0.85rem; margin-top: 0.1rem; }

    /* Footer */
    .footer {
        text-align: center;
        margin-top: 2rem;
        padding: 1rem;
        color: #64748b;
        font-size: 0.88rem;
        border-top: 1px solid #e2e8f0;
    }
    .footer .dev-name { color: #2563eb; font-weight: 700; }

    /* Stat Cards */
    .stat-card {
        background: white;
        border-left: 4px solid #2563eb;
        padding: 1rem 1.2rem;
        border-radius: 10px;
        margin-bottom: 0.5rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .stat-card .label { font-size: 0.78rem; color: #64748b; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }
    .stat-card .value { font-size: 1.4rem; color: #1e3a8a; font-weight: 700; margin-top: 0.2rem; }

    /* Status Boxes */
    .success-box {
        background: #ecfdf5;
        border: 1px solid #10b981;
        border-radius: 10px;
        padding: 1rem 1.5rem;
        margin: 1rem 0;
        color: #065f46;
    }
    .warning-box {
        background: #fffbeb;
        border: 1px solid #f59e0b;
        border-radius: 10px;
        padding: 0.85rem 1.25rem;
        margin: 0.5rem 0;
        color: #92400e;
    }

    /* Download Button */
    div[data-testid="stDownloadButton"] button {
        width: 100%;
        border-radius: 10px;
        font-weight: 600;
        padding: 0.7rem;
        background: #2563eb;
        color: white;
        border: none;
    }
    div[data-testid="stDownloadButton"] button:hover {
        background: #1e40af;
    }

    /* Progress */
    .stProgress > div > div > div {
        background: linear-gradient(90deg, #1e3a8a, #2563eb);
    }

    /* File Uploader Style */
    div[data-testid="stFileUploader"] section {
        background: #f8fafc;
        border: 2px dashed #cbd5e1;
        border-radius: 12px;
        padding: 1.5rem;
    }
    div[data-testid="stFileUploader"] section:hover {
        border-color: #2563eb;
        background: #eff6ff;
    }

    /* Text Input */
    div[data-testid="stTextInput"] input {
        border-radius: 10px;
        border: 1px solid #cbd5e1;
        padding: 0.75rem 1rem;
        font-size: 0.95rem;
    }
    div[data-testid="stTextInput"] input:focus {
        border-color: #2563eb;
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1);
    }

    /* Expander */
    div[data-testid="stExpander"] {
        background: white;
        border-radius: 12px;
        border: 1px solid #e2e8f0;
        margin-bottom: 0.75rem;
    }
    div[data-testid="stExpander"] summary {
        font-weight: 600;
        color: #0f172a;
        padding: 0.75rem 1rem;
    }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# COLUMN ALIASES
# ─────────────────────────────────────────────
COLUMN_ALIASES = {
    'start':            ['start', 'datetime', 'date', 'timestamp', 'call_date',
                         'call date', 'date/time', 'date time', 'starttime',
                         'start time', 'start_datetime', 'start_dttime', 'startdttime',
                         'call_datetime', 'calldatetime', 'datetime_start'],
    'operator':         ['operator', 'network', 'telco', 'carrier',
                         'provider name', 'provider_name', 'service provider',
                         # Banglalink CDR
                         'providername'],
    'party_a':          ['party a', 'party_a', 'a_number', 'msisdn_a', 'a-number',
                         'caller', 'originating', 'a number', 'partya', 'msisdn',
                         'aparty', 'a party', 'a_party',
                         # Banglalink CDR
                         'aparty'],
    'party_b':          ['party b', 'party_b', 'b_number', 'msisdn_b', 'b-number',
                         'called', 'terminating', 'b number', 'partyb', 'callee',
                         'bparty', 'b party', 'b_party',
                         # Banglalink CDR
                         'bparty'],
    'party_b_original': ['party b original', 'party_b_original', 'original_b',
                         'b_original', 'partyb_original'],
    'duration':         ['call duration', 'call_duration', 'duration',
                         'call_length', 'duration_sec', 'duration(sec)',
                         # Banglalink CDR
                         'callduration'],
    'usage_type':       ['usage type', 'usage_type', 'call_type', 'call type',
                         'type', 'direction', 'service_type',
                         # Banglalink CDR
                         'usagetype'],
    'cell_type':        ['cell type', 'cell_type', 'network_type', 'network type',
                         'technology', 'rat',
                         # Banglalink CDR
                         'networktype'],
    'lac':              ['lac id', 'lac_id', 'lac', 'location_area_code',
                         'lacstarta', 'lacstart'],
    'cell_id':          ['cell id', 'cell_id', 'cell', 'bts_id', 'bts id',
                         'site_id', 'tower_id', 'cistarta', 'cisstarta'],
    'imei':             ['imei', 'device_id', 'handset_id'],
    'imsi':             ['imsi', 'subscriber_id', 'imsia'],
    'address':          ['address', 'location', 'tower_location', 'tower location',
                         'site_name', 'cell_name', 'area', 'thana', 'district',
                         'bts address', 'bts_address'],
}

CALL_OUT_TYPES = ['moc', 'mo', 'outgoing', 'out', 'call-mo', 'call_mo', 'callmo']
CALL_IN_TYPES  = ['mtc', 'mt', 'incoming', 'in', 'call-mt', 'call_mt', 'callmt']
SMS_OUT_TYPES  = ['mo-sms', 'sms-mo', 'smsmo', 'sms_mo', 'sms-mo', 'sms out',
                  # Banglalink CDR
                  'smsmo']
SMS_IN_TYPES   = ['mt-sms', 'sms-mt', 'smsmt', 'sms_mt', 'sms-mt', 'sms in',
                  # Banglalink CDR
                  'smsmt']


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def detect_column(df_columns, alias_key):
    aliases = COLUMN_ALIASES.get(alias_key, [])
    cols_lower = {c.lower().strip(): c for c in df_columns}
    for alias in aliases:
        if alias.lower() in cols_lower:
            return cols_lower[alias.lower()]
    return None


def is_valid_number(val):
    return bool(re.match(r'^\d{10,20}$', str(val).strip()))


def fig_to_base64(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=100, bbox_inches='tight')
    buf.seek(0)
    b64 = base64.b64encode(buf.read()).decode('utf-8')
    plt.close(fig)
    return b64


# ─────────────────────────────────────────────
# LOAD & CLEAN
# ─────────────────────────────────────────────
def normalize_number(val):
    """
    Normalize a Bangladesh mobile number to last-10-digit form (internal key).
    880XXXXXXXXXX / 0XXXXXXXXXX / XXXXXXXXXX → last 10 digits
    Used for grouping/matching — not for display.
    """
    s = str(val).strip()
    digits = re.sub(r'[^0-9]', '', s)
    if len(digits) >= 10:
        return digits[-10:]
    return s


def display_number(val):
    """
    Format a number for display as 11-digit Bangladesh format: 01XXXXXXXXX
    8801XXXXXXXXX → 01XXXXXXXXX
    1XXXXXXXXX (10 digits) → 01XXXXXXXXX
    01XXXXXXXXX → 01XXXXXXXXX (unchanged)
    Non-mobile values returned as-is.
    """
    s = str(val).strip()
    digits = re.sub(r'[^0-9]', '', s)
    if len(digits) == 13 and digits.startswith('880'):
        return '0' + digits[3:]        # 8801XXXXXXXXX → 01XXXXXXXXX
    if len(digits) == 10:
        return '0' + digits            # 1XXXXXXXXX → 01XXXXXXXXX
    if len(digits) == 11 and digits.startswith('0'):
        return digits                  # already 01XXXXXXXXX
    return s                           # non-mobile (service names etc.)


def _extract_merged_rows(merged_val, cols):
    """
    Extract multiple hidden rows from a single merged cell value.
    Handles: tab-separated columns, _x000D_ row separators.
    Returns list of dicts, each dict = one recovered row.
    """
    # Split by _x000D_ (Excel carriage-return encoding) with optional newline
    raw_lines = re.split(r'_x000D_\r?\n?', merged_val)
    raw_lines = [l.strip() for l in raw_lines if l.strip()]

    recovered = []
    for line in raw_lines:
        parts = line.split('\t')
        # Pad missing cells with '-'
        if len(parts) < len(cols):
            parts = parts + ['-'] * (len(cols) - len(parts))
        elif len(parts) > len(cols):
            parts = parts[:len(cols)]
        recovered.append(dict(zip(cols, parts)))

    return recovered


def _try_fix_merged_row(row_series, expected_cols):
    """
    Try to detect and fix a row where data is merged into fewer cells.
    Returns fixed Series if fixable (simple case), else None.
    (Complex _x000D_ cases are handled separately in load_and_clean.)
    """
    vals = [str(v) for v in row_series.values if pd.notna(v) and str(v).strip() not in ('', 'nan')]
    if len(vals) == 0:
        return None

    # Simple pipe/semicolon separated single-row merges
    for sep in ['|', ';;']:
        if any(sep in v for v in vals):
            parts = []
            for v in vals:
                parts.extend([x.strip() for x in v.split(sep)])
            if len(parts) >= len(expected_cols) - 2:
                padded = (parts + ['-'] * len(expected_cols))[:len(expected_cols)]
                return pd.Series(padded, index=expected_cols)

    return None


def load_and_clean(file_bytes):
    # ── Detect best sheet ──
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets = xl.sheet_names
    best_sheet, best_score = sheets[0], 0
    cdr_keywords = ['start', 'party', 'duration', 'usage', 'operator',
                    'msisdn', 'lac', 'cell', 'imei', 'imsi', 'address',
                    'aparty', 'bparty']
    for s in sheets:
        try:
            tmp = pd.read_excel(io.BytesIO(file_bytes), sheet_name=s, dtype=str, nrows=3)
            cols_lower = [c.lower() for c in tmp.columns]
            score = sum(any(kw in c for c in cols_lower) for kw in cdr_keywords)
            if score > best_score:
                best_score, best_sheet = score, s
        except Exception:
            continue

    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=best_sheet, dtype=str)

    # ── No header detection ──
    non_str_cols = [c for c in df.columns if not isinstance(c, str)]
    if non_str_cols:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=best_sheet,
                           header=None, dtype=str)
        default_cols = ['Start', 'Operator', 'Party A', 'Party B', 'Call Duration',
                        'Usage Type', 'Cell Type', 'LAC ID', 'Cell ID', 'IMEI',
                        'IMSI', 'Address', 'Party B Original']
        df.columns = (default_cols[:len(df.columns)] if len(df.columns) <= len(default_cols)
                      else default_cols + [f'Extra_{i}' for i in range(len(df.columns)-len(default_cols))])

    total_raw = len(df)
    expected_cols = df.columns.tolist()

    # ── Merged Cell Fix ──
    # Detect rows where any cell contains _x000D_ (multi-row merge) or tab-separated data
    all_rows_out = []
    merged_fixed_count = 0

    for i, row in df.iterrows():
        found_merge = False
        for col in expected_cols:
            val = str(row.get(col, ''))
            if '_x000D_' in val and '\t' in val and len(val) > 100:
                # Multi-row merged cell: extract all hidden rows
                found_merge = True
                merged_fixed_count += 1

                # Build the first partial row from columns before the merged cell
                first_row = {}
                for c in expected_cols:
                    if c == col:
                        break
                    first_row[c] = str(row.get(c, '-'))

                # Extract hidden rows from merged cell
                recovered = _extract_merged_rows(val, expected_cols)

                # Merge first_row prefix into first recovered row
                if recovered:
                    for k, v in first_row.items():
                        recovered[0][k] = v
                    all_rows_out.extend(recovered)
                else:
                    all_rows_out.append(row.to_dict())
                break

        if not found_merge:
            # Simple merge check (pipe/semicolon)
            fixed = _try_fix_merged_row(row, expected_cols)
            if fixed is not None:
                all_rows_out.append(fixed.to_dict())
            else:
                all_rows_out.append(row.to_dict())

    if merged_fixed_count > 0:
        df = pd.DataFrame(all_rows_out, columns=expected_cols).reset_index(drop=True)

    # ── Rename columns ──
    col_map = {}
    for key in COLUMN_ALIASES:
        found = detect_column(df.columns, key)
        if found:
            col_map[key] = found
    rename = {v: k for k, v in col_map.items()}
    df = df.rename(columns=rename)

    # ── Parse datetime ──
    if 'start' in df.columns:
        raw = df['start'].astype(str).str.strip()

        # Try YYYYMMDDHHMMSS (14 digit compact format: 20260220092533)
        compact_mask = raw.str.match(r'^\d{14}$')
        if compact_mask.sum() > len(df) * 0.5:
            df['start'] = pd.to_datetime(raw, format='%Y%m%d%H%M%S', errors='coerce')
        # Try YYYYMMDD (8 digit date only)
        elif raw.str.match(r'^\d{8}$').sum() > len(df) * 0.5:
            df['start'] = pd.to_datetime(raw, format='%Y%m%d', errors='coerce')
        else:
            # Standard formats: ISO, DD/MM/YYYY, etc.
            df['start'] = pd.to_datetime(raw, errors='coerce', dayfirst=False)
            # Try dayfirst if too many failed
            if df['start'].isna().sum() > len(df) * 0.3:
                df['start'] = pd.to_datetime(raw, errors='coerce', dayfirst=True)

        df = df.dropna(subset=['start'])
        df = df.sort_values('start').reset_index(drop=True)

    # ── Parse duration ──
    if 'duration' in df.columns:
        df['duration'] = pd.to_numeric(df['duration'], errors='coerce').fillna(0).astype(int)

    # ── Fix Party B ──
    if 'party_b_original' in df.columns:
        pb_orig = df['party_b_original'].replace(['nan', 'None'], pd.NA)
        if pb_orig.notna().sum() > 0:
            df['party_b_clean'] = pb_orig.fillna(
                df.get('party_b', pb_orig)).astype(str).str.strip()
        elif 'party_b' in df.columns:
            df['party_b_clean'] = df['party_b'].astype(str).str.strip()
    elif 'party_b' in df.columns:
        df['party_b_clean'] = df['party_b'].astype(str).str.strip()

    # ── Normalize phone numbers: 880XXXXXXXXXX / 0XXXXXXXXXX / XXXXXXXXXX → last 10 digits ──
    # This ensures the same number stored differently is treated as one contact
    if 'party_b_clean' in df.columns:
        df['party_b_norm'] = df['party_b_clean'].apply(
            lambda v: normalize_number(v) if is_valid_number(v) else v
        )
    if 'party_a' in df.columns:
        df['party_a'] = df['party_a'].astype(str).str.strip()

    # ── Mark anomalies (DO NOT REMOVE — keep for location analysis) ──
    # is_anomaly = True means invalid party_b for CALL records
    # SMS records with service/app party_b (WhatsApp, service numbers) are VALID
    anomaly_count = 0
    if 'party_b_clean' in df.columns:
        invalid_pb = ~df['party_b_clean'].apply(is_valid_number)
        # For SMS rows, non-number party_b is normal (WhatsApp, bank OTP, etc.)
        is_sms_row = pd.Series(False, index=df.index)
        if 'usage_type' in df.columns:
            ut_tmp = df['usage_type'].astype(str).str.lower().str.strip()
            is_sms_row = ut_tmp.isin(SMS_OUT_TYPES + SMS_IN_TYPES)
        anomaly_mask = invalid_pb & ~is_sms_row
        anomaly_count = int(anomaly_mask.sum())
        df['is_anomaly'] = anomaly_mask
    else:
        df['is_anomaly'] = False

    if 'usage_type' in df.columns:
        df['usage_type'] = df['usage_type'].astype(str).str.strip()

    # ── Categorize ──
    ut = df['usage_type'].str.lower() if 'usage_type' in df.columns else pd.Series(['moc']*len(df))
    df['is_call_out'] = ut.isin(CALL_OUT_TYPES)
    df['is_call_in']  = ut.isin(CALL_IN_TYPES)
    df['is_sms_out']  = ut.isin(SMS_OUT_TYPES)
    df['is_sms_in']   = ut.isin(SMS_IN_TYPES)

    return df, col_map, total_raw, anomaly_count, best_sheet


def cdf(df):
    """Return clean df (non-anomaly rows only) for call/contact/SMS analysis.
    Also ensures party_b_norm exists (last-10-digit normalized number).
    """
    d = df[~df['is_anomaly']].copy() if 'is_anomaly' in df.columns else df.copy()
    # Ensure party_b_norm exists
    if 'party_b_norm' not in d.columns and 'party_b_clean' in d.columns:
        d['party_b_norm'] = d['party_b_clean'].apply(
            lambda v: normalize_number(v) if is_valid_number(v) else v
        )
    return d


# ─────────────────────────────────────────────
# ANALYSIS FUNCTIONS (same as cdr_analysis.py)
# ─────────────────────────────────────────────
def is_imei_cdr(df):
    """Return True if this CDR is IMEI-based (single dominant IMEI, multiple SIMs)."""
    if 'imei' not in df.columns or 'party_a' not in df.columns: return False
    imei_counts = df['imei'].dropna().value_counts()
    if imei_counts.empty: return False
    # Dominant IMEI covers >= 80% of records (filter garbage values like '-')
    real_imeis = imei_counts[imei_counts.index.map(
        lambda x: str(x).replace('-','').isdigit() and len(str(x)) >= 10
    )]
    if real_imeis.empty: return False
    top_imei_pct = real_imeis.iloc[0] / len(df)
    party_a_unique = df['party_a'].dropna().nunique()
    return top_imei_pct >= 0.80 and party_a_unique > 1

def get_phone(df):
    if 'party_a' not in df.columns: return 'N/A'
    if is_imei_cdr(df):
        # Show IMEI + all SIM numbers
        imei = df['imei'].dropna().mode()[0] if 'imei' in df.columns else 'N/A'
        sims = df['party_a'].dropna().value_counts()
        # Filter out garbage (too short or non-numeric)
        sims = [s for s in sims.index if str(s).replace('+','').isdigit() and len(str(s)) >= 10]
        sim_str = ' / '.join(str(s) for s in sims[:3])
        return f"IMEI: {imei} (SIMs: {sim_str})"
    return str(df['party_a'].mode()[0])

def get_operator(df):
    if 'operator' in df.columns:
        # Filter: keep only known operator names (not numeric garbage)
        ops = df['operator'].dropna().unique()
        known = [o for o in ops if str(o).strip() and not str(o).strip().isdigit()
                 and len(str(o).strip()) > 2]
        return ', '.join(str(o) for o in known) if known else 'N/A'
    return 'N/A'

def get_date_range(df):
    if 'start' in df.columns and len(df) > 0:
        return (f"{df['start'].min().strftime('%Y-%m-%d %H:%M:%S')} to "
                f"{df['start'].max().strftime('%Y-%m-%d %H:%M:%S')}")
    return 'N/A'

def call_summary(df):
    d = cdf(df)
    return pd.DataFrame({
        'Metric': ['Total Outgoing Calls','Total Incoming Calls',
                   'Total Sent SMS','Total Received SMS'],
        'Value':  [int(d['is_call_out'].sum()), int(d['is_call_in'].sum()),
                   int(d['is_sms_out'].sum()), int(d['is_sms_in'].sum())]
    })

def daily_call_count(df):
    if 'start' not in df.columns: return pd.DataFrame()
    calls = cdf(df); calls = calls[calls['is_call_out'] | calls['is_call_in']].copy()
    calls['hour'] = calls['start'].dt.hour.astype(int)
    rows = []
    for name, h_start, h_end, interval in [
        ('Morning', 5,  8,  '05:00-08:00'),
        ('Day',     8,  18, '08:00-18:00'),
        ('Evening', 18, 22, '18:00-22:00'),
        ('Night',   22, 29, '22:01-05:00')]:
        mask = (calls['hour']>=h_start)&(calls['hour']<h_end) if h_end<=24 \
               else (calls['hour']>=22)|(calls['hour']<5)
        sub = calls[mask]
        mc = sub['party_b_norm'].value_counts() if 'party_b_norm' in sub.columns and len(sub) else pd.Series()
        mv = sub['address'].dropna().value_counts() if 'address' in sub.columns and len(sub) else pd.Series()
        rows.append({'Time of day': name, 'Interval': interval,
                     'Total calls': len(sub),
                     'Most Contacted': f"{mc.index[0]} ({mc.iloc[0]})" if not mc.empty else 'N/A',
                     'Most Visited Place': f"{mv.index[0]} ({mv.iloc[0]})" if not mv.empty else 'N/A'})
    return pd.DataFrame(rows)

def weekly_call_count(df):
    if 'start' not in df.columns: return pd.DataFrame()
    calls = cdf(df); calls = calls[calls['is_call_out']|calls['is_call_in']].copy()
    calls['dow'] = calls['start'].dt.day_name()
    order = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday']
    rows = []
    for day in order:
        sub = calls[calls['dow']==day]
        mc = sub['party_b_clean'].value_counts() if 'party_b_clean' in sub.columns and len(sub) else pd.Series()
        mv = sub['address'].dropna().value_counts() if 'address' in sub.columns and len(sub) else pd.Series()
        rows.append({'Day': day, 'Total calls': len(sub),
                     'Most Contacted': f"{mc.index[0]} ({mc.iloc[0]})" if not mc.empty else 'N/A',
                     'Most Visited': f"{mv.index[0]} ({mv.iloc[0]})" if not mv.empty else 'N/A'})
    return pd.DataFrame(rows)

def monthly_call_count(df):
    if 'start' not in df.columns: return pd.DataFrame()
    calls = cdf(df); calls = calls[calls['is_call_out']|calls['is_call_in']].copy()
    calls['month'] = calls['start'].dt.strftime('%B')
    calls['mnum']  = calls['start'].dt.month
    order = calls[['month','mnum']].drop_duplicates().sort_values('mnum')
    rows = []
    for _, r in order.iterrows():
        sub = calls[calls['month']==r['month']]
        mc = sub['party_b_clean'].value_counts() if 'party_b_clean' in sub.columns and len(sub) else pd.Series()
        mv = sub['address'].dropna().value_counts() if 'address' in sub.columns and len(sub) else pd.Series()
        rows.append({'Month': r['month'], 'Total calls': len(sub),
                     'Most Contacted': f"{mc.index[0]} ({mc.iloc[0]})" if not mc.empty else 'N/A',
                     'Most Visited': f"{mv.index[0]} ({mv.iloc[0]})" if not mv.empty else 'N/A'})
    return pd.DataFrame(rows)

def contact_summary(df):
    if 'party_b_clean' not in df.columns: return pd.DataFrame()
    d = cdf(df)
    all_c = d[d['is_call_out']|d['is_call_in']]
    mc_all = all_c['party_b_norm'].value_counts() if 'party_b_norm' in all_c.columns else all_c['party_b_clean'].value_counts() if len(all_c) else pd.Series()
    mc_out = d[d['is_call_out']]['party_b_norm'].value_counts() if 'party_b_norm' in d.columns else d[d['is_call_out']]['party_b_clean'].value_counts() if d['is_call_out'].sum() else pd.Series()
    mc_in  = d[d['is_call_in']]['party_b_norm'].value_counts() if 'party_b_norm' in d.columns else d[d['is_call_in']]['party_b_clean'].value_counts()  if d['is_call_in'].sum()  else pd.Series()
    dur    = all_c.groupby('party_b_norm' if 'party_b_norm' in all_c.columns else 'party_b_clean')['duration'].sum() if 'duration' in d.columns and len(all_c) else pd.Series()
    return pd.DataFrame({
        'Metric': ['Total Unique Numbers','Most Called Number',
                   'Most Called Outgoing','Most Received Incoming',
                   'Most Total Call Time'],
        'Value':  [d['party_b_norm'].nunique() if 'party_b_norm' in d.columns else d['party_b_clean'].nunique(),
                   f"{display_number(mc_all.index[0])}, {mc_all.iloc[0]} times" if not mc_all.empty else 'N/A',
                   f"{display_number(mc_out.index[0])}, {mc_out.iloc[0]} times" if not mc_out.empty else 'N/A',
                   f"{display_number(mc_in.index[0])}, {mc_in.iloc[0]} times"   if not mc_in.empty  else 'N/A',
                   f"{display_number(dur.idxmax())}, {round(dur.max()/60,1)} min" if not dur.empty else 'N/A']
    })

def top_contacts(df, direction='out', n=10):
    if 'party_b_clean' not in df.columns: return pd.DataFrame()
    d = cdf(df)
    sub = d[d['is_call_out']] if direction=='out' else d[d['is_call_in']]
    if len(sub)==0: return pd.DataFrame()
    col = 'party_b_norm' if 'party_b_norm' in sub.columns else 'party_b_clean'
    c = sub[col].value_counts().head(n)
    return pd.DataFrame({'Party B': [display_number(x) for x in c.index],
                         'Total Calls': c.values,
                         'Percentage': (c.values/len(sub)*100).round(2)})

def top_lengthy(df, direction='out', n=10):
    if 'party_b_clean' not in df.columns or 'duration' not in df.columns: return pd.DataFrame()
    d = cdf(df)
    sub = d[d['is_call_out']] if direction=='out' else d[d['is_call_in']]
    if len(sub)==0: return pd.DataFrame()
    gcol = 'party_b_norm' if 'party_b_norm' in sub.columns else 'party_b_clean'
    g = sub.groupby(gcol).agg(
        Duration_Sec=('duration','sum'), Total_Calls=('duration','count')
    ).sort_values('Duration_Sec', ascending=False).head(n)
    td = sub['duration'].sum()
    g['Duration (min)'] = (g['Duration_Sec'] / 60).round(2)
    g['% of Call Time'] = (g['Duration_Sec']/td*100).round(2) if td>0 else 0
    g = g.drop(columns=['Duration_Sec'])
    g = g.reset_index().rename(columns={gcol:'Party B'})
    g['Party B'] = g['Party B'].apply(display_number)
    # Reorder columns
    g = g[['Party B', 'Total_Calls', 'Duration (min)', '% of Call Time']]
    return g

def top_call_overall(df, n=10):
    """Combined MOC + MTC table with Duration (Min) for all CDR types."""
    if 'party_b_clean' not in df.columns: return pd.DataFrame()
    d = cdf(df)
    gcol = 'party_b_norm' if 'party_b_norm' in d.columns else 'party_b_clean'
    out_df = d[d['is_call_out']]
    in_df  = d[d['is_call_in']]
    all_df = d[d['is_call_out'] | d['is_call_in']]
    if all_df.empty: return pd.DataFrame()
    moc = out_df.groupby(gcol).size().rename('MOC')
    mtc = in_df.groupby(gcol).size().rename('MTC')
    dur = all_df.groupby(gcol)['duration'].sum().rename('_dur') if 'duration' in all_df.columns else pd.Series(dtype=float)
    g = pd.concat([moc, mtc], axis=1).fillna(0).astype(int)
    g['Total Calls'] = g.get('MOC', 0) + g.get('MTC', 0)
    if not dur.empty:
        g = g.join(dur)
        g['Duration (Min)'] = (g['_dur'] / 60).round(1)
        g = g.drop(columns=['_dur'])
    else:
        g['Duration (Min)'] = 0.0
    g = g.sort_values('Total Calls', ascending=False).head(n).reset_index()
    g = g.rename(columns={gcol: 'Party B'})
    g['Party B'] = g['Party B'].apply(display_number)
    cols = ['Party B', 'MOC', 'MTC', 'Total Calls', 'Duration (Min)']
    return g[[c for c in cols if c in g.columns]]

def plot_top_call_overall(df, n=10):
    import matplotlib.ticker as ticker
    t = top_call_overall(df, n)
    if t.empty: return None
    fig, ax = plt.subplots(figsize=(14, 6))
    x      = list(range(len(t)))
    w      = 0.25
    labels = t['Party B'].tolist()
    moc_v  = t['MOC'].tolist()          if 'MOC'    in t.columns else [0]*len(t)
    mtc_v  = t['MTC'].tolist()          if 'MTC'    in t.columns else [0]*len(t)
    tot_v  = t['Total Calls'].tolist()
    ax.bar([i - w for i in x], moc_v, width=w, label='MOC (Outgoing)', color='#2196F3', zorder=3)
    ax.bar([i     for i in x], mtc_v, width=w, label='MTC (Incoming)', color='#4CAF50', zorder=3)
    ax.bar([i + w for i in x], tot_v, width=w, label='Total Calls',    color='#FF9800', zorder=3)
    ax2 = ax.twinx()
    if 'Duration (Min)' in t.columns:
        ax2.plot(x, t['Duration (Min)'].tolist(),
                 color='#E91E63', marker='o', linewidth=2, label='Duration (Min)', zorder=4)
        ax2.set_ylabel('Duration (Min)', fontsize=11, color='#E91E63')
        ax2.tick_params(axis='y', labelcolor='#E91E63')
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=30, ha='right', fontsize=9)
    ax.set_xlabel('Phone Number', fontsize=11)
    ax.set_ylabel('Call Count', fontsize=11)
    ax.set_title('5.6 Top Call Overall (MOC + MTC + Duration)', fontsize=13, fontweight='bold', pad=12)
    ax.yaxis.set_major_locator(ticker.MaxNLocator(integer=True))
    ax.grid(axis='y', linestyle='--', alpha=0.4, zorder=0)
    ax.set_axisbelow(True)
    lines1, labels1 = ax.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax.legend(lines1 + lines2, labels1 + labels2, loc='upper right', fontsize=9)
    plt.tight_layout()
    return fig

def top_locations(df, mask=None, n=10):
    if 'address' not in df.columns: return pd.DataFrame()
    data = df if mask is None else df[mask]
    valid = data['address'].dropna().apply(lambda a: a if _is_valid_address(a) else None).dropna()
    c = valid.value_counts().head(n)
    if c.empty: return pd.DataFrame()
    return pd.DataFrame({'Address': c.index, 'Count': c.values})

def location_summary(df):
    if 'address' not in df.columns: return pd.DataFrame()
    addrs = df['address'].dropna().apply(lambda a: a if _is_valid_address(a) else None).dropna()
    if len(addrs)==0: return pd.DataFrame()
    mv = addrs.value_counts()

    def top_addr(mask):
        if mask is None or 'start' not in df.columns: return ('N/A', 0)
        sub = df[mask]['address'].dropna().apply(
            lambda a: a if _is_valid_address(a) else None).dropna().value_counts()
        return (sub.index[0], int(sub.iloc[0])) if not sub.empty else ('N/A', 0)

    home_mask    = df['start'].dt.hour.astype(int).isin(list(range(0,6))+list(range(22,24))) if 'start' in df.columns else None
    work_mask    = (df['start'].dt.hour.astype(int)>=8)&(df['start'].dt.hour.astype(int)<18) if 'start' in df.columns else None
    weekend_mask = df['start'].dt.dayofweek.astype(int).isin([4,5]) if 'start' in df.columns else None

    h, hc = top_addr(home_mask)
    w, wc = top_addr(work_mask)
    e, ec = top_addr(weekend_mask)
    return pd.DataFrame({
        'Metric':  ['Total Towers Visited','Most Visited Place',
                    'Probable Home','Probable Work','Probable Weekend'],
        'Location': [df['cell_id'].nunique() if 'cell_id' in df.columns else 'N/A',
                     mv.index[0] if not mv.empty else 'N/A', h, w, e],
        'Count':    ['N/A', int(mv.iloc[0]) if not mv.empty else 0, hc, wc, ec]
    })


def top_sms_contacts(df, direction='out', n=5):
    """Top SMS contacts with strict valid number filter.
    Bangladesh mobile pattern: 8801XXXXXXXXX (13 digits) or 01XXXXXXXXX (11 digits).
    Also accepts other 10-15 digit international formats but excludes shortcodes/random hex.
    """
    if 'party_b_clean' not in df.columns:
        return pd.DataFrame()
    col = 'is_sms_out' if direction == 'out' else 'is_sms_in'
    if col not in df.columns:
        return pd.DataFrame()

    def is_real_mobile(val):
        """
        Strict Bangladesh mobile number filter.
        Valid formats:
          - 8801XXXXXXXXX  → 13 digits, starts with 8801
          - 01XXXXXXXXX    → 11 digits, starts with 01
          - 1XXXXXXXXX     → 10 digits, starts with 1 (bare BD number)
        All other patterns (service codes, intl shortcodes) are excluded.
        """
        s = str(val).strip()
        if not s.isdigit():
            return False
        # 13 digits: must start with 8801
        if len(s) == 13 and s.startswith('8801'):
            return True
        # 11 digits: must start with 01
        if len(s) == 11 and s.startswith('01'):
            return True
        # 10 digits: must start with 1 (bare Bangladesh number)
        if len(s) == 10 and s.startswith('1'):
            return True
        return False

    sub = df[df[col] & df['party_b_clean'].apply(is_real_mobile)]
    if len(sub) == 0:
        return pd.DataFrame()
    counts = sub['party_b_clean'].value_counts().head(n)
    return pd.DataFrame({'Party B': [display_number(x) for x in counts.index],
                         'SMS Count': counts.values})


def last_n_days_top_contacts(df, days=10, n=10):
    """Top contacts (MOC + MTC combined) in last N days."""
    if 'start' not in df.columns or 'party_b_clean' not in df.columns:
        return pd.DataFrame()
    d = cdf(df)
    max_date = d['start'].max()
    cutoff = max_date - pd.Timedelta(days=days)
    recent = d[(d['start'] >= cutoff) & (d['is_call_out'] | d['is_call_in'])].copy()
    if len(recent) == 0:
        return pd.DataFrame()
    gcol = 'party_b_norm' if 'party_b_norm' in recent.columns else 'party_b_clean'
    grouped = recent.groupby(gcol).agg(
        moc=('is_call_out', 'sum'),
        mtc=('is_call_in', 'sum'),
        total_duration=('duration', 'sum')
    )
    grouped['total_calls'] = grouped['moc'] + grouped['mtc']
    grouped = grouped.sort_values('total_calls', ascending=False).head(n).reset_index()
    grouped['Duration (min)'] = (grouped['total_duration'] / 60).round(1)
    result = grouped[[gcol, 'moc', 'mtc', 'total_calls', 'Duration (min)']].rename(
        columns={gcol: 'Party B', 'moc': 'MOC',
                 'mtc': 'MTC', 'total_calls': 'Total Calls'})
    result['Party B'] = result['Party B'].apply(display_number)
    return result


def last_n_days_top_locations(df, days=10, n=10):
    """Top locations in last N days."""
    if 'start' not in df.columns or 'address' not in df.columns:
        return pd.DataFrame()
    max_date = df['start'].max()
    cutoff = max_date - pd.Timedelta(days=days)
    recent = df[df['start'] >= cutoff]
    addrs = (recent['address'].dropna()
             .apply(lambda a: a if _is_valid_address(a) else None)
             .dropna()
             .value_counts().head(n))
    if addrs.empty:
        return pd.DataFrame()
    return pd.DataFrame({'Address': addrs.index, 'Count': addrs.values})


def specific_number_analysis(df, target_number):
    """Analyze interactions with a specific phone number (MOC, MTC, duration, SMS)."""
    if 'party_b_clean' not in df.columns or not target_number:
        return None
    target = re.sub(r'[^0-9]', '', str(target_number).strip())
    if len(target) < 10:
        return None
    d = cdf(df)
    candidates = {target}
    if target.startswith('880'):  candidates.add('0' + target[3:])
    elif target.startswith('0'):  candidates.add('880' + target[1:])
    if target.startswith('880'):  candidates.add(target[3:])
    candidates.add(target.lstrip('0'))
    # Match using normalized last-10-digit form
    norm_target = normalize_number(target)
    if 'party_b_norm' in d.columns:
        sub = d[d['party_b_norm'].astype(str) == norm_target]
    else:
        sub = d[d['party_b_clean'].astype(str).isin(candidates)]
    if len(sub) == 0:
        return None
    moc = int(sub.get('is_call_out', pd.Series([False]*len(sub))).sum())
    mtc = int(sub.get('is_call_in',  pd.Series([False]*len(sub))).sum())
    sms_out = int(sub.get('is_sms_out', pd.Series([False]*len(sub))).sum())
    sms_in  = int(sub.get('is_sms_in',  pd.Series([False]*len(sub))).sum())
    call_mask = sub.get('is_call_out', False) | sub.get('is_call_in', False)
    total_dur = int(sub.loc[call_mask, 'duration'].sum()) if 'duration' in sub.columns else 0
    return {
        'number': display_number(target_number),
        'moc': moc, 'mtc': mtc,
        'total_calls': moc + mtc,
        'total_duration_sec': total_dur,
        'total_duration_min': round(total_dur / 60, 2),
        'sms_sent': sms_out, 'sms_received': sms_in,
        'total_sms': sms_out + sms_in,
        'first_contact': sub['start'].min().strftime('%Y-%m-%d %H:%M:%S') if 'start' in sub.columns and len(sub) else 'N/A',
        'last_contact':  sub['start'].max().strftime('%Y-%m-%d %H:%M:%S') if 'start' in sub.columns and len(sub) else 'N/A',
    }

# ─────────────────────────────────────────────
# GRAPH FUNCTIONS
# ─────────────────────────────────────────────
def plot_hourly(df):
    if 'start' not in df.columns: return None
    calls = cdf(df); calls = calls[calls['is_call_out']|calls['is_call_in']].copy()
    calls['hour'] = calls['start'].dt.hour.astype(int)
    hourly = calls.groupby('hour').size().reindex(range(24), fill_value=0)
    fig, ax = plt.subplots(figsize=(12,4))
    ax.bar(hourly.index, hourly.values, color='steelblue', edgecolor='white')
    ax.set_xlabel('Hour of Day'); ax.set_ylabel('Calls')
    ax.set_title('Hourly Call Count')
    ax.set_xticks(range(24))
    ax.set_xticklabels([f'{h:02d}:00' for h in range(24)], rotation=45, ha='right', fontsize=8)
    ax.yaxis.set_major_locator(ticker.MaxNLocator(integer=True))
    plt.tight_layout()
    return fig

def plot_weekly(df):
    if 'start' not in df.columns: return None
    calls = cdf(df); calls = calls[calls['is_call_out']|calls['is_call_in']].copy()
    calls['dow'] = calls['start'].dt.day_name()
    order = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday']
    counts = calls['dow'].value_counts().reindex(order, fill_value=0)
    fig, ax = plt.subplots(figsize=(10,4))
    ax.bar(counts.index, counts.values, color='coral', edgecolor='white')
    ax.set_ylabel('Calls'); ax.set_title('Weekly Call Count')
    ax.yaxis.set_major_locator(ticker.MaxNLocator(integer=True))
    plt.tight_layout()
    return fig

def plot_monthly(df):
    if 'start' not in df.columns: return None
    calls = cdf(df); calls = calls[calls['is_call_out']|calls['is_call_in']].copy()
    calls['mp'] = calls['start'].dt.to_period('M')
    counts = calls.groupby('mp').size()
    fig, ax = plt.subplots(figsize=(max(8,len(counts)*2), 4))
    ax.bar([str(p) for p in counts.index], counts.values,
           color='mediumseagreen', edgecolor='white')
    ax.set_ylabel('Calls'); ax.set_title('Monthly Call Count')
    ax.yaxis.set_major_locator(ticker.MaxNLocator(integer=True))
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    return fig

def plot_contacts(df, direction='out', n=10, title='Top Contacts'):
    if 'party_b_clean' not in df.columns: return None
    d = cdf(df)
    sub = d[d['is_call_out']] if direction=='out' else d[d['is_call_in']]
    if len(sub)==0: return None
    c = sub['party_b_clean'].value_counts().head(n)
    fig, ax = plt.subplots(figsize=(12,5))
    bars = ax.barh(c.index[::-1], c.values[::-1], color='royalblue', edgecolor='white')
    ax.set_xlabel('Calls'); ax.set_title(title)
    ax.xaxis.set_major_locator(ticker.MaxNLocator(integer=True))
    for bar in bars:
        ax.text(bar.get_width()+0.1, bar.get_y()+bar.get_height()/2,
                f'{int(bar.get_width())}', va='center', fontsize=9)
    plt.tight_layout()
    return fig

def plot_locations(df, mask=None, title='Top Locations', n=10):
    if 'address' not in df.columns: return None
    data = df if mask is None else df[mask]
    c = (data['address'].dropna()
         .apply(lambda a: a if _is_valid_address(a) else None)
         .dropna().value_counts().head(n))
    if c.empty: return None
    labels = [str(a)[:40]+'...' if len(str(a))>40 else str(a) for a in c.index[::-1]]
    fig, ax = plt.subplots(figsize=(12, max(4, n*0.5)))
    ax.barh(labels, c.values[::-1], color='darkorange', edgecolor='white')
    ax.set_xlabel('Visits'); ax.set_title(title)
    ax.xaxis.set_major_locator(ticker.MaxNLocator(integer=True))
    plt.tight_layout()
    return fig


# ─────────────────────────────────────────────
# HTML GENERATOR
# ─────────────────────────────────────────────
CSS = """
body{margin:40px;font-family:Arial,sans-serif;font-size:14px;}
table{width:100%;border-collapse:collapse;margin-bottom:20px;}
table,th,td{border:1px solid #ccc;padding:8px;text-align:left;}
th{background:#2E74B5;color:white;font-weight:bold;}
tr:nth-child(even){background:#EBF3FB;}
h1{text-align:center;color:#1F3864;}
h2{border-bottom:2px solid #2E74B5;padding-bottom:4px;color:#1F3864;}
h3{color:#2E74B5;}
img{max-width:100%;}
.info-box{background:#e8f4f8;padding:12px;border-radius:6px;margin-bottom:15px;border-left:4px solid #2E74B5;}
.warning{color:#cc6600;font-style:italic;}
"""

def df_to_html(df):
    if df is None or df.empty: return '<p class="warning">No data available.</p>'
    return df.to_html(index=False, border=0)

def fig_to_html_img(fig):
    if fig is None: return '<p class="warning">Graph not available (insufficient data).</p>'
    return f'<img src="data:image/png;base64,{fig_to_base64(fig)}">'



# ─────────────────────────────────────────────
# MOVEMENT PATTERN ANALYSIS
# ─────────────────────────────────────────────

def _extract_district(address_str):
    """Extract district name from BTS address string."""
    if not address_str or str(address_str).strip() in ('', 'nan', '-'):
        return None
    addr = str(address_str).upper()

    # Bangladesh districts
    districts = [
        'DHAKA','CHITTAGONG','CHATTOGRAM','RAJSHAHI','KHULNA','SYLHET',
        'BARISHAL','BARISAL','MYMENSINGH','RANGPUR','COMILLA','CUMILLA',
        'GAZIPUR','NARAYANGANJ','NARSINGDI','MANIKGANJ','MUNSHIGANJ',
        'TANGAIL','KISHOREGANJ','NETROKONA','JAMALPUR','SHERPUR',
        'FARIDPUR','GOPALGANJ','MADARIPUR','SHARIATPUR','RAJBARI',
        'BOGURA','BOGRA','JOYPURHAT','NAOGAON','NATORE','CHAPAINAWABGANJ',
        'SIRAJGANJ','PABNA','KURIGRAM','LALMONIRHAT','NILPHAMARI',
        'GAIBANDHA','THAKURGAON','DINAJPUR','PANCHAGARH','RANGPUR',
        'CHUADANGA','JHENAIDAH','JESSORE','JASHORE','MAGURA','NARAIL',
        'SATKHIRA','BAGERHAT','KHULNA','KUSHTIA','MEHERPUR',
        'SUNAMGANJ','MOULVIBAZAR','HABIGANJ','BRAHMANBARIA',
        'CHANDPUR','LAKSHMIPUR','NOAKHALI','FENI','COX\'S BAZAR',
        "COX'S BAZAR",'BANDARBAN','RANGAMATI','KHAGRACHHARI',
        'PIROJPUR','JHALOKATHI','BARGUNA','PATUAKHALI','BHOLA','BARISAL',
        'NETRAKONA','MYMENSINGH','COX'
    ]
    for d in districts:
        if d in addr:
            return d.title()
    return None


def _home_district(df):
    """Detect home district from night-time BTS addresses."""
    if 'address' not in df.columns or 'start' not in df.columns:
        return None
    night = df[df['start'].dt.hour.astype(int).isin(list(range(0,6))+list(range(22,24)))]
    if night.empty:
        night = df
    addr_counts = (night['address'].dropna()
                   .apply(lambda a: a if _is_valid_address(a) else None)
                   .dropna().value_counts())
    for addr in addr_counts.index:
        d = _extract_district(addr)
        if d:
            return d
    return None


def _work_district(df):
    """Detect work district from daytime BTS addresses."""
    if 'address' not in df.columns or 'start' not in df.columns:
        return None
    day = df[(df['start'].dt.hour.astype(int) >= 8) & (df['start'].dt.hour.astype(int) < 18)]
    if day.empty:
        return None
    addr_counts = (day['address'].dropna()
                   .apply(lambda a: a if _is_valid_address(a) else None)
                   .dropna().value_counts())
    for addr in addr_counts.index:
        d = _extract_district(addr)
        if d:
            return d
    return None


def _is_valid_address(addr):
    """Check if BTS address is meaningful (not just dashes, commas, or empty)."""
    if not addr or pd.isna(addr):
        return False
    s = str(addr).strip()
    # Remove common invalid patterns: -, -,, --,  nan, empty
    cleaned = s.replace('-', '').replace(',', '').replace('.', '').replace(' ', '')
    return len(cleaned) >= 5  # must have at least 5 meaningful chars


def generate_overall_comment(df, phone=None, operator=None, date_range=None, total_raw=None):
    """
    Generates a structured, data-driven overall comment for the CDR report.
    Covers: identity, activity, contacts, time behaviour, location, network, anomalies.
    """
    d = cdf(df)
    lines = []

    # ── 1. Identity & Period ─────────────────────────────────────────────
    ph   = phone    or get_phone(df)
    op   = operator or get_operator(df)
    dr   = date_range or get_date_range(df)
    raw  = total_raw or len(df)
    days_span = 0
    if 'start' in df.columns and df['start'].notna().any():
        dmin = df['start'].min()
        dmax = df['start'].max()
        days_span = max((dmax - dmin).days + 1, 1)

    if is_imei_cdr(df):
        # IMEI-based CDR — special identity paragraph
        imei_val = df['imei'].dropna().mode()[0] if 'imei' in df.columns else 'N/A'
        sims = [s for s in df['party_a'].dropna().value_counts().index
                if str(s).replace('+','').isdigit() and len(str(s)) >= 10]
        sim_detail = '; '.join(
            f"{s} ({op_v})" for s, op_v in
            [(s, df[df['party_a']==s]['operator'].dropna().mode()[0]
              if len(df[df['party_a']==s]['operator'].dropna()) > 0 else 'N/A')
             for s in sims[:3]]
        )
        lines.append(
            f"This CDR report is IMEI-based, pertaining to device with IMEI {imei_val}. "
            f"The device was used with {len(sims)} SIM card(s) during the analysis period: {sim_detail}. "
            f"The analysis covers {days_span} days ({dr}), "
            f"with a total of {raw:,} raw records processed across all SIMs."
        )
    else:
        lines.append(
            f"This CDR report pertains to subscriber {ph} operating on the {op} network. "
            f"The analysis covers a period of {days_span} days ({dr}), "
            f"with a total of {raw:,} raw records processed."
        )

    # ── 2. Activity Overview ─────────────────────────────────────────────
    n_out  = int(d['is_call_out'].sum())  if 'is_call_out' in d.columns else 0
    n_in   = int(d['is_call_in'].sum())   if 'is_call_in'  in d.columns else 0
    n_sout = int(d['is_sms_out'].sum())   if 'is_sms_out'  in d.columns else 0
    n_sin  = int(d['is_sms_in'].sum())    if 'is_sms_in'   in d.columns else 0
    n_total_calls = n_out + n_in
    n_total_sms   = n_sout + n_sin

    dur_total = 0
    if 'duration' in d.columns:
        dur_total = int(d.loc[d['is_call_out'] | d['is_call_in'], 'duration'].sum()) if n_total_calls > 0 else 0

    daily_avg = round(n_total_calls / days_span, 1) if days_span > 0 else 0

    lines.append(
        f"The subscriber made {n_out:,} outgoing calls (MOC) and received {n_in:,} incoming calls (MTC), "
        f"totalling {n_total_calls:,} call events with a combined talk time of "
        f"{round(dur_total/60, 1):,} minutes ({round(dur_total/3600, 1)} hours). "
        f"Average daily call activity was {daily_avg} calls/day. "
        f"SMS activity recorded {n_sout} sent and {n_sin} received messages."
    )

    # ── 3. Contact Behaviour ─────────────────────────────────────────────
    gcol = 'party_b_norm' if 'party_b_norm' in d.columns else 'party_b_clean'
    contact_comment = ""
    if gcol in d.columns:
        all_calls = d[d['is_call_out'] | d['is_call_in']]
        if not all_calls.empty:
            vc = all_calls[gcol].value_counts()
            unique_contacts = len(vc)
            top_num   = display_number(vc.index[0]) if len(vc) > 0 else 'N/A'
            top_count = int(vc.iloc[0])             if len(vc) > 0 else 0
            top_pct   = round(top_count / n_total_calls * 100, 1) if n_total_calls > 0 else 0
            top2_pct  = round(vc.iloc[:3].sum() / n_total_calls * 100, 1) if len(vc) >= 3 and n_total_calls > 0 else 0

            dep_note = ""
            if top_pct >= 25:
                dep_note = (f" This indicates a high dependency on a single contact, "
                            f"with the top number accounting for {top_pct}% of all calls.")
            elif top_pct >= 15:
                dep_note = f" The top contact accounts for {top_pct}% of total call activity."

            contact_comment = (
                f"The subscriber communicated with {unique_contacts} unique numbers. "
                f"The most frequently contacted number is {top_num} with {top_count:,} call events ({top_pct}% of total).{dep_note} "
                f"The top 3 contacts collectively account for {top2_pct}% of all calls."
            )
    if contact_comment:
        lines.append(contact_comment)

    # ── 4. Time-of-Day Behaviour ─────────────────────────────────────────
    if 'start' in d.columns and d['start'].notna().any():
        d2 = d.copy()
        d2['hour'] = d2['start'].dt.hour
        slot_map = {
            'Night (22:00–05:00)':   d2['hour'].apply(lambda h: h >= 22 or h < 5),
            'Morning (05:00–08:00)': d2['hour'].apply(lambda h: 5 <= h < 8),
            'Day (08:00–18:00)':     d2['hour'].apply(lambda h: 8 <= h < 18),
            'Evening (18:00–22:00)': d2['hour'].apply(lambda h: 18 <= h < 22),
        }
        slot_counts = {k: int(v.sum()) for k, v in slot_map.items()}
        peak_slot   = max(slot_counts, key=slot_counts.get)
        peak_count  = slot_counts[peak_slot]
        peak_pct    = round(peak_count / len(d2) * 100, 1) if len(d2) > 0 else 0
        night_pct   = round(slot_counts.get('Night (22:00–05:00)', 0) / len(d2) * 100, 1) if len(d2) > 0 else 0

        night_note = ""
        if night_pct >= 10:
            night_note = (f" Notably, {night_pct}% of activity occurred during night hours (22:00–05:00), "
                          f"which may warrant further attention.")

        lines.append(
            f"Peak communication activity was recorded during {peak_slot}, "
            f"accounting for {peak_pct}% of all events.{night_note}"
        )

    # ── 5. Location & Movement ───────────────────────────────────────────
    mv = movement_pattern_analysis(df)
    if mv:
        home   = mv.get('home_district') or 'unknown'
        work   = mv.get('work_district') or 'unknown'
        trips  = mv.get('trips', [])
        gaps   = mv.get('gaps', [])
        same   = home.lower() == work.lower()
        loc_comment = (
            f"Location analysis indicates the subscriber's estimated home district as {home}"
            + (f" and work district as {work}." if not same else ", with work activity also centred in the same district.")
        )
        if trips:
            districts = list({t['district'] for t in trips})
            loc_comment += (f" The subscriber travelled outside the home district on {len(trips)} occasion(s), "
                            f"visiting: {', '.join(districts[:5])}.")
        if gaps:
            loc_comment += (f" {len(gaps)} network disconnection period(s) of more than 4 consecutive days "
                            f"were detected, which may indicate travel, device change, or SIM inactivity.")
        lines.append(loc_comment)

    # ── 6. Network Technology ────────────────────────────────────────────
    if 'cell_type' in df.columns and df['cell_type'].notna().any():
        # Filter valid tech values: 2G, 3G, 4G, 5G only
        valid_techs = ['2G','3G','4G','5G','LTE','WCDMA','GSM','NR']
        tech_vc = (df['cell_type'].dropna().str.upper().str.strip()
                   .apply(lambda x: x if any(t in x for t in valid_techs) else None)
                   .dropna().value_counts())
        if not tech_vc.empty:
            tech_str = ', '.join([f"{v} ({int(c):,} records)" for v, c in tech_vc.head(3).items()])
            lines.append(f"Network technology usage: {tech_str}.")

    # ── 7. IMEI / Device Note ────────────────────────────────────────────
    if 'imei' in df.columns:
        imei_vals = df['imei'].dropna().unique()
        imei_vals = [str(i) for i in imei_vals if str(i).strip() not in ('', '-', 'nan')]
        if len(imei_vals) > 1:
            lines.append(
                f"Multiple IMEI values ({len(imei_vals)}) detected for this subscriber, "
                f"suggesting possible device changes or use of multiple handsets during the analysis period."
            )
        elif len(imei_vals) == 1:
            lines.append(f"A single device (IMEI: {imei_vals[0]}) was used throughout the analysis period.")

    # ── 8. Closing ───────────────────────────────────────────────────────
    lines.append(
        "The above observations are derived solely from Call Detail Records (CDR) provided for analysis. "
        "This report is intended for investigative/analytical purposes and should be interpreted "
        "in conjunction with other available evidence."
    )

    return lines


def generate_recommendation(df):
    """Returns a list of recommendation strings based on CDR patterns."""
    d = cdf(df)
    recs = []

    # Top contact dependency
    gcol = 'party_b_norm' if 'party_b_norm' in d.columns else 'party_b_clean'
    if gcol in d.columns:
        all_calls  = d[d['is_call_out'] | d['is_call_in']]
        n_total    = len(all_calls)
        if n_total > 0:
            vc       = all_calls[gcol].value_counts()
            top_pct  = round(vc.iloc[0] / n_total * 100, 1) if len(vc) > 0 else 0
            if top_pct >= 20:
                recs.append(
                    f"Investigate the relationship between the subscriber and the top contact "
                    f"({display_number(vc.index[0])}) — {top_pct}% of all calls directed to/from "
                    f"a single number indicates a significant association."
                )

    # Night activity
    if 'start' in d.columns and d['start'].notna().any():
        d2 = d.copy()
        d2['hour'] = d2['start'].dt.hour
        night_pct = round(((d2['hour'] >= 22) | (d2['hour'] < 5)).sum() / len(d2) * 100, 1)
        if night_pct >= 10:
            recs.append(
                f"Significant night-time activity ({night_pct}% of events) detected. "
                f"Cross-reference night-time contacts and locations with case context."
            )

    # Movement
    mv = movement_pattern_analysis(df)
    if mv and mv.get('trips'):
        recs.append(
            f"Out-of-district travel recorded ({len(mv['trips'])} trip(s)). "
            f"Verify travel dates against case timeline and obtain tower dump data for visited districts if required."
        )
    if mv and mv.get('gaps'):
        recs.append(
            f"Network disconnection gap(s) of 4+ days detected. "
            f"Verify whether subscriber was using an alternate SIM or was unreachable during these periods."
        )

    # Multiple IMEI
    if 'imei' in df.columns:
        imei_vals = [str(i) for i in df['imei'].dropna().unique() if str(i).strip() not in ('', '-', 'nan')]
        if len(imei_vals) > 1:
            recs.append(
                f"Multiple devices (IMEI count: {len(imei_vals)}) detected. "
                f"Obtain CDR for all associated IMEIs to ensure complete communication picture."
            )

    if not recs:
        recs.append("No specific anomalies detected. Continue routine monitoring as required.")

    return recs


def movement_pattern_analysis(df):
    """
    Distance-based movement pattern analysis.
    - Home coord from most frequent night-time BTS address.
    - Out-of-home: any location >= 35 km from home coord.
    - Transit < 6 hours at intermediate stop: ignored.
    - Destination = furthest location in each away-session (upazila preferred).
    - Table in REVERSE chronological order.
    - Network gaps > 4 days flagged.
    """
    if "address" not in df.columns or "start" not in df.columns:
        return None

    import math

    # ── Comprehensive Bangladesh coordinates (all 64 districts + key upazilas) ──
    BD_COORDS = {
        # Kurigram
        "Rowmari":(25.5964,89.7662),"Chilmari":(25.5555,89.6836),
        "Rajibpur":(25.6580,89.8401),"Ulipur":(25.6717,89.5718),
        "Nageshwari":(25.9711,89.7039),"Bhurungamari":(26.0688,89.7164),
        "Rajarhat":(25.7594,89.4952),"Phulbari":(25.8654,89.4620),
        "Kurigram Sadar":(25.8057,89.6360),
        # Gaibandha
        "Sundarganj":(25.3810,89.4670),"Sadullapur":(25.1580,89.4887),
        "Gaibandha Sadar":(25.3288,89.5288),"Gobindaganj":(25.1175,89.3590),
        "Palashbari":(25.2116,89.3918),"Fulchhari":(25.1780,89.5420),
        # Rangpur
        "Rangpur Sadar":(25.7439,89.2752),"Pirganj":(25.8538,89.0346),
        "Pirgacha":(25.7011,89.3840),"Mahiganj":(25.7671,89.2387),
        "Gangachara":(25.7208,89.2019),"Kaunia":(25.6452,89.2884),
        "Mithapukur":(25.6046,89.1961),"Badarganj":(25.6754,89.0548),
        "Taraganj":(25.9302,89.1630),
        # Lalmonirhat
        "Lalmonirhat Sadar":(25.9923,89.2847),"Aditmari":(25.9042,89.3521),
        "Kaliganj":(25.8622,89.4014),"Hatibandha":(26.0551,89.4688),
        "Patgram":(26.1800,89.5127),
        # Nilphamari
        "Nilphamari Sadar":(25.9315,88.8560),"Saidpur":(25.7778,88.8879),
        "Jaldhaka":(25.8596,89.0196),"Domar":(25.9963,88.9601),
        "Kishoreganj Nilphamari":(25.9992,88.8773),"Dimla":(25.9167,88.9931),
        # Dinajpur
        "Dinajpur Sadar":(25.6279,88.6333),"Birampur":(25.4857,88.6987),
        "Birganj":(25.8344,88.7248),"Bochaganj":(25.5519,88.6993),
        "Chirirbandar":(25.6719,88.5633),"Ghoraghat":(25.3456,88.9993),
        "Hakimpur":(25.5248,88.9333),"Kaharole":(25.7256,88.5867),
        "Khansama":(25.9024,88.7122),"Nawabganj Dinajpur":(24.5972,88.2819),
        "Parbatipur":(25.6496,88.9140),"Phulbari Dinajpur":(25.1982,88.6349),
        "Fulbari":(25.1982,88.6349),
        # Thakurgaon
        "Thakurgaon Sadar":(26.0318,88.4582),"Baliadangi":(26.1500,88.3750),
        "Haripur":(26.2073,88.4300),"Pirganj Thakurgaon":(26.0167,88.3500),
        "Ranisankail":(26.0706,88.6500),
        # Panchagarh
        "Panchagarh Sadar":(26.3406,88.5549),"Atwari":(26.5833,88.5667),
        "Boda":(26.3167,88.7333),"Debiganj":(26.0333,88.5333),
        "Tetulia":(26.6337,88.6303),
        # Rajshahi
        "Rajshahi Sadar":(24.3745,88.6042),"Bagha":(24.3167,88.8333),
        "Bagmara":(24.4500,88.6833),"Charghat":(24.2667,88.7833),
        "Durgapur":(24.8500,88.7500),"Godagari":(24.4833,88.3833),
        "Mohanpur":(24.3500,88.6833),"Paba":(24.3667,88.5833),
        "Puthia":(24.3667,88.8500),"Tanore":(24.5333,88.5667),
        # Chapainawabganj
        "Chapainawabganj Sadar":(24.5965,88.2787),"Bholahat":(24.6667,88.2833),
        "Gomastapur":(24.8500,88.2000),"Nachole":(24.7333,88.3000),
        "Shibganj Chapai":(24.7667,88.1500),
        # Naogaon
        "Naogaon Sadar":(24.9131,88.7527),"Atrai":(24.6000,88.9000),
        "Badalgachhi":(25.0167,88.6833),"Dhamoirhat":(25.2167,88.7833),
        "Manda":(24.7667,88.8500),"Mahadebpur":(25.0167,88.5833),
        "Mohadevpur":(25.0167,88.5833),"Niamatpur":(24.9667,88.5000),
        "Patnitala":(25.0833,88.6333),"Porsha":(25.2333,88.5833),
        "Raninagar":(24.6833,88.6833),"Sapahar":(25.0667,88.4167),
        # Natore
        "Natore Sadar":(24.4200,89.0019),"Bagatipara":(24.3833,89.2500),
        "Baraigram":(24.4167,89.1833),"Gurudaspur":(24.2833,89.0333),
        "Lalpur":(24.1500,89.0167),"Singra":(24.4667,89.1500),
        # Sirajganj
        "Sirajganj Sadar":(24.4534,89.7006),"Belkuchi":(24.5333,89.5833),
        "Chauhali":(24.3167,89.8667),"Kamarkhanda":(24.4333,89.5333),
        "Kazipur":(24.6167,89.7000),"Raiganj":(24.5833,89.8167),
        "Shahjadpur":(24.2167,89.6500),"Tarash":(24.2333,89.4833),
        "Ullahpara":(24.3333,89.5333),
        # Pabna
        "Pabna Sadar":(24.0063,89.2372),"Atgharia":(24.0833,89.3667),
        "Bera":(24.0833,89.6667),"Bhangura":(24.2500,89.2500),
        "Chatmohar":(24.1500,89.3167),"Faridpur Pabna":(24.1667,89.1833),
        "Ishwardi":(24.1333,89.0667),"Santhia":(24.1167,89.3500),
        "Sujanagar":(23.9000,89.3833),
        # Joypurhat
        "Joypurhat Sadar":(25.0964,89.0222),"Akkelpur":(25.0333,89.0500),
        "Kalai":(25.0833,88.9333),"Khetlal":(25.0000,89.1500),
        "Panchbibi":(25.1833,88.9500),
        # Bogura
        "Bogura Sadar":(24.8465,89.3776),"Adamdighi":(24.9167,89.4667),
        "Dhunat":(24.7000,89.5167),"Dhupchanchia":(24.9000,89.2500),
        "Gabtali":(24.8500,89.5833),"Kahaloo":(24.9500,89.1833),
        "Nandigram":(24.7833,89.4333),"Sarial":(24.9833,89.4667),
        "Sariakandi":(24.9167,89.6000),"Shajahanpur":(24.8667,89.3000),
        "Sherpur Bogura":(24.9167,89.5167),"Shibganj Bogura":(25.0167,89.1833),
        "Sonatala":(25.0167,89.6000),
        # Tangail
        "Tangail Sadar":(24.2512,89.9167),"Basail":(24.2167,90.0500),
        "Bhuapur":(24.5333,89.8833),"Delduar":(24.1667,89.9667),
        "Ghatail":(24.4500,89.9667),"Gopalpur":(24.5167,90.0833),
        "Kalihati":(24.3167,89.9833),"Madhupur":(24.6333,90.0333),
        "Mirzapur":(24.0833,90.0500),"Nagarpur":(24.1333,89.8333),
        "Sakhipur":(24.2500,90.1833),"Dhanbari":(24.5000,90.1500),
        # Jamalpur
        "Jamalpur Sadar":(24.8966,89.9441),"Bakshiganj":(25.0667,89.7333),
        "Dewanganj":(25.0500,89.7833),"Islampur":(24.9333,89.7000),
        "Madarganj":(24.8833,89.7333),"Melandaha":(24.9833,89.8667),
        "Sarishabari":(24.6500,89.6500),
        # Mymensingh
        "Mymensingh Sadar":(24.7471,90.4203),"Trishal":(24.5469,90.3455),
        "Bhaluka":(24.4005,90.3715),"Gaffargaon":(24.4667,90.5333),
        "Gauripur":(24.7962,90.2638),"Gouripur":(24.7962,90.2638),
        "Haluaghat":(25.0667,90.5833),"Ishwarganj":(24.5667,90.6667),
        "Iswarganj":(24.5667,90.6667),"Ishwargonj":(24.5667,90.6667),
        "Muktagacha":(24.7667,90.2667),"Nandail":(24.4667,90.7500),
        "Phulbaria":(24.7833,90.2500),"Phulpur":(25.0000,90.5333),
        "Fulbaria":(24.7833,90.2500),
        # Netrokona
        "Netrokona Sadar":(24.8704,90.7270),"Atpara":(24.9500,91.0167),
        "Barhatta":(24.9500,90.8500),"Durgapur":(25.0833,90.6000),
        "Khaliajuri":(24.7000,91.0833),"Kalmakanda":(25.0167,91.1500),
        "Kendua":(24.7167,90.9667),"Kolmakanda":(25.0167,91.1500),
        "Madan":(24.6500,90.9333),"Mohanganj":(24.6667,91.0333),
        "Purbadhala":(24.9167,90.8500),"Chandua":(24.7167,90.9667),
        # Kishoreganj
        "Kishoreganj Sadar":(24.4449,90.7766),"Austagram":(24.3333,91.0167),
        "Bajitpur":(24.2167,90.9500),"Bhairab":(24.0667,90.9833),
        "Hossainpur":(24.4333,90.6167),"Itna":(24.5500,91.1167),
        "Karimganj":(24.5667,90.9667),"Katiadi":(24.3500,90.7667),
        "Kuliarchar":(24.3000,90.7000),"Mithamain":(24.6500,91.1333),
        "Nikli":(24.3500,90.9667),"Pakundia":(24.3667,90.6500),
        "Tarail":(24.4667,90.6333),
        # Dhaka + surroundings
        "Dhaka Sadar":(23.7104,90.4074),"Mirpur":(23.8223,90.3654),
        "Savar":(23.8580,90.2664),"Dhanmondi":(23.7461,90.3742),
        "Uttara":(23.8759,90.3795),"Motijheel":(23.7272,90.4093),
        "Demra":(23.7167,90.4833),"Jatrabari":(23.6951,90.4494),
        "Lalbag":(23.7167,90.3833),"Mohammadpur":(23.7667,90.3500),
        "Tejgaon":(23.7500,90.3833),"Rayer Bazar":(23.7333,90.3500),
        "Kamrangirchar":(23.7000,90.3833),"Keraniganj":(23.6833,90.3667),
        "Nawabganj Dhaka":(23.6500,90.2500),"Dohar":(23.5667,90.1500),
        "Dhamrai":(23.9000,90.2000),"Kalatia":(23.7167,90.3833),
        "Badda":(23.7833,90.4333),"Gulshan":(23.7833,90.4167),
        "Banani":(23.7946,90.4028),"Wari":(23.7167,90.4167),
        "Sutrapur":(23.7167,90.4167),"Kotwali":(23.7167,90.4167),
        "Ramna":(23.7333,90.4000),"Lalbagh":(23.7167,90.3833),
        "Adabor":(23.7667,90.3500),"Khilkhet":(23.8333,90.4167),
        "Cantonment":(23.7667,90.3833),"Pallabi":(23.8333,90.3667),
        "Turag":(23.8667,90.3667),"Shyampur":(23.7000,90.4500),
        "Kadamtoli":(23.7000,90.4333),"Sabujbagh":(23.7333,90.4500),
        "Airport":(23.8500,90.4000),
        # Gazipur
        "Gazipur Sadar":(23.9999,90.4203),"Kaliakair":(24.0833,90.2333),
        "Kaliganj Gazipur":(24.0000,90.5000),"Kapasia":(24.1333,90.5833),
        "Sreepur":(24.1833,90.4833),"Tongi":(23.8833,90.3833),
        # Narayanganj
        "Narayanganj Sadar":(23.6238,90.4998),"Araihazar":(23.7333,90.6333),
        "Bandar":(23.5833,90.5667),"Rupganj":(23.7667,90.5833),
        "Sonargaon":(23.6500,90.6167),
        # Narsingdi
        "Narsingdi Sadar":(23.9167,90.7167),"Belabo":(24.0000,90.8333),
        "Monohardi":(24.0667,90.7000),"Palash":(23.8333,90.6833),
        "Raipura":(23.8833,90.8833),"Shibpur":(24.0167,90.6167),
        # Manikganj
        "Manikganj Sadar":(23.8667,90.0167),"Daulatpur":(23.9333,89.9167),
        "Ghior":(23.9000,90.0333),"Harirampur":(23.6833,89.8167),
        "Saturia":(23.9167,90.0833),"Shivalaya":(23.7667,89.9000),
        "Singair":(23.8333,90.1333),
        # Munshiganj
        "Munshiganj Sadar":(23.5500,90.5333),"Gazaria":(23.5000,90.5833),
        "Lohajang":(23.4833,90.4167),"Sirajdikhan":(23.5000,90.4000),
        "Sreenagar":(23.5833,90.3667),"Tongibari":(23.4500,90.5833),
        # Faridpur
        "Faridpur Sadar":(23.6069,89.8431),"Alfadanga":(23.4000,89.7167),
        "Bhanga":(23.3833,90.0000),"Boalmari":(23.5000,89.7833),
        "Char Bhadrasan":(23.4500,89.5833),"Madhukhali":(23.5167,89.7333),
        "Nagarkanda":(23.5167,89.7000),"Sadarpur":(23.5333,89.7000),
        "Saltha":(23.5000,89.9500),
        # Gopalganj
        "Gopalganj Sadar":(23.0056,89.8264),"Kashiani":(23.1333,89.7167),
        "Kotalipara":(22.9667,90.0333),"Muksudpur":(23.2167,89.7000),
        "Tungipara":(23.0500,89.9833),
        # Madaripur
        "Madaripur Sadar":(23.1621,90.2012),"Kalkini":(22.9833,90.3000),
        "Rajoir":(23.1333,90.1167),"Shibchar":(23.1833,90.3333),
        # Shariatpur
        "Shariatpur Sadar":(23.2431,90.4361),"Bhedarganj":(23.1667,90.5833),
        "Damudhya":(23.2833,90.4000),"Gosairhat":(23.1333,90.5000),
        "Jajira":(23.4333,90.3333),"Naria":(23.3167,90.5333),
        "Zanjira":(23.4333,90.3333),
        # Rajbari
        "Rajbari Sadar":(23.7578,89.6414),"Baliakandi":(23.5667,89.8167),
        "Goalandaghat":(23.7000,90.0000),"Kalukhali":(23.6667,89.7167),
        "Pangsha":(23.6167,89.7000),
        # Cumilla/Comilla
        "Cumilla Sadar":(23.4607,91.1809),"Barura":(23.3667,91.0500),
        "Brahmanpara":(23.7000,91.1500),"Burichang":(23.5667,91.2000),
        "Chandina":(23.5500,90.9833),"Chauddagram":(23.3000,91.2500),
        "Daudkandi":(23.5500,90.9167),"Debidwar":(23.6833,91.0833),
        "Homna":(23.6333,90.8667),"Laksam":(23.2333,91.1167),
        "Lalmai":(23.3500,91.1000),"Meghna":(23.5500,90.8167),
        "Muradnagar":(23.7500,91.0500),"Nangalkot":(23.2667,91.3167),
        "Titas":(23.5667,90.8833),
        # Brahmanbaria
        "Brahmanbaria Sadar":(23.9602,91.1095),"Akhaura":(23.8833,91.2167),
        "Ashuganj":(24.0667,90.9833),"Banchharampur":(23.7333,90.9333),
        "Bijoynagar":(23.8000,91.2333),"Kasba":(23.8333,91.1500),
        "Nabinagar":(23.8833,91.0000),"Nasirnagar":(24.0667,91.2167),
        "Sarail":(24.0333,91.1833),
        # Chandpur
        "Chandpur Sadar":(23.2333,90.6667),"Faridganj":(23.1000,90.7333),
        "Haimchar":(23.1500,90.7667),"Haziganj":(23.3500,90.8333),
        "Kachua":(23.3333,91.0000),"Matlab North":(23.4833,90.7000),
        "Matlab South":(23.4167,90.7167),"Shahrasti":(23.2167,91.0333),
        # Lakshmipur
        "Lakshmipur Sadar":(22.9431,90.8278),"Kamalnagar":(22.8000,90.7000),
        "Raipur Lakshmipur":(22.9167,90.9833),"Ramganj":(23.0667,90.8167),
        "Ramgati":(22.8000,90.8500),
        # Noakhali
        "Noakhali Sadar":(22.8696,91.0996),"Begumganj":(22.9167,91.1333),
        "Chatkhil":(22.8833,91.2667),"Companiganj Noakhali":(22.7000,91.2333),
        "Hatiya":(22.4167,91.1167),"Kabirhat":(22.9667,91.2167),
        "Senbagh":(22.8500,91.2000),"Sonaimuri":(23.0167,91.0833),
        "Subarnachar":(22.7167,91.1500),
        # Feni
        "Feni Sadar":(23.0167,91.3967),"Chhagalnaiya":(23.1167,91.3167),
        "Daganbhuiyan":(23.1667,91.3500),"Fulgazi":(22.9667,91.4167),
        "Parshuram":(22.9333,91.4000),"Sonagazi":(22.9167,91.4000),
        # Chittagong/Chattogram
        "Chittagong Sadar":(22.3569,91.7832),"Anwara":(22.2167,91.8667),
        "Banshkhali":(22.0500,92.0167),"Boalkhali":(22.3333,91.9833),
        "Chandanaish":(22.2333,92.0167),"Fatikchhari":(22.6833,91.7500),
        "Hathazari":(22.5000,91.8333),"Karnaphuli":(22.3000,91.8333),
        "Lohagara":(22.0833,92.0833),"Mirsarai":(22.7500,91.5833),
        "Patiya":(22.2833,91.9667),"Rangunia":(22.4667,92.1000),
        "Raozan":(22.4167,91.9167),"Sandwip":(22.4833,91.6333),
        "Sitakunda":(22.6667,91.6667),"Satkania":(22.1167,92.0500),
        # Cox's Bazar
        "Cox'S Bazar Sadar":(21.4272,92.0058),"Chakaria":(21.7333,92.0833),
        "Cutubdia":(21.7500,91.8667),"Kutubdia":(21.7500,91.8667),
        "Maheshkhali":(21.6167,91.9833),"Pekua":(21.8167,91.9667),
        "Ramu":(21.4500,92.1000),"Teknaf":(20.8667,92.3000),
        "Ukhia":(21.1000,92.2000),
        # Bandarban
        "Bandarban Sadar":(22.1933,92.2183),"Alikadam":(21.5167,92.4333),
        "Lama":(21.8833,92.3000),"Naikhangchhari":(21.1167,92.3000),
        "Rowangchhari":(22.0000,92.3333),"Ruma":(22.0167,92.3500),
        "Thanchi":(21.7167,92.3667),
        # Rangamati
        "Rangamati Sadar":(22.6333,92.2000),"Baghaichhari":(22.7833,92.2667),
        "Barkal":(22.8667,92.5333),"Belaichhari":(22.6500,92.4333),
        "Juraichhari":(22.8833,92.2833),"Kaptai":(22.5000,92.2833),
        "Kaukhali Rangamati":(22.5833,92.1833),"Langadu":(23.1667,92.2333),
        "Naniarchar":(22.5667,92.2167),"Rajasthali":(22.5000,92.3333),
        # Khagrachhari
        "Khagrachhari Sadar":(23.1193,91.9847),"Dighinala":(23.2500,92.0167),
        "Guimara":(23.0167,91.9167),"Lakshmichhari":(22.9500,92.0500),
        "Mahalchhari":(23.0167,92.0000),"Manikchhari":(22.8500,91.8833),
        "Matiranga":(23.0167,91.8333),"Panchhari":(23.2833,92.1167),
        "Ramgarh":(22.8167,91.9833),
        # Sylhet
        "Sylhet Sadar":(24.8949,91.8687),"Beanibazar":(24.7000,92.0167),
        "Bishwanath":(24.7833,91.7833),"Companiganj Sylhet":(25.1000,91.7667),
        "Dakshin Surma":(24.8333,91.8000),"Fenchuganj":(24.6833,91.9833),
        "Golapganj":(24.6667,92.0000),"Gowainghat":(25.0500,92.0500),
        "Jaintiapur":(24.9833,92.1833),"Kanaighat":(24.9833,92.2833),
        "Osmani Nagar":(24.8000,91.9500),"Zakiganj":(24.6667,92.3000),
        # Moulvibazar
        "Moulvibazar Sadar":(24.4833,91.7833),"Barlekha":(24.5667,92.1500),
        "Juri":(24.4000,92.1000),"Kamalganj":(24.3167,91.9167),
        "Kulaura":(24.5167,92.0333),"Rajnagar":(24.3667,91.9000),
        "Sreemangal":(24.3000,91.7333),
        # Habiganj
        "Habiganj Sadar":(24.3739,91.4149),"Ajmiriganj":(24.4333,91.3833),
        "Baniachong":(24.5167,91.3667),"Bahubal":(24.3333,91.5667),
        "Chunarughat":(24.2167,91.6667),"Lakhai":(24.4167,91.2833),
        "Madhabpur":(24.2667,91.7167),"Nabiganj":(24.2167,91.3167),
        "Shaistaganj":(24.2833,91.4667),
        # Sunamganj
        "Sunamganj Sadar":(25.0694,91.3983),"Bishwamvarpur":(24.9167,91.5000),
        "Chhatak":(25.0333,91.6667),"Derai":(24.7500,91.4167),
        "Dharmapasha":(24.9333,91.0167),"Dowarabazar":(25.0833,91.7167),
        "Jagannathpur":(24.7833,91.4500),"Jamalganj":(25.0000,91.1167),
        "Sullah":(24.8667,91.3500),"Tahirpur":(25.1167,91.1000),
        # Khulna
        "Khulna Sadar":(22.8456,89.5403),"Batiaghata":(22.7500,89.7333),
        "Dacope":(22.5833,89.5333),"Daulatpur Khulna":(22.8500,89.5167),
        "Dighalia":(22.9500,89.5833),"Dumuria":(22.8000,89.4500),
        "Koyra":(22.3667,89.3167),"Paikgachha":(22.6667,89.3333),
        "Phultala":(22.9000,89.5000),"Rupsha":(22.7833,89.5500),
        "Terokhada":(22.9833,89.7167),
        # Bagerhat
        "Bagerhat Sadar":(22.6500,89.7833),"Chitalmari":(22.8167,89.6833),
        "Fakirhat":(22.7667,89.6833),"Kachua Bagerhat":(22.7500,89.9500),
        "Mollahat":(22.8167,89.8333),"Mongla":(22.4833,89.6000),
        "Morrelganj":(22.5167,89.8833),"Rampal":(22.7000,89.6167),
        "Sarankhola":(22.4333,89.9500),"Sharankhola":(22.4333,89.9500),
        # Jessore/Jashore
        "Jessore Sadar":(23.1667,89.2167),"Jashore Sadar":(23.1667,89.2167),
        "Abhaynagar":(23.0833,89.4167),"Bagherpara":(23.2833,89.2000),
        "Chaugachha":(23.1167,89.0333),"Jhikargachha":(23.1000,89.1500),
        "Keshabpur":(22.9167,89.2167),"Manirampur":(23.0167,89.2667),
        "Sharsha":(23.0500,89.0167),
        # Satkhira
        "Satkhira Sadar":(22.7167,89.0667),"Assasuni":(22.5667,89.1667),
        "Debhata":(22.4667,89.0000),"Kalaroa":(22.8833,89.0667),
        "Kaliganj Satkhira":(22.4833,89.1167),"Shyamnagar":(22.1833,89.0667),
        "Tala":(22.6833,89.1333),
        # Narail
        "Narail Sadar":(23.1667,89.5000),"Kalia":(23.1833,89.6833),
        "Lohagara Narail":(23.0167,89.5833),
        # Magura
        "Magura Sadar":(23.4833,89.4167),"Mohammadpur Magura":(23.4000,89.4833),
        "Shalikha":(23.4333,89.4667),"Sreepur Magura":(23.5833,89.3833),
        # Jhenaidah
        "Jhenaidah Sadar":(23.5444,89.1544),"Harinakunda":(23.5500,88.9833),
        "Kaliganj Jhenaidah":(23.3500,89.1333),"Kotchandpur":(23.4000,88.9833),
        "Maheshpur":(23.6667,88.9000),"Shailkupa":(23.6167,89.0833),
        # Chuadanga
        "Chuadanga Sadar":(23.6406,88.8417),"Alamdanga":(23.7167,88.9500),
        "Damurhuda":(23.7833,88.9333),"Jibannagar":(23.5333,88.8833),
        # Meherpur
        "Meherpur Sadar":(23.7619,88.6306),"Gangni":(23.8500,88.7333),
        "Mujibnagar":(23.7833,88.6667),
        # Kushtia
        "Kushtia Sadar":(23.9010,89.1208),"Bheramara":(24.0333,89.0167),
        "Daulatpur Kushtia":(24.1167,88.9667),"Khoksa":(23.8333,89.0667),
        "Kumarkhali":(23.8667,89.2167),"Mirpur Kushtia":(23.8667,88.9667),
        # Barishal/Barisal
        "Barishal Sadar":(22.7010,90.3535),"Barisal Sadar":(22.7010,90.3535),
        "Agailjhara":(22.9333,90.2000),"Babuganj":(22.6833,90.3333),
        "Bakerganj":(22.5833,90.2667),"Banaripara":(22.8333,90.3167),
        "Gaurnadi":(22.8667,90.2667),"Hizla":(22.5833,90.4333),
        "Mehendiganj":(22.4500,90.5167),"Muladi":(22.6000,90.3500),
        "Uzirpur":(22.8000,90.2000),"Wazirpur":(22.8000,90.2000),
        # Pirojpur
        "Pirojpur Sadar":(22.5794,89.9757),"Bhandaria":(22.4667,90.0500),
        "Kawkhali Pirojpur":(22.5833,89.9500),"Mathbaria":(22.2833,89.9500),
        "Nazirpur":(22.5167,89.9000),"Nesarabad":(22.6167,89.9833),
        "Zianagar":(22.5667,89.9167),
        # Jhalokathi
        "Jhalokathi Sadar":(22.6393,90.1986),"Kathi":(22.6167,90.1333),
        "Nalchity":(22.5500,90.0333),"Rajapur":(22.5167,90.0833),
        # Barguna
        "Barguna Sadar":(22.1500,90.1120),"Amtali":(22.0167,90.1333),
        "Bamna":(22.2667,89.9500),"Betagi":(22.1667,90.0167),
        "Patharghata":(22.0000,90.0333),"Taltali":(21.9667,90.2333),
        # Patuakhali
        "Patuakhali Sadar":(22.3596,90.3298),"Bauphal":(22.4667,90.5000),
        "Dashmina":(22.3833,90.5333),"Dumki":(22.4000,90.3833),
        "Galachipa":(22.1500,90.4667),"Kalapara":(21.9667,90.3167),
        "Mirzaganj":(22.5000,90.3333),"Rangabali":(22.1667,90.6167),
        # Bhola
        "Bhola Sadar":(22.6860,90.6480),"Burhanuddin":(22.5000,90.7167),
        "Char Fasson":(22.1667,90.7667),"Daulatkhan":(22.5333,90.7667),
        "Lalmohan":(22.4167,90.7000),"Manpura":(22.0833,90.8167),
        "Tazumuddin":(22.4167,90.8000),
        # Naogaon/Rajshahi division misc
        "Atrai Sadar":(24.6000,88.9000),
        # Gazipur
        "Gazipur":(23.9999,90.4203),
        # Narayanganj
        "Narayanganj":(23.6238,90.4998),
        # Netrokona (variant)
        "Netrokona":(24.8704,90.7270),"Netrakona":(24.8704,90.7270),
        # Mymensingh district coord (fallback)
        "Mymensingh":(24.7471,90.4203),"Maymensingh":(24.7471,90.4203),
        # Cumilla/Comilla variant
        "Comilla Sadar":(23.4607,91.1809),"Comilla":(23.4607,91.1809),
        # Rangpur district coord
        "Rangpur":(25.7439,89.2752),
        # Others
        "Sherpur Sadar":(25.0167,90.0167),"Nakla":(24.9333,90.1667),
        "Nalitabari":(25.1000,90.1500),"Jhenaigati":(25.0167,90.0833),
        "Sreebardi":(25.0667,90.0333),
    }

    # ── Upazila name normalization ──
    UPA_NORM = {
        # Kurigram
        "roumary":"Rowmari","raomari":"Rowmari","rowmari":"Rowmari","raumari":"Rowmari",
        "chilmari":"Chilmari","chilmare":"Chilmari","rajibpur":"Rajibpur",
        "ulipur":"Ulipur","nageshwari":"Nageshwari","bhurungamari":"Bhurungamari",
        "rajarhat":"Rajarhat","phulbari":"Phulbari","kurigram sadar":"Kurigram Sadar",
        # Gaibandha
        "sundarganj":"Sundarganj","sundorganj":"Sundarganj","sadullapur":"Sadullapur",
        "gaibandha sadar":"Gaibandha Sadar","gobindaganj":"Gobindaganj",
        "palashbari":"Palashbari","fulchhari":"Fulchhari","gaibandha":"Gaibandha Sadar",
        # Rangpur
        "rangpur sadar":"Rangpur Sadar","pirganj":"Pirganj","pirgonj":"Pirganj",
        "pirgacha":"Pirgacha","mahiganj":"Mahiganj","satmatha":"Mahiganj",
        "gangachara":"Gangachara","kaunia":"Kaunia","mithapukur":"Mithapukur",
        "badarganj":"Badarganj","taraganj":"Taraganj",
        # Lalmonirhat
        "lalmonirhat sadar":"Lalmonirhat Sadar","hatibandha":"Hatibandha",
        "kaliganj":"Kaliganj","aditmari":"Aditmari","patgram":"Patgram",
        # Nilphamari
        "nilphamari sadar":"Nilphamari Sadar","saidpur":"Saidpur",
        "jaldhaka":"Jaldhaka","domar":"Domar","dimla":"Dimla",
        # Dinajpur
        "dinajpur sadar":"Dinajpur Sadar","birampur":"Birampur",
        "parbatipur":"Parbatipur","fulbari":"Fulbari","phulbari":"Phulbari",
        # Bogura
        "bogura sadar":"Bogura Sadar","bogra sadar":"Bogura Sadar",
        "gabtali":"Gabtali","sariakandi":"Sariakandi","sonatala":"Sonatala",
        "dhunat":"Dhunat","sherpur bogura":"Sherpur Bogura",
        # Joypurhat
        "joypurhat sadar":"Joypurhat Sadar","kalai":"Kalai","khetlal":"Khetlal",
        # Sirajganj
        "sirajganj sadar":"Sirajganj Sadar","belkuchi":"Belkuchi",
        "shahjadpur":"Shahjadpur","ullapara":"Ullahpara","ullahpara":"Ullahpara",
        # Pabna
        "pabna sadar":"Pabna Sadar","ishwardi":"Ishwardi","santhia":"Santhia",
        # Rajshahi
        "rajshahi sadar":"Rajshahi Sadar","godagari":"Godagari",
        "tanore":"Tanore","puthia":"Puthia","charghat":"Charghat",
        # Naogaon
        "naogaon sadar":"Naogaon Sadar","atrai":"Atrai","manda":"Manda",
        "dhamoirhat":"Dhamoirhat","sapahar":"Sapahar","patnitala":"Patnitala",
        # Chapainawabganj
        "chapainawabganj sadar":"Chapainawabganj Sadar",
        "shibganj":"Shibganj Chapai","gomastapur":"Gomastapur",
        # Natore
        "natore sadar":"Natore Sadar","singra":"Singra","lalpur":"Lalpur",
        "baraigram":"Baraigram","gurudaspur":"Gurudaspur",
        # Tangail
        "tangail sadar":"Tangail Sadar","ghatail":"Ghatail","madhupur":"Madhupur",
        "mirzapur":"Mirzapur","sakhipur":"Sakhipur","kalihati":"Kalihati",
        "basail":"Basail","delduar":"Delduar","nagarpur":"Nagarpur",
        # Jamalpur
        "jamalpur sadar":"Jamalpur Sadar","islampur":"Islampur",
        "melandaha":"Melandaha","dewanganj":"Dewanganj","bakshiganj":"Bakshiganj",
        "sarishabari":"Sarishabari","madarganj":"Madarganj",
        # Mymensingh
        "mymensingh sadar":"Mymensingh Sadar","trishal":"Trishal","bhaluka":"Bhaluka",
        "gaffargaon":"Gaffargaon","gauripur":"Gauripur","gouripur":"Gouripur",
        "haluaghat":"Haluaghat","ishwarganj":"Ishwarganj","iswarganj":"Ishwarganj",
        "ishwargonj":"Ishwarganj","muktagacha":"Muktagacha","nandail":"Nandail",
        "phulbaria":"Phulbaria","phulpur":"Phulpur","fulbaria mymensingh":"Fulbaria",
        # Netrokona
        "netrokona sadar":"Netrokona Sadar","atpara":"Atpara","barhatta":"Barhatta",
        "durgapur":"Durgapur","kendua":"Kendua","madan":"Madan",
        "mohanganj":"Mohanganj","kalmakanda":"Kalmakanda","chandua":"Chandua",
        "khaliajuri":"Khaliajuri","purbadhala":"Purbadhala",
        # Kishoreganj
        "kishoreganj sadar":"Kishoreganj Sadar","bhairab":"Bhairab",
        "bajitpur":"Bajitpur","kuliarchar":"Kuliarchar","katiadi":"Katiadi",
        "pakundia":"Pakundia","tarail":"Tarail","karimganj":"Karimganj",
        "hossainpur":"Hossainpur","itna":"Itna","nikli":"Nikli",
        "austagram":"Austagram","mithamain":"Mithamain",
        # Dhaka
        "dhaka sadar":"Dhaka Sadar","mirpur":"Mirpur","savar":"Savar",
        "dhanmondi":"Dhanmondi","uttara":"Uttara","motijheel":"Motijheel",
        "demra":"Demra","jatrabari":"Jatrabari","tejgaon":"Tejgaon",
        "gulshan":"Gulshan","banani":"Banani","badda":"Badda",
        "keraniganj":"Keraniganj","dohar":"Dohar","dhamrai":"Dhamrai",
        "mohammadpur":"Mohammadpur","wari":"Wari","ramna":"Ramna",
        "lalbagh":"Lalbagh","sutrapur":"Sutrapur","kotwali":"Kotwali",
        "kakrail":"Dhaka Sadar","shahbag":"Dhaka Sadar","segunbagicha":"Dhaka Sadar",
        "rajarbag":"Dhaka Sadar","purana paltan":"Dhaka Sadar",
        # Islampur Dhaka (old town area) — different from Islampur Jamalpur
        "islampur road":"Kotwali","kumartuli":"Kotwali","islampur, dhaka":"Kotwali",
        "islampur dhaka":"Kotwali","islampur,dhaka":"Kotwali",
        # Gazipur
        "gazipur sadar":"Gazipur Sadar","tongi":"Tongi","kaliakair":"Kaliakair",
        "sreepur":"Sreepur","kapasia":"Kapasia",
        # Narayanganj
        "narayanganj sadar":"Narayanganj Sadar","rupganj":"Rupganj",
        "sonargaon":"Sonargaon","araihazar":"Araihazar","bandar":"Bandar",
        # Narsingdi
        "narsingdi sadar":"Narsingdi Sadar","shibpur":"Shibpur","belabo":"Belabo",
        "monohardi":"Monohardi","raipura":"Raipura","palash":"Palash",
        # Faridpur
        "faridpur sadar":"Faridpur Sadar","bhanga":"Bhanga","saltha":"Saltha",
        "nagarkanda":"Nagarkanda","alfadanga":"Alfadanga",
        # Gopalganj
        "gopalganj sadar":"Gopalganj Sadar","tungipara":"Tungipara",
        "kotalipara":"Kotalipara","kashiani":"Kashiani","muksudpur":"Muksudpur",
        # Madaripur
        "madaripur sadar":"Madaripur Sadar","shibchar":"Shibchar",
        "kalkini":"Kalkini","rajoir":"Rajoir",
        # Barishal/Barisal
        "barishal sadar":"Barishal Sadar","barisal sadar":"Barishal Sadar",
        "agailjhara":"Agailjhara","babuganj":"Babuganj","bakerganj":"Bakerganj",
        "banaripara":"Banaripara","gaurnadi":"Gaurnadi","hizla":"Hizla",
        "mehendiganj":"Mehendiganj","muladi":"Muladi","uzirpur":"Uzirpur",
        "wazirpur":"Uzirpur","bimanbandor":"Barishal Sadar",
        "barishal bimanbandor":"Barishal Sadar",
        # Pirojpur
        "pirojpur sadar":"Pirojpur Sadar","bhandaria":"Bhandaria",
        "mathbaria":"Mathbaria","nazirpur":"Nazirpur","nesarabad":"Nesarabad",
        # Jhalokathi
        "jhalokathi sadar":"Jhalokathi Sadar","nalchity":"Nalchity",
        "rajapur":"Rajapur","kathi":"Kathi",
        # Barguna
        "barguna sadar":"Barguna Sadar","amtali":"Amtali","bamna":"Bamna",
        "betagi":"Betagi","patharghata":"Patharghata",
        # Patuakhali
        "patuakhali sadar":"Patuakhali Sadar","bauphal":"Bauphal",
        "galachipa":"Galachipa","kalapara":"Kalapara","mirzaganj":"Mirzaganj",
        # Bhola
        "bhola sadar":"Bhola Sadar","burhanuddin":"Burhanuddin",
        "daulatkhan":"Daulatkhan","lalmohan":"Lalmohan","manpura":"Manpura",
        # Cumilla
        "cumilla sadar":"Cumilla Sadar","comilla sadar":"Cumilla Sadar",
        "barura":"Barura","brahmanpara":"Brahmanpara","burichang":"Burichang",
        "chandina":"Chandina","chauddagram":"Chauddagram","debidwar":"Debidwar",
        "daudkandi":"Daudkandi","homna":"Homna","laksam":"Laksam",
        "muradnagar":"Muradnagar","meghna":"Meghna","titas":"Titas",
        # Brahmanbaria
        "brahmanbaria sadar":"Brahmanbaria Sadar","akhaura":"Akhaura",
        "ashuganj":"Ashuganj","kasba":"Kasba","sarail":"Sarail",
        "nabinagar":"Nabinagar","nasirnagar":"Nasirnagar",
        # Chandpur
        "chandpur sadar":"Chandpur Sadar","haimchar":"Haimchar",
        "haziganj":"Haziganj","kachua":"Kachua","matlab north":"Matlab North",
        "matlab south":"Matlab South","shahrasti":"Shahrasti",
        # Chittagong/Chattogram
        "chittagong sadar":"Chittagong Sadar","chattogram sadar":"Chittagong Sadar",
        "hathazari":"Hathazari","fatikchhari":"Fatikchhari","mirsarai":"Mirsarai",
        "sandwip":"Sandwip","sitakunda":"Sitakunda","rangunia":"Rangunia",
        "anwara":"Anwara","patiya":"Patiya","chandanaish":"Chandanaish",
        # Sylhet
        "sylhet sadar":"Sylhet Sadar","bishwanath":"Bishwanath",
        "golapganj":"Golapganj","gowainghat":"Gowainghat",
        "jaintiapur":"Jaintiapur","kanaighat":"Kanaighat","zakiganj":"Zakiganj",
        "fenchuganj":"Fenchuganj","beanibazar":"Beanibazar",
        # Moulvibazar
        "moulvibazar sadar":"Moulvibazar Sadar","sreemangal":"Sreemangal",
        "kamalganj":"Kamalganj","kulaura":"Kulaura","barlekha":"Barlekha",
        # Habiganj
        "habiganj sadar":"Habiganj Sadar","chunarughat":"Chunarughat",
        "madhabpur":"Madhabpur","nabiganj":"Nabiganj","shaistaganj":"Shaistaganj",
        "baniachong":"Baniachong","lakhai":"Lakhai","bahubal":"Bahubal",
        # Sunamganj
        "sunamganj sadar":"Sunamganj Sadar","chhatak":"Chhatak","derai":"Derai",
        "dowarabazar":"Dowarabazar","jagannathpur":"Jagannathpur",
        # Khulna
        "khulna sadar":"Khulna Sadar","dumuria":"Dumuria","koyra":"Koyra",
        "paikgachha":"Paikgachha","batiaghata":"Batiaghata","dacope":"Dacope",
        # Jessore
        "jessore sadar":"Jessore Sadar","jashore sadar":"Jashore Sadar",
        "keshabpur":"Keshabpur","manirampur":"Manirampur","sharsha":"Sharsha",
        "abhaynagar":"Abhaynagar","jhikargachha":"Jhikargachha",
        # Satkhira
        "satkhira sadar":"Satkhira Sadar","kalaroa":"Kalaroa",
        "assasuni":"Assasuni","shyamnagar":"Shyamnagar","tala":"Tala",
        # Jhenaidah
        "jhenaidah sadar":"Jhenaidah Sadar","shailkupa":"Shailkupa",
        "kotchandpur":"Kotchandpur","maheshpur":"Maheshpur",
        # Chuadanga
        "chuadanga sadar":"Chuadanga Sadar","alamdanga":"Alamdanga",
        "damurhuda":"Damurhuda","jibannagar":"Jibannagar",
        # Kushtia
        "kushtia sadar":"Kushtia Sadar","bheramara":"Bheramara",
        "ishwardi":"Ishwardi","kumarkhali":"Kumarkhali",
        # Cox's Bazar
        "cox's bazar sadar":"Cox'S Bazar Sadar","chakaria":"Chakaria",
        "teknaf":"Teknaf","ukhia":"Ukhia","ramu":"Ramu",
        "maheshkhali":"Maheshkhali","kutubdia":"Kutubdia",
        # Bandarban
        "bandarban sadar":"Bandarban Sadar","lama":"Lama","ruma":"Ruma",
        "alikadam":"Alikadam","rowangchhari":"Rowangchhari",
        # Rangamati
        "rangamati sadar":"Rangamati Sadar","kaptai":"Kaptai",
        "baghaichhari":"Baghaichhari","barkal":"Barkal",
        # Khagrachhari
        "khagrachhari sadar":"Khagrachhari Sadar","dighinala":"Dighinala",
        "matiranga":"Matiranga","ramgarh":"Ramgarh","panchhari":"Panchhari",
        # Sherpur
        "sherpur sadar":"Sherpur Sadar","nakla":"Nakla",
        "nalitabari":"Nalitabari","sreebardi":"Sreebardi",
        # Misc village/area → upazila mappings often seen in CDR
        "pasar":"Gouripur","sohagi":"Ishwarganj","sahanati":"Gauripur",
        "maoha":"Gouripur","ishwargonj":"Ishwarganj",
    }

    # ── District name correction/normalization ──
    DIST_CORR = {
        # Kurigram variants
        "kuregram":"Kurigram","kurigrame":"Kurigram","kurigram":"Kurigram",
        # Rangpur variants
        "rongpur":"Rangpur","rangpur":"Rangpur","rangpur sadar":"Rangpur",
        # Gaibandha variants
        "gaibanda":"Gaibandha","gaibandha":"Gaibandha",
        # Major districts
        "dhaka":"Dhaka","mymensingh":"Mymensingh","mymensingh.":"Mymensingh",
        "mymensing":"Mymensingh","maymensingh":"Mymensingh","maimensingh":"Mymensingh",
        "rajshahi":"Rajshahi","khulna":"Khulna","sylhet":"Sylhet",
        "chittagong":"Chittagong","chattogram":"Chittagong",
        "barishal":"Barishal","barisal":"Barishal",
        "cumilla":"Cumilla","comilla":"Cumilla",
        "tangail":"Tangail","faridpur":"Faridpur","gopalganj":"Gopalganj",
        "narsingdi":"Narsingdi","narayanganj":"Narayanganj","gazipur":"Gazipur",
        "manikganj":"Manikganj","munshiganj":"Munshiganj","kishoreganj":"Kishoreganj",
        "netrokona":"Netrokona","netrakona":"Netrokona",
        "jamalpur":"Jamalpur","sherpur":"Sherpur","bogura":"Bogura","bogra":"Bogura",
        "sirajganj":"Sirajganj","pabna":"Pabna","natore":"Natore",
        "naogaon":"Naogaon","rajbari":"Rajbari","shariatpur":"Shariatpur",
        "madaripur":"Madaripur","joypurhat":"Joypurhat",
        "chapainawabganj":"Chapainawabganj","chapai nawabganj":"Chapainawabganj",
        "jhenaidah":"Jhenaidah","jessore":"Jessore","jashore":"Jessore",
        "magura":"Magura","narail":"Narail","satkhira":"Satkhira",
        "bagerhat":"Bagerhat","kushtia":"Kushtia","meherpur":"Meherpur",
        "chuadanga":"Chuadanga","lalmonirhat":"Lalmonirhat",
        "nilphamari":"Nilphamari","dinajpur":"Dinajpur","thakurgaon":"Thakurgaon",
        "panchagarh":"Panchagarh","habiganj":"Habiganj",
        "moulvibazar":"Moulvibazar","sunamganj":"Sunamganj",
        "brahmanbaria":"Brahmanbaria","chandpur":"Chandpur",
        "lakshmipur":"Lakshmipur","noakhali":"Noakhali","feni":"Feni",
        "cox's bazar":"Cox'S Bazar","coxs bazar":"Cox'S Bazar",
        "bandarban":"Bandarban","rangamati":"Rangamati","khagrachhari":"Khagrachhari",
        "pirojpur":"Pirojpur","jhalokathi":"Jhalokathi","barguna":"Barguna",
        "patuakhali":"Patuakhali","bhola":"Bhola",
        # With trailing punctuation
        "mymensingh.":"Mymensingh","dhaka.":"Dhaka","chittagong.":"Chittagong",
        # Dist : prefix cleanup
        "dist : barishal":"Barishal","dist : dhaka":"Dhaka",
        "dist : mymensingh":"Mymensingh","dist : chittagong":"Chittagong",
        "dist- netrokhona":"Netrokona","dist mymensingh":"Mymensingh",
    }

    def haversine(la1,lo1,la2,lo2):
        R=6371.0; p1,p2=math.radians(la1),math.radians(la2)
        dp=math.radians(la2-la1); dl=math.radians(lo2-lo1)
        a=math.sin(dp/2)**2+math.cos(p1)*math.cos(p2)*math.sin(dl/2)**2
        return R*2*math.atan2(math.sqrt(a),math.sqrt(1-a))

    def parse_ud(addr):
        """
        Extract (upazila, district) from a BTS address string.
        Handles: 'DIST : BARISHAL', 'DIST. MYMENSINGH', 'P.S- GOURIPUR', 'PS:ISHWARGANJ'
        Falls back to comma-split parts if regex fails.
        """
        if not addr or str(addr).strip() in ("","-","nan"): return None,None
        s = str(addr); d = upa = None

        # ── District: multiple patterns ──
        # Pattern 1: DIST : MYMENSINGH / DIST. DHAKA / DIST- CHITTAGONG / DIS- NETROKHONA
        dm = re.search(r"\bDIS[T]?\s*[:\.\-]\s*([A-Za-z][A-Za-z\s']+?)(?:[,\.\n;]|$)", s, re.IGNORECASE)
        if dm:
            d = dm.group(1).strip().rstrip(".")
        # Pattern 2: District Mymensingh / District: Dhaka
        if not d:
            dm2 = re.search(r"\bDistrict\s*[:\-]?\s*([A-Za-z][A-Za-z\s']+?)(?:[,\.\n;]|$)", s, re.IGNORECASE)
            if dm2: d = dm2.group(1).strip().rstrip(".")

        # ── Upazila/PS: multiple patterns ──
        # Pattern: P.S- GOURIPUR / P/S CHANDUA / PS:ISHWARGANJ / Thana- / Upazila-
        um = re.search(
            r"(?:P[\./]?\s*S[\s:\.\-]+|[Pp]s\s*[:\-]+|[Tt]hana\s*[:\-]+|[Uu]pazill?a\s*[:\-]+)"
            r"([A-Za-z][A-Za-z\s]+?)(?:[,\.\n;]|$)", s, re.IGNORECASE)
        if um: upa = um.group(1).strip().rstrip(".")

        # ── Fallback: comma-split last parts ──
        if not d:
            parts = [re.sub(r"\d+","",p).strip(" -.") for p in s.split(",")]
            parts = [p.strip() for p in parts if len(p.strip()) > 2]
            if parts:
                last = parts[-1].strip()
                # Strip leading "DIST :" or "DIS-" prefix if present
                last = re.sub(r"^DIS[T]?\s*[:\.\-]\s*", "", last, flags=re.IGNORECASE).strip()
                d = last
            if len(parts) >= 2 and not upa:
                upa = parts[-2].strip()

        # ── Normalize district ──
        if d:
            d_clean = d.lower().strip().rstrip(".")
            d = DIST_CORR.get(d_clean, d.title().strip())

        # ── Normalize upazila ──
        if upa:
            upa = re.sub(r"\s+"," ", upa).strip().rstrip(".")
            upa_norm = UPA_NORM.get(upa.lower().strip(), None)
            if upa_norm:
                upa = upa_norm
            else:
                # Try partial match: first word of upazila
                first_word = upa.lower().split()[0] if upa else ""
                upa = UPA_NORM.get(first_word, upa.title())

        return upa, d

    def get_coord(upa, dist):
        """
        Look up coordinates for a upazila/district.
        Uses district context to disambiguate same-name upazilas in different districts.
        e.g. Islampur in Dhaka (old town) vs Islampur in Jamalpur district.
        """
        # District-aware upazila disambiguation
        DIST_UPA_OVERRIDE = {
            # (upazila_lower, district_lower): canonical upazila name with correct coords
            ("islampur", "dhaka"):      "Kotwali",       # Islampur old Dhaka → Kotwali coords
            ("islampur", ""):           "Kotwali",       # if dist unknown but addr says Dhaka
        }

        if upa and dist:
            key = (upa.lower(), dist.lower())
            if key in DIST_UPA_OVERRIDE:
                upa = DIST_UPA_OVERRIDE[key]
        elif upa:
            # Check if address context helps (e.g. "Islampur" only used as Dhaka old town)
            pass

        if upa:
            if upa in BD_COORDS: return BD_COORDS[upa]
            # Try case-insensitive
            for k, v in BD_COORDS.items():
                if k.lower() == upa.lower(): return v
        if dist:
            # Try "District Sadar"
            k = dist + " Sadar"
            if k in BD_COORDS: return BD_COORDS[k]
            # Try exact district name
            if dist in BD_COORDS: return BD_COORDS[dist]
            # Try prefix match (first 6 chars)
            dist_low = dist.lower()
            for kk, vv in BD_COORDS.items():
                if kk.lower().startswith(dist_low[:6]): return vv
            # Try if district name is contained in key
            for kk, vv in BD_COORDS.items():
                if dist_low in kk.lower() and "Sadar" in kk: return vv
        return None

    df_loc=df[df["address"].notna()&df["address"].apply(_is_valid_address)].copy()
    df_loc=df_loc.sort_values("start").reset_index(drop=True)
    if df_loc.empty: return None

    home_dist = _home_district(df_loc)
    work_dist = None  # Not used in new logic

    # Home coord: most frequent address whose coord can be resolved
    home_coord = None
    home_label = None
    for addr in df_loc["address"].value_counts().index:
        upa, dist = parse_ud(addr)
        coord = get_coord(upa, dist)
        if coord:
            home_coord = coord
            home_label = upa or dist
            break
    if not home_coord:
        home_coord = (25.5964, 89.7662)  # fallback Rowmari
        home_label = "Rowmari"

    HOME_KM=35.0; TRANSIT_H=4; MIN_HOME_STAY_H=5

    # ── Enrich rows with distance from home ──
    rows_e=[]
    last_known_km = 0.0
    last_known_coord = home_coord
    for _,row in df_loc.iterrows():
        upa,dist=parse_ud(row["address"])
        coord=get_coord(upa,dist)
        if coord:
            km=haversine(home_coord[0],home_coord[1],coord[0],coord[1])
            last_known_km=km
            last_known_coord=coord
        else:
            km=last_known_km
            coord=last_known_coord
        rows_e.append({"ts":row["start"],"address":row["address"],
                       "upazila":upa,"district":dist,"coord":coord,"km":km})
    df_e=pd.DataFrame(rows_e).sort_values("ts").reset_index(drop=True)
    df_e["away"]=df_e["km"]>=HOME_KM

    # ── Group consecutive away runs into sessions ──
    # Session breaks ONLY on a CONFIRMED overnight home return:
    #   - km < HOME_KM (clearly home area)
    #   - AND consecutive home records span >= MIN_HOME_STAY_H hours
    # Brief daytime visits to home area do NOT break the trip.

    def _is_real_home_return(chunk_ts_list):
        """Return True if home buffer spans >= MIN_HOME_STAY_H hours (real overnight return)."""
        if not chunk_ts_list:
            return False
        span_h = (chunk_ts_list[-1] - chunk_ts_list[0]).total_seconds() / 3600
        return span_h >= MIN_HOME_STAY_H

    sessions=[]
    sess_rows=[]
    in_sess=False
    home_buffer=[]   # accumulate consecutive home rows to check if real return

    for idx, r in df_e.iterrows():
        if r["away"]:
            # Flush home buffer — was it a real return or brief visit?
            if home_buffer and in_sess:
                if _is_real_home_return([x["ts"] for x in home_buffer]):
                    # Real home return — end current session
                    if sess_rows:
                        sessions.append(sess_rows)
                    sess_rows=[]
                    in_sess=False
                else:
                    # Brief visit — fold home buffer into ongoing session
                    sess_rows.extend(home_buffer)
            home_buffer=[]
            in_sess=True
            sess_rows.append(r)
        else:
            if r["km"] < HOME_KM:
                home_buffer.append(r)
            else:
                # Unknown coord carrying forward away km — keep in session
                if home_buffer and in_sess:
                    sess_rows.extend(home_buffer)
                    home_buffer=[]
                if in_sess:
                    sess_rows.append(r)

    # Trailing home buffer
    if home_buffer and in_sess:
        if _is_real_home_return([x["ts"] for x in home_buffer]):
            if sess_rows: sessions.append(sess_rows)
        else:
            sess_rows.extend(home_buffer)
            if sess_rows: sessions.append(sess_rows)
    elif in_sess and sess_rows:
        sessions.append(sess_rows)

    trips=[]
    for sess in sessions:
        sdf=pd.DataFrame(sess)
        s_start=sdf["ts"].min(); s_end=sdf["ts"].max()
        s_days=(s_end.date()-s_start.date()).days+1

        # Sub-group by district (not upazila) to avoid over-fragmentation
        loc_groups=[]; cur_key=None; cur_rows=[]
        for _,r in sdf.iterrows():
            # Group by district only — same district = same location group
            lk = r["district"] or r["upazila"] or ""
            if lk != cur_key:
                if cur_rows: loc_groups.append((cur_key, cur_rows))
                cur_key=lk; cur_rows=[r]
            else:
                cur_rows.append(r)
        if cur_rows: loc_groups.append((cur_key, cur_rows))

        stay_locs=[]
        for dist_k, grp in loc_groups:
            gdf=pd.DataFrame(grp)
            hrs=(gdf["ts"].max()-gdf["ts"].min()).total_seconds()/3600
            max_km=round(gdf["km"].max(),1)
            # Best upazila for this district group (most frequent non-null)
            upa_vals = gdf["upazila"].dropna()
            upa_k = upa_vals.value_counts().index[0] if not upa_vals.empty else None
            # Best address sample
            best_addr = gdf["address"].value_counts().index[0] if not gdf.empty else ""
            # Include if: stayed >= TRANSIT_H OR only location OR far enough
            if hrs>=TRANSIT_H or len(loc_groups)==1 or max_km>=HOME_KM*1.5:
                stay_locs.append({
                    "upazila": upa_k,
                    "district": dist_k or None,
                    "km": max_km, "hours": round(hrs,1),
                    "first": gdf["ts"].min(), "last": gdf["ts"].max(),
                    "sample_addr": str(best_addr)[:80],
                })
        if not stay_locs: continue

        # Show the farthest district as the main trip destination
        dest=max(stay_locs,key=lambda x:x["km"])
        dest_label=(dest["upazila"] if dest["upazila"] and str(dest["upazila"]).strip().lower()!="nan" else None) or dest["district"] or "Unknown"
        dist_label=dest["district"] or "Unknown"

        trips.append({
            "upazila":dest_label,"district":dist_label,
            "km":dest["km"],"start_date":str(s_start.date()),
            "end_date":str(s_end.date()),"days":s_days,
            "hours":round((s_end-s_start).total_seconds()/3600,1),
            "stay_locs":stay_locs,
            "address":str(dest["sample_addr"])[:80],
        })

    trips.sort(key=lambda x:x["start_date"],reverse=True)

    all_dates=sorted(df["start"].dt.date.unique())
    gaps=[]
    for k in range(len(all_dates)-1):
        d1=pd.Timestamp(all_dates[k]); d2=pd.Timestamp(all_dates[k+1])
        gd=(d2-d1).days-1
        if gd>4: gaps.append({"gap_start":str(all_dates[k]),"gap_end":str(all_dates[k+1]),"days":gd})

    total_days=((pd.Timestamp(all_dates[-1])-pd.Timestamp(all_dates[0])).days+1 if all_dates else 0)
    return {
        "home_district":home_dist,"work_district":None,
        "home_coord":home_coord,"home_label":home_label,
        "base_districts":list(set(filter(None,[home_dist]))),
        "trips":trips,"gaps":gaps,"total_days":total_days,
        "out_of_home_days":sum(t["days"] for t in trips),
        "total_records":len(df),
        "date_from":str(all_dates[0]) if all_dates else "N/A",
        "date_to":str(all_dates[-1]) if all_dates else "N/A",
    }

def _movement_html(mv):
    """Generate HTML for movement pattern section (new distance-based logic)."""
    if not mv:
        return '<p style="color:#64748b;">Location data insufficient for movement analysis.</p>'

    trip_count = len(mv["trips"])
    gap_count  = len(mv["gaps"])

    home_label = mv.get("home_label") or mv.get("home_district") or "N/A"

    cards_html = f"""
    <div style="display:grid; grid-template-columns:repeat(4,1fr); gap:1rem; margin-bottom:1.5rem;">
        <div style="background:white; border-top:4px solid #2563eb; border-radius:10px;
                    padding:1rem; box-shadow:0 1px 3px rgba(0,0,0,0.06);">
            <div style="font-size:0.75rem; color:#94a3b8; font-weight:600;
                        text-transform:uppercase;">Total Records</div>
            <div style="font-size:1.8rem; font-weight:800; color:#0f172a;
                        margin:0.3rem 0;">{mv["total_records"]:,}</div>
            <div style="font-size:0.8rem; color:#64748b;">Call + SMS | {mv["total_days"]} days</div>
        </div>
        <div style="background:white; border-top:4px solid #16a34a; border-radius:10px;
                    padding:1rem; box-shadow:0 1px 3px rgba(0,0,0,0.06);">
            <div style="font-size:0.75rem; color:#94a3b8; font-weight:600;
                        text-transform:uppercase;">Home Location</div>
            <div style="font-size:1.1rem; font-weight:800; color:#0f172a;
                        margin:0.3rem 0;">{home_label}</div>
            <div style="font-size:0.78rem; color:#16a34a;">{mv["home_district"] or ""} District (most frequent)</div>
        </div>
        <div style="background:white; border-top:4px solid #dc2626; border-radius:10px;
                    padding:1rem; box-shadow:0 1px 3px rgba(0,0,0,0.06);">
            <div style="font-size:0.75rem; color:#94a3b8; font-weight:600;
                        text-transform:uppercase;">Network Gaps</div>
            <div style="font-size:1.8rem; font-weight:800; color:#dc2626;
                        margin:0.3rem 0;">{gap_count}</div>
            <div style="font-size:0.8rem; color:#64748b;">Disconnected more than 4 days</div>
        </div>
        <div style="background:white; border-top:4px solid #7c3aed; border-radius:10px;
                    padding:1rem; box-shadow:0 1px 3px rgba(0,0,0,0.06);">
            <div style="font-size:0.75rem; color:#94a3b8; font-weight:600;
                        text-transform:uppercase;">Out-of-Home Trips</div>
            <div style="font-size:1.8rem; font-weight:800; color:#7c3aed;
                        margin:0.3rem 0;">{trip_count}</div>
            <div style="font-size:0.8rem; color:#64748b;">{mv["out_of_home_days"]} days total</div>
        </div>
    </div>"""

    # ── Trips table (reverse chronological) ──────────────────────────────
    trips_html = ""
    if mv["trips"]:
        def _stay_detail(stay_locs):
            if len(stay_locs) <= 1: return ""
            parts = []
            for sl in stay_locs[:-1]:
                lbl = sl["upazila"] or sl["district"] or "?"
                parts.append(f"{lbl} ({sl['hours']}h, {sl['km']}km)")
            return "Via: " + " -> ".join(parts) if parts else ""

        rows = "".join([
            f"""<tr style="background:{'#f8fafc' if i%2==0 else 'white'};">
                <td style="padding:0.7rem 1rem; font-weight:600; color:#1e3a8a;">
                    {i+1}</td>
                <td style="padding:0.7rem 1rem; font-weight:700;">
                    {t["upazila"]}<br>
                    <span style="font-size:0.8rem;font-weight:400;color:#64748b;">
                        {t["district"]} District</span></td>
                <td style="padding:0.7rem 1rem; text-align:center;">
                    <span style="background:#dbeafe;color:#1e40af;border-radius:12px;
                                 padding:0.2rem 0.7rem;font-weight:700;">
                        {t["km"]} km</span></td>
                <td style="padding:0.7rem 1rem;">{t["start_date"]}</td>
                <td style="padding:0.7rem 1rem;">{t["end_date"]}</td>
                <td style="padding:0.7rem 1rem; text-align:center;">
                    <span style="background:#ede9fe;color:#5b21b6;border-radius:12px;
                                 padding:0.2rem 0.6rem;font-weight:700;">{t["days"]}</span>
                </td>
                <td style="padding:0.7rem 1rem; font-size:0.82rem; color:#475569;">
                    {str(t["address"])[:55]}
                    {"<br><span style='color:#94a3b8;font-size:0.78rem;'>" + _stay_detail(t.get("stay_locs",[])) + "</span>" if _stay_detail(t.get("stay_locs",[])) else ""}
                </td>
            </tr>"""
            for i, t in enumerate(mv["trips"])
        ])

        trips_html = f"""
        <div style="font-weight:700; color:#92400e; margin:0.5rem 0 0.75rem 0;
                    background:#fffbeb; border-left:4px solid #f59e0b;
                    border-radius:6px; padding:0.6rem 1rem;">
            Out-of-Home Travel Detected — {trip_count} trip(s) |
            Home Reference: {home_label}, {mv.get("home_district","?")} District
            (35 km radius threshold)
        </div>
        <table style="width:100%; border-collapse:collapse; margin-bottom:1rem;
                      border:1px solid #e2e8f0; border-radius:10px; overflow:hidden;">
            <thead>
                <tr style="background:#1e3a8a; color:white;">
                    <th style="padding:0.7rem 1rem; text-align:left;">#</th>
                    <th style="padding:0.7rem 1rem; text-align:left;">Destination (Upazila / District)</th>
                    <th style="padding:0.7rem 1rem; text-align:center;">Distance</th>
                    <th style="padding:0.7rem 1rem; text-align:left;">Departure</th>
                    <th style="padding:0.7rem 1rem; text-align:left;">Return</th>
                    <th style="padding:0.7rem 1rem; text-align:center;">Days</th>
                    <th style="padding:0.7rem 1rem; text-align:left;">BTS Location / Route</th>
                </tr>
            </thead>
            <tbody>{rows}</tbody>
        </table>"""
    else:
        trips_html = """<div style="background:#ecfdf5; border-left:4px solid #10b981;
                    border-radius:10px; padding:1rem 1.5rem; margin-bottom:1rem; color:#065f46;">
            No out-of-home travel detected (&ge;35 km from home location) within the analysis period.
        </div>"""

    # ── Gaps ──────────────────────────────────────────────────────────────
    gaps_html = ""
    if mv["gaps"]:
        gap_rows = "".join([
            f"""<tr style="background:{'#fff1f2' if i%2==0 else 'white'};">
                <td style="padding:0.7rem 1rem;">{g["gap_start"]}</td>
                <td style="padding:0.7rem 1rem;">{g["gap_end"]}</td>
                <td style="padding:0.7rem 1rem; text-align:center;">
                    <span style="background:#fee2e2;color:#dc2626;border-radius:12px;
                                 padding:0.2rem 0.7rem;font-weight:700;">{g["days"]} days</span>
                </td>
            </tr>"""
            for i, g in enumerate(mv["gaps"])
        ])
        gaps_html = f"""
        <div style="font-weight:700; color:#dc2626; margin:1rem 0 0.5rem 0;">
            Network Disconnection (No Activity &gt; 4 Days)
        </div>
        <table style="width:100%; border-collapse:collapse; border:1px solid #fecaca;
                      border-radius:10px; overflow:hidden;">
            <thead>
                <tr style="background:#dc2626; color:white;">
                    <th style="padding:0.7rem 1rem; text-align:left;">Last Seen</th>
                    <th style="padding:0.7rem 1rem; text-align:left;">Next Seen</th>
                    <th style="padding:0.7rem 1rem; text-align:center;">Gap Duration</th>
                </tr>
            </thead>
            <tbody>{gap_rows}</tbody>
        </table>"""

    return cards_html + trips_html + gaps_html

def imsi_change_analysis(df):
    """
    Track IMSI changes over time.
    Only considers IMSI numbers starting with '470' (Bangladesh).
    Returns list of {imsi, from_date, to_date, days, records} or None.
    """
    if 'imsi' not in df.columns or 'start' not in df.columns:
        return None

    # Filter valid Bangladesh IMSI (starts with 470)
    df_imsi = df[
        df['imsi'].notna() &
        df['imsi'].astype(str).str.strip().str.startswith('470')
    ].copy()

    if df_imsi.empty:
        return None

    unique_imsi = df_imsi['imsi'].astype(str).str.strip().unique()
    if len(unique_imsi) <= 1:
        return None   # Only one IMSI — no change

    # Sort by time and detect transitions
    df_imsi = df_imsi.sort_values('start').reset_index(drop=True)
    df_imsi['imsi_clean'] = df_imsi['imsi'].astype(str).str.strip()

    periods = []
    current_imsi  = df_imsi.iloc[0]['imsi_clean']
    period_start  = df_imsi.iloc[0]['start']
    period_end    = df_imsi.iloc[0]['start']
    period_count  = 1

    for _, row in df_imsi.iloc[1:].iterrows():
        imsi = row['imsi_clean']
        ts   = row['start']
        if imsi == current_imsi:
            period_end  = ts
            period_count += 1
        else:
            # IMSI changed
            days = max(1, (period_end - period_start).days + 1)
            periods.append({
                'IMSI':        current_imsi,
                'From':        str(period_start)[:19],
                'To':          str(period_end)[:19],
                'Days Active': days,
                'Records':     period_count,
            })
            current_imsi = imsi
            period_start = ts
            period_end   = ts
            period_count = 1

    # Last period
    days = max(1, (period_end - period_start).days + 1)
    periods.append({
        'IMSI':        current_imsi,
        'From':        str(period_start)[:19],
        'To':          str(period_end)[:19],
        'Days Active': days,
        'Records':     period_count,
    })

    return periods if len(periods) > 1 else None


def _imsi_change_html(df):
    """Generate HTML section for IMSI change tracking."""
    periods = imsi_change_analysis(df)
    if not periods:
        return ''   # No change — skip section entirely

    rows = ''.join([
        f"""<tr style="background:{'#f8fafc' if i%2==0 else 'white'};">
            <td style="padding:0.7rem 1rem; font-family:monospace; font-weight:600;">{p['IMSI']}</td>
            <td style="padding:0.7rem 1rem;">{p['From']}</td>
            <td style="padding:0.7rem 1rem;">{p['To']}</td>
            <td style="padding:0.7rem 1rem; text-align:center;">
                <span style="background:#dbeafe; color:#1e40af; border-radius:12px;
                             padding:0.2rem 0.7rem; font-weight:700;">{p['Days Active']}</span>
            </td>
            <td style="padding:0.7rem 1rem; text-align:center;">{p['Records']}</td>
        </tr>"""
        for i, p in enumerate(periods)
    ])

    return f"""
    <h2>2a. IMSI Change Analysis</h2>
    <div style="background:#fef3c7; border-left:4px solid #f59e0b; border-radius:8px;
                padding:0.75rem 1.25rem; margin-bottom:1rem; color:#92400e;">
        <strong>Multiple IMSI Detected!</strong> This SIM was used in {len(periods)} different
        IMSI periods — indicating possible SIM swap or dual-SIM activity.
    </div>
    <table style="width:100%; border-collapse:collapse; border:1px solid #e2e8f0; border-radius:10px; overflow:hidden;">
        <thead>
            <tr style="background:#1e3a8a; color:white;">
                <th style="padding:0.7rem 1rem; text-align:left;">IMSI Number</th>
                <th style="padding:0.7rem 1rem; text-align:left;">Active From</th>
                <th style="padding:0.7rem 1rem; text-align:left;">Active To</th>
                <th style="padding:0.7rem 1rem; text-align:center;">Days</th>
                <th style="padding:0.7rem 1rem; text-align:center;">Records</th>
            </tr>
        </thead>
        <tbody>{rows}</tbody>
    </table>"""


def imei_change_analysis(df):
    """
    Track IMEI changes over time.
    Only considers IMEI numbers with more than 8 digits.
    Returns list of {imei, from_date, to_date, days, records} or None.
    """
    if 'imei' not in df.columns or 'start' not in df.columns:
        return None

    # Filter valid IMEI (digits only, more than 8 digits)
    df_imei = df[
        df['imei'].notna() &
        df['imei'].astype(str).str.strip().str.match(r'^\d{9,}$')
    ].copy()

    if df_imei.empty:
        return None

    unique_imei = df_imei['imei'].astype(str).str.strip().unique()
    if len(unique_imei) <= 1:
        return None   # Only one IMEI — no change

    df_imei = df_imei.sort_values('start').reset_index(drop=True)
    df_imei['imei_clean'] = df_imei['imei'].astype(str).str.strip()

    periods = []
    current_imei  = df_imei.iloc[0]['imei_clean']
    period_start  = df_imei.iloc[0]['start']
    period_end    = df_imei.iloc[0]['start']
    period_count  = 1

    for _, row in df_imei.iloc[1:].iterrows():
        imei = row['imei_clean']
        ts   = row['start']
        if imei == current_imei:
            period_end   = ts
            period_count += 1
        else:
            days = max(1, (period_end - period_start).days + 1)
            periods.append({
                'IMEI':        current_imei,
                'From':        str(period_start)[:19],
                'To':          str(period_end)[:19],
                'Days Active': days,
                'Records':     period_count,
            })
            current_imei = imei
            period_start = ts
            period_end   = ts
            period_count = 1

    days = max(1, (period_end - period_start).days + 1)
    periods.append({
        'IMEI':        current_imei,
        'From':        str(period_start)[:19],
        'To':          str(period_end)[:19],
        'Days Active': days,
        'Records':     period_count,
    })

    return periods if len(periods) > 1 else None


def _imei_change_html(df):
    """Generate HTML section for IMEI change tracking (section 2b)."""
    periods = imei_change_analysis(df)
    if not periods:
        return ''

    rows = ''.join([
        f"""<tr style="background:{'#f8fafc' if i%2==0 else 'white'};">
            <td style="padding:0.7rem 1rem; font-family:monospace; font-weight:600;">{p['IMEI']}</td>
            <td style="padding:0.7rem 1rem;">{p['From']}</td>
            <td style="padding:0.7rem 1rem;">{p['To']}</td>
            <td style="padding:0.7rem 1rem; text-align:center;">
                <span style="background:#dbeafe; color:#1e40af; border-radius:12px;
                             padding:0.2rem 0.7rem; font-weight:700;">{p['Days Active']}</span>
            </td>
            <td style="padding:0.7rem 1rem; text-align:center;">{p['Records']}</td>
        </tr>"""
        for i, p in enumerate(periods)
    ])

    return f"""
    <h2>2b. IMEI Change Analysis</h2>
    <div style="background:#fef3c7; border-left:4px solid #f59e0b; border-radius:8px;
                padding:0.75rem 1.25rem; margin-bottom:1rem; color:#92400e;">
        <strong>Multiple IMEI Detected!</strong> This number was used in {len(periods)} different
        devices — indicating possible handset change.
    </div>
    <table style="width:100%; border-collapse:collapse; border:1px solid #e2e8f0;
                  border-radius:10px; overflow:hidden;">
        <thead>
            <tr style="background:#1e3a8a; color:white;">
                <th style="padding:0.7rem 1rem; text-align:left;">IMEI Number</th>
                <th style="padding:0.7rem 1rem; text-align:left;">Active From</th>
                <th style="padding:0.7rem 1rem; text-align:left;">Active To</th>
                <th style="padding:0.7rem 1rem; text-align:center;">Days</th>
                <th style="padding:0.7rem 1rem; text-align:center;">Records</th>
            </tr>
        </thead>
        <tbody>{rows}</tbody>
    </table>"""

def _target_number_html(df, target_number):
    if not target_number:
        return ''
    res = specific_number_analysis(df, target_number)
    if not res:
        return f'''<h2>10. Specific Number Analysis</h2>
        <div class="info-box">
        <strong>Target Number:</strong> {target_number}<br>
        <span class="warning">No communication found with this number.</span>
        </div>'''
    table_df = pd.DataFrame({
        'Metric': ['Target Number','Outgoing Calls (MOC)','Incoming Calls (MTC)',
                   'Total Calls','Total Call Duration (sec)','Total Call Duration (min)',
                   'Sent SMS','Received SMS','Total SMS',
                   'First Contact','Last Contact'],
        'Value':  [res['number'], res['moc'], res['mtc'], res['total_calls'],
                   f"{res['total_duration_sec']:,}", f"{res['total_duration_min']:,}",
                   res['sms_sent'], res['sms_received'], res['total_sms'],
                   res['first_contact'], res['last_contact']]
    })
    return f'''<h2>13. Specific Number Analysis</h2>
    <p>Communication summary between the subscriber and <strong>{target_number}</strong>.</p>
    {df_to_html(table_df)}'''

def build_html(df, phone, operator, date_range, total_raw, anomaly_count, target_number=None):
    import re as _re_html
    def _clean_id(series):
        result = []
        for val in series.dropna().unique():
            digits = _re_html.sub(r'[^0-9]', '', str(val).strip())
            if len(digits) >= 10:
                result.append(digits)
        return sorted(set(result))

    imei = _clean_id(df['imei']) if 'imei' in df.columns else []
    imsi = _clean_id(df['imsi']) if 'imsi' in df.columns else []

    home_mask    = df['start'].dt.hour.astype(int).isin(list(range(0,6))+list(range(22,24))) if 'start' in df.columns else None
    work_mask    = (df['start'].dt.hour.astype(int)>=8)&(df['start'].dt.hour.astype(int)<18)       if 'start' in df.columns else None
    weekend_mask = df['start'].dt.dayofweek.astype(int).isin([4,5])                               if 'start' in df.columns else None

    return f"""<html><head><title>CDR Analysis Report</title>
    <style>{CSS}</style></head><body>
    <h1>📞 CDR Analysis Report</h1>
    <h2>1. Executive Summary</h2>
    <div class="info-box">
    <strong>{"Device IMEI / SIM(s)" if is_imei_cdr(df) else "Phone Number"}:</strong> {phone}<br>
    <strong>Operator:</strong> {operator}<br>
    <strong>Analysis Period:</strong> {date_range}<br>
    <strong>Total Raw Records:</strong> {total_raw:,}<br>
    <strong>Anomalies Removed:</strong> {anomaly_count:,}<br>
    <strong>Records Analyzed:</strong> {len(df):,}
    </div>
    <h2>2. Device Information</h2>
    <table><tr><th>Field</th><th>Value</th></tr>
    <tr><td>IMEI</td><td>{', '.join(str(i) for i in imei) if imei else 'N/A'}</td></tr>
    <tr><td>IMSI</td><td>{', '.join(str(i) for i in imsi) if imsi else 'N/A'}</td></tr>
    <tr><td>{"Device IMEI / SIM(s)" if is_imei_cdr(df) else "Phone Number"}</td><td>{phone}</td></tr></table>
    {_imsi_change_html(df)}
    {_imei_change_html(df)}
    <h2>3. Call Analysis</h2>
    <h3>3.1 Call Analysis Summary</h3>{df_to_html(call_summary(df))}
    <h2>4. Call Count Analysis</h2>
    <h3>4.1 Daily Call Count</h3>{df_to_html(daily_call_count(df))}
    <h3>4.2 Hourly Call Count Graph</h3>{fig_to_html_img(plot_hourly(df))}
    <h3>4.3 Weekly Call Count</h3>{df_to_html(weekly_call_count(df))}
    <h3>4.3a Weekly Graph</h3>{fig_to_html_img(plot_weekly(df))}
    <h3>4.4 Monthly Call Count</h3>{df_to_html(monthly_call_count(df))}
    <h3>4.4a Monthly Graph</h3>{fig_to_html_img(plot_monthly(df))}
    <h2>5. Contact Analysis</h2>
    <h3>5.1 Contact Summary</h3>{df_to_html(contact_summary(df))}
    <h3>5.2 Top 10 Frequent Outgoing</h3>{df_to_html(top_contacts(df,'out',10))}
    <h3>5.3 Top 10 Frequent Incoming</h3>{df_to_html(top_contacts(df,'in',10))}
    <h3>5.4 Top 10 Lengthy Outgoing</h3>{df_to_html(top_lengthy(df,'out',10))}
    <h3>5.5 Top 10 Lengthy Incoming</h3>{df_to_html(top_lengthy(df,'in',10))}
    <h3>5.6 Top Call Overall</h3>{df_to_html(top_call_overall(df,10))}
    <h3>5.6a Top Call Overall Chart</h3>{fig_to_html_img(plot_top_call_overall(df,10))}
    <h2>6. Location Analysis</h2>
    <h3>6.1 Location Summary</h3>{df_to_html(location_summary(df))}
    <h3>6.2 Top 10 Frequent Locations</h3>{df_to_html(top_locations(df,None,10))}
    <h3>6.3 Frequent Locations Graph</h3>{fig_to_html_img(plot_locations(df,None,'Frequent Locations'))}
    <h3>6.4 Possible Home Locations</h3>{df_to_html(top_locations(df,home_mask,10))}
    <h3>6.5 Home Locations Graph</h3>{fig_to_html_img(plot_locations(df,home_mask,'Home Locations'))}
    <h3>6.6 Possible Work Locations</h3>{df_to_html(top_locations(df,work_mask,10))}
    <h3>6.7 Work Locations Graph</h3>{fig_to_html_img(plot_locations(df,work_mask,'Work Locations'))}
    <h3>6.8 Possible Weekend Locations</h3>{df_to_html(top_locations(df,weekend_mask,10))}
    <h3>6.9 Weekend Locations Graph</h3>{fig_to_html_img(plot_locations(df,weekend_mask,'Weekend Locations'))}

    <h2>7. SMS Contact Analysis</h2>
    <h3>7.1 Top 5 Sent SMS Contacts</h3>{df_to_html(top_sms_contacts(df,'out',5))}
    <h3>7.2 Top 5 Received SMS Contacts</h3>{df_to_html(top_sms_contacts(df,'in',5))}

    <h2>8. Last 10 Days Analysis</h2>
    <h3>8.1 Top Contacts in Last 10 Days (MOC + MTC)</h3>{df_to_html(last_n_days_top_contacts(df, 10, 10))}
    <h3>8.2 Top Locations in Last 10 Days</h3>{df_to_html(last_n_days_top_locations(df, 10, 10))}


    <h2>9. Movement Pattern Analysis</h2>
    <p>Analysis of movement outside estimated home/work district and network disconnection periods.</p>
    {_movement_html(movement_pattern_analysis(df))}

    {_target_number_html(df, target_number)}

    <h2>10. Overall Comment</h2>
    {''.join(f'<p>{ln}</p>' for ln in generate_overall_comment(df, phone, operator, date_range, total_raw))}
    <h2>11. Recommendation</h2>
    <ol>{''.join(f'<li>{r}</li>' for r in generate_recommendation(df))}</ol>
    <hr><p style="text-align:center;color:gray;font-size:11px;">
    Report generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | CDR Analysis Tool v1.0</p>
    </body></html>"""


# ─────────────────────────────────────────────
# WORD (DOCX) GENERATOR
# ─────────────────────────────────────────────
def build_docx(df, phone, operator, date_range, total_raw, anomaly_count, target_number=None):
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    doc = Document()
    sec = doc.sections[0]
    sec.page_width=Cm(21); sec.page_height=Cm(29.7)
    sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
    sec.left_margin=Cm(2.5); sec.right_margin=Cm(2.5)
    doc.styles['Normal'].font.name='Arial'
    doc.styles['Normal'].font.size=Pt(10)

    def add_h(text, lvl=1):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold=True; run.font.name='Arial'
        run.font.size=Pt(14 if lvl==1 else 11)
        run.font.color.rgb = RGBColor(0x1F,0x38,0x64) if lvl==1 else RGBColor(0x2E,0x74,0xB5)
        p.paragraph_format.space_before=Pt(10); p.paragraph_format.space_after=Pt(3)
        pPr=p._p.get_or_add_pPr(); pBdr=OxmlElement('w:pBdr')
        bot=OxmlElement('w:bottom'); bot.set(qn('w:val'),'single')
        bot.set(qn('w:sz'),'4'); bot.set(qn('w:space'),'1')
        bot.set(qn('w:color'),'2E74B5'); pBdr.append(bot); pPr.append(pBdr)

    def add_df_table(data):
        if data is None or data.empty:
            doc.add_paragraph('No data available.'); return
        cols = data.columns.tolist()
        tbl = doc.add_table(rows=1, cols=len(cols))
        tbl.style='Table Grid'
        for i,col in enumerate(cols):
            c=tbl.rows[0].cells[i]; c.text=str(col)
            r=c.paragraphs[0].runs[0]; r.bold=True
            r.font.size=Pt(9); r.font.name='Arial'
            r.font.color.rgb=RGBColor(255,255,255)
            tc=c._tc; tcPr=tc.get_or_add_tcPr()
            shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),'2E74B5')
            shd.set(qn('w:color'),'auto'); shd.set(qn('w:val'),'clear')
            tcPr.append(shd)
        for idx,row in data.iterrows():
            tr=tbl.add_row(); fill='EBF3FB' if idx%2==0 else 'FFFFFF'
            for i,col in enumerate(cols):
                c=tr.cells[i]; c.text=str(row[col]) if pd.notna(row[col]) else ''
                r=c.paragraphs[0].runs[0]; r.font.size=Pt(9); r.font.name='Arial'
                tc=c._tc; tcPr=tc.get_or_add_tcPr()
                shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),fill)
                shd.set(qn('w:color'),'auto'); shd.set(qn('w:val'),'clear')
                tcPr.append(shd)
        doc.add_paragraph()

    def add_fig(fig):
        if fig is None:
            doc.add_paragraph('Graph not available.'); return
        b64 = fig_to_base64(fig)
        img_data = base64.b64decode(b64)
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
            tmp.write(img_data); tmp_path=tmp.name
        doc.add_picture(tmp_path, width=Inches(6))
        doc.paragraphs[-1].alignment=WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph()

    # Title
    t=doc.add_paragraph(); t.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=t.add_run('CDR ANALYSIS REPORT'); r.bold=True
    r.font.size=Pt(20); r.font.name='Arial'
    r.font.color.rgb=RGBColor(0x1F,0x38,0x64)
    doc.add_paragraph()

    # Summary table
    add_h('1. Executive Summary')
    id_label = 'Device IMEI / SIM(s)' if is_imei_cdr(df) else 'Phone Number'
    info=[( id_label, phone),('Operator',operator),('Analysis Period',date_range),
          ('Total Raw Records',f'{total_raw:,}'),('Anomalies Removed',f'{anomaly_count:,}'),
          ('Records Analyzed',f'{len(df):,}')]
    tbl=doc.add_table(rows=len(info),cols=2); tbl.style='Table Grid'
    for i,(k,v) in enumerate(info):
        tbl.rows[i].cells[0].text=k
        tbl.rows[i].cells[0].paragraphs[0].runs[0].bold=True
        tbl.rows[i].cells[0].paragraphs[0].runs[0].font.size=Pt(9)
        tbl.rows[i].cells[1].text=str(v)
        tbl.rows[i].cells[1].paragraphs[0].runs[0].font.size=Pt(9)
    doc.add_paragraph()

    def _clean_id_list_d(series):
        import re as _re2
        result = []
        for val in series.dropna().unique():
            s = str(val).strip()
            digits = _re2.sub(r'[^0-9]', '', s)
            if len(digits) >= 10:
                result.append(digits)
        return sorted(set(result))
    imei=_clean_id_list_d(df['imei']) if 'imei' in df.columns else []
    imsi=_clean_id_list_d(df['imsi']) if 'imsi' in df.columns else []
    add_h('2. Device Information')
    dev=[('IMEI',', '.join(str(i) for i in imei) if imei else 'N/A'),
         ('IMSI',', '.join(str(i) for i in imsi) if imsi else 'N/A'),
         ('Phone Number',phone)]
    tbl=doc.add_table(rows=len(dev),cols=2); tbl.style='Table Grid'
    for i,(k,v) in enumerate(dev):
        tbl.rows[i].cells[0].text=k
        tbl.rows[i].cells[0].paragraphs[0].runs[0].bold=True
        tbl.rows[i].cells[0].paragraphs[0].runs[0].font.size=Pt(9)
        tbl.rows[i].cells[1].text=str(v)
        tbl.rows[i].cells[1].paragraphs[0].runs[0].font.size=Pt(9)
    doc.add_paragraph()

    home_mask    = df['start'].dt.hour.astype(int).isin(list(range(0,6))+list(range(22,24))) if 'start' in df.columns else None
    work_mask    = (df['start'].dt.hour.astype(int)>=8)&(df['start'].dt.hour.astype(int)<18)       if 'start' in df.columns else None
    weekend_mask = df['start'].dt.dayofweek.astype(int).isin([4,5])                               if 'start' in df.columns else None

    # IMSI Change section
    imsi_periods = imsi_change_analysis(df)
    if imsi_periods:
        add_h('2a. IMSI Change Analysis')
        doc.add_paragraph(
            f'Multiple IMSI Detected! This SIM was used in {len(imsi_periods)} '
            f'different IMSI periods — indicating possible SIM swap or dual-SIM activity.'
        )
        add_df_table(pd.DataFrame(imsi_periods))

    # IMEI Change section
    imei_periods = imei_change_analysis(df)
    if imei_periods:
        add_h('2b. IMEI Change Analysis')
        doc.add_paragraph(
            f'Multiple IMEI Detected! This number was used in {len(imei_periods)} '
            f'different devices — indicating possible handset change.'
        )
        add_df_table(pd.DataFrame(imei_periods))

    add_h('3. Call Analysis'); add_h('3.1 Call Analysis Summary',2); add_df_table(call_summary(df))
    add_h('4. Call Count Analysis')
    add_h('4.1 Daily Call Count',2);      add_df_table(daily_call_count(df))
    add_h('4.2 Hourly Graph',2);          add_fig(plot_hourly(df))
    add_h('4.3 Weekly Call Count',2);     add_df_table(weekly_call_count(df))
    add_h('4.3a Weekly Graph',2);         add_fig(plot_weekly(df))
    add_h('4.4 Monthly Call Count',2);    add_df_table(monthly_call_count(df))
    add_h('4.4a Monthly Graph',2);        add_fig(plot_monthly(df))
    add_h('5. Contact Analysis')
    add_h('5.1 Contact Summary',2);       add_df_table(contact_summary(df))
    add_h('5.2 Top 10 Outgoing',2);       add_df_table(top_contacts(df,'out',10))
    add_h('5.3 Top 10 Incoming',2);       add_df_table(top_contacts(df,'in',10))
    add_h('5.4 Lengthy Outgoing',2);      add_df_table(top_lengthy(df,'out',10))
    add_h('5.5 Lengthy Incoming',2);      add_df_table(top_lengthy(df,'in',10))
    add_h('5.6 Top Call Overall',2);      add_df_table(top_call_overall(df,10))
    add_h('5.6a Top Call Overall Chart',2); add_fig(plot_top_call_overall(df,10))
    add_h('6. Location Analysis')
    add_h('6.1 Location Summary',2);     add_df_table(location_summary(df))
    add_h('6.2 Frequent Locations',2);   add_df_table(top_locations(df,None,10))
    add_h('6.3 Locations Graph',2);      add_fig(plot_locations(df,None,'Frequent Locations'))
    add_h('6.4 Home Locations',2);       add_df_table(top_locations(df,home_mask,10))
    add_h('6.5 Home Graph',2);           add_fig(plot_locations(df,home_mask,'Home Locations'))
    add_h('6.6 Work Locations',2);       add_df_table(top_locations(df,work_mask,10))
    add_h('6.7 Work Graph',2);           add_fig(plot_locations(df,work_mask,'Work Locations'))
    add_h('6.8 Weekend Locations',2);    add_df_table(top_locations(df,weekend_mask,10))
    add_h('6.9 Weekend Graph',2);        add_fig(plot_locations(df,weekend_mask,'Weekend Locations'))

    add_h('7. SMS Contact Analysis')
    add_h('7.1 Top 5 Sent SMS Contacts',2);     add_df_table(top_sms_contacts(df,'out',5))
    add_h('7.2 Top 5 Received SMS Contacts',2); add_df_table(top_sms_contacts(df,'in',5))

    add_h('8. Last 10 Days Analysis')
    add_h('8.1 Top Contacts in Last 10 Days (MOC + MTC)',2)
    add_df_table(last_n_days_top_contacts(df, 10, 10))
    add_h('8.2 Top Locations in Last 10 Days',2)
    add_df_table(last_n_days_top_locations(df, 10, 10))

    # Movement Pattern section
    add_h('9. Movement Pattern Analysis')
    mv = movement_pattern_analysis(df)
    if mv:
        home_lbl = mv.get('home_label') or mv.get('home_district') or 'N/A'
        mv_summary = pd.DataFrame({
            'Metric': ['Home Location','Home District','Total Days',
                       'Out-of-Home Trips','Out-of-Home Days','Network Gaps (>4 days)'],
            'Value':  [home_lbl, mv['home_district'] or 'N/A',
                       str(mv['total_days']), str(len(mv['trips'])),
                       str(mv['out_of_home_days']), str(len(mv['gaps']))]
        })
        add_df_table(mv_summary)
        if mv['trips']:
            add_h('13.1 Out-of-Home Travel (35km+ from Home)', 2)
            trip_df = pd.DataFrame([{
                'Destination': t['upazila'], 'District': t['district'],
                'Distance(km)': t['km'], 'From': t['start_date'],
                'To': t['end_date'], 'Days': t['days'],
                'BTS Location': str(t.get('address',''))[:80]
            } for t in mv['trips']])
            add_df_table(trip_df)
        if mv['gaps']:
            add_h('13.2 Network Disconnection Periods', 2)
            gap_df = pd.DataFrame([{
                'Last Seen': g['gap_start'], 'Next Seen': g['gap_end'],
                'Gap (days)': g['days']
            } for g in mv['gaps']])
            add_df_table(gap_df)
    else:
        doc.add_paragraph('Insufficient location data for movement analysis.')

    if target_number:
        add_h('10. Specific Number Analysis')
        res = specific_number_analysis(df, target_number)
        if res is None:
            doc.add_paragraph(f'Target Number: {target_number}')
            doc.add_paragraph(f'No communication found with target number: {target_number}.')
        else:
            doc.add_paragraph(f'Communication summary between the subscriber and {target_number}.')
            target_df = pd.DataFrame({
                'Metric': ['Target Number','Outgoing Calls (MOC)','Incoming Calls (MTC)',
                           'Total Calls','Total Call Duration (sec)','Total Call Duration (min)',
                           'Sent SMS','Received SMS','Total SMS',
                           'First Contact','Last Contact'],
                'Value':  [str(res['number']), str(res['moc']), str(res['mtc']), str(res['total_calls']),
                           f"{res['total_duration_sec']:,}", f"{res['total_duration_min']:,}",
                           str(res['sms_sent']), str(res['sms_received']), str(res['total_sms']),
                           res['first_contact'], res['last_contact']]
            })
            add_df_table(target_df)

    add_h('11. Overall Comment')
    for ln in generate_overall_comment(df, phone, operator, date_range, total_raw):
        doc.add_paragraph(ln)

    add_h('12. Recommendation')
    for i, rec in enumerate(generate_recommendation(df), 1):
        doc.add_paragraph(f"{i}. {rec}")

    fp=doc.add_paragraph(f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")} | CDR Analysis Tool v1.0')
    fp.alignment=WD_ALIGN_PARAGRAPH.CENTER
    fp.runs[0].font.size=Pt(8); fp.runs[0].font.color.rgb=RGBColor(0x80,0x80,0x80)

    buf=io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────
def main():
    # ── Top App Header ──
    st.markdown("""
    <div class="app-header">
        <div class="app-header-left">
            <div class="app-header-logo">📊</div>
            <div class="app-header-title">CDR Analysis Platform</div>
        </div>
        <div class="app-header-nav">
            <span>📈 Dashboard Preview</span>
            <span>ℹ️ How It Works</span>
            <span>❓ Help</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Hero Section ──
    st.markdown("""
    <div class="hero">
        <div class="hero-icon-area">📱📊</div>
        <h1>CDR Analysis Platform</h1>
        <p>Upload your Call Detail Records (CDR) and generate actionable insights with automated reports in seconds.</p>
        <div class="hero-supports">✓ Supports All Major Operators</div>
        <div class="operator-row">
            <div class="op-badge"><div class="op-dot" style="background:#0073cf;"></div>Grameenphone</div>
            <div class="op-badge"><div class="op-dot" style="background:#e60028;"></div>Robi</div>
            <div class="op-badge"><div class="op-dot" style="background:#f57c00;"></div>Banglalink</div>
            <div class="op-badge"><div class="op-dot" style="background:#00a651;"></div>Teletalk</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Upload + Target Number Row ──
    up_col, num_col = st.columns([1.4, 1])

    with up_col:
        st.markdown("""
        <div class="card-title">
            <div class="card-icon-blue">📂</div>
            <div>
                <div>Upload CDR File</div>
                <div class="card-subtitle" style="font-weight:400;">Upload your Excel file (.XLSX or .XLS)</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        uploaded = st.file_uploader(
            label=" ",
            type=['xlsx', 'xls'],
            help="যেকোনো অপারেটরের CDR Excel ফাইল",
            label_visibility="collapsed"
        )
        st.caption("Maximum file size: 200MB")

    with num_col:
        st.markdown("""
        <div class="card-title">
            <div class="card-icon-blue">🎯</div>
            <div>
                <div>Target Number <span style="color:#94a3b8; font-weight:500; font-size:0.85rem;">(Optional)</span></div>
                <div class="card-subtitle" style="font-weight:400;">Enter a number to perform focused analysis.</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        target_number = st.text_input(
            label=" ",
            placeholder="e.g. 8801712345678",
            label_visibility="collapsed"
        )
        target_number = target_number.strip() if target_number else None

        st.markdown("""
        <div style="background:#eff6ff; border-radius:10px; padding:0.75rem 1rem; margin-top:0.5rem; display:flex; gap:0.6rem; align-items:flex-start;">
            <div style="color:#2563eb; font-size:1.1rem;">ℹ️</div>
            <div style="color:#1e40af; font-size:0.85rem; line-height:1.4;">
                Providing a target number helps generate detailed insights for that specific number.
            </div>
        </div>
        """, unsafe_allow_html=True)

    # ── Show landing content if no file uploaded ──
    if uploaded is None:
        # Info Banner
        st.markdown("""
        <div class="info-banner">
            <div class="info-banner-icon">⚡</div>
            <div style="flex:1;">
                <div class="info-banner-title">Instant Analysis & Reports</div>
                <div class="info-banner-text">After upload, the system will instantly generate interactive dashboards and downloadable HTML & Word reports with comprehensive insights.</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Key Capabilities
        st.markdown("""
        <div class="cap-card">
            <div class="cap-title">⭐ Key Capabilities</div>
            <div class="cap-grid">
                <div class="cap-item">
                    <div class="cap-icon" style="background:#dcfce7; color:#16a34a;">✓</div>
                    <div class="cap-text">Processes CDR data from any operator</div>
                </div>
                <div class="cap-item">
                    <div class="cap-icon" style="background:#dbeafe; color:#2563eb;">👥</div>
                    <div class="cap-text">Top contacts &amp; communication patterns</div>
                </div>
                <div class="cap-item">
                    <div class="cap-icon" style="background:#e0e7ff; color:#4f46e5;">🕒</div>
                    <div class="cap-text">Last 10 days activity tracking</div>
                </div>
                <div class="cap-item">
                    <div class="cap-icon" style="background:#fef3c7; color:#d97706;">⚠</div>
                    <div class="cap-text">Detects anomalies and invalid numbers</div>
                </div>
                <div class="cap-item">
                    <div class="cap-icon" style="background:#dcfce7; color:#16a34a;">📍</div>
                    <div class="cap-text">Location and movement analysis</div>
                </div>
                <div class="cap-item">
                    <div class="cap-icon" style="background:#fce7f3; color:#db2777;">🎯</div>
                    <div class="cap-text">Target-based number analysis</div>
                </div>
                <div class="cap-item">
                    <div class="cap-icon" style="background:#ede9fe; color:#7c3aed;">📅</div>
                    <div class="cap-text">Call summary (Daily, Weekly, Monthly)</div>
                </div>
                <div class="cap-item">
                    <div class="cap-icon" style="background:#fee2e2; color:#dc2626;">💬</div>
                    <div class="cap-text">SMS activity insights</div>
                </div>
                <div class="cap-item">
                    <div class="cap-icon" style="background:#cffafe; color:#0891b2;">📊</div>
                    <div class="cap-text">Auto-generated reports with graphs</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Trust Badges
        st.markdown("""
        <div class="trust-row">
            <div class="trust-item">
                <div class="trust-icon" style="background:#dbeafe; color:#2563eb;">🛡️</div>
                <div>
                    <div class="trust-title">Secure Processing</div>
                    <div class="trust-sub">Your data is processed securely</div>
                </div>
            </div>
            <div class="trust-item">
                <div class="trust-icon" style="background:#dcfce7; color:#16a34a;">🔒</div>
                <div>
                    <div class="trust-title">No Data Stored</div>
                    <div class="trust-sub">We don't store or share your data</div>
                </div>
            </div>
            <div class="trust-item">
                <div class="trust-icon" style="background:#ede9fe; color:#7c3aed;">⚡</div>
                <div>
                    <div class="trust-title">Fast &amp; Reliable</div>
                    <div class="trust-sub">Get results in seconds</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Footer
        st.markdown("""
        <div class="footer">
            🛡️ Developed By <span class="dev-name">Md. Omar Faruk Mazumder</span>
        </div>
        """, unsafe_allow_html=True)
        return

    # ── Process ──
    progress = st.progress(0, text="📥 Reading file...")

    try:
        file_bytes = uploaded.read()
        progress.progress(15, text="🔍 Analyzing data structure...")

        df, col_map, total_raw, anomaly_count, sheet = load_and_clean(file_bytes)
        df_clean = cdf(df)
        progress.progress(35, text="📊 Generating report...")

        phone      = get_phone(df)
        operator   = get_operator(df)
        date_range = get_date_range(df)
        base_name  = os.path.splitext(uploaded.name)[0]

        # ── Result Header ──
        st.markdown(f"""
        <div style="background:white; border-radius:14px; padding:1.5rem 2rem;
                    box-shadow:0 1px 3px rgba(0,0,0,0.05); margin-bottom:1.25rem;">
            <div style="font-size:0.85rem; color:#64748b; font-weight:600;
                        text-transform:uppercase; letter-spacing:0.5px; margin-bottom:0.75rem;">
                📊 Analysis Results
            </div>
            <div style="display:grid; grid-template-columns:repeat(4,1fr); gap:1rem;">
                <div>
                    <div style="font-size:0.78rem; color:#94a3b8; font-weight:600;
                                text-transform:uppercase; letter-spacing:0.5px;">Phone Number</div>
                    <div style="font-size:1.1rem; color:#0f172a; font-weight:700;
                                margin-top:0.2rem;">{phone}</div>
                </div>
                <div>
                    <div style="font-size:0.78rem; color:#94a3b8; font-weight:600;
                                text-transform:uppercase; letter-spacing:0.5px;">Operator</div>
                    <div style="font-size:1.1rem; color:#0f172a; font-weight:700;
                                margin-top:0.2rem;">{operator}</div>
                </div>
                <div>
                    <div style="font-size:0.78rem; color:#94a3b8; font-weight:600;
                                text-transform:uppercase; letter-spacing:0.5px;">Total Records</div>
                    <div style="font-size:1.5rem; color:#1e3a8a; font-weight:800;
                                margin-top:0.2rem;">{total_raw:,}</div>
                </div>
                <div>
                    <div style="font-size:0.78rem; color:#94a3b8; font-weight:600;
                                text-transform:uppercase; letter-spacing:0.5px;">Records Analyzed</div>
                    <div style="font-size:1.5rem; color:#16a34a; font-weight:800;
                                margin-top:0.2rem;">{len(df_clean):,}</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # ── Anomaly Info Banner ──
        if anomaly_count > 0:
            st.markdown(f"""
            <div style="background:#fffbeb; border-left:4px solid #f59e0b; border-radius:10px;
                        padding:0.85rem 1.25rem; margin-bottom:1rem; display:flex;
                        align-items:center; gap:0.75rem;">
                <div style="font-size:1.3rem;">⚠️</div>
                <div>
                    <strong style="color:#92400e;">{anomaly_count:,} Anomalous Records Detected</strong>
                    <div style="color:#b45309; font-size:0.88rem; margin-top:0.15rem;">
                        Service messages, promotional SMS, and invalid numbers were excluded from
                        call/contact analysis. However, their <strong>BTS location data is retained</strong>
                        for location analysis.
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

        progress.progress(55, text="📄 Generating HTML report...")
        html_content = build_html(df, phone, operator, date_range, total_raw, anomaly_count, target_number)
        html_bytes   = html_content.encode('utf-8')

        progress.progress(80, text="📝 Generating Word report...")
        docx_bytes = build_docx(df, phone, operator, date_range, total_raw, anomaly_count, target_number)

        progress.progress(100, text="✅ Complete!")

        # ── Success + Download ──
        st.markdown("""
        <div style="background:#ecfdf5; border:1px solid #10b981; border-radius:10px;
                    padding:1rem 1.5rem; margin:1rem 0; color:#065f46;">
            ✅ <strong>Reports generated successfully!</strong>
            Click the buttons below to download.
        </div>
        """, unsafe_allow_html=True)

        dl1, dl2 = st.columns(2)
        with dl1:
            st.download_button(
                label="⬇️ Download HTML Report",
                data=html_bytes,
                file_name=f"{base_name}_Report.html",
                mime="text/html",
                use_container_width=True
            )
        with dl2:
            st.download_button(
                label="⬇️ Download Word Report",
                data=docx_bytes,
                file_name=f"{base_name}_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

        # ── Section Divider ──
        st.markdown("""
        <div style="margin:1.5rem 0 1rem 0; display:flex; align-items:center; gap:1rem;">
            <div style="height:2px; flex:1; background:linear-gradient(90deg,#2563eb,#e2e8f0);
                        border-radius:2px;"></div>
            <div style="font-size:1rem; font-weight:700; color:#0f172a; white-space:nowrap;">
                📋 Quick Analysis Dashboard
            </div>
            <div style="height:2px; flex:1; background:linear-gradient(90deg,#e2e8f0,transparent);
                        border-radius:2px;"></div>
        </div>
        """, unsafe_allow_html=True)

        # ── 1. Device Info + Call Summary ──
        with st.expander("📊 Device Information & Call Analysis Summary", expanded=True):
            import re as _re_ui
            def _clean_ids(col):
                result = []
                for val in df[col].dropna().unique():
                    s = str(val).strip()
                    digits = _re_ui.sub(r'[^0-9]', '', s)
                    if len(digits) >= 10:
                        result.append(digits)
                return sorted(set(result))
            imei_list = _clean_ids("imei") if "imei" in df.columns else []
            imsi_list = _clean_ids("imsi") if "imsi" in df.columns else []

            st.markdown("""
            <div style="font-size:0.9rem; font-weight:700; color:#475569;
                        text-transform:uppercase; letter-spacing:0.5px; margin-bottom:0.75rem;">
                📱 Device Information
            </div>
            """, unsafe_allow_html=True)

            dev1, dev2, dev3, dev4 = st.columns(4)
            with dev1:
                st.markdown(f'<div class="stat-card"><div class="label">Phone Number</div><div class="value" style="font-size:0.95rem;">{phone}</div></div>', unsafe_allow_html=True)
            with dev2:
                st.markdown(f'<div class="stat-card"><div class="label">Operator</div><div class="value">{operator}</div></div>', unsafe_allow_html=True)
            with dev3:
                imei_val = imei_list[0] if imei_list else "N/A"
                st.markdown(f'<div class="stat-card"><div class="label">IMEI</div><div class="value" style="font-size:0.85rem;">{imei_val}</div></div>', unsafe_allow_html=True)
            with dev4:
                imsi_val = imsi_list[0] if imsi_list else "N/A"
                st.markdown(f'<div class="stat-card"><div class="label">IMSI</div><div class="value" style="font-size:0.85rem;">{imsi_val}</div></div>', unsafe_allow_html=True)

            st.markdown("<hr style='margin:1rem 0; border-color:#f1f5f9;'>", unsafe_allow_html=True)
            st.markdown("""
            <div style="font-size:0.9rem; font-weight:700; color:#475569;
                        text-transform:uppercase; letter-spacing:0.5px; margin-bottom:0.75rem;">
                📞 Call Analysis Summary
            </div>
            """, unsafe_allow_html=True)

            cs = call_summary(df)
            ca1, ca2, ca3, ca4 = st.columns(4)
            for i, (col_st, row) in enumerate(zip([ca1, ca2, ca3, ca4], cs.itertuples())):
                with col_st:
                    st.markdown(f'<div class="stat-card"><div class="label">{row.Metric}</div><div class="value">{row.Value:,}</div></div>', unsafe_allow_html=True)

            st.markdown("<hr style='margin:1rem 0; border-color:#f1f5f9;'>", unsafe_allow_html=True)

            # IMSI Change Detection
            imsi_periods = imsi_change_analysis(df)
            if imsi_periods:
                st.markdown("""
                <div style="background:#fef3c7; border-left:4px solid #f59e0b;
                            border-radius:8px; padding:0.75rem 1.25rem; margin-bottom:0.75rem; color:#92400e;">
                    <strong>2a. Multiple IMSI Detected!</strong>
                    Possible SIM swap or dual-SIM activity found.
                </div>""", unsafe_allow_html=True)
                imsi_df = pd.DataFrame(imsi_periods)
                st.dataframe(imsi_df, use_container_width=True, hide_index=True)

            # IMEI Change Detection
            imei_periods = imei_change_analysis(df)
            if imei_periods:
                st.markdown("""
                <div style="background:#ffe4e6; border-left:4px solid #f43f5e;
                            border-radius:8px; padding:0.75rem 1.25rem; margin-bottom:0.75rem; color:#881337;">
                    <strong>2b. Multiple IMEI Detected!</strong>
                    Possible device/handset change found.
                </div>""", unsafe_allow_html=True)
                imei_df = pd.DataFrame(imei_periods)
                st.dataframe(imei_df, use_container_width=True, hide_index=True)

            st.markdown("<hr style='margin:1rem 0; border-color:#f1f5f9;'>", unsafe_allow_html=True)
            st.markdown("""
            <div style="font-size:0.9rem; font-weight:700; color:#475569;
                        text-transform:uppercase; letter-spacing:0.5px; margin-bottom:0.75rem;">
                📅 Analysis Period
            </div>
            """, unsafe_allow_html=True)
            st.markdown(f"""
            <div style="background:#f8fafc; border-radius:8px; padding:0.75rem 1rem;
                        color:#334155; font-size:0.95rem;">
                🗓️ <strong>{date_range}</strong>
            </div>
            """, unsafe_allow_html=True)

        # ── 2. Top 10 Contacts ──
        with st.expander("📞 Top 10 Contacts — Outgoing & Incoming", expanded=True):
            t1, t2 = st.columns(2)
            with t1:
                st.markdown("""<div style="font-weight:700; color:#2563eb; margin-bottom:0.5rem;">
                    📤 Top 10 Outgoing Contacts (MOC)</div>""", unsafe_allow_html=True)
                out_df = top_contacts(df, 'out', 10)
                if not out_df.empty:
                    st.dataframe(out_df, use_container_width=True, hide_index=True)
                    fig = plot_contacts(df, 'out', 10, 'Top 10 Outgoing Contacts')
                    if fig: st.pyplot(fig)
                else:
                    st.info("No outgoing call data available.")
            with t2:
                st.markdown("""<div style="font-weight:700; color:#16a34a; margin-bottom:0.5rem;">
                    📥 Top 10 Incoming Contacts (MTC)</div>""", unsafe_allow_html=True)
                in_df = top_contacts(df, 'in', 10)
                if not in_df.empty:
                    st.dataframe(in_df, use_container_width=True, hide_index=True)
                    fig = plot_contacts(df, 'in', 10, 'Top 10 Incoming Contacts')
                    if fig: st.pyplot(fig)
                else:
                    st.info("No incoming call data available.")

        # ── 3. Top Contacts by Call Duration ──
        with st.expander("⏱️ Top 10 Contacts by Call Duration", expanded=True):
            d1, d2 = st.columns(2)
            with d1:
                st.markdown("""<div style="font-weight:700; color:#2563eb; margin-bottom:0.5rem;">
                    📤 Outgoing — Longest Call Duration</div>""", unsafe_allow_html=True)
                lo_df = top_lengthy(df, 'out', 10)
                if not lo_df.empty:
                    st.dataframe(lo_df, use_container_width=True, hide_index=True)
                else:
                    st.info("No outgoing call duration data available.")
            with d2:
                st.markdown("""<div style="font-weight:700; color:#16a34a; margin-bottom:0.5rem;">
                    📥 Incoming — Longest Call Duration</div>""", unsafe_allow_html=True)
                li_df = top_lengthy(df, 'in', 10)
                if not li_df.empty:
                    st.dataframe(li_df, use_container_width=True, hide_index=True)
                else:
                    st.info("No incoming call duration data available.")

        # ── 4. Top Stay Locations (uses FULL df including anomalies for BTS data) ──
        with st.expander("📍 Top Stay Locations", expanded=True):
            if 'address' in df.columns:
                home_mask    = df['start'].dt.hour.astype(int).isin(list(range(0,6))+list(range(22,24))) if 'start' in df.columns else None
                work_mask    = (df['start'].dt.hour.astype(int)>=8)&(df['start'].dt.hour.astype(int)<18) if 'start' in df.columns else None
                weekend_mask = df['start'].dt.dayofweek.astype(int).isin([4,5])                         if 'start' in df.columns else None

                st.markdown("""
                <div style="background:#eff6ff; border-radius:8px; padding:0.6rem 1rem;
                            margin-bottom:1rem; font-size:0.85rem; color:#1e40af;">
                    ℹ️ Location analysis includes <strong>all records</strong> (including service messages)
                    to ensure complete BTS coverage data.
                </div>""", unsafe_allow_html=True)

                l1, l2, l3 = st.columns(3)
                with l1:
                    st.markdown("""<div style="font-weight:700; color:#1e3a8a; margin-bottom:0.25rem;">
                        🏠 Probable Home Location</div>
                        <div style="font-size:0.8rem; color:#64748b; margin-bottom:0.5rem;">
                        Night (10 PM – 6 AM)</div>""", unsafe_allow_html=True)
                    hl = top_locations(df, home_mask, 3)
                    if not hl.empty:
                        st.dataframe(hl, use_container_width=True, hide_index=True)
                    else:
                        st.info("No location data available.")

                with l2:
                    st.markdown("""<div style="font-weight:700; color:#1e3a8a; margin-bottom:0.25rem;">
                        🏢 Probable Work Location</div>
                        <div style="font-size:0.8rem; color:#64748b; margin-bottom:0.5rem;">
                        Daytime (8 AM – 6 PM)</div>""", unsafe_allow_html=True)
                    wl = top_locations(df, work_mask, 3)
                    if not wl.empty:
                        st.dataframe(wl, use_container_width=True, hide_index=True)
                    else:
                        st.info("No location data available.")

                with l3:
                    st.markdown("""<div style="font-weight:700; color:#1e3a8a; margin-bottom:0.25rem;">
                        🕌 Probable Weekend Location</div>
                        <div style="font-size:0.8rem; color:#64748b; margin-bottom:0.5rem;">
                        Friday & Saturday</div>""", unsafe_allow_html=True)
                    el = top_locations(df, weekend_mask, 3)
                    if not el.empty:
                        st.dataframe(el, use_container_width=True, hide_index=True)
                    else:
                        st.info("No location data available.")
            else:
                st.info("No location data available in this CDR file.")

        # ── 5. Top 5 SMS Contacts ──
        with st.expander("💬 Top 5 SMS Contacts", expanded=True):
            if "is_sms_out" in df.columns and "party_b_clean" in df.columns:
                s1, s2 = st.columns(2)
                with s1:
                    st.markdown("""<div style="font-weight:700; color:#2563eb; margin-bottom:0.5rem;">
                        📤 Top 5 Sent SMS Contacts</div>""", unsafe_allow_html=True)
                    sms_out_df = top_sms_contacts(df, 'out', 5)
                    if not sms_out_df.empty:
                        st.dataframe(sms_out_df, use_container_width=True, hide_index=True)
                    else:
                        st.info("No sent SMS data available.")
                with s2:
                    st.markdown("""<div style="font-weight:700; color:#16a34a; margin-bottom:0.5rem;">
                        📥 Top 5 Received SMS Contacts</div>""", unsafe_allow_html=True)
                    sms_in_df = top_sms_contacts(df, 'in', 5)
                    if not sms_in_df.empty:
                        st.dataframe(sms_in_df, use_container_width=True, hide_index=True)
                    else:
                        st.info("No received SMS data available.")
            else:
                st.info("No SMS data available in this CDR file.")

        # ── 6. Movement Pattern Analysis ──
        with st.expander("🗺️ Movement Pattern Analysis", expanded=True):
            mv = movement_pattern_analysis(df)
            if mv:
                # Summary cards
                mv1, mv2, mv3, mv4, mv5 = st.columns(5)
                with mv1:
                    st.markdown(f'''<div class="stat-card" style="border-left-color:#2563eb;">
                        <div class="label">Total Records</div>
                        <div class="value">{mv["total_records"]:,}</div>
                        <div style="font-size:0.75rem;color:#94a3b8;">{mv["total_days"]} days</div>
                    </div>''', unsafe_allow_html=True)
                with mv2:
                    st.markdown(f'''<div class="stat-card" style="border-left-color:#16a34a;">
                        <div class="label">Home Location</div>
                        <div class="value" style="font-size:1rem;">{mv.get("home_label") or mv["home_district"] or "N/A"}</div>
                        <div style="font-size:0.75rem;color:#94a3b8;">{mv["home_district"] or ""} District (most frequent)</div>
                    </div>''', unsafe_allow_html=True)
                with mv4:
                    gap_color = "#dc2626" if mv["gaps"] else "#16a34a"
                    st.markdown(f'''<div class="stat-card" style="border-left-color:{gap_color};">
                        <div class="label">Network Gaps</div>
                        <div class="value" style="color:{gap_color};">{len(mv["gaps"])}</div>
                        <div style="font-size:0.75rem;color:#94a3b8;">&gt;4 days offline</div>
                    </div>''', unsafe_allow_html=True)
                with mv5:
                    st.markdown(f'''<div class="stat-card" style="border-left-color:#7c3aed;">
                        <div class="label">Out-of-Home Trips</div>
                        <div class="value" style="color:#7c3aed;">{len(mv["trips"])}</div>
                        <div style="font-size:0.75rem;color:#94a3b8;">{mv["out_of_home_days"]} days total</div>
                    </div>''', unsafe_allow_html=True)

                st.markdown("")

                # Trips
                if mv['trips']:
                    st.markdown("""<div style="background:#fffbeb; border-left:4px solid #f59e0b;
                        border-radius:8px; padding:0.75rem 1.25rem; color:#92400e; margin-bottom:0.75rem;">
                        <strong>Out-of-Home District Travel Detected</strong></div>""",
                        unsafe_allow_html=True)
                    trip_df = pd.DataFrame([{
                        'District': t['district'],
                        'From': t['start_date'], 'To': t['end_date'],
                        'Days': t['days'],
                        'BTS Location': str(t.get('address',''))[:80]
                    } for t in mv['trips']])
                    st.dataframe(trip_df, use_container_width=True, hide_index=True)
                else:
                    st.success("✅ No out-of-home-district travel detected.")

                # Network gaps
                if mv['gaps']:
                    st.markdown("""<div style="font-weight:700; color:#dc2626; margin-top:1rem;">
                        📵 Network Disconnection Periods (>4 Days)</div>""", unsafe_allow_html=True)
                    gap_df = pd.DataFrame([{
                        'Last Seen': g['gap_start'],
                        'Next Seen': g['gap_end'],
                        'Gap Duration (days)': g['days']
                    } for g in mv['gaps']])
                    st.dataframe(gap_df, use_container_width=True, hide_index=True)
                else:
                    st.success("✅ No network disconnection gaps (>4 days) detected.")
            else:
                st.info("Location data insufficient for movement analysis.")

        # ── 7. Last 10 Days Analysis ──
        with st.expander("📅 Last 10 Days Activity", expanded=True):
            ld1, ld2 = st.columns(2)
            with ld1:
                st.markdown("""<div style="font-weight:700; color:#2563eb; margin-bottom:0.5rem;">
                    📞 Top Contacts — Last 10 Days (MOC + MTC)</div>""", unsafe_allow_html=True)
                last_calls = last_n_days_top_contacts(df, 10, 10)
                if not last_calls.empty:
                    st.dataframe(last_calls, use_container_width=True, hide_index=True)
                else:
                    st.info("No data available for the last 10 days.")
            with ld2:
                st.markdown("""<div style="font-weight:700; color:#7c3aed; margin-bottom:0.5rem;">
                    📍 Top Locations — Last 10 Days</div>""", unsafe_allow_html=True)
                last_loc = last_n_days_top_locations(df, 10, 10)
                if not last_loc.empty:
                    st.dataframe(last_loc, use_container_width=True, hide_index=True)
                else:
                    st.info("No location data for the last 10 days.")

        # ── 7. Specific Number Analysis ──
        if target_number:
            with st.expander(f"🎯 Specific Number Analysis — {target_number}", expanded=True):
                res = specific_number_analysis(df, target_number)
                if res is None:
                    st.warning(f"No communication found with {target_number} in this CDR.")
                else:
                    sn1, sn2, sn3, sn4 = st.columns(4)
                    with sn1:
                        st.markdown(f'<div class="stat-card"><div class="label">Outgoing Calls (MOC)</div><div class="value">{res["moc"]}</div></div>', unsafe_allow_html=True)
                    with sn2:
                        st.markdown(f'<div class="stat-card"><div class="label">Incoming Calls (MTC)</div><div class="value">{res["mtc"]}</div></div>', unsafe_allow_html=True)
                    with sn3:
                        st.markdown(f'<div class="stat-card"><div class="label">Total Call Duration</div><div class="value" style="font-size:1.1rem;">{res["total_duration_min"]} min</div></div>', unsafe_allow_html=True)
                    with sn4:
                        st.markdown(f'<div class="stat-card"><div class="label">Total SMS</div><div class="value">{res["total_sms"]}</div></div>', unsafe_allow_html=True)
                    st.markdown("")
                    detail_df = pd.DataFrame({
                        'Metric': ['Total Calls','Sent SMS','Received SMS','First Contact','Last Contact'],
                        'Value':  [res['total_calls'], res['sms_sent'], res['sms_received'],
                                   res['first_contact'], res['last_contact']]
                    })
                    st.dataframe(detail_df, use_container_width=True, hide_index=True)

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        st.code(str(e))

    # ── Footer ──
    st.markdown("""
    <div class="footer">
        🛡️ Developed By <span class="dev-name">Md. Omar Faruk Mazumder</span>
    </div>
    """, unsafe_allow_html=True)


if __name__ == '__main__':
    main()
