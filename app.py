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
# CUSTOM CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1F3864 0%, #2E74B5 100%);
        padding: 2rem;
        border-radius: 12px;
        text-align: center;
        margin-bottom: 2rem;
        color: white;
    }
    .main-header h1 { color: white; margin: 0; font-size: 2.2rem; }
    .main-header p  { color: #cce0ff; margin: 0.5rem 0 0 0; font-size: 1rem; }

    .stat-card {
        background: #f0f4ff;
        border-left: 4px solid #2E74B5;
        padding: 1rem 1.2rem;
        border-radius: 8px;
        margin-bottom: 0.5rem;
    }
    .stat-card .label { font-size: 0.8rem; color: #666; font-weight: 600; text-transform: uppercase; }
    .stat-card .value { font-size: 1.4rem; color: #1F3864; font-weight: 700; }

    .success-box {
        background: #e8f5e9;
        border: 1px solid #4caf50;
        border-radius: 8px;
        padding: 1rem 1.5rem;
        margin: 1rem 0;
    }
    .warning-box {
        background: #fff8e1;
        border: 1px solid #ffc107;
        border-radius: 8px;
        padding: 0.8rem 1.2rem;
        margin: 0.5rem 0;
    }
    .download-btn {
        display: inline-block;
        padding: 0.6rem 1.2rem;
        border-radius: 8px;
        font-weight: 600;
        text-decoration: none;
        margin: 0.3rem;
    }
    div[data-testid="stDownloadButton"] button {
        width: 100%;
        border-radius: 8px;
        font-weight: 600;
        padding: 0.6rem;
    }
    .stProgress > div > div > div {
        background: linear-gradient(90deg, #1F3864, #2E74B5);
    }
    footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# COLUMN ALIASES
# ─────────────────────────────────────────────
COLUMN_ALIASES = {
    'start':            ['start', 'datetime', 'date', 'timestamp', 'call_date',
                         'call date', 'date/time', 'date time', 'starttime',
                         'start time', 'start_datetime'],
    'operator':         ['operator', 'network', 'telco', 'carrier',
                         'provider name', 'provider_name', 'service provider'],
    'party_a':          ['party a', 'party_a', 'a_number', 'msisdn_a', 'a-number',
                         'caller', 'originating', 'a number', 'partya', 'msisdn',
                         'aparty', 'a party', 'a_party'],
    'party_b':          ['party b', 'party_b', 'b_number', 'msisdn_b', 'b-number',
                         'called', 'terminating', 'b number', 'partyb', 'callee',
                         'bparty', 'b party', 'b_party'],
    'party_b_original': ['party b original', 'party_b_original', 'original_b',
                         'b_original', 'partyb_original'],
    'duration':         ['call duration', 'call_duration', 'duration',
                         'call_length', 'duration_sec', 'duration(sec)'],
    'usage_type':       ['usage type', 'usage_type', 'call_type', 'call type',
                         'type', 'direction', 'service_type'],
    'cell_type':        ['cell type', 'cell_type', 'network_type', 'network type',
                         'technology', 'rat'],
    'lac':              ['lac id', 'lac_id', 'lac', 'location_area_code',
                         'lacstarta', 'lacstart'],
    'cell_id':          ['cell id', 'cell_id', 'cell', 'bts_id', 'bts id',
                         'site_id', 'tower_id', 'cistarta'],
    'imei':             ['imei', 'device_id', 'handset_id'],
    'imsi':             ['imsi', 'subscriber_id', 'imsia'],
    'address':          ['address', 'location', 'tower_location', 'tower location',
                         'site_name', 'cell_name', 'area', 'thana', 'district',
                         'bts address', 'bts_address'],
}

CALL_OUT_TYPES = ['moc', 'mo', 'outgoing', 'out', 'call-mo', 'call_mo', 'callmo']
CALL_IN_TYPES  = ['mtc', 'mt', 'incoming', 'in', 'call-mt', 'call_mt', 'callmt']
SMS_OUT_TYPES  = ['mo-sms', 'sms-mo', 'smsmo', 'sms_mo', 'sms-mo', 'sms out']
SMS_IN_TYPES   = ['mt-sms', 'sms-mt', 'smsmt', 'sms_mt', 'sms-mt', 'sms in']


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
def load_and_clean(file_bytes):
    # Detect best sheet
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

    # No header detection
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

    # Rename columns
    col_map = {}
    for key in COLUMN_ALIASES:
        found = detect_column(df.columns, key)
        if found:
            col_map[key] = found
    rename = {v: k for k, v in col_map.items()}
    df = df.rename(columns=rename)

    # Parse datetime
    if 'start' in df.columns:
        df['start'] = pd.to_datetime(df['start'], errors='coerce')
        df = df.dropna(subset=['start'])
        df = df.sort_values('start').reset_index(drop=True)

    # Parse duration
    if 'duration' in df.columns:
        df['duration'] = pd.to_numeric(df['duration'], errors='coerce').fillna(0).astype(int)

    # Fix Party B
    if 'party_b_original' in df.columns:
        pb_orig = df['party_b_original'].replace(['nan', 'None'], pd.NA)
        if pb_orig.notna().sum() > 0:
            df['party_b_clean'] = pb_orig.fillna(
                df.get('party_b', pb_orig)).astype(str).str.strip()
        elif 'party_b' in df.columns:
            df['party_b_clean'] = df['party_b'].astype(str).str.strip()
    elif 'party_b' in df.columns:
        df['party_b_clean'] = df['party_b'].astype(str).str.strip()

    anomaly_count = 0
    if 'party_b_clean' in df.columns:
        anomaly_mask = ~df['party_b_clean'].apply(is_valid_number)
        anomaly_count = anomaly_mask.sum()
        df = df[~anomaly_mask].reset_index(drop=True)

    if 'party_a' in df.columns:
        df['party_a'] = df['party_a'].astype(str).str.strip()
    if 'usage_type' in df.columns:
        df['usage_type'] = df['usage_type'].astype(str).str.strip()

    # Categorize
    ut = df['usage_type'].str.lower() if 'usage_type' in df.columns else pd.Series(['moc']*len(df))
    df['is_call_out'] = ut.isin(CALL_OUT_TYPES)
    df['is_call_in']  = ut.isin(CALL_IN_TYPES)
    df['is_sms_out']  = ut.isin(SMS_OUT_TYPES)
    df['is_sms_in']   = ut.isin(SMS_IN_TYPES)

    return df, col_map, total_raw, anomaly_count, best_sheet


# ─────────────────────────────────────────────
# ANALYSIS FUNCTIONS (same as cdr_analysis.py)
# ─────────────────────────────────────────────
def get_phone(df):
    if 'party_a' in df.columns and len(df) > 0:
        return str(df['party_a'].mode()[0])
    return 'N/A'

def get_operator(df):
    if 'operator' in df.columns:
        ops = df['operator'].dropna().unique()
        return ', '.join(str(o) for o in ops) if len(ops) > 0 else 'N/A'
    return 'N/A'

def get_date_range(df):
    if 'start' in df.columns and len(df) > 0:
        return (f"{df['start'].min().strftime('%Y-%m-%d %H:%M:%S')} to "
                f"{df['start'].max().strftime('%Y-%m-%d %H:%M:%S')}")
    return 'N/A'

def call_summary(df):
    return pd.DataFrame({
        'Metric': ['Total Outgoing Calls','Total Incoming Calls',
                   'Total Sent SMS','Total Received SMS'],
        'Value':  [int(df['is_call_out'].sum()), int(df['is_call_in'].sum()),
                   int(df['is_sms_out'].sum()), int(df['is_sms_in'].sum())]
    })

def daily_call_count(df):
    if 'start' not in df.columns: return pd.DataFrame()
    calls = df[df['is_call_out'] | df['is_call_in']].copy()
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
        mc = sub['party_b_clean'].value_counts() if 'party_b_clean' in sub.columns and len(sub) else pd.Series()
        mv = sub['address'].dropna().value_counts() if 'address' in sub.columns and len(sub) else pd.Series()
        rows.append({'Time of day': name, 'Interval': interval,
                     'Total calls': len(sub),
                     'Most Contacted': f"{mc.index[0]} ({mc.iloc[0]})" if not mc.empty else 'N/A',
                     'Most Visited Place': f"{mv.index[0]} ({mv.iloc[0]})" if not mv.empty else 'N/A'})
    return pd.DataFrame(rows)

def weekly_call_count(df):
    if 'start' not in df.columns: return pd.DataFrame()
    calls = df[df['is_call_out']|df['is_call_in']].copy()
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
    calls = df[df['is_call_out']|df['is_call_in']].copy()
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
    all_c = df[df['is_call_out']|df['is_call_in']]
    mc_all = all_c['party_b_clean'].value_counts() if len(all_c) else pd.Series()
    mc_out = df[df['is_call_out']]['party_b_clean'].value_counts() if df['is_call_out'].sum() else pd.Series()
    mc_in  = df[df['is_call_in']]['party_b_clean'].value_counts()  if df['is_call_in'].sum()  else pd.Series()
    dur    = all_c.groupby('party_b_clean')['duration'].sum() if 'duration' in df.columns and len(all_c) else pd.Series()
    return pd.DataFrame({
        'Metric': ['Total Unique Numbers','Most Called Number',
                   'Most Called Outgoing','Most Received Incoming',
                   'Most Total Call Time'],
        'Value':  [df['party_b_clean'].nunique(),
                   f"{mc_all.index[0]}, {mc_all.iloc[0]} times" if not mc_all.empty else 'N/A',
                   f"{mc_out.index[0]}, {mc_out.iloc[0]} times" if not mc_out.empty else 'N/A',
                   f"{mc_in.index[0]},  {mc_in.iloc[0]} times"  if not mc_in.empty  else 'N/A',
                   f"{dur.idxmax()}, {round(dur.max()/60,1)} min" if not dur.empty else 'N/A']
    })

def top_contacts(df, direction='out', n=10):
    if 'party_b_clean' not in df.columns: return pd.DataFrame()
    sub = df[df['is_call_out']] if direction=='out' else df[df['is_call_in']]
    if len(sub)==0: return pd.DataFrame()
    c = sub['party_b_clean'].value_counts().head(n)
    return pd.DataFrame({'Party B': c.index,
                         'Total Calls': c.values,
                         'Percentage': (c.values/len(sub)*100).round(2)})

def top_lengthy(df, direction='out', n=10):
    if 'party_b_clean' not in df.columns or 'duration' not in df.columns: return pd.DataFrame()
    sub = df[df['is_call_out']] if direction=='out' else df[df['is_call_in']]
    if len(sub)==0: return pd.DataFrame()
    g = sub.groupby('party_b_clean').agg(
        Total_Duration=('duration','sum'), Total_Calls=('duration','count')
    ).sort_values('Total_Duration', ascending=False).head(n)
    td = sub['duration'].sum()
    g['Pct_CallTime'] = (g['Total_Duration']/td*100).round(2) if td>0 else 0
    return g.reset_index().rename(columns={'party_b_clean':'Party B'})

def top_locations(df, mask=None, n=10):
    if 'address' not in df.columns: return pd.DataFrame()
    data = df if mask is None else df[mask]
    c = data['address'].dropna().value_counts().head(n)
    if c.empty: return pd.DataFrame()
    return pd.DataFrame({'Address': c.index, 'Count': c.values})

def location_summary(df):
    if 'address' not in df.columns: return pd.DataFrame()
    addrs = df['address'].dropna()
    if len(addrs)==0: return pd.DataFrame()
    mv = addrs.value_counts()

    def top_addr(mask):
        if mask is None or 'start' not in df.columns: return ('N/A', 0)
        sub = df[mask]['address'].dropna().value_counts()
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
        s = str(val).strip()
        if not s.isdigit():
            return False
        # Bangladesh mobile: 8801XXXXXXXXX (13 digits) or 01XXXXXXXXX (11 digits)
        if s.startswith('8801') and len(s) == 13:
            return True
        if s.startswith('01') and len(s) == 11:
            return True
        # International mobile: 10-15 digits, NOT starting with 0 or 88 (with non-1 third digit)
        if 10 <= len(s) <= 15 and not s.startswith('0') and not s.startswith('880'):
            return True
        return False

    sub = df[df[col] & df['party_b_clean'].apply(is_real_mobile)]
    if len(sub) == 0:
        return pd.DataFrame()
    counts = sub['party_b_clean'].value_counts().head(n)
    return pd.DataFrame({'Party B': counts.index, 'SMS Count': counts.values})


def last_n_days_top_contacts(df, days=10, n=10):
    """Top contacts (MOC + MTC combined) in last N days."""
    if 'start' not in df.columns or 'party_b_clean' not in df.columns:
        return pd.DataFrame()
    max_date = df['start'].max()
    cutoff = max_date - pd.Timedelta(days=days)
    recent = df[(df['start'] >= cutoff) & (df['is_call_out'] | df['is_call_in'])].copy()
    if len(recent) == 0:
        return pd.DataFrame()
    grouped = recent.groupby('party_b_clean').agg(
        moc=('is_call_out', 'sum'),
        mtc=('is_call_in', 'sum'),
        total_duration=('duration', 'sum')
    )
    grouped['total_calls'] = grouped['moc'] + grouped['mtc']
    grouped = grouped.sort_values('total_calls', ascending=False).head(n).reset_index()
    grouped['Duration (min)'] = (grouped['total_duration'] / 60).round(1)
    return grouped[['party_b_clean', 'moc', 'mtc', 'total_calls', 'Duration (min)']].rename(
        columns={'party_b_clean': 'Party B', 'moc': 'MOC',
                 'mtc': 'MTC', 'total_calls': 'Total Calls'})


def last_n_days_top_locations(df, days=10, n=10):
    """Top locations in last N days."""
    if 'start' not in df.columns or 'address' not in df.columns:
        return pd.DataFrame()
    max_date = df['start'].max()
    cutoff = max_date - pd.Timedelta(days=days)
    recent = df[df['start'] >= cutoff]
    addrs = recent['address'].dropna().value_counts().head(n)
    if addrs.empty:
        return pd.DataFrame()
    return pd.DataFrame({'Address': addrs.index, 'Count': addrs.values})


def specific_number_analysis(df, target_number):
    """Analyze interactions with a specific phone number (MOC, MTC, duration, SMS)."""
    if 'party_b_clean' not in df.columns or not target_number:
        return None
    target = re.sub(r'\D', '', str(target_number).strip())
    if len(target) < 10:
        return None
    # Match flexibly: 8801712345678 == 01712345678 == 1712345678
    candidates = {target}
    if target.startswith('880'):  candidates.add('0' + target[3:])
    elif target.startswith('0'):  candidates.add('880' + target[1:])
    if target.startswith('880'):  candidates.add(target[3:])
    candidates.add(target.lstrip('0'))
    sub = df[df['party_b_clean'].astype(str).isin(candidates)]
    if len(sub) == 0:
        return None
    moc = int(sub.get('is_call_out', pd.Series([False]*len(sub))).sum())
    mtc = int(sub.get('is_call_in',  pd.Series([False]*len(sub))).sum())
    sms_out = int(sub.get('is_sms_out', pd.Series([False]*len(sub))).sum())
    sms_in  = int(sub.get('is_sms_in',  pd.Series([False]*len(sub))).sum())
    call_mask = sub.get('is_call_out', False) | sub.get('is_call_in', False)
    total_dur = int(sub.loc[call_mask, 'duration'].sum()) if 'duration' in sub.columns else 0
    return {
        'number': target_number,
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
    calls = df[df['is_call_out']|df['is_call_in']].copy()
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
    calls = df[df['is_call_out']|df['is_call_in']].copy()
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
    calls = df[df['is_call_out']|df['is_call_in']].copy()
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
    sub = df[df['is_call_out']] if direction=='out' else df[df['is_call_in']]
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
    c = data['address'].dropna().value_counts().head(n)
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


def _target_number_html(df, target_number):
    if not target_number:
        return ''
    res = specific_number_analysis(df, target_number)
    if not res:
        return f'''<h2>13. Specific Number Analysis</h2>
        <div class="info-box">
        <strong>Target Number:</strong> {target_number}<br>
        <span class="warning">⚠️ এই নম্বরের সাথে কোনো communication পাওয়া যায়নি।</span>
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
    <p>এই section-এ <strong>{target_number}</strong> নম্বরের সাথে subscriber-এর সকল communication-এর সারসংক্ষেপ।</p>
    {df_to_html(table_df)}'''

def build_html(df, phone, operator, date_range, total_raw, anomaly_count, target_number=None):
    imei = sorted(df['imei'].dropna().unique().tolist()) if 'imei' in df.columns else []
    imsi = sorted(df['imsi'].dropna().unique().tolist()) if 'imsi' in df.columns else []

    home_mask    = df['start'].dt.hour.astype(int).isin(list(range(0,6))+list(range(22,24))) if 'start' in df.columns else None
    work_mask    = (df['start'].dt.hour.astype(int)>=8)&(df['start'].dt.hour.astype(int)<18)       if 'start' in df.columns else None
    weekend_mask = df['start'].dt.dayofweek.astype(int).isin([4,5])                               if 'start' in df.columns else None

    return f"""<html><head><title>CDR Analysis Report</title>
    <style>{CSS}</style></head><body>
    <h1>📞 CDR Analysis Report</h1>
    <h2>1. Executive Summary</h2>
    <div class="info-box">
    <strong>Phone Number:</strong> {phone}<br>
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
    <tr><td>Phone Number</td><td>{phone}</td></tr></table>
    <h2>7. Call Analysis</h2>
    <h3>7.1 Call Analysis Summary</h3>{df_to_html(call_summary(df))}
    <h2>8. Call Count Analysis</h2>
    <h3>8.1 Daily Call Count</h3>{df_to_html(daily_call_count(df))}
    <h3>8.2 Hourly Call Count Graph</h3>{fig_to_html_img(plot_hourly(df))}
    <h3>8.3 Weekly Call Count</h3>{df_to_html(weekly_call_count(df))}
    <h3>8.3a Weekly Graph</h3>{fig_to_html_img(plot_weekly(df))}
    <h3>8.4 Monthly Call Count</h3>{df_to_html(monthly_call_count(df))}
    <h3>8.4a Monthly Graph</h3>{fig_to_html_img(plot_monthly(df))}
    <h2>9. Contact Analysis</h2>
    <h3>9.1 Contact Summary</h3>{df_to_html(contact_summary(df))}
    <h3>9.2 Top 10 Frequent Outgoing</h3>{df_to_html(top_contacts(df,'out',10))}
    <h3>9.3 Top 10 Frequent Incoming</h3>{df_to_html(top_contacts(df,'in',10))}
    <h3>9.4 Frequent Call Graph</h3>{fig_to_html_img(plot_contacts(df,'out',10,'Top 10 Outgoing Contacts'))}
    <h3>9.5 Top 10 Lengthy Outgoing</h3>{df_to_html(top_lengthy(df,'out',10))}
    <h3>9.6 Top 10 Lengthy Incoming</h3>{df_to_html(top_lengthy(df,'in',10))}
    <h3>9.7 Lengthy Call Graph</h3>{fig_to_html_img(plot_contacts(df,'out',10,'Top 10 Lengthy Contacts'))}
    <h2>10. Location Analysis</h2>
    <h3>10.1 Location Summary</h3>{df_to_html(location_summary(df))}
    <h3>10.2 Top 10 Frequent Locations</h3>{df_to_html(top_locations(df,None,10))}
    <h3>10.3 Frequent Locations Graph</h3>{fig_to_html_img(plot_locations(df,None,'Frequent Locations'))}
    <h3>10.4 Possible Home Locations</h3>{df_to_html(top_locations(df,home_mask,10))}
    <h3>10.5 Home Locations Graph</h3>{fig_to_html_img(plot_locations(df,home_mask,'Home Locations'))}
    <h3>10.5 Possible Work Locations</h3>{df_to_html(top_locations(df,work_mask,10))}
    <h3>10.6 Work Locations Graph</h3>{fig_to_html_img(plot_locations(df,work_mask,'Work Locations'))}
    <h3>10.5 Possible Weekend Locations</h3>{df_to_html(top_locations(df,weekend_mask,10))}
    <h3>10.6 Weekend Locations Graph</h3>{fig_to_html_img(plot_locations(df,weekend_mask,'Weekend Locations'))}

    <h2>11. SMS Contact Analysis</h2>
    <h3>11.1 Top 5 Sent SMS Contacts</h3>{df_to_html(top_sms_contacts(df,'out',5))}
    <h3>11.2 Top 5 Received SMS Contacts</h3>{df_to_html(top_sms_contacts(df,'in',5))}

    <h2>12. Last 10 Days Analysis</h2>
    <h3>12.1 Top Contacts in Last 10 Days (MOC + MTC)</h3>{df_to_html(last_n_days_top_contacts(df, 10, 10))}
    <h3>12.2 Top Locations in Last 10 Days</h3>{df_to_html(last_n_days_top_locations(df, 10, 10))}

    {_target_number_html(df, target_number)}

    <h2>Overall Comment</h2><p>N/A</p>
    <h2>Recommendation</h2><p>N/A</p>
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
    info=[('Phone Number',phone),('Operator',operator),('Analysis Period',date_range),
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

    imei=sorted(df['imei'].dropna().unique().tolist()) if 'imei' in df.columns else []
    imsi=sorted(df['imsi'].dropna().unique().tolist()) if 'imsi' in df.columns else []
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

    add_h('7. Call Analysis'); add_h('7.1 Call Analysis Summary',2); add_df_table(call_summary(df))
    add_h('8. Call Count Analysis')
    add_h('8.1 Daily Call Count',2);      add_df_table(daily_call_count(df))
    add_h('8.2 Hourly Graph',2);          add_fig(plot_hourly(df))
    add_h('8.3 Weekly Call Count',2);     add_df_table(weekly_call_count(df))
    add_h('8.3a Weekly Graph',2);         add_fig(plot_weekly(df))
    add_h('8.4 Monthly Call Count',2);    add_df_table(monthly_call_count(df))
    add_h('8.4a Monthly Graph',2);        add_fig(plot_monthly(df))
    add_h('9. Contact Analysis')
    add_h('9.1 Contact Summary',2);       add_df_table(contact_summary(df))
    add_h('9.2 Top 10 Outgoing',2);       add_df_table(top_contacts(df,'out',10))
    add_h('9.3 Top 10 Incoming',2);       add_df_table(top_contacts(df,'in',10))
    add_h('9.4 Frequent Call Graph',2);   add_fig(plot_contacts(df,'out',10,'Top 10 Outgoing'))
    add_h('9.5 Lengthy Outgoing',2);      add_df_table(top_lengthy(df,'out',10))
    add_h('9.6 Lengthy Incoming',2);      add_df_table(top_lengthy(df,'in',10))
    add_h('9.7 Lengthy Call Graph',2);    add_fig(plot_contacts(df,'out',10,'Top 10 Lengthy'))
    add_h('10. Location Analysis')
    add_h('10.1 Location Summary',2);     add_df_table(location_summary(df))
    add_h('10.2 Frequent Locations',2);   add_df_table(top_locations(df,None,10))
    add_h('10.3 Locations Graph',2);      add_fig(plot_locations(df,None,'Frequent Locations'))
    add_h('10.4 Home Locations',2);       add_df_table(top_locations(df,home_mask,10))
    add_h('10.5 Home Graph',2);           add_fig(plot_locations(df,home_mask,'Home Locations'))
    add_h('10.5 Work Locations',2);       add_df_table(top_locations(df,work_mask,10))
    add_h('10.6 Work Graph',2);           add_fig(plot_locations(df,work_mask,'Work Locations'))
    add_h('10.5 Weekend Locations',2);    add_df_table(top_locations(df,weekend_mask,10))
    add_h('10.6 Weekend Graph',2);        add_fig(plot_locations(df,weekend_mask,'Weekend Locations'))

    add_h('11. SMS Contact Analysis')
    add_h('11.1 Top 5 Sent SMS Contacts',2);     add_df_table(top_sms_contacts(df,'out',5))
    add_h('11.2 Top 5 Received SMS Contacts',2); add_df_table(top_sms_contacts(df,'in',5))

    add_h('12. Last 10 Days Analysis')
    add_h('12.1 Top Contacts in Last 10 Days (MOC + MTC)',2)
    add_df_table(last_n_days_top_contacts(df, 10, 10))
    add_h('12.2 Top Locations in Last 10 Days',2)
    add_df_table(last_n_days_top_locations(df, 10, 10))

    if target_number:
        add_h('13. Specific Number Analysis')
        res = specific_number_analysis(df, target_number)
        if res is None:
            doc.add_paragraph(f'Target Number: {target_number}')
            doc.add_paragraph('⚠️ এই নম্বরের সাথে কোনো communication পাওয়া যায়নি।')
        else:
            doc.add_paragraph(f'এই section-এ {target_number} নম্বরের সাথে subscriber-এর সকল communication-এর সারসংক্ষেপ।')
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

    add_h('Overall Comment');             doc.add_paragraph('N/A')
    add_h('Recommendation');              doc.add_paragraph('N/A')

    fp=doc.add_paragraph(f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")} | CDR Analysis Tool v1.0')
    fp.alignment=WD_ALIGN_PARAGRAPH.CENTER
    fp.runs[0].font.size=Pt(8); fp.runs[0].font.color.rgb=RGBColor(0x80,0x80,0x80)

    buf=io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────
def main():
    st.markdown("""
    <div class="main-header">
        <h1>📞 CDR Analysis Tool</h1>
        <p>Upload any CDR Excel file → Get HTML & Word Report instantly</p>
        <p style="font-size:0.85rem; margin-top:0.3rem;">
        Supports: Grameenphone · Robi · Banglalink · Teletalk · Airtel
        </p>
    </div>
    """, unsafe_allow_html=True)

    # ── File Upload + Target Number Input ──
    up_col, num_col = st.columns([2, 1])
    with up_col:
        uploaded = st.file_uploader(
            "📂 CDR Excel ফাইল আপলোড করুন",
            type=['xlsx', 'xls'],
            help="যেকোনো অপারেটরের CDR Excel ফাইল (.xlsx / .xls)"
        )
    with num_col:
        target_number = st.text_input(
            "🎯 Target Number (Optional)",
            placeholder="e.g. 8801712345678",
            help="যদি কোনো নির্দিষ্ট নম্বরের সাথে interaction জানতে চান, এখানে দিন। Optional — খালি রাখলেও চলবে।"
        )
        target_number = target_number.strip() if target_number else None

    if uploaded is None:
        st.info("⬆️ উপরে CDR Excel ফাইল upload করুন — HTML ও Word Report তৈরি হবে।")
        st.markdown("""
        **এই tool যা করে:**
        - ✅ যেকোনো Operator-এর CDR ফাইল process করে
        - ✅ Anomaly (Service SMS, Invalid numbers) সরিয়ে দেয়
        - ✅ Call Summary, Daily/Weekly/Monthly Analysis
        - ✅ Top Contacts, Location Analysis, SMS Analysis
        - ✅ Last 10 Days Activity (Calls + Locations)
        - ✅ Specific Number Analysis (যদি Target Number দেন)
        - ✅ Graphs সহ HTML + Word Report generate করে
        """)
        return

    # ── Process ──
    progress = st.progress(0, text="📥 ফাইল পড়ছি...")

    try:
        file_bytes = uploaded.read()
        progress.progress(15, text="🔍 ডেটা বিশ্লেষণ করছি...")

        df, col_map, total_raw, anomaly_count, sheet = load_and_clean(file_bytes)
        progress.progress(35, text="📊 Report তৈরি করছি...")

        phone     = get_phone(df)
        operator  = get_operator(df)
        date_range= get_date_range(df)
        base_name = os.path.splitext(uploaded.name)[0]

        # ── Stats ──
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown(f'<div class="stat-card"><div class="label">ফোন নম্বর</div><div class="value" style="font-size:1rem;">{phone}</div></div>', unsafe_allow_html=True)
        with col2:
            st.markdown(f'<div class="stat-card"><div class="label">Operator</div><div class="value">{operator}</div></div>', unsafe_allow_html=True)
        with col3:
            st.markdown(f'<div class="stat-card"><div class="label">মোট Records</div><div class="value">{total_raw:,}</div></div>', unsafe_allow_html=True)
        with col4:
            st.markdown(f'<div class="stat-card"><div class="label">বিশ্লেষিত Records</div><div class="value">{len(df):,}</div></div>', unsafe_allow_html=True)

        if anomaly_count > 0:
            st.markdown(f'<div class="warning-box">⚠️ <strong>{anomaly_count:,} টি Anomaly</strong> (Service SMS / Invalid Numbers) সরানো হয়েছে।</div>', unsafe_allow_html=True)

        progress.progress(55, text="📄 HTML Report তৈরি করছি...")
        html_content = build_html(df, phone, operator, date_range, total_raw, anomaly_count, target_number)
        html_bytes   = html_content.encode('utf-8')

        progress.progress(80, text="📝 Word Report তৈরি করছি...")
        docx_bytes = build_docx(df, phone, operator, date_range, total_raw, anomaly_count, target_number)

        progress.progress(100, text="✅ সম্পন্ন!")

        st.markdown('<div class="success-box">✅ <strong>Report তৈরি হয়েছে!</strong> নিচের বাটনে ক্লিক করে Download করুন।</div>', unsafe_allow_html=True)

        # ── Download Buttons ──
        dl1, dl2 = st.columns(2)
        with dl1:
            st.download_button(
                label="⬇️ HTML Report Download করুন",
                data=html_bytes,
                file_name=f"{base_name}_Report.html",
                mime="text/html",
                use_container_width=True
            )
        with dl2:
            st.download_button(
                label="⬇️ Word Report Download করুন",
                data=docx_bytes,
                file_name=f"{base_name}_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

        st.markdown("---")
        st.markdown("### 📋 Quick Preview")

        # ── 1. Call Analysis Summary + Device Info ──
        with st.expander("📊 Call Analysis Summary", expanded=True):
            st.markdown("#### 📱 Device Information")
            imei_list = sorted(df["imei"].dropna().unique().tolist()) if "imei" in df.columns else []
            imsi_list = sorted(df["imsi"].dropna().unique().tolist()) if "imsi" in df.columns else []
            d1, d2, d3 = st.columns(3)
            with d1:
                st.markdown(f'<div class="stat-card"><div class="label">Phone Number</div><div class="value" style="font-size:0.95rem;">{phone}</div></div>', unsafe_allow_html=True)
            with d2:
                imei_val = ", ".join(str(i) for i in imei_list) if imei_list else "N/A"
                st.markdown(f'<div class="stat-card"><div class="label">IMEI</div><div class="value" style="font-size:0.85rem;">{imei_val}</div></div>', unsafe_allow_html=True)
            with d3:
                imsi_val = ", ".join(str(i) for i in imsi_list) if imsi_list else "N/A"
                st.markdown(f'<div class="stat-card"><div class="label">IMSI</div><div class="value" style="font-size:0.85rem;">{imsi_val}</div></div>', unsafe_allow_html=True)
            st.markdown("---")
            st.markdown("#### 📞 Call Analysis Summary")
            cs = call_summary(df)
            c1, c2, c3, c4 = st.columns(4)
            cols_exp = [c1, c2, c3, c4]
            for i, row in cs.iterrows():
                with cols_exp[i]:
                    st.markdown(f'''
                    <div class="stat-card">
                        <div class="label">{row["Metric"]}</div>
                        <div class="value">{row["Value"]:,}</div>
                    </div>''', unsafe_allow_html=True)

        # ── 2. Top 10 Contacts Incoming & Outgoing ──
        with st.expander("📞 Top 10 Contacts — Incoming & Outgoing", expanded=True):
            t1, t2 = st.columns(2)
            with t1:
                st.markdown("#### 📤 Outgoing (MOC)")
                out_df = top_contacts(df, 'out', 10)
                if not out_df.empty:
                    st.dataframe(out_df, use_container_width=True, hide_index=True)
                    fig = plot_contacts(df, 'out', 10, 'Top 10 Outgoing Contacts')
                    if fig: st.pyplot(fig)
                else:
                    st.info("ডেটা নেই")
            with t2:
                st.markdown("#### 📥 Incoming (MTC)")
                in_df = top_contacts(df, 'in', 10)
                if not in_df.empty:
                    st.dataframe(in_df, use_container_width=True, hide_index=True)
                    fig = plot_contacts(df, 'in', 10, 'Top 10 Incoming Contacts')
                    if fig: st.pyplot(fig)
                else:
                    st.info("ডেটা নেই")

        # ── 3. Top 10 Contacts by Call Duration ──
        with st.expander("⏱️ Top 10 Contacts — Call Duration", expanded=True):
            d1, d2 = st.columns(2)
            with d1:
                st.markdown("#### 📤 Outgoing — সবচেয়ে বেশি কথা বলেছে")
                lo_df = top_lengthy(df, 'out', 10)
                if not lo_df.empty:
                    lo_df['Total_Duration_Min'] = (lo_df['Total_Duration'] / 60).round(1)
                    st.dataframe(
                        lo_df[['Party B','Total_Calls','Total_Duration_Min','Pct_CallTime']].rename(columns={
                            'Total_Calls':'Calls',
                            'Total_Duration_Min':'Duration (min)',
                            'Pct_CallTime':'% of Time'
                        }),
                        use_container_width=True, hide_index=True
                    )
                else:
                    st.info("ডেটা নেই")
            with d2:
                st.markdown("#### 📥 Incoming — সবচেয়ে বেশি কথা বলেছে")
                li_df = top_lengthy(df, 'in', 10)
                if not li_df.empty:
                    li_df['Total_Duration_Min'] = (li_df['Total_Duration'] / 60).round(1)
                    st.dataframe(
                        li_df[['Party B','Total_Calls','Total_Duration_Min','Pct_CallTime']].rename(columns={
                            'Total_Calls':'Calls',
                            'Total_Duration_Min':'Duration (min)',
                            'Pct_CallTime':'% of Time'
                        }),
                        use_container_width=True, hide_index=True
                    )
                else:
                    st.info("ডেটা নেই")

        # ── 4. Top Stay Locations ──
        with st.expander("📍 Top Stay Locations", expanded=True):
            if 'address' in df.columns:
                home_mask    = df['start'].dt.hour.astype(int).isin(list(range(0,6))+list(range(22,24))) if 'start' in df.columns else None
                work_mask    = (df['start'].dt.hour.astype(int)>=8)&(df['start'].dt.hour.astype(int)<18) if 'start' in df.columns else None
                weekend_mask = df['start'].dt.dayofweek.astype(int).isin([4,5])                         if 'start' in df.columns else None

                l1, l2, l3 = st.columns(3)
                with l1:
                    st.markdown("#### 🏠 Probable Home Location")
                    st.caption("রাত ১০টা – সকাল ৬টা")
                    hl = top_locations(df, home_mask, 3)
                    if not hl.empty:
                        st.dataframe(hl, use_container_width=True, hide_index=True)
                    else:
                        st.info("Location ডেটা নেই")

                with l2:
                    st.markdown("#### 🏢 Probable Work Location")
                    st.caption("সকাল ৮টা – সন্ধ্যা ৬টা")
                    wl = top_locations(df, work_mask, 3)
                    if not wl.empty:
                        st.dataframe(wl, use_container_width=True, hide_index=True)
                    else:
                        st.info("Location ডেটা নেই")

                with l3:
                    st.markdown("#### 🕌 Probable Weekend Location")
                    st.caption("শুক্র ও শনিবার")
                    el = top_locations(df, weekend_mask, 3)
                    if not el.empty:
                        st.dataframe(el, use_container_width=True, hide_index=True)
                    else:
                        st.info("Location ডেটা নেই")
            else:
                st.info("এই CDR ফাইলে Location ডেটা নেই।")

        # ── 5. Top 5 SMS Contacts ──
        with st.expander("💬 Top 5 SMS Contacts", expanded=True):
            if "is_sms_out" in df.columns and "party_b_clean" in df.columns:
                s1, s2 = st.columns(2)
                with s1:
                    st.markdown("#### 📤 Sent SMS — সবচেয়ে বেশি SMS পাঠিয়েছে")
                    sms_out_df = top_sms_contacts(df, 'out', 5)
                    if not sms_out_df.empty:
                        st.dataframe(sms_out_df, use_container_width=True, hide_index=True)
                    else:
                        st.info("Sent SMS ডেটা নেই")
                with s2:
                    st.markdown("#### 📥 Received SMS — সবচেয়ে বেশি SMS এসেছে")
                    sms_in_df = top_sms_contacts(df, 'in', 5)
                    if not sms_in_df.empty:
                        st.dataframe(sms_in_df, use_container_width=True, hide_index=True)
                    else:
                        st.info("Received SMS ডেটা নেই")
            else:
                st.info("এই CDR ফাইলে SMS ডেটা নেই।")

        # ── 6. Last 10 Days Analysis ──
        with st.expander("📅 Last 10 Days Analysis", expanded=True):
            ld1, ld2 = st.columns(2)
            with ld1:
                st.markdown("#### 📞 Top Contacts — Last 10 Days (MOC + MTC)")
                last_calls = last_n_days_top_contacts(df, 10, 10)
                if not last_calls.empty:
                    st.dataframe(last_calls, use_container_width=True, hide_index=True)
                else:
                    st.info("শেষ ১০ দিনের ডেটা নেই")
            with ld2:
                st.markdown("#### 📍 Top Locations — Last 10 Days")
                last_loc = last_n_days_top_locations(df, 10, 10)
                if not last_loc.empty:
                    st.dataframe(last_loc, use_container_width=True, hide_index=True)
                else:
                    st.info("শেষ ১০ দিনের Location ডেটা নেই")

        # ── 7. Specific Number Analysis (if target_number provided) ──
        if target_number:
            with st.expander(f"🎯 Specific Number Analysis — {target_number}", expanded=True):
                res = specific_number_analysis(df, target_number)
                if res is None:
                    st.warning(f"⚠️ **{target_number}** নম্বরের সাথে কোনো communication পাওয়া যায়নি।")
                else:
                    sn1, sn2, sn3, sn4 = st.columns(4)
                    with sn1:
                        st.markdown(f'<div class="stat-card"><div class="label">MOC</div><div class="value">{res["moc"]}</div></div>', unsafe_allow_html=True)
                    with sn2:
                        st.markdown(f'<div class="stat-card"><div class="label">MTC</div><div class="value">{res["mtc"]}</div></div>', unsafe_allow_html=True)
                    with sn3:
                        st.markdown(f'<div class="stat-card"><div class="label">Total Duration</div><div class="value" style="font-size:1.1rem;">{res["total_duration_min"]} min</div></div>', unsafe_allow_html=True)
                    with sn4:
                        st.markdown(f'<div class="stat-card"><div class="label">Total SMS</div><div class="value">{res["total_sms"]}</div></div>', unsafe_allow_html=True)
                    st.markdown("")
                    detail_df = pd.DataFrame({
                        'Metric': ['Sent SMS', 'Received SMS', 'First Contact', 'Last Contact'],
                        'Value':  [res['sms_sent'], res['sms_received'], res['first_contact'], res['last_contact']]
                    })
                    st.dataframe(detail_df, use_container_width=True, hide_index=True)

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        st.code(str(e))


if __name__ == '__main__':
    main()
