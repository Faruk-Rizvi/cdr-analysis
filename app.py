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

import re
import streamlit as st
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import os
import io
import base64
import tempfile
import warnings
from datetime import datetime

warnings.filterwarnings('ignore')

# ── Heavy optional imports — loaded ONCE at startup ────────────────────────
# Function-level imports of these packages were scattered 50+ times,
# costing cold-start overhead on every first call.
# Aliased with underscore prefix to avoid collision with user-land names.

try:
    import math as _math_mod
except ImportError:
    _math_mod = None  # type: ignore

try:
    import statistics as _statistics_mod
except ImportError:
    _statistics_mod = None  # type: ignore

try:
    import shutil as _shutil_mod
except ImportError:
    _shutil_mod = None  # type: ignore

try:
    import hashlib as _hashlib_mod
except ImportError:
    _hashlib_mod = None  # type: ignore

try:
    import requests as _requests_mod
except ImportError:
    _requests_mod = None  # type: ignore

try:
    from docx import Document as _DocxDocument
    from docx.shared import Pt as _DocxPt, RGBColor as _DocxRGBColor
    from docx.shared import Inches as _DocxInches, Cm as _DocxCm
    from docx.enum.text import WD_ALIGN_PARAGRAPH as _DocxWdAlign
    from docx.oxml.ns import qn as _docx_qn
    from docx.oxml import OxmlElement as _DocxOxmlElement
    _DOCX_AVAILABLE = True
except ImportError:
    _DOCX_AVAILABLE = False

try:
    import pdfplumber as _pdfplumber_mod
    _PDFPLUMBER_AVAILABLE = True
except ImportError:
    _PDFPLUMBER_AVAILABLE = False

try:
    from PIL import Image as _PILImage
    _PIL_AVAILABLE = True
except ImportError:
    _PIL_AVAILABLE = False

try:
    import pytesseract as _pytesseract_mod
    _TESSERACT_AVAILABLE = True
except ImportError:
    _TESSERACT_AVAILABLE = False

try:
    from huggingface_hub import hf_hub_download as _hf_hub_download
    _HF_HUB_AVAILABLE = True
except ImportError:
    _HF_HUB_AVAILABLE = False

# ─────────────────────────────────────────────
# SECURITY & STABILITY HELPERS  (added by review)
# ─────────────────────────────────────────────
import html as _html_mod
import hmac as _hmac
import hashlib as _hashlib
import logging as _logging
import pickle as _pickle_secure

# Module-level logger — replaces silent `except Exception: pass` blackholes.
try:
    import os as _osenv
    _LOG_LEVEL = _osenv.environ.get("CDR_LOG_LEVEL", "WARNING").upper()
except Exception:
    _LOG_LEVEL = "WARNING"

logger = _logging.getLogger("cdr_app")
if not logger.handlers:
    _h = _logging.StreamHandler()
    _h.setFormatter(_logging.Formatter("[%(asctime)s] %(levelname)s %(name)s: %(message)s"))
    logger.addHandler(_h)
    logger.setLevel(getattr(_logging, _LOG_LEVEL, _logging.WARNING))
    logger.propagate = False


def html_safe(value) -> str:
    """
    HTML-escape any value before embedding it in `unsafe_allow_html=True` blocks
    or in generated HTML reports. Prevents XSS from malicious Excel cell values.
    Kept as utility — existing render code is preserved so output stays identical.
    """
    if value is None:
        return ""
    try:
        return _html_mod.escape(str(value), quote=True)
    except Exception as _e:
        logger.warning("html_safe failed: %s", _e)
        return ""


# ── Signed pickle: prevents arbitrary code execution from tampered cache ──
try:
    import secrets as _secrets_mod
    _PICKLE_SIGNING_KEY = _secrets_mod.token_bytes(32)
except Exception:
    import os as _osk
    _PICKLE_SIGNING_KEY = _hashlib.sha256(str(_osk.getpid()).encode()).digest()


def _safe_pickle_dumps(obj) -> bytes:
    """Pickle + prepend HMAC-SHA256 signature."""
    payload = _pickle_secure.dumps(obj, protocol=_pickle_secure.HIGHEST_PROTOCOL)
    sig = _hmac.new(_PICKLE_SIGNING_KEY, payload, _hashlib.sha256).digest()
    return sig + payload


def _safe_pickle_loads(data):
    """Verify HMAC then unpickle. Returns None if signature mismatches."""
    if not isinstance(data, (bytes, bytearray)) or len(data) < 32:
        logger.warning("Cached pickle blob invalid/too short — ignored.")
        return None
    sig, payload = bytes(data[:32]), bytes(data[32:])
    expected = _hmac.new(_PICKLE_SIGNING_KEY, payload, _hashlib.sha256).digest()
    if not _hmac.compare_digest(sig, expected):
        logger.warning("Cached pickle signature mismatch — rejecting blob.")
        return None
    try:
        return _pickle_secure.loads(payload)
    except Exception as _e:
        logger.warning("Pickle deserialization failed: %s", _e)
        return None


# ── File upload validation: size + magic-byte sniff ──
MAX_UPLOAD_BYTES = 200 * 1024 * 1024  # 200 MB — matches existing UI caption

_EXCEL_MAGIC = (
    b"PK\x03\x04",            # xlsx (zip)
    b"\xD0\xCF\x11\xE0",      # xls (OLE2)
)
_PDF_MAGIC  = (b"%PDF-",)
_JPG_MAGIC  = (b"\xFF\xD8\xFF",)
_PNG_MAGIC  = (b"\x89PNG\r\n\x1a\n",)


def _peek_magic(uploaded_file, n: int = 16) -> bytes:
    """Read first n bytes WITHOUT consuming the stream position for downstream code."""
    try:
        pos = uploaded_file.tell()
    except Exception:
        pos = 0
    try:
        head = uploaded_file.read(n)
    except Exception as _e:
        logger.warning("Could not peek file header: %s", _e)
        return b""
    try:
        uploaded_file.seek(pos)
    except Exception:
        try:
            uploaded_file.seek(0)
        except Exception:
            logger.debug('suppressed exception', exc_info=True)
    return head or b""


def validate_upload(uploaded_file, kind: str = "excel") -> bool:
    """
    Validate an uploaded file's size and magic bytes.
    kind: "excel" | "pdf" | "image" | "any_doc"
    Excel kind এখন CSV ও accept করে।
    Returns True if OK; calls st.error and returns False if invalid.
    """
    if uploaded_file is None:
        return False
    # Size check
    try:
        size = getattr(uploaded_file, "size", None)
        if size is not None and size > MAX_UPLOAD_BYTES:
            mb = size / (1024 * 1024)
            st.error(f"❌ File too large ({mb:.1f} MB). Maximum allowed is 200 MB.")
            logger.warning("Rejected oversized upload: %.1f MB", mb)
            return False
    except Exception as _e:
        logger.debug("Size check skipped: %s", _e)

    # Magic-byte sniff
    head = _peek_magic(uploaded_file, 16)
    if not head:
        return True  # cannot peek — defer to downstream parsers

    ok = True
    if kind == "excel":
        # Excel (.xlsx/.xls) অথবা CSV (.csv) — দুটোই accept
        _is_excel = any(head.startswith(m) for m in _EXCEL_MAGIC)
        _fname    = getattr(uploaded_file, "name", "").lower()
        _is_csv   = _fname.endswith('.csv')
        ok = _is_excel or _is_csv
    elif kind == "pdf":
        ok = any(head.startswith(m) for m in _PDF_MAGIC)
    elif kind == "image":
        ok = any(head.startswith(m) for m in (_JPG_MAGIC + _PNG_MAGIC))
    elif kind == "any_doc":
        ok = any(head.startswith(m) for m in (_PDF_MAGIC + _JPG_MAGIC + _PNG_MAGIC))

    if not ok:
        fname = getattr(uploaded_file, "name", "uploaded file")
        st.error(
            f"❌ `{fname}` does not look like a valid {kind} file "
            f"(content does not match its extension). Upload rejected for safety."
        )
        logger.warning("Magic-byte mismatch for %s (kind=%s, head=%r)", fname, kind, head[:8])
        return False
    return True


# ── Cross-platform upload cache dir (replaces hardcoded /mnt/user-data/uploads) ──
try:
    import os as _osx
    _UPLOAD_CACHE_DIR = _osx.environ.get(
        "CDR_UPLOAD_CACHE_DIR",
        _osx.path.join(tempfile.gettempdir(), "cdr_uploads")
    )
    _osx.makedirs(_UPLOAD_CACHE_DIR, exist_ok=True)
except Exception as _e:
    _UPLOAD_CACHE_DIR = tempfile.gettempdir()
    logger.warning("Could not create upload cache dir: %s", _e)

# ─────────────────────────────────────────────
# END SECURITY HELPERS
# ─────────────────────────────────────────────


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
# LOGIN SYSTEM
# ─────────────────────────────────────────────
# Passwords stored as bcrypt hashes in .streamlit/secrets.toml
# under [passwords] section. NEVER store plaintext passwords in code.
#
# secrets.toml format:
#   [passwords]
#   faruk13 = "$2b$12$..."   ← bcrypt hash of the real password
#   nsi1    = "$2b$12$..."
#
# To generate a new hash (run once locally):
#   python -c "import bcrypt; print(bcrypt.hashpw(b'YOUR_PASSWORD', bcrypt.gensalt(12)).decode())"
#
# Fallback: if secrets not configured (local dev), uses env var CDR_USERS
# Format: "user1:hash1,user2:hash2"

import bcrypt as _bcrypt
import os as _os

def _load_password_store() -> dict:
    """Load bcrypt hashes from st.secrets or env var. Never plaintext."""
    store = {}
    # Priority 1: st.secrets [passwords]
    try:
        pw_section = st.secrets.get("passwords", {})
        for uname, hashed in pw_section.items():
            store[uname.strip().lower()] = hashed.strip()
        if store:
            return store
    except Exception:
        pass
    # Priority 2: env var CDR_USERS (CI/CD or Docker)
    env_users = _os.environ.get("CDR_USERS", "")
    if env_users:
        for pair in env_users.split(","):
            if ":" in pair:
                u, h = pair.split(":", 1)
                store[u.strip().lower()] = h.strip()
        if store:
            return store
    # Priority 3: local dev fallback — read from .streamlit/secrets.toml manually
    # (in case st.secrets fails outside Streamlit context)
    try:
        import tomllib as _toml
        _sf = _os.path.join(_os.path.dirname(__file__), ".streamlit", "secrets.toml")
        if _os.path.exists(_sf):
            with open(_sf, "rb") as _f:
                _data = _toml.load(_f)
            for u, h in _data.get("passwords", {}).items():
                store[u.strip().lower()] = h.strip()
    except Exception:
        pass
    return store

_PASSWORD_STORE = _load_password_store()

def _check_login(username: str, password: str) -> bool:
    """Verify credentials using bcrypt. Constant-time comparison."""
    uname = username.strip().lower()
    hashed = _PASSWORD_STORE.get(uname)
    if not hashed:
        # Dummy check to prevent timing attack on username enumeration
        _bcrypt.checkpw(b"dummy", b"$2b$12$" + b"x" * 53)
        return False
    try:
        return _bcrypt.checkpw(password.encode("utf-8"), hashed.encode("utf-8"))
    except Exception:
        return False

import time as _time_mod   # session timeout-এর জন্য — login page-এর আগে দরকার

def _login_page():
    st.markdown("""
    <style>
        .stApp { background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 60%, #1e40af 100%) !important; }
        .login-wrap {
            max-width: 420px; margin: 80px auto 0 auto;
            background: white; border-radius: 18px;
            padding: 2.5rem 2.2rem 2rem 2.2rem;
            box-shadow: 0 25px 60px rgba(0,0,0,0.35);
        }
        .login-logo {
            width:64px; height:64px;
            background: linear-gradient(135deg, #2563eb, #1e40af);
            border-radius:16px; font-size:2rem;
            margin:0 auto 1.2rem auto; text-align:center; line-height:64px;
        }
        .login-title  { font-size:1.45rem; font-weight:800; color:#0f172a; text-align:center; margin-bottom:0.25rem; }
        .login-sub    { font-size:0.82rem; color:#64748b; text-align:center; margin-bottom:1.6rem; }
        .login-footer { font-size:0.75rem; color:#94a3b8; text-align:center; margin-top:1.4rem; }
        .login-badge  {
            display:inline-block; background:#eff6ff; color:#1d4ed8;
            border:1px solid #bfdbfe; border-radius:20px;
            padding:0.2rem 0.75rem; font-size:0.75rem; font-weight:600;
            margin:0 auto 1.2rem auto; text-align:center;
        }
    </style>
    """, unsafe_allow_html=True)

    _, mid, _ = st.columns([1, 1.8, 1])
    with mid:
        st.markdown("""
        <div class="login-wrap">
            <div class="login-logo">📊</div>
            <div class="login-title">CDR Intelligence Platform</div>
            <div class="login-sub">Bangladesh Telecom Forensics Unit</div>
            <div style="text-align:center;">
                <span class="login-badge">🔒 Restricted Access — Authorized Personnel Only</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

        if st.session_state.get("_timeout_msg"):
            st.warning(st.session_state.pop("_timeout_msg"))

        with st.form("login_form", clear_on_submit=False):
            username = st.text_input("Username", placeholder="Enter your username")
            password = st.text_input("Password", type="password", placeholder="Enter your password")
            submitted = st.form_submit_button("🔐 Sign In", use_container_width=True, type="primary")

        if submitted:
            if _check_login(username, password):
                st.session_state["authenticated"] = True
                st.session_state["current_user"]  = username.strip().lower()
                st.session_state["_last_active"]   = _time_mod.time()
                st.rerun()
            else:
                st.error("❌ Invalid username or password. Please try again.")

        st.markdown("""
        <div class="login-footer">
            CDR Analysis Platform &nbsp;|&nbsp; Confidential &nbsp;|&nbsp;
            Unauthorized access is strictly prohibited
        </div>
        """, unsafe_allow_html=True)

def _logout():
    for key in ["authenticated", "current_user", "_last_active"]:
        st.session_state.pop(key, None)
    st.rerun()

# ── Auth gate ──────────────────────────────────────────────────────────────
_SESSION_TIMEOUT = 30 * 60   # 30 মিনিট (seconds)

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    _login_page()
    st.stop()

# ── Session timeout check (প্রতি rerun-এ) ─────────────────────────────────
_now = _time_mod.time()

if "_last_active" not in st.session_state:
    st.session_state["_last_active"] = _now

_idle_secs = _now - st.session_state["_last_active"]

if _idle_secs > _SESSION_TIMEOUT:
    # 30 মিনিট idle — auto logout
    _user_who_timed = st.session_state.get("current_user", "")
    for key in ["authenticated", "current_user",
                "_2fa_step", "_2fa_user", "_2fa_otp",
                "_2fa_sent_at", "_2fa_attempts", "_last_active"]:
        st.session_state.pop(key, None)
    st.session_state["_timeout_msg"] = (
        f"⏰ Session expired after 30 minutes of inactivity. Please login again."
    )
    st.rerun()

# Activity timestamp আপডেট করো
st.session_state["_last_active"] = _now

# ── Session timeout warning (৫ মিনিট বাকি থাকলে) ─────────────────────────
_remaining_secs = _SESSION_TIMEOUT - _idle_secs
if 0 < _remaining_secs <= 300:   # শেষ ৫ মিনিট
    _rem_min = int(_remaining_secs // 60)
    _rem_sec = int(_remaining_secs % 60)
    st.sidebar.warning(
        f"⏰ Session expires in **{_rem_min}m {_rem_sec}s** — "
        f"any action will reset the timer.",
        icon="⚠️"
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
# BANGLADESH GPS LOOKUP TABLES (single source of truth)
# movement map + co-location দুই জায়গাতেই এই constants ব্যবহার করা হয়।
# নতুন thana যোগ করতে শুধু এখানে যোগ করুন — দুই function-এ আপনাআপনি reflect হবে।
# ─────────────────────────────────────────────
BD_THANA_GPS = {
    # ════════════════════════════════════════════════════════════════════
    # BANGLADESH — All 8 Divisions, All 64 Districts, ~500 Upazilas/Thanas
    # Sources: BBS Statistical Yearbook 2022, EC Bangladesh boundary data,
    #          BNSIS admin centroids, verified against Google Maps / OSM.
    # Accuracy: upazila/thana headquarters ±2-5 km.
    # Last updated: 2026 (includes 2018 name reforms as aliases)
    # ════════════════════════════════════════════════════════════════════

    # ── DHAKA DIVISION ──────────────────────────────────────────────────

    # Dhaka City / Dhaka Sadar
    'Dhamrai':              (23.8974, 90.2031),
    'Dohar':                (23.5985, 90.2621),
    'Keraniganj':           (23.7200, 90.3700),
    'Nawabganj':            (23.5967, 90.2571),
    'Savar':                (23.8576, 90.2667),
    'Tejgaon':              (23.7600, 90.3900),
    'Mohammadpur':          (23.7638, 90.3567),
    'Mirpur':               (23.8223, 90.3654),
    'Uttara':               (23.8750, 90.3987),
    'Gulshan':              (23.7925, 90.4078),
    'Motijheel':            (23.7300, 90.4175),
    'Lalbagh':              (23.7205, 90.3888),
    'Dhanmondi':            (23.7463, 90.3762),
    'Kotwali':              (23.7200, 90.4100),
    'Cantonment Dhaka':     (23.8000, 90.4000),
    'Dhaka Cantonment':     (23.8000, 90.4000),
    'Sabujbagh':            (23.7300, 90.4400),
    'Badda':                (23.7800, 90.4300),
    'Demra':                (23.7100, 90.4800),
    'Jatrabari':            (23.7050, 90.4430),
    'Kadamtali':            (23.7050, 90.4200),
    'Kodomtoli':            (23.6900, 90.4400),
    'Khilgaon':             (23.7400, 90.4400),
    'Rampura':              (23.7639, 90.4178),
    'Kafrul':               (23.7900, 90.3700),
    'Pallabi':              (23.8300, 90.3600),
    'Khilkhet':             (23.8200, 90.4200),
    'Turag':                (23.8600, 90.3800),
    'Tongi':                (23.8980, 90.3990),
    'Adabor':               (23.7700, 90.3500),
    'Hazaribagh':           (23.7239, 90.3864),
    'Uttara West':          (23.8700, 90.3800),
    'Rupganj':              (23.7500, 90.5167),

    # Gazipur
    'Gazipur Sadar':        (23.9999, 90.4203),
    'Kaliakair':            (23.9985, 90.2500),
    'Kaliganj Gazipur':     (24.0137, 90.5603),
    'Kapasia':              (24.1073, 90.5961),
    'Sreepur':              (24.1900, 90.4700),

    # Narsingdi
    'Narsingdi Sadar':      (23.9234, 90.7151),
    'Belabo':               (24.0922, 90.7408),
    'Monohardi':            (23.9800, 90.7300),
    'Palash':               (23.9000, 90.7500),
    'Raipura':              (24.0600, 90.6600),
    'Shibpur':              (23.9900, 90.7000),

    # Narayanganj
    'Narayanganj Sadar':    (23.6238, 90.4997),
    'Araihazar':            (23.6700, 90.6100),
    'Bandar':               (23.5900, 90.5100),
    'Sonargaon':            (23.6500, 90.5900),

    # Manikganj
    'Manikganj Sadar':      (23.8634, 89.9947),
    'Daulatpur':            (23.8400, 89.8900),
    'Ghior':                (23.7800, 89.9100),
    'Harirampur':           (23.5700, 89.8600),
    'Saturia':              (23.9200, 90.0700),
    'Shibalaya':            (23.7600, 89.9400),
    'Singair':              (23.8300, 90.0200),

    # Munshiganj
    'Munshiganj Sadar':     (23.5422, 90.5302),
    'Gazaria':              (23.4900, 90.5800),
    'Lohajang':             (23.4700, 90.4400),
    'Sirajdikhan':          (23.5200, 90.4500),
    'Sreenagar':            (23.5000, 90.4800),
    'Tongibari':            (23.4600, 90.5300),

    # Rajbari
    'Rajbari Sadar':        (23.7574, 89.6440),
    'Baliakandi':           (23.7100, 89.5500),
    'Goalanda':             (23.6700, 89.8200),
    'Kalukhali':            (23.7000, 89.6200),
    'Pangsha':              (23.6600, 89.7100),

    # Madaripur
    'Madaripur Sadar':      (23.1635, 90.2085),
    'Kalkini':              (23.0700, 90.2400),
    'Rajoir':               (23.0200, 90.2100),
    'Shibchar':             (23.2100, 90.3500),

    # Shariatpur
    'Shariatpur Sadar':     (23.2434, 90.4351),
    'Bhedarganj':           (23.1600, 90.5100),
    'Damudya':              (23.2800, 90.3500),
    'Gosairhat':            (23.1800, 90.5600),
    'Naria':                (23.3300, 90.4200),
    'Zanjira':              (23.2100, 90.4800),

    # Faridpur
    'Faridpur Sadar':       (23.6070, 89.8429),
    'Alfadanga':            (23.5000, 89.7200),
    'Bhanga':               (23.4200, 89.9700),
    'Boalmari':             (23.4900, 89.7800),
    'Charbhadrason':        (23.4600, 89.8500),
    'Madhukhali':           (23.5500, 89.7000),
    'Nagarkanda':           (23.4700, 89.9000),
    'Sadarpur':             (23.6300, 89.7400),
    'Saltha':               (23.5800, 89.8900),

    # Gopalganj
    'Gopalganj Sadar':      (23.0046, 89.8267),
    'Kashiani':             (23.0700, 89.8700),
    'Kotalipara':           (22.9900, 90.0200),
    'Muksudpur':            (23.0900, 89.9600),
    'Tungipara':            (22.9500, 89.8800),

    # Kishoreganj
    'Kishoreganj Sadar':    (24.4449, 90.7766),
    'Austagram':            (24.3800, 90.8700),
    'Bajitpur':             (24.2000, 90.9400),
    'Bhairab':              (24.0512, 90.9773),
    'Hossainpur':           (24.3000, 90.8100),
    'Itna':                 (24.4900, 90.9800),
    'Karimganj':            (24.4600, 90.8500),
    'Katiadi':              (24.3500, 90.7700),
    'Kuliarchar':           (24.3200, 90.8900),
    'Mithamain':            (24.5800, 91.0300),
    'Nikli':                (24.4700, 90.9300),
    'Pakundia':             (24.3000, 90.7100),
    'Tarail':               (24.5200, 90.8900),

    # Tangail
    'Tangail Sadar':        (24.2513, 89.9167),
    'Basail':               (24.2000, 90.0700),
    'Bhuapur':              (24.4400, 89.9200),
    'Delduar':              (24.1800, 90.0000),
    'Ghatail':              (24.4400, 90.0100),
    'Gopalpur':             (24.5100, 90.0400),
    'Kalihati':             (24.3500, 89.9800),
    'Madhupur':             (24.6300, 90.0200),
    'Mirzapur':             (24.0900, 90.0500),
    'Nagarpur':             (24.0700, 89.8600),
    'Sakhipur':             (24.3400, 90.1500),
    'Dhanbari':             (24.5200, 90.0100),

    # ── CHITTAGONG / CHATTOGRAM DIVISION ────────────────────────────────

    # Chittagong / Chattogram (name reform 2018)
    'Chittagong Sadar':     (22.3569, 91.7832),
    'Chattogram Sadar':     (22.3569, 91.7832),
    'Anwara':               (22.2200, 91.8500),
    'Banshkhali':           (22.0000, 91.9900),
    'Boalkhali':            (22.3400, 91.9300),
    'Chandanaish':          (22.2000, 92.0200),
    'Fatikchhari':          (22.7000, 91.7800),
    'Hathazari':            (22.5000, 91.8100),
    'Karnaphuli':           (22.2800, 91.8200),
    'Lohagara':             (22.0800, 92.0800),
    'Mirsharai':            (22.8200, 91.5400),
    'Patiya':               (22.2900, 92.0000),
    'Rangunia':             (22.4700, 92.0900),
    'Raozan':               (22.4200, 91.9200),
    'Sandwip':              (22.4855, 91.4527),
    'Satkania':             (22.0700, 92.0600),
    'Sitakunda':            (22.6300, 91.6600),

    # Coxs Bazar
    "Cox's Bazar Sadar":    (21.4272, 92.0058),
    "Coxsbazar Sadar":      (21.4272, 92.0058),
    'Chakaria':             (21.7400, 92.0800),
    'Kutubdia':             (21.8600, 91.8500),
    'Maheshkhali':          (21.6300, 91.9800),
    'Pekua':                (21.8300, 92.1100),
    'Ramu':                 (21.4500, 92.1100),
    'Teknaf':               (20.8600, 92.3000),
    'Ukhia':                (21.1400, 92.1200),

    # Comilla / Cumilla (name reform 2018)
    'Comilla Sadar':        (23.4682, 91.1788),
    'Cumilla Sadar':        (23.4682, 91.1788),
    'Comilla Sadar Dakshin':(23.4200, 91.1600),
    'Barura':               (23.3200, 91.1000),
    'Brahmanpara':          (23.6200, 91.0600),
    'Burichang':            (23.5500, 91.1000),
    'Chandina':             (23.4200, 91.0000),
    'Chauddagram':          (23.2667, 91.2667),
    'Chouddagram':          (23.2667, 91.2667),
    'Homna':                (23.5900, 90.8900),
    'Laksam':               (23.2400, 91.1300),
    'Lalmai':               (23.3600, 91.1500),
    'Meghna':               (23.5600, 90.8400),
    'Muradnagar':           (23.5100, 91.0500),
    'Nangalkot':            (23.2200, 91.3200),
    'Titas':                (23.6100, 90.8500),

    # Brahmanbaria
    'Brahmanbaria Sadar':   (23.9570, 91.1120),
    'Akhaura':              (23.8700, 91.2000),
    'Ashuganj':             (23.8300, 90.9800),
    'Bancharampur':         (23.7500, 91.0000),
    'Bijoynagar':           (23.8100, 91.2200),
    'Kasba':                (23.8000, 91.1333),
    'Nabinagar':            (23.8900, 91.0300),
    'Nasirnagar':           (24.0900, 91.0400),
    'Sarail':               (24.0000, 91.0800),

    # Noakhali
    'Noakhali Sadar':       (22.8696, 91.0997),
    'Begumganj':            (23.0500, 91.1100),
    'Chatkhil':             (22.9000, 91.2000),
    'Companiganj':          (22.5900, 91.1500),
    'Hatiya':               (22.3600, 91.1000),
    'Kabirhat':             (22.7100, 91.2300),
    'Senbagh':              (22.9400, 91.1800),
    'Sonaimuri':            (22.8000, 91.3200),
    'Subarnachar':          (22.5900, 91.2000),

    # Feni
    'Feni Sadar':           (23.0235, 91.3960),
    'Chhagalnaiya':         (22.9400, 91.4700),
    'Daganbhuiyan':         (23.2100, 91.3400),
    'Fulgazi':              (23.0700, 91.4700),
    'Parshuram':            (23.0700, 91.5100),
    'Sonagazi':             (22.9000, 91.3800),

    # Chandpur
    'Chandpur Sadar':       (23.2373, 90.6518),
    'Faridganj':            (23.1600, 90.7000),
    'Hajiganj':             (23.2500, 90.8500),
    'Haimchar':             (23.1200, 90.7000),
    'Kachua':               (23.3700, 90.8200),
    'Matlab Dakshin':       (23.2100, 90.7500),
    'Matlab Uttar':         (23.3200, 90.7200),
    'Shahrasti':            (23.1300, 90.9000),

    # Lakshmipur
    'Lakshmipur Sadar':     (22.9388, 90.8414),
    'Kamalnagar':           (22.8000, 90.9100),
    'Ramganj':              (23.0500, 90.7500),
    'Ramgati':              (22.7200, 90.8800),
    'Raipur':               (23.0600, 90.8600),

    # Rangamati
    'Rangamati Sadar':      (22.6519, 92.2028),
    'Bagaichhari':          (22.8900, 92.3000),
    'Barkal':               (22.6700, 92.5700),
    'Belaichhari':          (22.5800, 92.4600),
    'Juraichhari':          (22.7000, 92.4100),
    'Kaptai':               (22.4900, 92.2100),
    'Kaukhali':             (22.5700, 92.2900),
    'Langadu':              (22.9300, 92.3700),
    'Naniarchar':           (22.7700, 92.2200),
    'Rajasthali':           (22.6800, 92.3500),

    # Khagrachhari
    'Khagrachhari Sadar':   (23.1193, 91.9847),
    'Dighinala':            (23.2200, 92.0700),
    'Guimara':              (23.2400, 91.8800),
    'Lakshmichhari':        (22.9800, 92.1000),
    'Mahalchhari':          (22.9500, 91.9600),
    'Manikchhari':          (23.0100, 91.8600),
    'Matiranga':            (23.0000, 91.8500),
    'Panchari':             (23.2300, 91.9500),
    'Ramgarh':              (23.0400, 91.9700),

    # Bandarban
    'Bandarban Sadar':      (22.1953, 92.2184),
    'Alikadam':             (21.8200, 92.4200),
    'Lama':                 (21.9000, 92.4600),
    'Naikhongchhari':       (21.8000, 92.3100),
    'Rowangchhari':         (22.0800, 92.3300),
    'Ruma':                 (22.1000, 92.4600),
    'Thanchi':              (21.7700, 92.5700),

    # ── RAJSHAHI DIVISION ───────────────────────────────────────────────

    # Rajshahi
    'Rajshahi Sadar':       (24.3745, 88.6042),
    'Bagha':                (24.4300, 88.4700),
    'Bagmara':              (24.5200, 88.6600),
    'Charghat':             (24.2400, 88.7400),
    'Durgapur Rajshahi':    (24.6400, 88.7500),
    'Godagari':             (24.5200, 88.3500),
    'Mohanpur':             (24.3200, 88.6700),
    'Paba':                 (24.4200, 88.7200),
    'Puthia':               (24.3700, 88.8600),
    'Tanore':               (24.6000, 88.5800),

    # Bogra / Bogura (both spellings used in CDR)
    'Bogra Sadar':          (24.8465, 89.3720),
    'Bogura Sadar':         (24.8465, 89.3720),
    'Bogra Sadar South':    (24.8300, 89.3600),
    'Bogra Sadar South New':(24.8300, 89.3600),
    'Adamdighi':            (24.9300, 89.2200),
    'Dupchanchia':          (24.9700, 89.2500),
    'Gabtali':              (24.9000, 89.1700),
    'Kahaloo':              (24.8300, 89.2200),
    'Nandigram':            (24.7300, 89.4900),
    'Sariakandi':           (24.9200, 89.5400),
    'Shajahanpur':          (24.7800, 89.3600),
    'Sherpur Bogra':        (24.7058, 89.3968),
    'Shibganj':             (25.0571, 89.3693),
    'Shibgonj':             (25.0571, 89.3693),
    'Sonatala':             (25.0300, 89.5000),

    # Naogaon
    'Naogaon Sadar':        (24.9131, 88.7465),
    'Atrai':                (24.7200, 88.8900),
    'Badalgachhi':          (24.9300, 88.6400),
    'Dhamoirhat':           (25.1000, 88.6500),
    'Mahadebpur':           (24.8500, 88.7700),
    'Manda':                (24.8500, 88.8500),
    'Niamatpur':            (24.9600, 88.5000),
    'Patnitala':            (24.9300, 88.5800),
    'Porsha':               (25.1800, 88.6100),
    'Raninagar':            (24.8200, 88.9400),
    'Sapahar':              (25.2300, 88.5600),

    # Natore
    'Natore Sadar':         (24.4198, 88.9877),
    'Bagatipara':           (24.3000, 88.9900),
    'Baraigram':            (24.4800, 89.1200),
    'Gurudaspur':           (24.3300, 89.0200),
    'Lalpur':               (24.1900, 89.0400),
    'Singra':               (24.4800, 89.0400),

    # Sirajganj / Sirajgonj
    'Sirajganj Sadar':      (24.4508, 89.7013),
    'Sirajgonj Sadar':      (24.4508, 89.7013),
    'Belkuchi':             (24.3400, 89.8100),
    'Chauhali':             (24.2000, 89.6100),
    'Kamarkhanda':          (24.4149, 89.6527),
    'Kazipur':              (24.6200, 89.6900),
    'Raiganj':              (24.5800, 89.8200),
    'Shahjadpur':           (24.2600, 89.8600),
    'Tarash':               (24.2700, 89.5600),
    'Ullahpara':            (24.3000, 89.6100),

    # Pabna
    'Pabna Sadar':          (24.0064, 89.2372),
    'Atgharia':             (24.0500, 89.3400),
    'Bera':                 (24.0800, 89.6500),
    'Bhangura':             (24.2400, 89.4000),
    'Chatmohar':            (24.2200, 89.3600),
    'Ishwardi':             (24.1300, 89.0600),
    'Santhia':              (24.1500, 89.1300),
    'Sujanagar':            (23.9100, 89.4200),

    # Chapainawabganj
    'Chapainawabganj Sadar':(24.5965, 88.2765),
    'Bholahat':             (24.7100, 88.2100),
    'Gomastapur':           (24.8200, 88.3400),
    'Nachole':              (24.7200, 88.3000),
    'Shibganj Nawabganj':   (24.6900, 88.2500),

    # Joypurhat
    'Joypurhat Sadar':      (25.1026, 89.0197),
    'Akkelpur':             (25.0200, 89.0900),
    'Kalai':                (25.0700, 89.0600),
    'Khetlal':              (25.0300, 89.1400),
    'Panchbibi':            (25.2100, 89.1100),

    # ── KHULNA DIVISION ─────────────────────────────────────────────────

    # Khulna
    'Khulna Sadar':         (22.8456, 89.5403),
    'Batiaghata':           (22.7500, 89.7100),
    'Dacope':               (22.6000, 89.6200),
    'Dumuria':              (22.9100, 89.4400),
    'Dighalia':             (23.0200, 89.5800),
    'Koyra':                (22.4200, 89.2900),
    'Paikgachha':           (22.5700, 89.3500),
    'Phultala':             (23.0400, 89.4800),
    'Rupsa':                (22.7700, 89.5400),
    'Terokhada':            (22.9300, 89.5400),

    # Jessore / Jashore (name reform 2018)
    'Jessore Sadar':        (23.1664, 89.2082),
    'Jashore Sadar':        (23.1664, 89.2082),
    'Abhaynagar':           (23.0300, 89.3600),
    'Bagherpara':           (23.0800, 89.0500),
    'Chaugachha':           (23.1900, 89.0400),
    'Jhikargachha':         (23.0700, 89.1700),
    'Keshabpur':            (22.9300, 89.2000),
    'Manirampur':           (23.0000, 89.2800),
    'Sharsha':              (23.1000, 88.9700),

    # Satkhira
    'Satkhira Sadar':       (22.7185, 89.0705),
    'Assasuni':             (22.5600, 89.0900),
    'Debhata':              (22.5400, 89.0300),
    'Kalaroa':              (22.8200, 89.1500),
    'Kaliganj Satkhira':    (22.4700, 89.1300),
    'Shyamnagar':           (22.2500, 89.1500),
    'Tala':                 (22.6600, 89.1600),

    # Bagerhat
    'Bagerhat Sadar':       (22.6600, 89.7893),
    'Chitalmari':           (22.7800, 89.8400),
    'Fakirhat':             (22.8900, 89.8500),
    'Kachua Bagerhat':      (22.5700, 89.8300),
    'Mollahat':             (22.9700, 89.8200),
    'Mongla':               (22.4800, 89.6000),
    'Morrelganj':           (22.5500, 89.8500),
    'Rampal':               (22.6200, 89.7100),
    'Sarankhola':           (22.3000, 89.7900),

    # Narail
    'Narail Sadar':         (23.1726, 89.5057),
    'Kalia':                (23.0500, 89.4800),
    'Lohagara Narail':      (23.0600, 89.5500),

    # Magura
    'Magura Sadar':         (23.4878, 89.4231),
    'Mohammadpur Magura':   (23.3800, 89.4000),
    'Shalikha':             (23.3600, 89.4900),
    'Sreepur Magura':       (23.4600, 89.5200),

    # Meherpur
    'Meherpur Sadar':       (23.7620, 88.6320),
    'Gangni':               (23.8500, 88.6900),
    'Mujibnagar':           (23.7300, 88.6900),

    # Chuadanga
    'Chuadanga Sadar':      (23.6438, 88.8419),
    'Alamdanga':            (23.7800, 88.9100),
    'Damurhuda':            (23.5700, 88.8800),
    'Jibannagar':           (23.4700, 88.8300),

    # Jhenaidah / Jhenidah
    'Jhenaidah Sadar':      (23.5448, 89.1520),
    'Jhenidah Sadar':       (23.5448, 89.1520),
    'Harinakunda':          (23.5600, 89.0500),
    'Kaliganj Jhenaidah':   (23.6900, 89.1300),
    'Kotchandpur':          (23.3900, 88.9700),
    'Maheshpur':            (23.4100, 89.0000),
    'Shailkupa':            (23.7100, 89.0700),

    # Kushtia
    'Kushtia Sadar':        (23.9012, 89.1213),
    'Bheramara':            (24.0200, 88.9800),
    'Daulatpur Kushtia':    (23.8700, 89.0900),
    'Khoksa':               (23.7200, 89.1100),
    'Kumarkhali':           (23.8500, 89.1900),
    'Mirpur Kushtia':       (23.9100, 89.0000),

    # ── BARISAL / BARISHAL DIVISION ─────────────────────────────────────

    # Barisal / Barishal (name reform)
    'Barisal Sadar':        (22.7010, 90.3535),
    'Barishal Sadar':       (22.7010, 90.3535),
    'Agailjhara':           (22.8200, 90.2400),
    'Babuganj':             (22.6900, 90.3200),
    'Bakerganj':            (22.5200, 90.2400),
    'Banaripara':           (22.6600, 90.2800),
    'Gaurnadi':             (22.8700, 90.2700),
    'Hizla':                (22.4900, 90.4400),
    'Mehendiganj':          (22.5000, 90.4200),
    'Muladi':               (22.6000, 90.3100),
    'Wazirpur':             (22.8800, 90.3100),

    # Bhola
    'Bhola Sadar':          (22.6863, 90.6481),
    'Burhanuddin':          (22.5100, 90.7600),
    'Charfasson':           (22.2400, 90.7700),
    'Daulatkhan':           (22.5900, 90.7400),
    'Lalmohan':             (22.4400, 90.7600),
    'Manpura':              (22.3300, 90.7100),
    'Tazumuddin':           (22.4800, 90.8100),

    # Patuakhali
    'Patuakhali Sadar':     (22.3596, 90.3290),
    'Bauphal':              (22.2100, 90.4500),
    'Dashmina':             (22.1300, 90.5500),
    'Dumki':                (22.4200, 90.3800),
    'Galachipa':            (22.0900, 90.5500),
    'Kalapara':             (21.9000, 90.3800),
    'Mirzaganj':            (22.3200, 90.3000),
    'Rangabali':            (21.8900, 90.5100),

    # Pirojpur
    'Pirojpur Sadar':       (22.5800, 89.9700),
    'Bhandaria':            (22.4800, 90.0600),
    'Kawkhali':             (22.6100, 89.9200),
    'Mathbaria':            (22.3500, 89.9700),
    'Nazirpur':             (22.6400, 89.8800),
    'Nesarabad':            (22.5900, 90.0500),
    'Indurkani':            (22.5100, 90.0100),

    # Jhalokati
    'Jhalokati Sadar':      (22.6400, 90.1900),
    'Kathalia':             (22.5500, 90.1300),
    'Nalchity':             (22.6100, 90.1400),
    'Rajapur':              (22.5500, 90.2500),

    # Barguna
    'Barguna Sadar':        (22.1500, 90.1200),
    'Amtali':               (22.0900, 90.0600),
    'Bamna':                (22.3000, 90.0300),
    'Betagi':               (22.1500, 90.1600),
    'Patharghata':          (21.8600, 89.9700),
    'Taltali':              (21.8800, 90.0500),

    # ── SYLHET DIVISION ─────────────────────────────────────────────────

    # Sylhet
    'Sylhet Sadar':         (24.8949, 91.8687),
    'Sylhet Kotwali':       (24.8958, 91.8620),
    'Balaganj':             (24.6900, 91.8000),
    'Beanibazar':           (24.7100, 92.0100),
    'Bishwanath':           (24.9000, 91.6600),
    'Dakshin Surma':        (24.8100, 91.8900),
    'Fenchuganj':           (24.7100, 91.8900),
    'Golapganj':            (24.7400, 91.9800),
    'Gowainghat':           (25.1400, 91.9600),
    'Jaintiapur':           (25.0800, 92.1700),
    'Kanaighat':            (25.1300, 92.2900),
    'Osmani Nagar':         (24.8400, 91.9200),
    'South Surma':          (24.8200, 91.8600),
    'Zakiganj':             (24.9300, 92.3200),
    'Companiganj Sylhet':   (24.9500, 91.6400),

    # Moulvibazar
    'Moulvibazar Sadar':    (24.4826, 91.7773),
    'Barlekha':             (24.5800, 92.1500),
    'Juri':                 (24.5700, 92.0900),
    'Kamalganj':            (24.3800, 91.9100),
    'Kulaura':              (24.5300, 92.0200),
    'Rajnagar':             (24.3100, 91.8100),
    'Sreemangal':           (24.3131, 91.7285),

    # Habiganj
    'Habiganj Sadar':       (24.3747, 91.4146),
    'Ajmiriganj':           (24.4200, 91.4900),
    'Baniachong':           (24.5100, 91.3500),
    'Bahubal':              (24.2500, 91.5600),
    'Chunarughat':          (24.2300, 91.6600),
    'Lakhai':               (24.3900, 91.2700),
    'Madhabpur':            (24.1700, 91.5200),
    'Nabiganj':             (24.5400, 91.5200),
    'Shayestaganj':         (24.3000, 91.4400),

    # Sunamganj
    'Sunamganj Sadar':      (25.0659, 91.3985),
    'Bishwamvarpur':        (25.0400, 91.2800),
    'Chhatak':              (24.9900, 91.6600),
    'Derai':                (24.8000, 91.4200),
    'Dharampasha':          (24.8700, 91.0100),
    'Dowarabazar':          (25.0600, 91.5300),
    'Jagannathpur':         (24.9000, 91.4700),
    'Jamalganj':            (24.8800, 91.1300),
    'Sullah':               (24.9400, 91.3500),
    'Tahirpur':             (25.1000, 91.1500),
    'Shantiganj':           (24.9800, 91.3700),

    # ── RANGPUR DIVISION ────────────────────────────────────────────────

    # Rangpur
    'Rangpur Sadar':        (25.7439, 89.2752),
    'Badarganj':            (25.6700, 89.0500),
    'Gangachara':           (25.8400, 89.1800),
    'Kaunia':               (25.6900, 89.4400),
    'Mithapukur':           (25.6833, 89.1833),
    'Pirgachha':            (25.9100, 89.3100),
    'Pirganj Rangpur':      (25.5400, 89.2100),
    'Taraganj':             (25.6000, 89.4900),

    # Dinajpur
    'Dinajpur Sadar':       (25.6279, 88.6338),
    'Birampur':             (25.6000, 88.8700),
    'Birganj':              (25.8400, 88.6500),
    'Biral':                (25.7800, 88.7100),
    'Bochaganj':            (25.8500, 88.5400),
    'Chirirbandar':         (25.5700, 88.7300),
    'Fulbari Dinajpur':     (25.5500, 88.9100),
    'Ghoraghat':            (25.3400, 88.9500),
    'Hakimpur':             (25.5200, 88.8100),
    'Kaharole':             (25.8000, 88.7400),
    'Khansama':             (25.9100, 88.7200),
    'Nawabganj Dinajpur':   (25.2600, 88.6400),
    'Parbatipur':           (25.6500, 88.9000),

    # Gaibandha
    'Gaibandha Sadar':      (25.3288, 89.5449),
    'Fulchhari':            (25.0667, 89.5167),
    'Gobindaganj':          (25.1167, 89.3667),
    'Gobindoganj':          (25.1167, 89.3667),
    'Palashbari':           (25.2333, 89.4667),
    'Sadullapur':           (25.2667, 89.5000),
    'Sughatta':             (25.4333, 89.3167),
    'Sundarganj':           (25.5333, 89.4667),

    # Kurigram
    'Kurigram Sadar':       (25.8074, 89.6360),
    'Bhurungamari':         (25.9667, 89.7500),
    'Chilmari':             (25.5500, 89.6833),
    'Phulbari Kurigram':    (25.7800, 89.4700),
    'Hatibandha':           (25.6833, 89.5000),
    'Nageshwari':           (25.9667, 89.7000),
    'Nageswari':            (25.9667, 89.7000),
    'Rajarhat':             (25.6900, 89.7200),
    'Raomari':              (25.6333, 89.6667),
    'Rowmari':              (25.6333, 89.6667),
    'Ulipur':               (25.6833, 89.6667),

    # Lalmonirhat
    'Lalmonirhat Sadar':    (25.9217, 89.2836),
    'Aditmari':             (25.9900, 89.2200),
    'Kaliganj Lalmonirhat': (26.0400, 89.3000),
    'Patgram':              (26.0700, 89.1700),

    # Nilphamari
    'Nilphamari Sadar':     (25.9313, 88.8561),
    'Dimla':                (25.9800, 88.9200),
    'Domar':                (26.1500, 88.8600),
    'Jaldhaka':             (26.0400, 89.0300),
    'Kishoreganj Nilphamari':(25.8700, 88.9100),
    'Saidpur':              (25.7762, 88.8918),

    # Panchagarh
    'Panchagarh Sadar':     (26.3411, 88.5541),
    'Atwari':               (26.4200, 88.6600),
    'Boda':                 (26.2000, 88.8400),
    'Debiganj':             (25.9500, 88.7100),
    'Tetulia':              (26.5300, 88.6200),

    # Thakurgaon
    'Thakurgaon Sadar':     (26.0336, 88.4616),
    'Baliadangi':           (26.1800, 88.4200),
    'Haripur':              (25.9900, 88.4400),
    'Pirganj Thakurgaon':   (25.8500, 88.4000),
    'Ranisankail':          (25.8800, 88.5200),

    # ── MYMENSINGH DIVISION ─────────────────────────────────────────────

    # Mymensingh
    'Mymensingh Sadar':     (24.7471, 90.4203),
    'Bhaluka':              (24.3600, 90.3600),
    'Gouripur':             (24.8200, 90.2600),
    'Gauripur':             (24.8200, 90.2600),
    'Dhobaura':             (25.0600, 90.5300),
    'Fulbaria':             (24.5600, 90.2900),
    'Gaffargaon':           (24.4200, 90.5000),
    'Gauripur':             (24.8200, 90.2600),
    'Haluaghat':            (25.0900, 90.5700),
    'Ishwarganj':           (24.5200, 90.6500),
    'Muktagachha':          (24.7600, 90.2600),
    'Nandail':              (24.5000, 90.7000),
    'Phulpur':              (24.9900, 90.5000),
    'Trishal':              (24.5700, 90.3500),
    'Tarakanda':            (24.8600, 90.5200),

    # Netrokona
    'Netrokona Sadar':      (24.8710, 90.7278),
    'Atpara':               (24.8100, 90.8300),
    'Barhatta':             (24.7500, 90.7400),
    'Durgapur Netrokona':   (24.8800, 90.6400),
    'Kalmakanda':           (25.0700, 90.5300),
    'Kendua':               (24.9200, 90.8400),
    'Khaliajuri':           (24.7000, 91.0500),
    'Madan':                (24.7800, 91.0500),
    'Mohanganj':            (24.7400, 91.1200),
    'Purbadhala':           (24.9700, 90.7600),

    # Jamalpur
    'Jamalpur Sadar':       (24.9373, 89.9373),
    'Baksiganj':            (25.0600, 89.9900),
    'Bakshiganj':           (25.0333, 89.8167),
    'Dewanganj':            (25.2200, 89.9200),
    'Islampur':             (25.0800, 89.8200),
    'Madarganj':            (25.1300, 89.8400),
    'Melandah':             (24.9000, 89.8800),
    'Sariszapur':           (25.1600, 89.9200),

    # Sherpur
    'Sherpur Sadar':        (25.0204, 90.0152),
    'Jhenaigati':           (25.1000, 90.2200),
    'Nalitabari':           (25.0300, 90.2900),
    'Nakla':                (25.0100, 90.1800),
    'Sreebardi':            (25.1600, 90.0500),
}

BD_DISTRICT_GPS = {
    # All 64 districts — administrative capital coordinates
    # Includes 2018 name reform spellings as aliases
    'Dhaka':           (23.8103, 90.4125),
    'Gazipur':         (23.9999, 90.4203),
    'Narsingdi':       (23.9234, 90.7151),
    'Narayanganj':     (23.6238, 90.4997),
    'Manikganj':       (23.8634, 89.9947),
    'Munshiganj':      (23.5422, 90.5302),
    'Rajbari':         (23.7574, 89.6440),
    'Madaripur':       (23.1635, 90.2085),
    'Shariatpur':      (23.2434, 90.4351),
    'Faridpur':        (23.6070, 89.8429),
    'Gopalganj':       (23.0046, 89.8267),
    'Kishoreganj':     (24.4449, 90.7766),
    'Tangail':         (24.2513, 89.9167),
    # Chittagong Division
    'Chittagong':      (22.3384, 91.8317),
    'Chattogram':      (22.3384, 91.8317),   # 2018 reform
    "Cox's Bazar":     (21.4272, 92.0058),
    'Coxsbazar':       (21.4272, 92.0058),
    'Comilla':         (23.4682, 91.1788),
    'Cumilla':         (23.4682, 91.1788),   # 2018 reform
    'Brahmanbaria':    (23.9570, 91.1120),
    'Noakhali':        (22.8696, 91.0997),
    'Feni':            (23.0235, 91.3960),
    'Chandpur':        (23.2373, 90.6518),
    'Lakshmipur':      (22.9388, 90.8414),
    'Rangamati':       (22.6519, 92.2028),
    'Khagrachhari':    (23.1193, 91.9847),
    'Bandarban':       (22.1953, 92.2184),
    # Rajshahi Division
    'Rajshahi':        (24.3745, 88.6042),
    'Bogra':           (24.8465, 89.3720),
    'Bogura':          (24.8465, 89.3720),   # 2018 reform
    'Naogaon':         (24.9131, 88.7465),
    'Natore':          (24.4198, 88.9877),
    'Sirajganj':       (24.4508, 89.7013),
    'Sirajgonj':       (24.4508, 89.7013),
    'Pabna':           (24.0064, 89.2372),
    'Chapainawabganj': (24.5965, 88.2765),
    'Joypurhat':       (25.1026, 89.0197),
    # Khulna Division
    'Khulna':          (22.8456, 89.5403),
    'Jessore':         (23.1664, 89.2082),
    'Jashore':         (23.1664, 89.2082),   # 2018 reform
    'Satkhira':        (22.7185, 89.0705),
    'Bagerhat':        (22.6600, 89.7893),
    'Narail':          (23.1726, 89.5057),
    'Magura':          (23.4878, 89.4231),
    'Meherpur':        (23.7620, 88.6320),
    'Chuadanga':       (23.6438, 88.8419),
    'Jhenaidah':       (23.5448, 89.1520),
    'Jhenidah':        (23.5448, 89.1520),
    'Kushtia':         (23.9012, 89.1213),
    # Barisal Division
    'Barisal':         (22.7010, 90.3535),
    'Barishal':        (22.7010, 90.3535),   # 2018 reform
    'Bhola':           (22.6863, 90.6481),
    'Patuakhali':      (22.3596, 90.3290),
    'Pirojpur':        (22.5800, 89.9700),
    'Jhalokati':       (22.6400, 90.1900),
    'Barguna':         (22.1500, 90.1200),
    # Sylhet Division
    'Sylhet':          (24.8949, 91.8687),
    'Moulvibazar':     (24.4826, 91.7773),
    'Habiganj':        (24.3747, 91.4146),
    'Sunamganj':       (25.0659, 91.3985),
    # Rangpur Division
    'Rangpur':         (25.7439, 89.2752),
    'Dinajpur':        (25.6279, 88.6338),
    'Gaibandha':       (25.3288, 89.5449),
    'Kurigram':        (25.8074, 89.6360),
    'Lalmonirhat':     (25.9217, 89.2836),
    'Nilphamari':      (25.9313, 88.8561),
    'Panchagarh':      (26.3411, 88.5541),
    'Thakurgaon':      (26.0336, 88.4616),
    # Mymensingh Division
    'Mymensingh':      (24.7471, 90.4203),
    'Netrokona':       (24.8710, 90.7278),
    'Jamalpur':        (24.9373, 89.9373),
    'Sherpur':         (25.0204, 90.0152),
}

# ─────────────────────────────────────────────
# COLUMN ALIASES
# ─────────────────────────────────────────────
COLUMN_ALIASES = {
    'start':            ['start', 'datetime', 'date', 'timestamp', 'call_date',
                         'call date', 'date/time', 'date time', 'starttime',
                         'start time', 'start_datetime', 'start_dttime', 'startdttime',
                         'call_datetime', 'calldatetime', 'datetime_start',
                         # Teletalk 14-column CSV format
                         'start_time', 'starttime', 'call_start_time'],
    'operator':         ['operator', 'network', 'telco', 'carrier',
                         'provider name', 'provider_name', 'service provider',
                         'providername'],
    'party_a':          ['party a', 'party_a', 'a_number', 'msisdn_a', 'a-number',
                         'caller', 'originating', 'a number', 'partya', 'msisdn',
                         'aparty', 'a party', 'a_party'],
    'party_b':          ['party b', 'party_b', 'b_number', 'msisdn_b', 'b-number',
                         'called', 'terminating', 'b number', 'partyb', 'callee',
                         'bparty', 'b party', 'b_party'],
    'party_b_original': ['party b original', 'party_b_original', 'original_b',
                         'b_original', 'partyb_original'],
    'duration':         ['call duration', 'call_duration', 'duration',
                         'call_length', 'duration_sec', 'duration(sec)',
                         'callduration'],
    'usage_type':       ['usage type', 'usage_type', 'call_type', 'call type',
                         'type', 'direction', 'service_type',
                         'usagetype'],
    'cell_type':        ['cell type', 'cell_type', 'network_type', 'network type',
                         'technology', 'rat',
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


@st.cache_data(show_spinner=False, max_entries=10)
def _apply_bts_enrichment(df):
    """
    Teletalk CGI/ECGI + Robi BTS CSV GPS enrichment।
    load_and_clean()-এর CSV ও Excel দুই path-এই call করা হয়।
    df-কে in-place modify করে এবং modified df return করে।
    """
    import os as _os_e, tempfile as _tf_e, glob as _glob_e

    _UPLOADS_DIR_E = '/mnt/user-data/uploads'
    _CELL_DIR_E    = _os_e.path.join(_tf_e.gettempdir(), 'celltower_cache')

    # ── Teletalk CGI/ECGI Enrichment ──────────────────────────────────────
    try:
        _is_tt = False
        if 'operator' in df.columns:
            _is_tt = any('teletalk' in str(v).lower()
                         for v in df['operator'].dropna().unique())

        if _is_tt and 'cell_id' in df.columns:
            _tt_path = None
            for _d in [_UPLOADS_DIR_E, _CELL_DIR_E]:
                _p = _os_e.path.join(_d, 'Teletalk.csv')
                if _os_e.path.isfile(_p) and _os_e.path.getsize(_p) > 1000:
                    _tt_path = _p; break

            if _tt_path:
                _tt = pd.read_csv(_tt_path, dtype=str, encoding='utf-8',
                                  low_memory=False, on_bad_lines='skip')
                _tt.columns = [c.strip().lower() for c in _tt.columns]

                _cgi_col   = next((c for c in _tt.columns if 'cgi' in c or 'ecgi' in c), None)
                _lat_col   = next((c for c in _tt.columns if c == 'latitude'), None)
                _lon_col   = next((c for c in _tt.columns if c == 'longitude'), None)
                _addr_col  = next((c for c in _tt.columns if 'full' in c and 'address' in c), None)
                _dist_col  = next((c for c in _tt.columns if 'district' in c), None)
                _thana_col = next((c for c in _tt.columns if 'thana' in c or 'upazila' in c), None)

                if _cgi_col and _lat_col and _lon_col:
                    # ── Vectorized lookup build (iterrows বাদ) ──────────────
                    _tt2 = _tt[[_cgi_col, _lat_col, _lon_col]
                               + ([_addr_col] if _addr_col else [])
                               + ([_thana_col] if _thana_col and _thana_col != 'site id' else [])
                               ].copy()
                    _tt2[_lat_col]  = pd.to_numeric(_tt2[_lat_col],  errors='coerce')
                    _tt2[_lon_col]  = pd.to_numeric(_tt2[_lon_col],  errors='coerce')
                    _tt2 = _tt2.dropna(subset=[_lat_col, _lon_col])
                    _tt2 = _tt2[(_tt2[_lat_col].between(19,27)) & (_tt2[_lon_col].between(87,93))]
                    _tt2[_cgi_col] = _tt2[_cgi_col].astype(str).str.strip()
                    _tt2 = _tt2[~_tt2[_cgi_col].isin(['nan','','0'])]
                    _tt2 = _tt2.drop_duplicates(subset=[_cgi_col])

                    if _addr_col:
                        _tt2['_ad'] = _tt2[_addr_col].astype(str).str.strip()
                        _tt2['_ad'] = _tt2['_ad'].replace({'nan':'','None':''})
                        _tt2['_di'] = _tt2['_ad'].apply(
                            lambda a: a.split(',')[-1].strip() if ',' in a else '')
                    else:
                        _tt2['_ad'] = ''; _tt2['_di'] = ''

                    if _thana_col and _thana_col != 'site id' and _thana_col in _tt2.columns:
                        _tt2['_th'] = _tt2[_thana_col].astype(str).str.strip()
                        _tt2['_th'] = _tt2['_th'].replace({'nan':'','None':''})
                    else:
                        _tt2['_th'] = ''

                    _tt_map = dict(zip(
                        _tt2[_cgi_col],
                        zip(_tt2[_lat_col], _tt2[_lon_col],
                            _tt2['_ad'], _tt2['_di'], _tt2['_th'])
                    ))

                    if _tt_map:
                        # ── Vectorized GPS inject (iterrows বাদ) ─────────────
                        _cid_series = df['cell_id'].astype(str).str.strip()
                        _cid_series = _cid_series.where(
                            ~(_cid_series.str.endswith('.0') & _cid_series.str[:-2].str.isdigit()),
                            _cid_series.str[:-2]
                        )
                        _matched_data = _cid_series.map(_tt_map)
                        _mask = _matched_data.notna()

                        if _mask.any():
                            _matched_df = pd.DataFrame(
                                _matched_data[_mask].tolist(),
                                index=_matched_data[_mask].index,
                                columns=['_lat','_lon','_addr','_dist','_thana']
                            )
                            if 'address' not in df.columns:
                                df['address'] = ''
                            # Address: blank row-এ CSV address দাও
                            _blank_mask = (
                                df['address'].astype(str).str.strip().isin(
                                    ['','nan','None','NaN']) | ~df['address'].apply(_is_valid_address)
                            )
                            df.loc[_mask & _blank_mask, 'address'] = \
                                _matched_df.loc[_blank_mask[_mask], '_addr'].values

                            df.loc[_mask, 'cell_lat']       = pd.to_numeric(_matched_df['_lat'], errors='coerce')
                            df.loc[_mask, 'cell_lon']       = pd.to_numeric(_matched_df['_lon'], errors='coerce')
                            df.loc[_mask, 'loc_method']     = 'cell_exact'
                            df.loc[_mask, 'csv_district']   = _matched_df['_dist'].values
                            df.loc[_mask, 'csv_thana']      = _matched_df['_thana'].values
                            df.loc[_mask, 'cell_csv_label'] = _matched_df['_addr'].values
    except Exception:
        logger.debug('Teletalk enrichment (shared) failed', exc_info=True)

    return df


def load_and_clean(file_bytes):
    # ── CSV format detect ──────────────────────────────────────────────────
    # Excel magic bytes: 50 4B (xlsx) বা D0 CF (xls)
    # CSV হলে directly read করো, Excel path skip করো
    _magic = file_bytes[:4]
    _is_csv = not (_magic[:2] in (b'PK', b'\xd0\xcf') or _magic[:4] == b'\xd0\xcf\x11\xe0')

    if _is_csv:
        # ── CSV auto-encoding detect ──
        _csv_df = None
        for _enc in ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252']:
            try:
                _csv_df = pd.read_csv(
                    io.BytesIO(file_bytes),
                    dtype=str,
                    encoding=_enc,
                    on_bad_lines='skip',
                    low_memory=False,
                )
                break
            except Exception:
                continue
        if _csv_df is None or _csv_df.empty:
            raise ValueError("CSV file পড়া যাচ্ছে না — encoding সমস্যা বা empty file।")

        total_raw = len(_csv_df)
        df = _csv_df.copy()
        # Rename columns via COLUMN_ALIASES (same as Excel path)
        _csv_col_map = {}
        for _key in COLUMN_ALIASES:
            _found = detect_column(df.columns, _key)
            if _found:
                _csv_col_map[_key] = _found
        _csv_rename = {v: k for k, v in _csv_col_map.items()}
        df = df.rename(columns=_csv_rename)

        # ── Parse datetime — CSV date: dd/MM/yyyy HH:mm:ss ──
        if 'start' in df.columns:
            _raw = df['start'].astype(str).str.strip()
            df['start'] = pd.to_datetime(_raw, dayfirst=True, errors='coerce')
            if df['start'].isna().sum() > len(df) * 0.3:
                df['start'] = pd.to_datetime(_raw, dayfirst=False, errors='coerce')
            df = df.dropna(subset=['start'])
            df = df.sort_values('start').reset_index(drop=True)

        # ── Parse duration ──
        if 'duration' in df.columns:
            df['duration'] = pd.to_numeric(df['duration'], errors='coerce').fillna(0).astype(int)

        # ── Fix Party B ──
        if 'party_b_original' in df.columns:
            pb_orig = df['party_b_original'].replace(['nan', 'None'], pd.NA)
            df['party_b_clean'] = pb_orig.fillna(df.get('party_b', pb_orig)).astype(str).str.strip()
        elif 'party_b' in df.columns:
            df['party_b_clean'] = df['party_b'].astype(str).str.strip()

        # ── Normalize phone numbers ──
        if 'party_b_clean' in df.columns:
            df['party_b_norm'] = df['party_b_clean'].apply(
                lambda v: normalize_number(v) if is_valid_number(v) else v
            )
        if 'party_a' in df.columns:
            df['party_a'] = df['party_a'].astype(str).str.strip()

        # ── Anomaly detection ──
        anomaly_count = 0
        if 'party_b_clean' in df.columns:
            invalid_pb = ~df['party_b_clean'].apply(is_valid_number)
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

        col_map = _csv_col_map

        # ── Teletalk + Robi GPS enrichment — CSV path ────────────────────────
        df = _apply_bts_enrichment(df)

        return df, col_map, total_raw, anomaly_count, None
    # ── End CSV path ───────────────────────────────────────────────────────

    # ── Detect best sheet (Excel path) ──

    # ── Detect best sheet (Excel path) ──
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

    # ── Teletalk: CGI/ECGI দিয়ে Address + GPS Enrichment ──────────────────────
    # Teletalk CDR-এ Cell ID কলামে full CGI/ECGI থাকে (e.g. 470040122737532)।
    # Address কলাম সাধারণত blank থাকে।
    # Teletalk.csv (uploads ফোল্ডারে বা CELL_DIR-এ) থেকে CGI/ECGI → Full Address,
    # Latitude, Longitude, District মিলিয়ে ফাঁকা Address এবং GPS পূরণ করা হয়।
    # অন্য operator-এর CDR-এ এই block কাজ করে না।
    try:
        _is_teletalk = False
        if 'operator' in df.columns:
            _ops_vals = df['operator'].dropna().astype(str).str.lower().unique()
            _is_teletalk = any('teletalk' in _ov for _ov in _ops_vals)

        if _is_teletalk and 'cell_id' in df.columns:
            import os as _os_tt, tempfile as _tf_tt
            _UPLOADS_DIR_TT  = '/mnt/user-data/uploads'
            _CELL_DIR_TT     = _os_tt.path.join(_tf_tt.gettempdir(), 'celltower_cache')
            _TELETALK_FNAME  = 'Teletalk.csv'

            # CSV খোঁজার অগ্রাধিকার: uploads ফোল্ডার → celltower_cache
            _tt_csv_path = None
            for _d in [_UPLOADS_DIR_TT, _CELL_DIR_TT]:
                _p = _os_tt.path.join(_d, _TELETALK_FNAME)
                if _os_tt.path.isfile(_p) and _os_tt.path.getsize(_p) > 1000:
                    _tt_csv_path = _p
                    break

            if _tt_csv_path:
                # CGI/ECGI → (lat, lon, full_address, district, thana) lookup map বানাই
                _tt_csv = pd.read_csv(_tt_csv_path, dtype=str, encoding='utf-8',
                                      low_memory=False, on_bad_lines='skip')
                _tt_csv.columns = [c.strip().lower() for c in _tt_csv.columns]

                # Column detect
                _cgi_col  = next((c for c in _tt_csv.columns
                                  if 'cgi' in c or 'ecgi' in c), None)
                _lat_col  = next((c for c in _tt_csv.columns
                                  if c == 'latitude'), None)
                _lon_col  = next((c for c in _tt_csv.columns
                                  if c == 'longitude'), None)
                _addr_col = next((c for c in _tt_csv.columns
                                  if 'full' in c and 'address' in c), None)
                _dist_col = next((c for c in _tt_csv.columns
                                  if 'district' in c or c == 'unnamed: 1'), None)
                _thana_col = next((c for c in _tt_csv.columns
                                   if 'thana' in c or 'upazila' in c or 'site id' in c), None)

                if _cgi_col and _lat_col and _lon_col:
                    _tt_map = {}  # CGI/ECGI string → (lat, lon, address, district, thana)
                    for _, _r in _tt_csv.iterrows():
                        try:
                            _cgi_v = str(_r[_cgi_col]).strip()
                            if not _cgi_v or _cgi_v in ('nan', '', '0'):
                                continue
                            _lat_v = float(_r[_lat_col])
                            _lon_v = float(_r[_lon_col])
                            if not (19 <= _lat_v <= 27 and 87 <= _lon_v <= 93):
                                continue
                            _addr_v = str(_r[_addr_col]).strip() if _addr_col else ''
                            if _addr_v in ('nan', 'None'): _addr_v = ''
                            # District: full address শেষ comma-part
                            _dist_v = ''
                            if _addr_v and ',' in _addr_v:
                                _dist_v = _addr_v.split(',')[-1].strip()
                            # Thana: thana/upazila column — site id বাদ (site code)
                            _thana_v = ''
                            if _thana_col and _thana_col != 'site id':
                                _thana_v = str(_r[_thana_col]).strip()
                                if _thana_v in ('nan', 'None'): _thana_v = ''
                            _tt_map[_cgi_v] = (_lat_v, _lon_v, _addr_v, _dist_v, _thana_v)
                        except Exception:
                            continue

                    if _tt_map:
                        # CDR-এর Cell ID = CGI/ECGI → মিলিয়ে Address ও GPS দাও
                        _new_addr    = df['address'].copy()   if 'address'      in df.columns else pd.Series([''] * len(df), dtype=str)
                        _new_lat     = pd.Series([None] * len(df), dtype=object)
                        _new_lon     = pd.Series([None] * len(df), dtype=object)
                        _new_method  = pd.Series(['none'] * len(df), dtype=str)
                        _new_dist    = pd.Series([''] * len(df),    dtype=str)
                        _new_thana   = pd.Series([''] * len(df),    dtype=str)
                        _new_label   = pd.Series([''] * len(df),    dtype=str)

                        for _idx, _row in df.iterrows():
                            _cid_raw = str(_row.get('cell_id', '')).strip()
                            # float suffix পরিষ্কার (e.g. "470040122737532.0")
                            if _cid_raw.endswith('.0') and _cid_raw[:-2].isdigit():
                                _cid_raw = _cid_raw[:-2]

                            if _cid_raw not in _tt_map:
                                continue

                            _t_lat, _t_lon, _t_addr, _t_dist, _t_thana = _tt_map[_cid_raw]

                            # Address: CDR-এ blank বা invalid থাকলে CSV-এর Full Address দাও
                            # invalid = '', 'nan', '-', ',', '-,-', ',,', '- -' ইত্যাদি
                            _cur_addr = str(_row.get('address', '') or '').strip()
                            _addr_is_blank = (
                                _cur_addr in ('', 'nan', 'None', 'NaN')
                                or pd.isna(_row.get('address'))
                                or not _is_valid_address(_cur_addr)
                            )
                            if _addr_is_blank:
                                _new_addr.at[_idx] = _t_addr if _t_addr else _cur_addr

                            # GPS সবসময় দাও (exact CSV match)
                            _new_lat.at[_idx]    = _t_lat
                            _new_lon.at[_idx]    = _t_lon
                            _new_method.at[_idx] = 'cell_exact'
                            _new_dist.at[_idx]   = _t_dist
                            _new_thana.at[_idx]  = _t_thana
                            _new_label.at[_idx]  = _t_addr  # CSV Full Address as label

                        # DataFrame-এ যোগ করি
                        if 'address' not in df.columns:
                            df['address'] = ''
                        df['address']       = _new_addr
                        df['cell_lat']      = pd.to_numeric(_new_lat,   errors='coerce')
                        df['cell_lon']      = pd.to_numeric(_new_lon,   errors='coerce')
                        df['loc_method']    = _new_method
                        df['csv_district']  = _new_dist
                        df['csv_thana']     = _new_thana
                        df['cell_csv_label']= _new_label
    except Exception:
        logger.debug('Teletalk CGI/ECGI address+GPS enrichment failed', exc_info=True)
    # ── End Teletalk Enrichment ────────────────────────────────────────────────

    # ── Robi BTS CSV Enrichment ───────────────────────────────────────────────
    # Robi 4G: ENODEBID//100 = CDR LAC_ID, Cell_ID শেষ ২ digit = CSV CELL_ID
    # Robi 2G: CSV LAC = CDR LAC_ID, CSV CELL_ID = CDR Cell_ID (direct match)
    # CSV ফাইল নাম: Robi_4G*.csv, Robi_2G*.csv (uploads বা celltower_cache)
    try:
        _is_robi = False
        if 'operator' in df.columns:
            _ops_vals_r = df['operator'].dropna().astype(str).str.lower().unique()
            _is_robi = any('robi' in _ov for _ov in _ops_vals_r)

        if _is_robi and 'cell_id' in df.columns and 'lac' in df.columns:
            import os as _os_r, glob as _glob_r, tempfile as _tf_r

            _UPLOADS_DIR_R = '/mnt/user-data/uploads'
            _CELL_DIR_R    = _os_r.path.join(_tf_r.gettempdir(), 'celltower_cache')

            def _find_robi_csv(tech_key):
                """Robi_4G*.csv বা Robi_2G*.csv খোঁজো।"""
                for _d in [_UPLOADS_DIR_R, _CELL_DIR_R]:
                    hits = _glob_r.glob(_os_r.path.join(_d, f'Robi_{tech_key}*.csv'))
                    hits += _glob_r.glob(_os_r.path.join(_d, f'robi_{tech_key.lower()}*.csv'))
                    hits = [h for h in hits if _os_r.path.getsize(h) > 1000]
                    if hits:
                        return hits[0]
                return None

            # ── Lookup map তৈরি ──────────────────────────────────────────────
            _robi_map = {}  # (lac_id, sector_cell_id) → (lat, lon, addr, dist, thana)

            # 4G CSV
            _r4g_path = _find_robi_csv('4G')
            if _r4g_path:
                try:
                    _r4g = pd.read_csv(_r4g_path, engine='python', encoding='latin-1',
                                       low_memory=False, on_bad_lines='skip')
                    _r4g.columns = [c.strip().lower() for c in _r4g.columns]
                    if all(c in _r4g.columns for c in ['enodebid','cell_id','latitude','longitude']):
                        # ── Vectorized 4G lookup build ────────────────────────
                        _r4g_v = _r4g[['enodebid','cell_id','latitude','longitude',
                                       'address','district','thana']].copy() \
                                 if all(c in _r4g.columns for c in ['address','district','thana']) \
                                 else _r4g[['enodebid','cell_id','latitude','longitude']].copy()
                        _r4g_v['enodebid']  = pd.to_numeric(_r4g_v['enodebid'], errors='coerce')
                        _r4g_v['cell_id']   = pd.to_numeric(_r4g_v['cell_id'],  errors='coerce')
                        _r4g_v['latitude']  = pd.to_numeric(_r4g_v['latitude'], errors='coerce')
                        _r4g_v['longitude'] = pd.to_numeric(_r4g_v['longitude'],errors='coerce')
                        _r4g_v = _r4g_v.dropna(subset=['enodebid','cell_id','latitude','longitude'])
                        _r4g_v = _r4g_v[
                            _r4g_v['latitude'].between(19,27) &
                            _r4g_v['longitude'].between(87,93)
                        ]
                        _r4g_v['_derived_lac'] = (_r4g_v['enodebid'] // 100).astype(int)
                        _r4g_v['_cid_int']     = _r4g_v['cell_id'].astype(int)
                        _r4g_v = _r4g_v.drop_duplicates(subset=['_derived_lac','_cid_int'])
                        for _fv in ('nan','None','NaN'):
                            for _col in ['address','district','thana']:
                                if _col in _r4g_v.columns:
                                    _r4g_v[_col] = _r4g_v[_col].astype(str).replace(_fv,'')
                        for _, _rr in _r4g_v.iterrows():
                            _key = (int(_rr['_derived_lac']), int(_rr['_cid_int']))
                            if _key not in _robi_map:
                                _robi_map[_key] = (
                                    float(_rr['latitude']), float(_rr['longitude']),
                                    str(_rr.get('address','')), str(_rr.get('district','')),
                                    str(_rr.get('thana',''))
                                )
                except Exception:
                    logger.debug('Suppressed exception', exc_info=True)

            # 2G CSV
            _r2g_path = _find_robi_csv('2G')
            _robi_2g_cell_only = {}   # cell_id → [(lat,lon,addr,dist,thana)] fallback
            if _r2g_path:
                try:
                    _r2g = pd.read_csv(_r2g_path, engine='python', encoding='latin-1',
                                       low_memory=False, on_bad_lines='skip')
                    _r2g.columns = [c.strip().lower() for c in _r2g.columns]
                    if all(c in _r2g.columns for c in ['lac','cell_id','latitude','longitude']):
                        # ── Vectorized 2G lookup build ────────────────────────
                        _r2g_v = _r2g[['lac','cell_id','latitude','longitude',
                                       'address','district','thana']].copy() \
                                 if all(c in _r2g.columns for c in ['address','district','thana']) \
                                 else _r2g[['lac','cell_id','latitude','longitude']].copy()
                        _r2g_v['lac']       = pd.to_numeric(_r2g_v['lac'],      errors='coerce')
                        _r2g_v['cell_id']   = pd.to_numeric(_r2g_v['cell_id'],  errors='coerce')
                        _r2g_v['latitude']  = pd.to_numeric(_r2g_v['latitude'], errors='coerce')
                        _r2g_v['longitude'] = pd.to_numeric(_r2g_v['longitude'],errors='coerce')
                        _r2g_v = _r2g_v.dropna(subset=['lac','cell_id','latitude','longitude'])
                        _r2g_v = _r2g_v[
                            _r2g_v['latitude'].between(19,27) &
                            _r2g_v['longitude'].between(87,93)
                        ]
                        for _fv in ('nan','None','NaN'):
                            for _col in ['address','district','thana']:
                                if _col in _r2g_v.columns:
                                    _r2g_v[_col] = _r2g_v[_col].astype(str).replace(_fv,'')
                        for _, _rr in _r2g_v.iterrows():
                            try:
                                _lac2 = int(_rr['lac']); _cid2 = int(_rr['cell_id'])
                                _lat2 = float(_rr['latitude']); _lon2 = float(_rr['longitude'])
                                _addr2 = str(_rr.get('address','')); _dist2 = str(_rr.get('district',''))
                                _thana2= str(_rr.get('thana',''))
                                _key2 = (_lac2, _cid2)
                                if _key2 not in _robi_map:
                                    _robi_map[_key2] = (_lat2,_lon2,_addr2,_dist2,_thana2)
                                if _cid2 not in _robi_2g_cell_only:
                                    _robi_2g_cell_only[_cid2] = []
                                _robi_2g_cell_only[_cid2].append(
                                    (_lac2,_lat2,_lon2,_addr2,_dist2.lower(),_thana2))
                            except Exception:
                                continue
                except Exception:
                    logger.debug('Suppressed exception', exc_info=True)

            # ── CDR rows-এ GPS ও Address ইনজেক্ট করো (Vectorized) ──────────
            if _robi_map:
                # cell_id normalize: '14025.0' → '14025'
                _cid_s = df['cell_id'].astype(str).str.strip()
                _cid_s = _cid_s.where(
                    ~(_cid_s.str.endswith('.0') & _cid_s.str[:-2].str.isdigit()),
                    _cid_s.str[:-2]
                )
                _cid_int_s = pd.to_numeric(_cid_s, errors='coerce').astype('Int64')
                _lac_s     = pd.to_numeric(df.get('lac', pd.Series(dtype=str)), errors='coerce').astype('Int64')
                _ctype_s   = df['cell_type'].astype(str).str.upper() if 'cell_type' in df.columns \
                             else pd.Series(['2G']*len(df))

                # 4G: sector = শেষ ২ digit
                _sector_s = _cid_s.apply(
                    lambda s: int(s[-2:]) if len(s) >= 2 and s.isdigit() else None
                )

                # Build key per row
                _keys = pd.Series([
                    (_lac_s.iloc[i], _sector_s.iloc[i])
                    if _ctype_s.iloc[i] == '4G'
                    else (_lac_s.iloc[i], _cid_int_s.iloc[i])
                    for i in range(len(df))
                ], index=df.index)

                # Map keys → GPS tuples
                _gps_mapped = _keys.map(lambda k: _robi_map.get(k) if pd.notna(k[0]) and pd.notna(k[1]) else None)
                _mask_r = _gps_mapped.notna()

                # 2G fallback for unmatched rows
                _no_match = ~_mask_r & (_ctype_s != '4G')
                if _no_match.any():
                    for _idx in df.index[_no_match]:
                        try:
                            _ci = int(_cid_int_s.at[_idx])
                            if _ci not in _robi_2g_cell_only: continue
                            _cands = _robi_2g_cell_only[_ci]
                            _addr_lo = str(df.at[_idx, 'address'] or '').lower() \
                                       if 'address' in df.columns else ''
                            if len(_cands) == 1:
                                _best = _cands[0]
                            else:
                                _best = next((c for c in _cands if c[4] and c[4] in _addr_lo), None)
                                if not _best:
                                    _dk = [c for c in _cands if 'dhaka' in c[4]]
                                    _best = _dk[0] if _dk else _cands[0]
                            if _best:
                                _gps_mapped.at[_idx] = (_best[1], _best[2], _best[3], _best[4], _best[5])
                                _mask_r.at[_idx] = True
                        except Exception:
                            continue

                if _mask_r.any():
                    _gps_df = pd.DataFrame(
                        _gps_mapped[_mask_r].tolist(),
                        index=_gps_mapped[_mask_r].index,
                        columns=['_lat','_lon','_addr','_dist','_thana']
                    )
                    if 'address' not in df.columns: df['address'] = ''
                    _blank_r = df['address'].astype(str).str.strip().isin(['','nan','None','NaN'])
                    df.loc[_mask_r & _blank_r, 'address'] = \
                        _gps_df.loc[_blank_r[_mask_r], '_addr'].values

                    df.loc[_mask_r, 'cell_lat']       = pd.to_numeric(_gps_df['_lat'], errors='coerce')
                    df.loc[_mask_r, 'cell_lon']       = pd.to_numeric(_gps_df['_lon'], errors='coerce')
                    df.loc[_mask_r, 'loc_method']     = 'cell_exact'
                    df.loc[_mask_r, 'csv_district']   = _gps_df['_dist'].values
                    df.loc[_mask_r, 'csv_thana']      = _gps_df['_thana'].values
                    df.loc[_mask_r, 'cell_csv_label'] = _gps_df['_addr'].values

    except Exception:
        logger.debug('Robi BTS CSV address+GPS enrichment failed', exc_info=True)
    # ── End Robi BTS Enrichment ───────────────────────────────────────────────

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
        mv = _best_address_series(sub).value_counts() if len(sub) else pd.Series()
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
        mv = _best_address_series(sub).value_counts() if len(sub) else pd.Series()
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
        mv = _best_address_series(sub).value_counts() if len(sub) else pd.Series()
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
    """
    Top N locations with columns:
    CDR Location (BTS Address) | Cell Tower Location (CSV) | GPS Coordinates | Count
    Priority: CSV exact GPS → BD_COORDS text-based fallback
    Teletalk: address-less rows যাদের cell_csv_label আছে সেগুলোও include করা হয়।
    """
    if 'address' not in df.columns and 'cell_csv_label' not in df.columns:
        return pd.DataFrame()
    data = df if mask is None else df[mask]

    # ── Rows with valid CDR address ──
    if 'address' in data.columns:
        valid_rows = data[data['address'].notna() & data['address'].apply(_is_valid_address)].copy()
    else:
        valid_rows = pd.DataFrame()

    # ── Teletalk fallback: address blank/invalid কিন্তু cell_csv_label আছে ──
    if 'cell_csv_label' in data.columns and 'loc_method' in data.columns:
        enriched_cands = data[data['loc_method'] == 'cell_exact'].copy()
        if 'address' in data.columns:
            enriched_cands = enriched_cands[
                ~(enriched_cands['address'].notna() &
                  enriched_cands['address'].apply(_is_valid_address))
            ]
        enriched_cands = enriched_cands[
            enriched_cands['cell_csv_label'].notna() &
            (enriched_cands['cell_csv_label'].str.strip() != '')
        ].copy()
        if not enriched_cands.empty:
            enriched_cands['address'] = enriched_cands['cell_csv_label']
            valid_rows = pd.concat([valid_rows, enriched_cands], ignore_index=True)

    if valid_rows.empty: return pd.DataFrame()

    has_gps      = ('cell_lat' in valid_rows.columns and 'cell_lon' in valid_rows.columns)
    has_csv_label = 'cell_csv_label' in valid_rows.columns

    rows = []
    for addr, grp in valid_rows.groupby('address'):
        cnt = len(grp)
        gps_coord = '—'
        csv_lbl   = '—'

        # ── Priority 1: CSV exact GPS ──
        if has_gps and 'loc_method' in grp.columns:
            exact = grp[grp['loc_method'] == 'cell_exact']
            src   = exact if not exact.empty else pd.DataFrame()
            if not src.empty:
                lats = src['cell_lat'].dropna()
                lons = src['cell_lon'].dropna()
                if not lats.empty and not lons.empty:
                    gps_coord = f"{round(float(lats.iloc[0]), 6)}, {round(float(lons.iloc[0]), 6)}"

        # ── Priority 2: BD_COORDS text-based fallback ──
        if gps_coord == '—':
            try:
                from cdr_funcs import movement_pattern_analysis as _mpa
            except Exception:
                logger.debug('suppressed exception', exc_info=True)
            # inline parse using same BD_COORDS logic via movement helper
            # Use _extract_district / parse_ud indirectly: just show '—' if no CSV match
            pass

        # ── CSV label ──
        if has_csv_label and 'loc_method' in grp.columns:
            csv_vals = grp[grp['loc_method'] == 'cell_exact']['cell_csv_label'].dropna()
            if not csv_vals.empty:
                csv_lbl = str(csv_vals.iloc[0]).strip('"').strip()

        rows.append({
            'CDR Location (BTS Address)': addr,
            'Cell Tower Location (CSV)':  csv_lbl,
            'GPS Coordinates':            gps_coord,
            'Count':                      cnt,
        })
    rows.sort(key=lambda x: x['Count'], reverse=True)
    return pd.DataFrame(rows[:n])

def location_summary(df):
    """Location summary with GPS Coordinates (lat, lon combined) from CSV first, text fallback.
    Teletalk: address blank/invalid হলে cell_csv_label দিয়ে address পূরণ করা হয়।
    """
    if 'address' not in df.columns and 'cell_csv_label' not in df.columns:
        return pd.DataFrame()

    # Teletalk fallback: address invalid rows-এ cell_csv_label বসাই (working copy)
    _df = df.copy()
    if 'cell_csv_label' in _df.columns and 'loc_method' in _df.columns:
        _mask_invalid = (
            _df['address'].isna() |
            ~_df['address'].apply(lambda a: _is_valid_address(str(a)) if pd.notna(a) else False)
        ) if 'address' in _df.columns else pd.Series([True] * len(_df), index=_df.index)
        _mask_enriched = (_df['loc_method'] == 'cell_exact') & _df['cell_csv_label'].notna() & (_df['cell_csv_label'].str.strip() != '')
        _fill_mask = _mask_invalid & _mask_enriched
        if 'address' not in _df.columns:
            _df['address'] = ''
        _df.loc[_fill_mask, 'address'] = _df.loc[_fill_mask, 'cell_csv_label']

    addrs = _df['address'].dropna().apply(lambda a: a if _is_valid_address(a) else None).dropna()
    if len(addrs)==0: return pd.DataFrame()
    mv = addrs.value_counts()

    has_gps = ('cell_lat' in _df.columns and 'cell_lon' in _df.columns)

    def get_gps(addr):
        """CSV exact GPS first, then text-based fallback → returns 'lat, lon' string or '—'"""
        if not addr or addr == 'N/A': return '—'
        rows = _df[_df['address'] == addr]
        # Priority 1: CSV exact
        if has_gps and 'loc_method' in rows.columns:
            exact = rows[rows['loc_method'] == 'cell_exact']
            if not exact.empty:
                lats = exact['cell_lat'].dropna()
                lons = exact['cell_lon'].dropna()
                if not lats.empty and not lons.empty:
                    return f"{round(float(lats.iloc[0]), 6)}, {round(float(lons.iloc[0]), 6)}"
        # Priority 2: text-based (cell_lat may still be set from BD_COORDS)
        if has_gps:
            lats = rows['cell_lat'].dropna()
            lons = rows['cell_lon'].dropna()
            if not lats.empty and not lons.empty:
                return f"{round(float(lats.iloc[0]), 6)}, {round(float(lons.iloc[0]), 6)}"
        return '—'

    def top_addr(mask):
        if mask is None or 'start' not in _df.columns: return ('N/A', 0)
        sub = _df[mask]['address'].dropna().apply(
            lambda a: a if _is_valid_address(a) else None).dropna().value_counts()
        return (sub.index[0], int(sub.iloc[0])) if not sub.empty else ('N/A', 0)

    home_mask    = _df['start'].dt.hour.astype(int).isin(list(range(0,6))+list(range(22,24))) if 'start' in _df.columns else None
    work_mask    = (_df['start'].dt.hour.astype(int)>=8)&(_df['start'].dt.hour.astype(int)<18) if 'start' in _df.columns else None
    weekend_mask = _df['start'].dt.dayofweek.astype(int).isin([4,5]) if 'start' in _df.columns else None

    h, hc = top_addr(home_mask)
    w, wc = top_addr(work_mask)
    e, ec = top_addr(weekend_mask)

    most_visited = mv.index[0] if not mv.empty else 'N/A'
    most_cnt     = int(mv.iloc[0]) if not mv.empty else 0
    total_towers = _df['cell_id'].nunique() if 'cell_id' in _df.columns else len(mv)

    return pd.DataFrame({
        'Metric':          ['Total Towers Visited', 'Most Visited Place',
                            'Probable Home', 'Probable Work', 'Probable Weekend'],
        'Location':        ['N/A', most_visited, h, w, e],
        'Count':           ['N/A', most_cnt, hc, wc, ec],
        'GPS Coordinates': ['N/A', get_gps(most_visited), get_gps(h), get_gps(w), get_gps(e)],
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
    """Top locations in last N days — with GPS columns."""
    if 'start' not in df.columns or 'address' not in df.columns:
        return pd.DataFrame()
    max_date = df['start'].max()
    cutoff = max_date - pd.Timedelta(days=days)
    recent = df[df['start'] >= cutoff]
    return top_locations(recent, mask=None, n=n)


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
    return df.to_html(index=False, border=0, classes="dataframe")

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


def _best_address_series(sub):
    """
    DataFrame subset থেকে best address Series বের করে।
    CDR address valid হলে সেটা নেয়, না হলে cell_csv_label (Teletalk CSV) নেয়।
    """
    if 'address' not in sub.columns and 'cell_csv_label' not in sub.columns:
        return pd.Series(dtype=str)
    if 'address' in sub.columns:
        addr = sub['address'].copy().astype(str)
    else:
        addr = pd.Series([''] * len(sub), index=sub.index, dtype=str)
    # Teletalk fallback: address invalid হলে cell_csv_label ব্যবহার করি
    if 'cell_csv_label' in sub.columns:
        invalid_mask = ~addr.apply(_is_valid_address)
        csv_lbl = sub['cell_csv_label'].fillna('').astype(str)
        addr = addr.where(~invalid_mask, csv_lbl.where(csv_lbl != '', addr))
    return addr.where(addr.apply(_is_valid_address)).dropna()

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
            districts = list({str(t['district']) for t in trips if t.get('district') and str(t.get('district','')) not in ('','nan','None')})
            loc_comment += (f" The subscriber travelled outside the home district on {len(trips)} occasion(s), "
                            f"visiting: {', '.join(str(d) for d in districts[:5])}.")
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

    # math → _math_mod (module-level import)

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
        R=6371.0; p1,p2=_math_mod.radians(la1),_math_mod.radians(la2)
        dp=_math_mod.radians(la2-la1); dl=_math_mod.radians(lo2-lo1)
        a=_math_mod.sin(dp/2)**2+_math_mod.cos(p1)*_math_mod.cos(p2)*_math_mod.sin(dl/2)**2
        return R*2*_math_mod.atan2(_math_mod.sqrt(a),_math_mod.sqrt(1-a))

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

    # ── df_loc: valid address rows + Teletalk cell_csv_label fallback ──────────
    _df_mv = df.copy()
    if 'cell_csv_label' in _df_mv.columns and 'loc_method' in _df_mv.columns:
        _mv_invalid = (
            _df_mv['address'].isna() |
            ~_df_mv['address'].apply(lambda a: _is_valid_address(str(a)) if pd.notna(a) else False)
        ) if 'address' in _df_mv.columns else pd.Series([True]*len(_df_mv), index=_df_mv.index)
        _mv_enriched = (
            (_df_mv['loc_method'] == 'cell_exact') &
            _df_mv['cell_csv_label'].notna() &
            (_df_mv['cell_csv_label'].str.strip() != '')
        )
        if 'address' not in _df_mv.columns:
            _df_mv['address'] = ''
        _df_mv.loc[_mv_invalid & _mv_enriched, 'address'] = _df_mv.loc[_mv_invalid & _mv_enriched, 'cell_csv_label']

    df_loc=_df_mv[_df_mv["address"].notna()&_df_mv["address"].apply(_is_valid_address)].copy()
    df_loc=df_loc.sort_values("start").reset_index(drop=True)
    if df_loc.empty: return None

    home_dist = _home_district(df_loc)
    work_dist = None  # Not used in new logic

    # ── Home coord = Most Frequent Location GPS ──────────────────────────
    # সবচেয়ে বেশি records যে BTS-এ সেটার GPS home হিসেবে ব্যবহার করা হবে।
    # Priority: CSV exact GPS (most frequent) → text-based BD_COORDS → fallback
    home_coord = None
    home_label = None
    home_addr  = None

    has_cell = ("cell_lat" in df_loc.columns and "cell_lon" in df_loc.columns
                and "loc_method" in df_loc.columns)

    if has_cell:
        # Most frequent address that has an exact CSV GPS match
        exact_df = df_loc[df_loc["loc_method"] == "cell_exact"]
        if not exact_df.empty:
            top_addr_exact = exact_df["address"].value_counts().index[0]
            top_rows = exact_df[exact_df["address"] == top_addr_exact]
            lats = top_rows["cell_lat"].dropna()
            lons = top_rows["cell_lon"].dropna()
            if not lats.empty:
                home_coord = (float(lats.iloc[0]), float(lons.iloc[0]))
                home_addr  = top_addr_exact
                # Priority: csv_thana → csv_district → parse_ud → raw address[:30]
                _csv_thana = ""
                _csv_dist2 = ""
                if "csv_thana" in top_rows.columns:
                    _v = top_rows["csv_thana"].dropna()
                    if not _v.empty: _csv_thana = str(_v.iloc[0]).strip()
                if "csv_district" in top_rows.columns:
                    _v = top_rows["csv_district"].dropna()
                    if not _v.empty: _csv_dist2 = str(_v.iloc[0]).strip()
                if _csv_thana and _csv_thana not in ("nan", "None", ""):
                    home_label = _csv_thana.title()
                    if not home_dist and _csv_dist2 and _csv_dist2 not in ("nan", "None", ""):
                        home_dist = _csv_dist2.title()
                elif _csv_dist2 and _csv_dist2 not in ("nan", "None", ""):
                    home_label = _csv_dist2.title()
                else:
                    upa_h, dist_h = parse_ud(top_addr_exact)
                    home_label = upa_h or dist_h or top_addr_exact[:30]

    # Fallback: text-based BD_COORDS from most frequent address
    if not home_coord:
        for addr in df_loc["address"].value_counts().index:
            upa, dist = parse_ud(addr)
            coord = get_coord(upa, dist)
            if coord:
                home_coord = coord
                home_label = upa or dist
                home_addr  = addr
                break

    # Last resort fallback
    if not home_coord:
        home_coord = (23.7104, 90.4074)  # Dhaka fallback
        home_label = "Dhaka"

    HOME_KM=35.0; TRANSIT_H=4; MIN_HOME_STAY_H=5

    rows_e=[]
    last_known_km = 0.0
    last_known_coord = home_coord
    for _,row in df_loc.iterrows():
        coord = None

        # Priority 1: exact GPS from cell tower CSV
        if has_cell and row.get("loc_method") == "cell_exact":
            try:
                clat = float(row["cell_lat"])
                clon = float(row["cell_lon"])
                if 20 <= clat <= 27 and 88 <= clon <= 93:
                    coord = (clat, clon)
            except (ValueError, TypeError):
                coord = None

        # Priority 2: text-based BD_COORDS lookup
        upa, dist = parse_ud(row["address"])
        if coord is None:
            coord = get_coord(upa, dist)

        if coord:
            km=haversine(home_coord[0],home_coord[1],coord[0],coord[1])
            last_known_km=km
            last_known_coord=coord
        else:
            km=last_known_km
            coord=last_known_coord

        # District: prefer CSV district (accurate) over text-based parse
        # IMPORTANT: if no exact GPS (coord is None), district must also be empty
        # to prevent last_known_km from triggering false trips (Brahmanbaria bug)
        csv_dist = str(row.get("csv_district","")).strip().title() if has_cell else ""
        has_exact_gps = (has_cell and row.get("loc_method") == "cell_exact")
        if has_exact_gps:
            # Exact GPS → use csv_district first, fallback to text parse
            final_dist = csv_dist if csv_dist and csv_dist not in ("","Nan","None") else dist
        else:
            # No exact GPS → only use csv_district if available, else empty
            # Text-based dist + last_known_km = false positive risk
            final_dist = csv_dist if csv_dist and csv_dist not in ("","Nan","None") else ""

        rows_e.append({"ts":row["start"],"address":row["address"],
                       "csv_label": row.get("cell_csv_label","") if has_cell else "",
                       "upazila":upa,"district":final_dist,"coord":coord,"km":km,
                       "lat":coord[0] if coord else None,
                       "lon":coord[1] if coord else None,
                       "is_exact": (has_cell and row.get("loc_method") == "cell_exact")})
    df_e=pd.DataFrame(rows_e).sort_values("ts").reset_index(drop=True)
    # "away" = km >= threshold AND different district from home
    # Priority: csv_district from home rows → home_label text
    home_dist_for_away = ""
    if has_cell and "csv_district" in df_loc.columns and home_addr:
        home_rows = df_loc[df_loc["address"] == home_addr]
        if not home_rows.empty:
            hd = home_rows["csv_district"].dropna()
            hd = hd[hd.str.strip().str.lower().isin(["", "nan", "none"]) == False]
            if not hd.empty:
                home_dist_for_away = str(hd.value_counts().index[0]).strip().title()
    if not home_dist_for_away:
        home_dist_for_away = str(home_label).strip().title()
    # Conservative away: km >= HOME_KM AND district known AND different from home
    # If district is unknown/empty → NOT away (avoid false positives)
    df_e["district_clean"] = df_e["district"].str.strip().str.title().fillna("")
    df_e["away"] = (
        (df_e["km"] >= HOME_KM) &
        (df_e["district_clean"] != "") &           # district must be known
        (df_e["district_clean"] != home_dist_for_away.strip().title())
    )

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
            # Best address: prefer csv_label (from cell tower CSV), fallback to CDR address
            best_addr = ""
            if "csv_label" in gdf.columns:
                csv_vals = gdf["csv_label"].dropna()
                csv_vals = csv_vals[csv_vals.astype(str).str.strip() != ""]
                if not csv_vals.empty:
                    best_addr = str(csv_vals.value_counts().index[0])
            if not best_addr and not gdf.empty:
                _gdf_addr_vc = gdf["address"].value_counts()
                best_addr = str(_gdf_addr_vc.index[0]) if not _gdf_addr_vc.empty else ''
            # Best exact GPS coord from CSV
            exact_rows = gdf[gdf["is_exact"] == True] if "is_exact" in gdf.columns else pd.DataFrame()
            best_coord = None
            if not exact_rows.empty and exact_rows.iloc[0].get("coord"):
                best_coord = exact_rows.iloc[0]["coord"]
            elif not gdf.empty and gdf.iloc[0].get("coord"):
                best_coord = gdf.iloc[0]["coord"]
            # Include if: stayed >= TRANSIT_H OR only location OR far enough
            if hrs>=TRANSIT_H or len(loc_groups)==1 or max_km>=HOME_KM*1.5:
                stay_locs.append({
                    "upazila": upa_k,
                    "district": dist_k or None,
                    "km": max_km, "hours": round(hrs,1),
                    "first": gdf["ts"].min(), "last": gdf["ts"].max(),
                    "sample_addr": best_addr[:100],
                    "coord": best_coord,
                    "is_exact": not exact_rows.empty,
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
            "lat": round(float(dest["coord"][0]), 6) if dest.get("coord") else None,
            "lon": round(float(dest["coord"][1]), 6) if dest.get("coord") else None,
            "is_exact": dest.get("is_exact", False),
        })

    trips.sort(key=lambda x:x["start_date"],reverse=True)

    all_dates=sorted(df["start"].dt.date.unique())
    gaps=[]
    for k in range(len(all_dates)-1):
        d1=pd.Timestamp(all_dates[k]); d2=pd.Timestamp(all_dates[k+1])
        gd=(d2-d1).days-1
        if gd>4: gaps.append({"gap_start":str(all_dates[k]),"gap_end":str(all_dates[k+1]),"days":gd})

    total_days=((pd.Timestamp(all_dates[-1])-pd.Timestamp(all_dates[0])).days+1 if all_dates else 0)

    # ── Top 3 Frequent Locations — Home + 2 others ────────────────────────
    top_locations = []
    try:
        # df_loc = cleaned location rows — address + cell_lat/cell_lon সব আছে
        _tl_src     = df_loc.copy()
        _has_cell_tl = ('cell_lat' in _tl_src.columns and 'cell_lon' in _tl_src.columns)
        _has_loc_tl  = 'loc_method' in _tl_src.columns
        _addr_freq   = _tl_src["address"].value_counts() if "address" in _tl_src.columns else pd.Series(dtype=int)

        for _addr, _cnt in _addr_freq.head(15).items():
            if not _is_valid_address(str(_addr)):
                continue
            _upa, _dist = parse_ud(str(_addr))
            _coord = None
            # Priority 1: CSV exact GPS from cell tower database
            _addr_rows = _tl_src[_tl_src["address"] == _addr]
            if _has_cell_tl:
                _exact_r = _addr_rows[_addr_rows["loc_method"] == "cell_exact"] if _has_loc_tl else _addr_rows
                _lats = _exact_r["cell_lat"].dropna()
                _lons = _exact_r["cell_lon"].dropna()
                if not _lats.empty and not _lons.empty:
                    try:
                        _clat = float(_lats.mean())
                        _clon = float(_lons.mean())
                        if 19 <= _clat <= 27 and 87 <= _clon <= 93:
                            _coord = (_clat, _clon)
                    except Exception:
                        logger.debug('Suppressed exception', exc_info=True)
            # Priority 2: text-based BD_COORDS
            if not _coord:
                _coord = get_coord(_upa, _dist)
            # Label: csv_thana > csv_district > upazila > district > address prefix
            _lbl = ""
            if "csv_thana" in _tl_src.columns:
                _ct = _addr_rows["csv_thana"].dropna()
                _ct = _ct[~_ct.astype(str).str.lower().isin(["","nan","none"])]
                if not _ct.empty:
                    _lbl = str(_ct.value_counts().index[0]).strip().title()
            if not _lbl and "csv_district" in _tl_src.columns:
                _cd = _addr_rows["csv_district"].dropna()
                _cd = _cd[~_cd.astype(str).str.lower().isin(["","nan","none"])]
                if not _cd.empty:
                    _lbl = str(_cd.value_counts().index[0]).strip().title()
            if not _lbl:
                _lbl = _upa or _dist or str(_addr)[:35]
            # district for display — csv_district > text parse
            _disp_dist = ""
            if "csv_district" in _tl_src.columns:
                _dd = _addr_rows["csv_district"].dropna()
                _dd = _dd[~_dd.astype(str).str.lower().isin(["","nan","none"])]
                if not _dd.empty:
                    _disp_dist = str(_dd.value_counts().index[0]).strip().title()
            if not _disp_dist:
                _disp_dist = _dist or ""
            top_locations.append({
                "name":     _lbl,
                "district": _disp_dist,
                "gps":      f"{round(_coord[0],6)}, {round(_coord[1],6)}" if _coord else "",
                "count":    int(_cnt),
                "address":  str(_addr)[:70],
            })
            if len(top_locations) >= 3:
                break
    except Exception:
        logger.debug('Suppressed exception', exc_info=True)

    return {
        "home_district":home_dist,"work_district":None,
        "home_coord":home_coord,"home_label":home_label,
        "base_districts":list(set(filter(None,[home_dist]))),
        "trips":trips,"gaps":gaps,"total_days":total_days,
        "out_of_home_days":sum(t["days"] for t in trips),
        "total_records":len(df),
        "date_from":str(all_dates[0]) if all_dates else "N/A",
        "date_to":str(all_dates[-1]) if all_dates else "N/A",
        "top_locations": top_locations,
    }

def _loc_accuracy_html(df):
    """Show GPS accuracy badge in location section."""
    if "loc_method" not in df.columns:
        return ""
    exact = int((df["loc_method"] == "cell_exact").sum())
    total = len(df)
    pct = round(exact / max(total, 1) * 100, 1)
    color = "#059669" if pct >= 70 else "#d97706" if pct >= 30 else "#dc2626"
    return f"""<div style="background:#f0fdf4; border-left:4px solid {color}; border-radius:8px;
                    padding:0.6rem 1.1rem; margin-bottom:0.75rem; font-size:0.88rem; color:#065f46;">
        <strong>📡 GPS Accuracy:</strong> {exact:,}/{total:,} records ({pct}%) matched to
        exact cell tower coordinates (±0.5–2 km).
        Remaining {total-exact:,} records use BTS address text-based location (±5–20 km).
    </div>"""


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

    # ── Top 3 Frequent Locations ──────────────────────────────────────────
    # Home (nighttime) + Top 2 other frequent locations
    freq_locs_html = ""
    if mv.get("top_locations"):
        top3 = mv["top_locations"][:3]
        icons = ["🏠", "📍", "📍"]
        labels = ["Home (Most Frequent)", "2nd Frequent Location", "3rd Frequent Location"]
        colors = ["#16a34a", "#2563eb", "#7c3aed"]
        cards = ""
        for i, loc in enumerate(top3):
            loc_name  = str(loc.get("name") or loc.get("label") or "Unknown")
            loc_dist  = str(loc.get("district") or "")
            loc_gps   = loc.get("gps") or ""
            loc_count = loc.get("count") or loc.get("records") or ""
            gps_tag   = f"<div style='font-family:monospace;font-size:0.78rem;color:#475569;margin-top:0.3rem;'>📡 {loc_gps}</div>" if loc_gps else ""
            dist_tag  = f"<div style='font-size:0.78rem;color:#64748b;'>{loc_dist} District</div>" if loc_dist else ""
            cnt_tag   = f"<div style='font-size:0.75rem;color:#94a3b8;margin-top:0.25rem;'>{loc_count} records</div>" if loc_count else ""
            cards += f"""
            <div style="background:white; border-top:4px solid {colors[i]}; border-radius:10px;
                        padding:1rem 1.1rem; box-shadow:0 1px 4px rgba(0,0,0,0.07); flex:1; min-width:0;">
                <div style="font-size:0.72rem; color:#94a3b8; font-weight:700;
                            text-transform:uppercase; letter-spacing:0.05em;">{icons[i]} {labels[i]}</div>
                <div style="font-size:1rem; font-weight:800; color:#0f172a;
                            margin:0.35rem 0; white-space:nowrap; overflow:hidden;
                            text-overflow:ellipsis;" title="{loc_name}">{loc_name}</div>
                {dist_tag}{gps_tag}{cnt_tag}
            </div>"""
        freq_locs_html = f"""
        <div style="margin-bottom:1.25rem;">
            <div style="font-size:0.85rem; font-weight:700; color:#374151;
                        margin-bottom:0.6rem;">📊 Most Frequent Locations</div>
            <div style="display:flex; gap:0.9rem; flex-wrap:wrap;">
                {cards}
            </div>
        </div>"""
    trips_html = ""
    if mv["trips"]:
        def _stay_detail(stay_locs):
            if len(stay_locs) <= 1: return ""
            parts = []
            for sl in stay_locs[:-1]:
                lbl = str(sl.get("upazila") or sl.get("district") or "?").strip() or "?"
                parts.append(f"{lbl} ({sl['hours']}h, {sl['km']}km)")
            return "Via: " + " -> ".join(parts) if parts else ""

        rows = "".join([
            f"""<tr style="background:{'#f8fafc' if i%2==0 else 'white'};">
                <td style="padding:0.7rem 1rem; font-weight:600; color:#1e3a8a;">
                    {i+1}</td>
                <td style="padding:0.7rem 1rem; font-weight:700;">
                    {t["district"] or t["upazila"]}</td>
                <td style="padding:0.7rem 1rem; text-align:center; font-family:monospace; font-size:0.82rem; color:#1e3a8a;">
                    {"✅ " + str(round(float(t["lat"]),6)) + "<br>" + str(round(float(t["lon"]),6)) if t.get("lat") else "—"}
                </td>
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
                    <th style="padding:0.7rem 1rem; text-align:center;">GPS Coordinates</th>
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

    return cards_html + freq_locs_html + trips_html + gaps_html

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

def target_location_analysis(df, target_location):
    """
    Target Location Analysis:
    Enter district or upazila name to see CDR activity dates in that area.
    Fuzzy match: partial name matching।
    Returns: list of dicts [{date, address, usage_type, count, lat, lon, gps_coord}]
    """
    if not target_location or "address" not in df.columns or "start" not in df.columns:
        return []

    tgt = target_location.strip().lower()
    # Remove common suffixes for broader matching
    tgt_clean = re.sub(r'(?i)[ ]*(sadar|district|upazila|thana|zila)[ ]*$', '', tgt).strip()

    results = []
    addr_df = df[df["address"].notna() & df["address"].apply(_is_valid_address)].copy()
    if addr_df.empty:
        return []

    addr_df["date"] = addr_df["start"].dt.date
    addr_df["addr_lower"] = addr_df["address"].str.lower()

    # Match: address contains target string (partial match)
    matched = addr_df[
        addr_df["addr_lower"].str.contains(tgt_clean, na=False, regex=False) |
        addr_df["addr_lower"].str.contains(tgt, na=False, regex=False)
    ]

    if matched.empty:
        return []

    # Group by date — each date: count, sample address, GPS if available
    has_gps = "cell_lat" in matched.columns and "cell_lon" in matched.columns

    daily = []
    for date, grp in matched.groupby("date"):
        cnt = len(grp)
        _addr_vc = grp["address"].value_counts()
        sample_addr = _addr_vc.index[0] if not _addr_vc.empty else '—'
        usage_types = grp["usage_type"].value_counts().to_dict() if "usage_type" in grp.columns else {}
        usage_str = ", ".join(f"{k}:{v}" for k, v in usage_types.items())

        gps_coord = "—"
        if has_gps:
            exact = grp[grp["loc_method"] == "cell_exact"] if "loc_method" in grp.columns else pd.DataFrame()
            src_gps = exact if not exact.empty else grp
            lats = src_gps["cell_lat"].dropna()
            if not lats.empty:
                lat = round(float(lats.iloc[0]), 6)
                lon = round(float(src_gps["cell_lon"].dropna().iloc[0]), 6)
                gps_coord = f"{lat}, {lon}"

        daily.append({
            "Date":            str(date),
            "BTS Address":     sample_addr[:80],
            "GPS Coordinates": gps_coord,
            "Usage Type":      usage_str,
            "Records":         cnt,
        })

    daily.sort(key=lambda x: x["Date"])
    return daily


def _target_location_html(df, target_location, sec_num=11):
    """HTML section for target location analysis."""
    if not target_location:
        return ""
    results = target_location_analysis(df, target_location)
    if not results:
        return f"""<h2>{sec_num}. Target Location Analysis</h2>
    <div class="info-box">
        <strong>Target Location:</strong> {html_safe(target_location)}<br>
        <span class="warning">No CDR activity found near '{html_safe(target_location)}'.</span>
    </div>"""

    rows_html = "".join([
        f"""<tr style="background:{'#EBF3FB' if i%2==0 else 'white'};">
            <td>{r['Date']}</td>
            <td>{r['BTS Address']}</td>
            <td style="font-family:monospace;font-size:0.82em;">{r['GPS Coordinates']}</td>
            <td>{r['Usage Type']}</td>
            <td style="text-align:center;font-weight:700;">{r['Records']}</td>
        </tr>"""
        for i, r in enumerate(results)
    ])

    return f"""<h2>{sec_num}. Target Location Analysis</h2>
    <div class="info-box">
        <strong>Target Location:</strong> {html_safe(target_location)}<br>
        <strong>Total Days with Activity:</strong> {len(results)} day(s)<br>
        <strong>Total Records:</strong> {sum(r['Records'] for r in results)}
    </div>
    <p>Dates when the subscriber's CDR activity was detected near <strong>{html_safe(target_location)}</strong>:</p>
    <table class="dataframe">
        <thead>
            <tr style="text-align:right;">
                <th>Date</th>
                <th>BTS Address</th>
                <th>GPS Coordinates</th>
                <th>Usage Type</th>
                <th>Records</th>
            </tr>
        </thead>
        <tbody>{rows_html}</tbody>
    </table>"""


def _target_number_html(df, target_number, sec_num=10):
    if not target_number:
        return ''
    res = specific_number_analysis(df, target_number)
    if not res:
        return f'''<h2>{sec_num}. Specific Number Analysis</h2>
        <div class="info-box">
        <strong>Target Number:</strong> {html_safe(target_number)}<br>
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
    return f'''<h2>{sec_num}. Specific Number Analysis</h2>
    <p>Communication summary between the subscriber and <strong>{html_safe(target_number)}</strong>.</p>
    {df_to_html(table_df)}'''



# ─────────────────────────────────────────────
# BURST & INTERNATIONAL ANALYSIS
# ─────────────────────────────────────────────

def burst_analysis(df, window_min=60, min_calls=15, daily_mult=4, max_show=10):
    """
    CDR burst detection — দুটো method একসাথে:
      1. Rolling Window: T→T+60min এ ≥15 call → 🔴 Rolling
      2. Daily Anomaly: দিনের total > daily_avg × 4 → 📊 Daily
      উভয় → ⚡ Both
    Returns list of event dicts or empty list (if no burst found).
    """
    if 'start' not in df.columns or df['start'].isna().all():
        return []

    # Only call rows (not SMS)
    CALL_TYPES = {'moc','mo','outgoing','call-mo','callmo',
                  'mtc','mt','incoming','call-mt','callmt','rcf'}
    call_df = df[df.get('usage_type',df.get('ut',pd.Series())).str.lower().str.strip().isin(CALL_TYPES)]
    if len(call_df) < min_calls:
        return []

    call_df = call_df[call_df['start'].notna()].sort_values('start').reset_index(drop=True)
    if len(call_df) < min_calls:
        return []

    # Daily counts & threshold
    daily_counts = call_df.groupby(call_df['start'].dt.date).size()
    if len(daily_counts) == 0:
        return []
    daily_avg   = float(daily_counts.mean())
    daily_thresh = daily_avg * daily_mult

    pb_col = 'party_b_norm' if 'party_b_norm' in call_df.columns else (
             'party_b_clean' if 'party_b_clean' in call_df.columns else
             'party_b' if 'party_b' in call_df.columns else None)

    times = call_df['start'].tolist()
    pbs   = call_df[pb_col].tolist() if pb_col else ['—'] * len(call_df)

    # Rolling window burst detection
    raw_bursts = []
    for i, t in enumerate(times):
        window_end = t + pd.Timedelta(minutes=window_min)
        idx = [j for j, t2 in enumerate(times) if t <= t2 <= window_end]
        if len(idx) >= min_calls:
            day_count = int(daily_counts.get(t.date(), 0))
            top_pb    = pd.Series([pbs[j] for j in idx]).value_counts()
            top_num   = top_pb.index[0] if not top_pb.empty else '—'
            raw_bursts.append({
                'start': t, 'end': times[idx[-1]],
                'count': len(idx), 'day_count': day_count,
                'top_contact': str(top_num)[:20],
            })

    if not raw_bursts:
        # Check daily anomaly alone (no rolling burst)
        anomaly_days = daily_counts[daily_counts > daily_thresh]
        results = []
        for day, cnt in anomaly_days.items():
            results.append({
                'Date & Time':       str(day),
                'Flag':              '📊 Daily',
                'Calls in Window':   f"(day: {cnt})",
                'Intensity':         f"{cnt/max(daily_avg,1):.1f}×",
                'Top Contact':       '—',
                'Daily Avg / Limit': f"avg: {daily_avg:.0f} · limit: {daily_thresh:.0f}",
            })
        return sorted(results, key=lambda x: x['Intensity'], reverse=True)[:max_show]

    # Deduplicate: merge events within 30-min gap
    deduped = []
    for b in raw_bursts:
        if not deduped or (b['start'] - deduped[-1]['start']).total_seconds() > 1800:
            deduped.append(b)
        elif b['count'] > deduped[-1]['count']:
            deduped[-1] = b

    # Build result rows
    results = []
    for b in deduped[:max_show]:
        rolling_flag = b['count'] >= min_calls
        daily_flag   = b['day_count'] > daily_thresh
        badge = '⚡ Both' if (rolling_flag and daily_flag) else (
                '🔴 Rolling' if rolling_flag else '📊 Daily')
        intensity = round(b['count'] / max(daily_avg, 1), 1)
        results.append({
            'Date & Time':       (f"{b['start']:%Y-%m-%d %a}"
                                  f" · {b['start']:%H:%M}–{b['end']:%H:%M}"),
            'Flag':              badge,
            'Calls in Window':   f"{b['count']} (day: {b['day_count']})",
            'Intensity':         f"{intensity}×",
            'Top Contact':       b['top_contact'],
            'Daily Avg / Limit': f"avg: {daily_avg:.0f} · limit: {daily_thresh:.0f}",
        })
    return results



# ─────────────────────────────────────────────────────────────────────────────
# FOREIGN VOICE CALL DETECTION — Updated Logic
# শুধুমাত্র Voice Call (MOC/MTC) — SMS বাদ।
# Raw E.164 format (+ বা 00 prefix ছাড়া) সাপোর্ট।
# Middle East, Pakistan, Afghanistan সহ সকল দেশ।
# ─────────────────────────────────────────────────────────────────────────────

_VOICE_TYPES_INTL = {'MOC', 'MTC'}

# (country_name, prefix_digits, min_total_len, max_total_len)
# ৩-digit prefix আগে — greedy match এর জন্য collision এড়ানো
_FOREIGN_PHONE_TABLE = [
    # ── Middle East (Gulf) ──────────────────────────────
    ('🇸🇦 Saudi Arabia',   '966', 12, 12),
    ('🇦🇪 UAE',            '971', 12, 12),
    ('🇰🇼 Kuwait',         '965', 11, 11),
    ('🇶🇦 Qatar',          '974', 11, 11),
    ('🇧🇭 Bahrain',        '973', 11, 11),
    ('🇴🇲 Oman',           '968', 11, 12),
    ('🇾🇪 Yemen',          '967', 12, 12),
    # ── Middle East (Levant / Mashreq) ──────────────────
    ('🇯🇴 Jordan',         '962', 12, 12),
    ('🇸🇾 Syria',          '963', 12, 12),
    ('🇱🇧 Lebanon',        '961', 11, 11),  # mobile only (70x,71x,76x,78x,79x,81x) — landline (10 digit) বাদ
    ('🇮🇶 Iraq',           '964', 12, 13),
    ('🇵🇸 Palestine',      '970', 12, 12),
    ('🇮🇱 Israel',         '972', 12, 12),
    # ── North Africa ────────────────────────────────────
    ('🇱🇾 Libya',          '218', 12, 12),
    ('🇹🇳 Tunisia',        '216', 11, 11),
    ('🇩🇿 Algeria',        '213', 12, 12),
    ('🇲🇦 Morocco',        '212', 12, 12),
    ('🇸🇩 Sudan',          '249', 12, 12),
    # ── Nigeria / South Africa ──────────────────────────
    ('🇳🇬 Nigeria',        '234', 13, 13),
    ('🇿🇦 South Africa',   '27',  11, 11),
    # ── South Asia ──────────────────────────────────────
    ('🇮🇳 India',          '91',  12, 12),
    ('🇵🇰 Pakistan',       '92',  12, 12),
    ('🇦🇫 Afghanistan',    '93',  11, 12),
    ('🇮🇷 Iran',           '98',  12, 12),
    ('🇱🇰 Sri Lanka',      '94',  11, 11),
    ('🇳🇵 Nepal',          '977', 12, 12),
    ('🇧🇹 Bhutan',         '975', 11, 11),
    ('🇲🇲 Myanmar',        '95',  11, 12),  # mobile only — landline (9-10 digit) বাদ
    # ── Europe ──────────────────────────────────────────
    ('🇹🇷 Turkey',         '90',  12, 12),
    ('🇬🇧 United Kingdom', '44',  12, 12),
    ('🇩🇪 Germany',        '49',  11, 12),
    ('🇫🇷 France',         '33',  11, 11),
    ('🇸🇪 Sweden',         '46',  11, 11),
    ('🇮🇹 Italy',          '39',  11, 12),
    ('🇪🇸 Spain',          '34',  11, 11),
    # ── East / Southeast Asia ───────────────────────────
    ('🇨🇳 China',          '86',  13, 13),
    ('🇯🇵 Japan',          '81',  11, 11),
    ('🇰🇷 South Korea',    '82',  11, 12),
    ('🇲🇾 Malaysia',       '60',  11, 12),
    ('🇦🇺 Australia',      '61',  11, 11),
    ('🇮🇩 Indonesia',      '62',  11, 12),
    ('🇸🇬 Singapore',      '65',  10, 10),
    ('🇹🇭 Thailand',       '66',  11, 11),
    ('🇵🇭 Philippines',    '63',  11, 12),
    # ── Americas / Russia (1-digit — সবার শেষে) ─────────
    ('🇺🇸 USA/Canada',     '1',   11, 11),
    ('🇷🇺 Russia',         '7',   11, 11),
]


def _is_bd_number(digits: str) -> bool:
    """Bangladesh mobile ও PSTN number চেনা।"""
    if digits.startswith('880'):
        return True
    # local format: 01XXXXXXXXX (11 digits)
    if digits.startswith('01') and len(digits) == 11:
        return True
    return False


def _detect_foreign_voice_country(party_b_raw: str, usage_type_raw: str):
    """
    Voice call (MOC/MTC) হলে foreign country name return করে।
    SMS হলে — None (check করা হয় না)।
    Raw E.164 format (447..., 917..., 1347...) সাপোর্ট করে।

    Returns: str (country name with flag) অথবা None
    """
    # ── শুধু Voice call ──────────────────────────────────
    if str(usage_type_raw).strip().upper() not in _VOICE_TYPES_INTL:
        return None

    digits = re.sub(r'[^0-9]', '', str(party_b_raw).strip())

    # ── খুব ছোট, খালি বা clearly invalid ───────────────
    if len(digits) < 7:
        return None

    # ── Bangladesh number বাদ ────────────────────────────
    if _is_bd_number(digits):
        return None

    # ── '00' prefix normalize: 0091XXX → 91XXX ──────────
    # '0088' = BD, বাদ দেওয়া হয়েছে
    if digits.startswith('00') and not digits.startswith('0088'):
        digits = digits[2:]

    # ── Operator routing / SMSC prefix বাদ ───────────────
    # 475... = GP SMSC prefix, numeric only
    if digits.startswith('475') and len(digits) > 14:
        return None

    # ── Country prefix match ─────────────────────────────
    # ৩-digit আগে, তারপর ২-digit, তারপর ১-digit
    for country, prefix, min_len, max_len in _FOREIGN_PHONE_TABLE:
        if digits.startswith(prefix) and min_len <= len(digits) <= max_len:
            return country

    return None


def international_call_analysis(df):
    """
    Foreign VOICE CALL detection — শুধু MOC/MTC।
    SMS বাদ (promotional/OTP false positive এড়ানো)।
    Raw E.164 format সাপোর্ট।
    Returns dict with summary + per-country table, or None if no intl found.
    """
    # ── Column lookup ────────────────────────────────────────────────────────
    # party_b_norm বাদ — last-10-digit truncate করে foreign prefix হারায়।
    # party_b_clean বাদ — Party B Original (routing/callback number) দিয়ে
    #   overwrite হয়, তাই Indian/foreign নম্বর হারিয়ে যায়।
    # party_b (raw) সবার আগে — CDR-এর actual dialed/received number।

    raw_pb_col = next((c for c in ['party_b', 'Party B'] if c in df.columns), None)
    dis_pb_col = next((c for c in ['party_b_clean', 'party_b', 'Party B'] if c in df.columns), None)
    ut_col     = next((c for c in ['usage_type', 'ut', 'Usage Type'] if c in df.columns), None)

    if raw_pb_col is None or ut_col is None:
        return None

    tmp = df[df['start'].notna()].copy()

    # Detection: raw party_b দিয়ে (original dialed number)
    # Display: party_b_clean দিয়ে (formatted, but raw হলে raw)
    tmp['_country'] = tmp.apply(
        lambda r: _detect_foreign_voice_country(r[raw_pb_col], r[ut_col]),
        axis=1
    )
    # Display column — groupby-তে raw number দেখাবে
    tmp['_pb_display'] = tmp[raw_pb_col].astype(str).str.strip()

    intl = tmp[tmp['_country'].notna()].copy()
    if intl.empty:
        return None

    # Duration column
    dur_col = next((c for c in ['duration', 'call_duration', 'dur', 'Call Duration']
                    if c in intl.columns), None)
    intl['_dur'] = pd.to_numeric(intl[dur_col], errors='coerce').fillna(0) if dur_col else 0

    intl['_ut'] = intl[ut_col].astype(str).str.upper().str.strip().fillna('')

    # Aggregate per country + number (_pb_display = raw party_b)
    grp = intl.groupby(['_country', '_pb_display']).agg(
        Calls     = ('start', 'count'),
        Duration  = ('_dur', lambda x: int(x.sum()) // 60),
        Direction = ('_ut', lambda x:
                     'MOC' if (x == 'MOC').all() else
                     'MTC' if (x == 'MTC').all() else 'Both'),
        First     = ('start', lambda x: str(x.min())[:10]),
        Last      = ('start', lambda x: str(x.max())[:10]),
    ).reset_index()

    grp.columns = ['Country', 'Number', 'Calls', 'Duration (min)',
                   'Direction', 'First Contact', 'Last Contact']
    grp = grp.sort_values('Calls', ascending=False).head(30)

    unique_countries = intl['_country'].nunique()
    unique_numbers   = intl['_pb_display'].nunique()
    total_calls      = len(intl)
    total_dur        = int(intl['_dur'].sum()) // 60

    return {
        'total_calls':      total_calls,
        'unique_countries': unique_countries,
        'unique_numbers':   unique_numbers,
        'total_duration':   total_dur,
        'table':            grp,
    }


def _burst_html(burst_results, sec_num=10):
    """Burst analysis HTML section."""
    if not burst_results:
        return ''
    both  = sum(1 for r in burst_results if '⚡' in r['Flag'])
    roll  = sum(1 for r in burst_results if '🔴' in r['Flag'])
    daily = sum(1 for r in burst_results if '📊' in r['Flag'])

    rows = ''.join(
        f"""<tr style="background:{'#FFF3CD' if '⚡' in r['Flag'] else
                                    '#FDDEDE' if '🔴' in r['Flag'] else '#F0F0FF'}">
            <td>{r['Date & Time']}</td>
            <td style="text-align:center;font-weight:700">{r['Flag']}</td>
            <td style="text-align:center">{r['Calls in Window']}</td>
            <td style="text-align:center;font-weight:700;color:#c00">{r['Intensity']}</td>
            <td>{html_safe(r['Top Contact'])}</td>
            <td style="color:#666;font-size:0.85em">{r['Daily Avg / Limit']}</td>
        </tr>"""
        for r in burst_results
    )
    return f"""<h2>{sec_num}. Burst Activity Detection</h2>
    <div class="info-box">
        ⚡ <strong>{len(burst_results)} Burst Event(s) Detected</strong>
        &nbsp;|&nbsp; Rolling Window (60 min / 15+ calls) &amp; Daily Anomaly (4× avg)
        &nbsp;&nbsp;
        <span style="color:#c00;font-weight:700">{both} Both</span> &nbsp;
        <span style="color:#e00">{roll} Rolling</span> &nbsp;
        <span style="color:#44f">{daily} Daily</span>
    </div>
    <table>
    <tr><th>Date &amp; Time</th><th>Flag</th><th>Calls in Window</th>
        <th>Intensity</th><th>Top Contact</th><th>Daily Avg / Limit</th></tr>
    {rows}
    </table>"""


def _intl_html(intl, sec_num=11):
    """International call analysis HTML section."""
    if not intl:
        return ''
    tbl = intl['table']
    rows = ''.join(
        f"""<tr style="background:{'#f8fafc' if i%2==0 else 'white'}">
            <td>{html_safe(r['Country'])}</td>
            <td style="font-family:monospace">{html_safe(r['Number'])}</td>
            <td style="text-align:center">{r['Direction']}</td>
            <td style="text-align:center;font-weight:700">{r['Calls']}</td>
            <td style="text-align:center">{r['Duration (min)']}</td>
            <td>{r['First Contact']}</td>
            <td>{r['Last Contact']}</td>
        </tr>"""
        for i, (_, r) in enumerate(tbl.iterrows())
    )
    return f"""<h2>{sec_num}. International Call Analysis</h2>
    <div class="info-box">
        🌐 <strong>International Activity Detected</strong>
        &nbsp;&nbsp;
        Total: <strong>{intl['total_calls']}</strong> records &nbsp;|&nbsp;
        Countries: <strong>{intl['unique_countries']}</strong> &nbsp;|&nbsp;
        Unique Numbers: <strong>{intl['unique_numbers']}</strong> &nbsp;|&nbsp;
        Duration: <strong>{intl['total_duration']} min</strong>
    </div>
    <table>
    <tr><th>Country</th><th>Number</th><th>Direction</th>
        <th>Calls</th><th>Duration (min)</th>
        <th>First Contact</th><th>Last Contact</th></tr>
    {rows}
    </table>"""


def build_html(df, phone, operator, date_range, total_raw, anomaly_count, target_number=None, target_location=None, profile_data=None):
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

    # ── Pre-compute optional sections (data থাকলেই দেখাবে) ──────────────
    _burst_data   = burst_analysis(df)
    _intl_data    = international_call_analysis(df)
    _has_burst    = bool(_burst_data)
    _has_intl     = bool(_intl_data)

    # Dynamic section counter — conditional sections skip করলে gap হবে না
    _s  = [9]  # sections 1-9 fixed
    def _sec():
        _s[0] += 1
        return _s[0]

    _sec_burst  = _sec() if _has_burst else None
    _sec_intl   = _sec() if _has_intl  else None
    _sec_tnum   = _sec() if target_number   else None
    _sec_tloc   = _sec() if target_location else None
    _sec_oc     = _sec()   # Overall Comment
    _sec_rec    = _sec()   # Recommendation

    return f"""<html><head><title>CDR Analysis Report</title>
    <style>{CSS}</style></head><body>
    <h1>📞 CDR Analysis Report</h1>
    <h2>1. Executive Summary</h2>
    <div class="info-box">
    <strong>{"Device IMEI / SIM(s)" if is_imei_cdr(df) else "Phone Number"}:</strong> {html_safe(phone)}<br>
    <strong>Operator:</strong> {html_safe(operator)}<br>
    <strong>Analysis Period:</strong> {html_safe(date_range)}<br>
    <strong>Total Raw Records:</strong> {total_raw:,}<br>
    <strong>Anomalies Removed:</strong> {anomaly_count:,}<br>
    <strong>Records Analyzed:</strong> {len(df):,}
    </div>
    <h2>2. Device Information</h2>
    <table><tr><th>Field</th><th>Value</th></tr>
    <tr><td>IMEI</td><td>{', '.join(str(i) for i in imei) if imei else 'N/A'}</td></tr>
    <tr><td>IMSI</td><td>{', '.join(str(i) for i in imsi) if imsi else 'N/A'}</td></tr>
    <tr><td>{"Device IMEI / SIM(s)" if is_imei_cdr(df) else "Phone Number"}</td><td>{html_safe(phone)}</td></tr></table>
    {_imsi_change_html(df)}
    {_imei_change_html(df)}
    <h2>3. Call Analysis</h2>
    <h3>3.1 Call Analysis Summary</h3>{df_to_html(call_summary(df))}
    <h2>4. Call Count Analysis</h2>
    <h3>4.1 Daily Call Count</h3>{df_to_html(daily_call_count(df))}
    <h3>4.2 Hourly Call Count Graph</h3>{fig_to_html_img(plot_hourly(df))}
    <h3>4.3 Weekly Call Count</h3>{df_to_html(weekly_call_count(df))}
    <h3>4.4 Monthly Call Count</h3>{df_to_html(monthly_call_count(df))}
    <h2>5. Contact Analysis</h2>
    <h3>5.1 Contact Summary</h3>{df_to_html(contact_summary(df))}
    <h3>5.2 Top 10 Frequent Outgoing</h3>{df_to_html(top_contacts(df,'out',10))}
    <h3>5.3 Top 10 Frequent Incoming</h3>{df_to_html(top_contacts(df,'in',10))}
    <h3>5.4 Top 10 Lengthy Outgoing</h3>{df_to_html(top_lengthy(df,'out',10))}
    <h3>5.5 Top 10 Lengthy Incoming</h3>{df_to_html(top_lengthy(df,'in',10))}
    <h3>5.6 Top Call Overall</h3>{df_to_html(top_call_overall(df,10))}
    <h3>5.6a Top Call Overall Chart</h3>{fig_to_html_img(plot_top_call_overall(df,10))}
    <h2>6. Location Analysis</h2>
    {_loc_accuracy_html(df)}
    <h3>6.1 Location Summary</h3>{df_to_html(location_summary(df))}
    <h3>6.2 Top 10 Frequent Locations</h3>{df_to_html(top_locations(df,None,10))}
    <h3>6.4 Possible Home Locations</h3>{df_to_html(top_locations(df,home_mask,10))}
    <h3>6.6 Possible Work Locations</h3>{df_to_html(top_locations(df,work_mask,10))}
    <h3>6.8 Possible Weekend Locations</h3>{df_to_html(top_locations(df,weekend_mask,10))}

    <h2>7. SMS Contact Analysis</h2>
    <h3>7.1 Top 5 Sent SMS Contacts</h3>{df_to_html(top_sms_contacts(df,'out',5))}
    <h3>7.2 Top 5 Received SMS Contacts</h3>{df_to_html(top_sms_contacts(df,'in',5))}

    <h2>8. Last 10 Days Analysis</h2>
    <h3>8.1 Top Contacts in Last 10 Days (MOC + MTC)</h3>{df_to_html(last_n_days_top_contacts(df, 10, 10))}
    <h3>8.2 Top Locations in Last 10 Days</h3>{df_to_html(last_n_days_top_locations(df, 10, 10))}

    <h2>9. Movement Pattern Analysis</h2>
    <p>Analysis of movement outside estimated home/work district and network disconnection periods.</p>
    {_movement_html(movement_pattern_analysis(df))}

    {_burst_html(_burst_data, _sec_burst) if _has_burst else ''}
    {_intl_html(_intl_data, _sec_intl)   if _has_intl  else ''}

    {_target_number_html(df, target_number, _sec_tnum)     if target_number   else ''}
    {_target_location_html(df, target_location, _sec_tloc) if target_location else ''}

    <h2>{_sec_oc}. Overall Comment</h2>
    {''.join(f'<p>{ln}</p>' for ln in generate_overall_comment(df, phone, operator, date_range, total_raw))}
    <h2>{_sec_rec}. Recommendation</h2>
    <ol>{''.join(f'<li>{r}</li>' for r in generate_recommendation(df))}</ol>
    <hr><p style="text-align:center;color:gray;font-size:11px;">
    Report generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | CDR Analysis Tool v1.0</p>
    </body></html>"""


# ─────────────────────────────────────────────
# WORD (DOCX) GENERATOR
# ─────────────────────────────────────────────
def build_docx(df, phone, operator, date_range, total_raw, anomaly_count, target_number=None, target_location=None, profile_data=None):

    doc = _DocxDocument()

    # ── Page Setup: A4, Calibri, proper margins ───────────────────────────
    sec = doc.sections[0]
    sec.page_width   = _DocxCm(21.0)
    sec.page_height  = _DocxCm(29.7)
    sec.top_margin   = _DocxCm(2.54)
    sec.bottom_margin= _DocxCm(2.54)
    sec.left_margin  = _DocxCm(2.54)
    sec.right_margin = _DocxCm(2.54)

    # Default font: Calibri 11pt
    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = _DocxPt(11)
    from docx.oxml.ns import qn as _qn2
    doc.styles['Normal'].element.rPr.rFonts.set(_qn2('w:ascii'),   'Calibri')
    doc.styles['Normal'].element.rPr.rFonts.set(_qn2('w:hAnsi'),   'Calibri')
    doc.styles['Normal'].element.rPr.rFonts.set(_qn2('w:eastAsia'),'Calibri')

    FONT      = 'Calibri'
    C_NAVY    = _DocxRGBColor(0x1F, 0x38, 0x64)
    C_BLUE    = _DocxRGBColor(0x2E, 0x74, 0xB5)
    C_WHITE   = _DocxRGBColor(0xFF, 0xFF, 0xFF)
    C_GRAY    = _DocxRGBColor(0x80, 0x80, 0x80)
    C_RED     = _DocxRGBColor(0xDC, 0x26, 0x26)
    C_HDR_BG  = '2E74B5'
    C_ROW_ALT = 'EBF3FB'

    # ── Helper: set cell font ─────────────────────────────────────────────
    def _cell_font(cell, size=10, bold=False, color=None, bg=None):
        for para in cell.paragraphs:
            para.paragraph_format.space_before = _DocxPt(1)
            para.paragraph_format.space_after  = _DocxPt(1)
            for run in para.runs:
                run.font.name = FONT
                run.font.size = _DocxPt(size)
                run.bold = bold
                if color: run.font.color.rgb = color
        if bg:
            tc  = cell._tc
            tcPr= tc.get_or_add_tcPr()
            shd = _DocxOxmlElement('w:shd')
            shd.set(_docx_qn('w:fill'),  bg)
            shd.set(_docx_qn('w:color'), 'auto')
            shd.set(_docx_qn('w:val'),   'clear')
            tcPr.append(shd)

    # ── Helper: add heading ───────────────────────────────────────────────
    def add_h(text, lvl=1):
        p   = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = True
        run.font.name = FONT
        if lvl == 1:
            run.font.size      = _DocxPt(13)
            run.font.color.rgb = C_NAVY
            p.paragraph_format.space_before = _DocxPt(14)
            p.paragraph_format.space_after  = _DocxPt(4)
        else:
            run.font.size      = _DocxPt(11)
            run.font.color.rgb = C_BLUE
            p.paragraph_format.space_before = _DocxPt(8)
            p.paragraph_format.space_after  = _DocxPt(3)
        # Bottom border
        pPr  = p._p.get_or_add_pPr()
        pBdr = _DocxOxmlElement('w:pBdr')
        bot  = _DocxOxmlElement('w:bottom')
        bot.set(_docx_qn('w:val'),   'single')
        bot.set(_docx_qn('w:sz'),    '4')
        bot.set(_docx_qn('w:space'), '1')
        bot.set(_docx_qn('w:color'), '2E74B5' if lvl == 1 else 'BDD7EE')
        pBdr.append(bot); pPr.append(pBdr)

    # ── Helper: add DataFrame as table ────────────────────────────────────
    def add_df_table(data):
        if data is None or (hasattr(data, 'empty') and data.empty):
            p = doc.add_paragraph('No data available.')
            p.runs[0].font.name = FONT
            p.runs[0].font.size = _DocxPt(10)
            return
        cols = list(data.columns)
        tbl  = doc.add_table(rows=1, cols=len(cols))
        tbl.style = 'Table Grid'
        # Set equal column widths within A4 text area (~16cm)
        col_w_cm = 15.92 / len(cols)
        for i, col in enumerate(cols):
            c = tbl.rows[0].cells[i]
            c.text = str(col)
            c.width = _DocxCm(col_w_cm)
            _cell_font(c, size=9, bold=True, color=C_WHITE, bg=C_HDR_BG)
        for idx, row in data.iterrows():
            tr   = tbl.add_row()
            fill = C_ROW_ALT if idx % 2 == 0 else 'FFFFFF'
            for i, col in enumerate(cols):
                c = tr.cells[i]
                c.text  = str(row[col]) if pd.notna(row[col]) else ''
                c.width = _DocxCm(col_w_cm)
                _cell_font(c, size=9, bg=fill)
        doc.add_paragraph()

    # ── Helper: add chart image ───────────────────────────────────────────
    def add_fig(fig):
        if fig is None:
            p = doc.add_paragraph('Graph not available.')
            p.runs[0].font.name = FONT; return
        b64      = fig_to_base64(fig)
        img_data = base64.b64decode(b64)
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
            tmp.write(img_data); tmp_path = tmp.name
        doc.add_picture(tmp_path, width=_DocxCm(15.92))
        doc.paragraphs[-1].alignment = _DocxWdAlign.CENTER
        doc.add_paragraph()

    # ── Helper: add key-value info table ─────────────────────────────────
    def add_kv_table(rows_data):
        tbl = doc.add_table(rows=len(rows_data), cols=2)
        tbl.style = 'Table Grid'
        tbl.columns[0].width = _DocxCm(5.5)
        tbl.columns[1].width = _DocxCm(10.42)
        for i, (k, v) in enumerate(rows_data):
            fill = C_ROW_ALT if i % 2 == 0 else 'FFFFFF'
            c0 = tbl.rows[i].cells[0]; c0.text = str(k)
            _cell_font(c0, size=10, bold=True, bg=fill)
            c1 = tbl.rows[i].cells[1]; c1.text = str(v)
            _cell_font(c1, size=10, bg=fill)
        doc.add_paragraph()

    # ════════════════════════════════════════════════════════════════════
    # COVER / TITLE
    # ════════════════════════════════════════════════════════════════════
    t = doc.add_paragraph(); t.alignment = _DocxWdAlign.CENTER
    r = t.add_run('CDR ANALYSIS REPORT')
    r.bold = True; r.font.name = FONT
    r.font.size = _DocxPt(20); r.font.color.rgb = C_NAVY
    doc.add_paragraph()

    # ════════════════════════════════════════════════════════════════════
    # PROFILE ANALYSIS
    # ════════════════════════════════════════════════════════════════════
    _has_profile_d = bool(
        profile_data and (
            profile_data.get('name') or profile_data.get('nid') or
            profile_data.get('docs_found') or profile_data.get('photo_b64')
        )
    )
    if _has_profile_d:
        def _src_lbl(fk):
            sl = profile_data.get(f'{fk}_src', [])
            return f" [{sl[0][1]}]" if sl else ''
        add_h('Profile Analysis')
        _pf_rows = [
            ('নাম',           profile_data.get('name','')             + _src_lbl('name')),
            ('পিতা',          profile_data.get('father','')           + _src_lbl('father')),
            ('মাতা',          profile_data.get('mother','')           + _src_lbl('mother')),
            ('স্ত্রী/স্বামী', profile_data.get('spouse','')           + _src_lbl('spouse')),
            ('জন্মতারিখ',    profile_data.get('dob','')              + _src_lbl('dob')),
            ('লিঙ্গ',         profile_data.get('gender','')           + _src_lbl('gender')),
            ('পেশা',          profile_data.get('profession','')       + _src_lbl('profession')),
            ('রক্তের গ্রুপ',  profile_data.get('blood_group','')      + _src_lbl('blood_group')),
            ('মোবাইল',        profile_data.get('mobile','')           + _src_lbl('mobile')),
            ('NID',           profile_data.get('nid','')              + _src_lbl('nid')),
            ('Passport',      profile_data.get('passport','')         + _src_lbl('passport')),
            ('TIN',           profile_data.get('tin','')              + _src_lbl('tin')),
            ('ড্রাইভিং লাইসেন্স', profile_data.get('license_no','') + _src_lbl('license_no')),
            ('গাড়ি',          profile_data.get('vehicle_reg','')      + _src_lbl('vehicle_reg')),
            ('স্থায়ী ঠিকানা', profile_data.get('address_permanent','')+ _src_lbl('address_permanent')),
            ('বর্তমান ঠিকানা', profile_data.get('address_present','') + _src_lbl('address_present')),
            ('নথি', ', '.join(profile_data.get('docs_found', []))),
        ]
        add_kv_table([(k, v) for k, v in _pf_rows if v and v.strip()])
        if profile_data.get('mismatches'):
            add_h('তথ্য অসঙ্গতি (Mismatch)', 2)
            for _mm in profile_data['mismatches']:
                _mp  = doc.add_paragraph(style='List Bullet')
                _r2  = _mp.add_run(_mm)
                _r2.font.color.rgb = C_RED
                _r2.font.name = FONT
        doc.add_paragraph()

    # ════════════════════════════════════════════════════════════════════
    # 1. EXECUTIVE SUMMARY
    # ════════════════════════════════════════════════════════════════════
    add_h('1. Executive Summary')
    id_label = 'Device IMEI / SIM(s)' if is_imei_cdr(df) else 'Phone Number'
    add_kv_table([
        (id_label,         phone),
        ('Operator',       operator),
        ('Analysis Period',date_range),
        ('Total Raw Records', f'{total_raw:,}'),
        ('Anomalies Removed', f'{anomaly_count:,}'),
        ('Records Analyzed',  f'{len(df):,}'),
    ])

    # ════════════════════════════════════════════════════════════════════
    # 2. DEVICE INFORMATION
    # ════════════════════════════════════════════════════════════════════
    def _clean_ids(series):
        import re as _re2
        result = []
        for val in series.dropna().unique():
            d = _re2.sub(r'[^0-9]', '', str(val).strip())
            if len(d) >= 10: result.append(d)
        return sorted(set(result))

    imei = _clean_ids(df['imei']) if 'imei' in df.columns else []
    imsi = _clean_ids(df['imsi']) if 'imsi' in df.columns else []
    add_h('2. Device Information')
    add_kv_table([
        ('IMEI',         ', '.join(imei) if imei else 'N/A'),
        ('IMSI',         ', '.join(imsi) if imsi else 'N/A'),
        ('Phone Number', phone),
    ])

    # IMSI Change
    imsi_periods = imsi_change_analysis(df)
    if imsi_periods:
        add_h('2a. IMSI Change Analysis', 2)
        p = doc.add_paragraph(
            f'⚠ Multiple IMSI Detected! This SIM was used in {len(imsi_periods)} '
            f'different IMSI periods — indicating possible SIM swap or dual-SIM activity.'
        )
        p.runs[0].font.name = FONT; p.runs[0].font.size = _DocxPt(10)
        add_df_table(pd.DataFrame(imsi_periods))

    # IMEI Change
    imei_periods = imei_change_analysis(df)
    if imei_periods:
        add_h('2b. IMEI Change Analysis', 2)
        p = doc.add_paragraph(
            f'⚠ Multiple IMEI Detected! This number was used in {len(imei_periods)} '
            f'different devices — indicating possible handset change.'
        )
        p.runs[0].font.name = FONT; p.runs[0].font.size = _DocxPt(10)
        add_df_table(pd.DataFrame(imei_periods))

    # ════════════════════════════════════════════════════════════════════
    # 3. CALL ANALYSIS
    # ════════════════════════════════════════════════════════════════════
    add_h('3. Call Analysis')
    add_h('3.1 Call Analysis Summary', 2);  add_df_table(call_summary(df))

    # ════════════════════════════════════════════════════════════════════
    # 4. CALL COUNT ANALYSIS
    # ════════════════════════════════════════════════════════════════════
    add_h('4. Call Count Analysis')
    add_h('4.1 Daily Call Count', 2);       add_df_table(daily_call_count(df))
    add_h('4.2 Hourly Call Count Graph', 2);add_fig(plot_hourly(df))
    add_h('4.3 Weekly Call Count', 2);      add_df_table(weekly_call_count(df))
    add_h('4.4 Monthly Call Count', 2);     add_df_table(monthly_call_count(df))

    # ════════════════════════════════════════════════════════════════════
    # 5. CONTACT ANALYSIS
    # ════════════════════════════════════════════════════════════════════
    add_h('5. Contact Analysis')
    add_h('5.1 Contact Summary', 2);        add_df_table(contact_summary(df))
    add_h('5.2 Top 10 Frequent Outgoing', 2);add_df_table(top_contacts(df, 'out', 10))
    add_h('5.3 Top 10 Frequent Incoming', 2);add_df_table(top_contacts(df, 'in',  10))
    add_h('5.4 Top 10 Lengthy Outgoing', 2); add_df_table(top_lengthy(df, 'out', 10))
    add_h('5.5 Top 10 Lengthy Incoming', 2); add_df_table(top_lengthy(df, 'in',  10))
    add_h('5.6 Top Call Overall', 2);        add_df_table(top_call_overall(df, 10))
    add_h('5.6a Top Call Overall Chart', 2); add_fig(plot_top_call_overall(df, 10))

    # ════════════════════════════════════════════════════════════════════
    # 6. LOCATION ANALYSIS
    # ════════════════════════════════════════════════════════════════════
    home_mask    = df['start'].dt.hour.astype(int).isin(list(range(0,6))+list(range(22,24))) if 'start' in df.columns else None
    work_mask    = (df['start'].dt.hour.astype(int)>=8)&(df['start'].dt.hour.astype(int)<18)if 'start' in df.columns else None
    weekend_mask = df['start'].dt.dayofweek.astype(int).isin([4,5])                         if 'start' in df.columns else None

    add_h('6. Location Analysis')
    add_h('6.1 Location Summary', 2);       add_df_table(location_summary(df))
    add_h('6.2 Top 10 Frequent Locations', 2); add_df_table(top_locations(df, None, 10))
    add_h('6.4 Possible Home Locations', 2);   add_df_table(top_locations(df, home_mask, 10))
    add_h('6.6 Possible Work Locations', 2);   add_df_table(top_locations(df, work_mask, 10))
    add_h('6.8 Possible Weekend Locations', 2);add_df_table(top_locations(df, weekend_mask, 10))

    # ════════════════════════════════════════════════════════════════════
    # 7. SMS CONTACT ANALYSIS
    # ════════════════════════════════════════════════════════════════════
    add_h('7. SMS Contact Analysis')
    add_h('7.1 Top 5 Sent SMS Contacts', 2);     add_df_table(top_sms_contacts(df, 'out', 5))
    add_h('7.2 Top 5 Received SMS Contacts', 2); add_df_table(top_sms_contacts(df, 'in',  5))

    # ════════════════════════════════════════════════════════════════════
    # 8. LAST 10 DAYS ANALYSIS
    # ════════════════════════════════════════════════════════════════════
    add_h('8. Last 10 Days Analysis')
    add_h('8.1 Top Contacts in Last 10 Days (MOC + MTC)', 2)
    add_df_table(last_n_days_top_contacts(df, 10, 10))
    add_h('8.2 Top Locations in Last 10 Days', 2)
    add_df_table(last_n_days_top_locations(df, 10, 10))

    # ════════════════════════════════════════════════════════════════════
    # 9. MOVEMENT PATTERN ANALYSIS
    # ════════════════════════════════════════════════════════════════════
    add_h('9. Movement Pattern Analysis')
    mv = movement_pattern_analysis(df)
    if mv:
        home_lbl = mv.get('home_label') or mv.get('home_district') or 'N/A'
        add_kv_table([
            ('Home Location',          home_lbl),
            ('Home District',          mv.get('home_district') or 'N/A'),
            ('Total Days',             str(mv.get('total_days', 0))),
            ('Out-of-Home Trips',      str(len(mv.get('trips', [])))),
            ('Out-of-Home Days',       str(mv.get('out_of_home_days', 0))),
            ('Network Gaps (>4 days)', str(len(mv.get('gaps', [])))),
        ])
        if mv.get('trips'):
            add_h('9.1 Out-of-Home Travel (35km+ from Home)', 2)
            add_df_table(pd.DataFrame([{
                '#':             i + 1,
                'Destination':   t.get('upazila', ''),
                'District':      t.get('district', ''),
                'Distance (km)': t.get('km', 0),
                'Departure':     t.get('start_date', ''),
                'Return':        t.get('end_date', ''),
                'Days':          t.get('days', 0),
                'GPS':           f"{t.get('lat','')}, {t.get('lon','')}" if t.get('lat') else '',
                'BTS Location':  str(t.get('address', ''))[:80],
            } for i, t in enumerate(mv['trips'])]))
        if mv.get('gaps'):
            add_h('9.2 Network Disconnection Periods', 2)
            add_df_table(pd.DataFrame([{
                'Last Seen':  g['gap_start'],
                'Next Seen':  g['gap_end'],
                'Gap (days)': g['days'],
            } for g in mv['gaps']]))
    else:
        doc.add_paragraph('Insufficient location data for movement analysis.')

    # ════════════════════════════════════════════════════════════════════
    # 10. BURST ACTIVITY DETECTION
    # ════════════════════════════════════════════════════════════════════
    burst_res = burst_analysis(df)
    if burst_res:
        add_h('10. Burst Activity Detection')
        add_df_table(pd.DataFrame(burst_res))

    # ════════════════════════════════════════════════════════════════════
    # 11. INTERNATIONAL CALL ANALYSIS
    # ════════════════════════════════════════════════════════════════════
    intl_res = international_call_analysis(df)
    if intl_res:
        add_h('11. International Call Analysis')
        add_kv_table([
            ('Total Records',        str(intl_res['total_calls'])),
            ('Unique Countries',     str(intl_res['unique_countries'])),
            ('Unique Numbers',       str(intl_res['unique_numbers'])),
            ('Total Duration (min)', str(intl_res['total_duration'])),
        ])
        add_h('11.1 Details by Country / Number', 2)
        add_df_table(intl_res['table'])

    # ════════════════════════════════════════════════════════════════════
    # 12. SPECIFIC NUMBER ANALYSIS (optional)
    # ════════════════════════════════════════════════════════════════════
    if target_number:
        add_h('12. Specific Number Analysis')
        res = specific_number_analysis(df, target_number)
        if res is None:
            p = doc.add_paragraph(f'No communication found with target number: {target_number}.')
            p.runs[0].font.name = FONT
        else:
            add_kv_table([
                ('Target Number',            str(res['number'])),
                ('Outgoing Calls (MOC)',      str(res['moc'])),
                ('Incoming Calls (MTC)',      str(res['mtc'])),
                ('Total Calls',              str(res['total_calls'])),
                ('Total Duration (sec)',      f"{res['total_duration_sec']:,}"),
                ('Total Duration (min)',      f"{res['total_duration_min']:,}"),
                ('Sent SMS',                 str(res['sms_sent'])),
                ('Received SMS',             str(res['sms_received'])),
                ('Total SMS',                str(res['total_sms'])),
                ('First Contact',            res['first_contact']),
                ('Last Contact',             res['last_contact']),
            ])

    # ════════════════════════════════════════════════════════════════════
    # 13. TARGET LOCATION ANALYSIS (optional)
    # ════════════════════════════════════════════════════════════════════
    if target_location:
        add_h('13. Target Location Analysis')
        tloc_df = target_location_analysis(df, target_location)
        if tloc_df is not None and not tloc_df.empty:
            add_df_table(tloc_df)
        else:
            p = doc.add_paragraph(f'No activity found near "{target_location}".')
            p.runs[0].font.name = FONT

    # ════════════════════════════════════════════════════════════════════
    # 14. OVERALL COMMENT
    # ════════════════════════════════════════════════════════════════════
    add_h('14. Overall Comment')
    for ln in generate_overall_comment(df, phone, operator, date_range, total_raw):
        p = doc.add_paragraph(ln)
        p.runs[0].font.name = FONT
        p.runs[0].font.size = _DocxPt(10)
        p.paragraph_format.space_after = _DocxPt(4)

    # ════════════════════════════════════════════════════════════════════
    # 15. RECOMMENDATION
    # ════════════════════════════════════════════════════════════════════
    add_h('15. Recommendation')
    for i, rec in enumerate(generate_recommendation(df), 1):
        p = doc.add_paragraph(style='List Number')
        r = p.add_run(rec)
        r.font.name = FONT; r.font.size = _DocxPt(10)

    # ── Footer ────────────────────────────────────────────────────────
    doc.add_paragraph()
    fp = doc.add_paragraph(
        f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}  |  '
        f'CDR Intelligence Analysis Platform  |  Confidential'
    )
    fp.alignment = _DocxWdAlign.CENTER
    if fp.runs:
        fp.runs[0].font.size = _DocxPt(8)
        fp.runs[0].font.name = FONT
        fp.runs[0].font.color.rgb = C_GRAY

    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────






def _keyword_scan_gps(addr_upper):
    """
    Fallback GPS: P.S keyword ছাড়া full address-এ known area name scan করে।
    Returns (lat, lon, district, thana, radius_m, method) or None.
    ±2-3km accuracy — P.S parse-এর চেয়ে কম accurate কিন্তু miss কমায়।
    """
    # Sub-thana / area / mohalla level GPS for Bangladesh
    # Common in GP/Robi/BL BTS address without P.S: prefix
    BD_AREA_GPS = {
        # ── Dhaka ──
        'Shegunbagicha':  (23.7300, 90.4080, 'Dhaka', 'Ramna'),
        'Kakrail':        (23.7331, 90.4075, 'Dhaka', 'Ramna'),
        'Purana Paltan':  (23.7327, 90.4170, 'Dhaka', 'Motijheel'),
        'Paltan':         (23.7327, 90.4170, 'Dhaka', 'Motijheel'),
        'Dhanmondi':      (23.7463, 90.3762, 'Dhaka', 'Dhanmondi'),
        'Gulshan':        (23.7925, 90.4078, 'Dhaka', 'Gulshan'),
        'Banani':         (23.7937, 90.4066, 'Dhaka', 'Gulshan'),
        'Baridhara':      (23.8063, 90.4244, 'Dhaka', 'Gulshan'),
        'Uttara':         (23.8750, 90.3987, 'Dhaka', 'Uttara'),
        'Mirpur':         (23.8223, 90.3654, 'Dhaka', 'Mirpur'),
        'Mohammadpur':    (23.7638, 90.3567, 'Dhaka', 'Mohammadpur'),
        'Agargaon':       (23.7760, 90.3809, 'Dhaka', 'Sher-e-Bangla Nagar'),
        'Shyamoli':       (23.7742, 90.3564, 'Dhaka', 'Mohammadpur'),
        'Kalabagan':      (23.7529, 90.3820, 'Dhaka', 'Dhanmondi'),
        'Hatirjheel':     (23.7537, 90.4188, 'Dhaka', 'Rampura'),
        'Rampura':        (23.7639, 90.4178, 'Dhaka', 'Rampura'),
        'Badda':          (23.7800, 90.4300, 'Dhaka', 'Badda'),
        'Demra':          (23.7100, 90.4800, 'Dhaka', 'Demra'),
        'Jatrabari':      (23.7050, 90.4430, 'Dhaka', 'Jatrabari'),
        'Dania':          (23.7000, 90.4380, 'Dhaka', 'Jatrabari'),
        'Shampur':        (23.6900, 90.4900, 'Dhaka', 'Demra'),
        'Kodomtoli':      (23.6900, 90.4400, 'Dhaka', 'Kodomtoli'),
        'Shiddhirganj':   (23.6500, 90.5100, 'Narayanganj', 'Siddhirganj'),
        'Siddhirganj':    (23.6500, 90.5100, 'Narayanganj', 'Siddhirganj'),
        'Narayanganj':    (23.6238, 90.4997, 'Narayanganj', 'Narayanganj Sadar'),
        'Aminbazar':      (23.8300, 90.3100, 'Dhaka', 'Savar'),
        'Tejgaon':        (23.7600, 90.3900, 'Dhaka', 'Tejgaon'),
        'Farmgate':       (23.7582, 90.3887, 'Dhaka', 'Tejgaon'),
        'Karwan Bazar':   (23.7507, 90.3930, 'Dhaka', 'Tejgaon'),
        'Bangshal':       (23.7178, 90.4083, 'Dhaka', 'Bangshal'),
        'Nawabpur':       (23.7218, 90.4130, 'Dhaka', 'Kotwali'),
        'Sadarghat':      (23.7098, 90.4062, 'Dhaka', 'Kotwali'),
        'Wari':           (23.7179, 90.4258, 'Dhaka', 'Wari'),
        'Lalbagh':        (23.7205, 90.3888, 'Dhaka', 'Lalbagh'),
        'Hazaribagh':     (23.7239, 90.3864, 'Dhaka', 'Hazaribagh'),
        'Kamrangirchar':  (23.7063, 90.3699, 'Dhaka', 'Kamrangirchar'),
        'Postagola':      (23.7050, 90.4200, 'Dhaka', 'Kadamtali'),
        'Rayerbazar':     (23.7500, 90.3700, 'Dhaka', 'Mohammadpur'),
        'Manik Mia':      (23.7600, 90.3650, 'Dhaka', 'Sher-e-Bangla Nagar'),
        'Bijoy Nagar':    (23.7419, 90.4098, 'Dhaka', 'Ramna'),
        'Segunbagicha':   (23.7300, 90.4080, 'Dhaka', 'Ramna'),
        'Eskaton':        (23.7450, 90.4050, 'Dhaka', 'Ramna'),
        'Naya Paltan':    (23.7390, 90.4173, 'Dhaka', 'Motijheel'),
        'Dilkusha':       (23.7300, 90.4200, 'Dhaka', 'Motijheel'),
        # Gazipur
        'Tongi':          (23.8980, 90.3990, 'Gazipur', 'Tongi'),
        'Konabari':       (23.9300, 90.3600, 'Gazipur', 'Kaliakair'),
        'Chandra':        (23.9800, 90.2800, 'Gazipur', 'Kaliakair'),
        # ── Chittagong ──
        'Agrabad':        (22.3300, 91.8200, 'Chittagong', 'Chittagong Sadar'),
        'Khulshi':        (22.3700, 91.8200, 'Chittagong', 'Chittagong Sadar'),
        'Halishahar':     (22.3500, 91.7800, 'Chittagong', 'Chittagong Sadar'),
        'Pahartali':      (22.4100, 91.7900, 'Chittagong', 'Sitakunda'),
        'Patenga':        (22.2400, 91.8000, 'Chittagong', 'Chittagong Sadar'),
        'Nasirabad':      (22.3600, 91.8100, 'Chittagong', 'Chittagong Sadar'),
        'Chawkbazar':     (22.3470, 91.8270, 'Chittagong', 'Chittagong Sadar'),
        # ── Sylhet ──
        'Amberkhana':     (24.8992, 91.8626, 'Sylhet', 'Sylhet Sadar'),
        'Zindabazar':     (24.8944, 91.8593, 'Sylhet', 'Sylhet Sadar'),
        'Subhanighat':    (24.9000, 91.8700, 'Sylhet', 'Sylhet Sadar'),
        # ── Khulna ──
        'Khalispur':      (22.8300, 89.5700, 'Khulna', 'Khulna Sadar'),
        'Boyra':          (22.8500, 89.5500, 'Khulna', 'Khulna Sadar'),
        # ── Rajshahi ──
        'Shaheb Bazar':   (24.3700, 88.5900, 'Rajshahi', 'Rajshahi Sadar'),
        'Uposhohor':      (24.3800, 88.6100, 'Rajshahi', 'Rajshahi Sadar'),
    }

    # Scan: address-এ যেকোনো area name match হলে GPS দাও
    # Longest match first to avoid partial false match (e.g. 'Paltan' before 'Purana Paltan')
    for area in sorted(BD_AREA_GPS.keys(), key=len, reverse=True):
        if area.upper() in addr_upper:
            lat, lon, dist, thana = BD_AREA_GPS[area]
            # BD bounding box check
            if 19 <= lat <= 27 and 87 <= lon <= 93:
                return (lat, lon, dist, thana, 3000, 'keyword_area')
    return None


# Nominatim geocoding cache (in-memory, same session)
_NOMINATIM_CACHE = {}

def _nominatim_geocode(addr_str):
    """
    OpenStreetMap Nominatim geocoding — free, no API key.
    CDR address → GPS. Last resort যখন অন্য সব fail করে।
    Rate limit: 1 req/sec. Same address cache করা হয়।
    NOMINATIM_ENABLED = False হলে skip করা হবে।
    Returns (lat, lon, district, thana, radius_m, method) or None.
    """
    # Global toggle — UI থেকে on/off করা যাবে
    if not _NOMINATIM_CACHE.get('__enabled__', True):
        return None

    import re as _re_gc
    if not addr_str or len(str(addr_str).strip()) < 5:
        return None

    cache_key = _re_gc.sub(r'\s+', ' ', str(addr_str).strip().upper())
    if cache_key in _NOMINATIM_CACHE:
        return _NOMINATIM_CACHE[cache_key]

    try:
        import urllib.request as _urlr, json as _json, time as _time
        # Always append Bangladesh context for better results
        query = str(addr_str).strip()
        if 'bangladesh' not in query.lower() and 'dhaka' not in query.lower():
            query = query + ', Bangladesh'

        encoded = _urlr.quote(query)
        url = (f"https://nominatim.openstreetmap.org/search"
               f"?q={encoded}&format=json&limit=1&countrycodes=bd"
               f"&addressdetails=1")
        headers = {'User-Agent': 'CDR-Analysis-Bangladesh/1.0 (contact@ntmc.gov.bd)'}
        req = _urlr.Request(url, headers=headers)
        with _urlr.urlopen(req, timeout=5) as resp:
            data = _json.loads(resp.read())

        if not data:
            _NOMINATIM_CACHE[cache_key] = None
            return None

        top = data[0]
        lat = float(top['lat'])
        lon = float(top['lon'])

        # BD bounding box check
        if not (19 <= lat <= 27 and 87 <= lon <= 93):
            _NOMINATIM_CACHE[cache_key] = None
            return None

        # Extract district/thana from addressdetails
        addr_detail = top.get('address', {})
        district = (addr_detail.get('county') or
                    addr_detail.get('state_district') or
                    addr_detail.get('city') or '')
        thana = (addr_detail.get('suburb') or
                 addr_detail.get('quarter') or
                 addr_detail.get('neighbourhood') or '')

        _time.sleep(1.0)  # Nominatim rate limit: max 1 req/sec
        result = (lat, lon, str(district), str(thana), 2000, 'geocoded')
        _NOMINATIM_CACHE[cache_key] = result
        return result

    except Exception:
        _NOMINATIM_CACHE[cache_key] = None
        return None


def build_movement_map(df, phone, operator, mv_data=None):
    """
    GPS Location Logic:

    PRIORITY 1 — CSV LAC+CID match:
        → CDR address vs CSV address token similarity check
        → High sim (>=0.6): confirmed CSV GPS + district/thana
        → Low sim (<0.6) but same LAC match: GPS used with caution flag
        → CSV GPS vs text-parsed GPS distance > 40km: mark as suspicious

    PRIORITY 2 — Text parse from CDR BTS address:
        → P.S:/P/S: keyword → thana extract → THANA_GPS (±3-8km)
        → DIST: keyword → district → DISTRICT_GPS (±10-20km)
        → Last comma token fallback
        → Error/placeholder address → skip GPS

    HOME:
        → Most frequent GPS location (by coordinate cluster)
        → Not time-based — purely frequency

    FALSE POSITIVE RULES:
        → Single record AND >35km from home → ⚠️ suspicious
        → Same-day records in 2+ districts AND far → transit day
        → CSV GPS vs text GPS distance >40km → suspicious
        → "MOUZA NOT FOUND" / error address → no GPS assigned
    """
    import json as _json, re as _re  # math → _math_mod
    import pandas as _pd

    # ── Thana/District GPS — module-level BD_THANA_GPS / BD_DISTRICT_GPS ──
    # (দুই জায়গায় duplicate না রেখে একটা single source of truth থেকে নেওয়া হচ্ছে)
    THANA_GPS    = BD_THANA_GPS
    DISTRICT_GPS = BD_DISTRICT_GPS

    INVALID_PATTERNS = [
        'MOUZA NOT FOUND','NOT FOUND IN AG','CAAB PERMISSION',
        'PERMISSION FOUND','MOUZA-MOUZA',
    ]

    def hav(a,b,c,d):
        R=6371; dlat=_math_mod.radians(c-a); dlon=_math_mod.radians(d-b)
        x=_math_mod.sin(dlat/2)**2+_math_mod.cos(_math_mod.radians(a))*_math_mod.cos(_math_mod.radians(c))*_math_mod.sin(dlon/2)**2
        return round(R*2*_math_mod.asin(_math_mod.sqrt(max(0.0,x))),1)

    def bearing(la1,lo1,la2,lo2):
        dlo=_math_mod.radians(lo2-lo1); la1r=_math_mod.radians(la1); la2r=_math_mod.radians(la2)
        x=_math_mod.sin(dlo)*_math_mod.cos(la2r)
        y=_math_mod.cos(la1r)*_math_mod.sin(la2r)-_math_mod.sin(la1r)*_math_mod.cos(la2r)*_math_mod.cos(dlo)
        return round((_math_mod.degrees(_math_mod.atan2(x,y))+360)%360,1)

    def get_color(km, home_dist, this_dist):
        if km < 3:               return '#c62828'
        if this_dist==home_dist: return '#1565C0'
        if km < 50:              return '#E65100'
        if km < 150:             return '#7B1FA2'
        return '#1B5E20'

    def addr_tokens(s):
        noise={'vill','village','road','ward','house','the','and','plot','dist',
               'p.o','p.s','p','o','s','no','num','po','ps','mouza','moza',
               'union','para','gram','gram','bazar','hat','ghat','more'}
        return set(t for t in re.sub(r'[^a-z0-9]',' ',str(s).lower()).split()
                   if len(t)>=3 and t not in noise)

    def addr_sim(cdr_addr, csv_addr):
        ct=addr_tokens(cdr_addr); st=addr_tokens(csv_addr)
        if not ct or not st: return 0.0
        return round(len(ct&st)/max(len(ct),len(st)),2)

    def is_invalid(addr):
        a=str(addr).upper()
        return any(p in a for p in INVALID_PATTERNS)

    def text_gps(addr_str):
        """
        CDR address text → GPS.
        Priority chain (conditional):
          1. CSV LAC+CID (caller handles this before text_gps)
          2a. P.S: আছে → Thana parse → District parse
          2b. P.S: নেই → Area keyword scan → District fallback
          3.  উভয় path miss → Nominatim geocoding (last resort)
        """
        if is_invalid(addr_str): return None
        s = str(addr_str).upper()

        # ── P.S: keyword detect করো ──────────────────────────────────────
        mt = re.search(
            r'P[\.\s]*/?\s*S[\.\:\s\-/]+([A-Z][A-Z\s\-]{2,}?)(?:[,\.\n]|DIST|$)', s)
        has_ps = mt is not None
        thana  = mt.group(1).strip().rstrip('.,- ').title() if mt else ''

        # DIST: → district (উভয় path-এ দরকার)
        district = ''
        md = re.search(r'DIST[\.\:\s]+([A-Z][A-Z\s\-]{2,}?)(?:[,\.\n]|$|\s+BD)', s)
        if md:
            district = md.group(1).strip().rstrip('.,- ').title()

        # comma fallback for district/thana (শুধু P.S: path-এ)
        if has_ps and not district:
            parts = [p.strip() for p in re.split(r'[,،]', s) if len(p.strip()) > 2]
            parts = [re.sub(r'\b(BD|BANGLADESH|\d{4,})\b', '', p).strip() for p in parts]
            parts = [p for p in parts if p and not p.isdigit()]
            if parts:
                last = parts[-1].rstrip('.').title()
                if last.replace(' ', '').isalpha() and len(last) >= 4:
                    district = last
                if len(parts) >= 2 and not thana:
                    sl = parts[-2].rstrip('.').title()
                    if len(sl) >= 4: thana = sl

        dist_norm = {'Gaibanda':'Gaibandha','Bogura':'Bogra','Sirajgonj':'Sirajganj',
                     'Cumilla':'Comilla','Bogra Sadar South New':'Bogra'}

        if has_ps:
            # ── PATH A: P.S: আছে → Thana → District ─────────────────────
            if thana:
                for tk in [thana, thana.replace(' Sadar', '').strip()]:
                    if tk in THANA_GPS:
                        g = THANA_GPS[tk]
                        d = dist_norm.get(district, district) or tk.split()[0]
                        return (g[0], g[1], d, thana, 5000, 'text_thana')

            if district:
                d = dist_norm.get(district, district)
                if d in DISTRICT_GPS:
                    g = DISTRICT_GPS[d]
                    return (g[0], g[1], d, thana, 15000, 'text_district')
                for k, v in DISTRICT_GPS.items():
                    if k.lower() in d.lower() or d.lower() in k.lower():
                        return (v[0], v[1], k, thana, 15000, 'text_district')

        else:
            # ── PATH B: P.S: নেই → Keyword scan → District fallback ──────
            # Keyword scan — P.S: ছাড়া address-এ area name সরাসরি খোঁজো
            res = _keyword_scan_gps(s)
            if res:
                return res

            # District fallback (comma-last-token)
            parts = [p.strip() for p in re.split(r'[,،]', s) if len(p.strip()) > 2]
            parts = [re.sub(r'\b(BD|BANGLADESH|\d{4,})\b', '', p).strip() for p in parts]
            parts = [p for p in parts if p and not p.isdigit()]
            if parts:
                last = parts[-1].rstrip('.').title()
                if last.replace(' ', '').isalpha() and len(last) >= 4:
                    district = last
            if district:
                d = dist_norm.get(district, district)
                if d in DISTRICT_GPS:
                    g = DISTRICT_GPS[d]
                    return (g[0], g[1], d, '', 15000, 'text_district')
                for k, v in DISTRICT_GPS.items():
                    if k.lower() in d.lower() or d.lower() in k.lower():
                        return (v[0], v[1], k, '', 15000, 'text_district')

        # ── LAST RESORT: Nominatim geocoding ─────────────────────────────
        # উভয় path miss হলে শুধু তখনই Nominatim call হবে।
        # Rate limit ১ req/sec — cache করা থাকে, repeat call নেই।
        res2 = _nominatim_geocode(str(addr_str))
        if res2:
            return res2

        return None

    # ── Validate inputs ──
    if 'start' not in df.columns or ('address' not in df.columns and 'cell_csv_label' not in df.columns):
        return None

    df=df.copy()

    # ── Teletalk fallback: address blank/invalid → cell_csv_label ──────────────
    if 'cell_csv_label' in df.columns and 'loc_method' in df.columns:
        _mv2_invalid = (
            df['address'].isna() |
            ~df['address'].apply(lambda a: _is_valid_address(str(a)) if _pd.notna(a) else False)
        ) if 'address' in df.columns else _pd.Series([True]*len(df), index=df.index)
        _mv2_enriched = (
            (df['loc_method'] == 'cell_exact') &
            df['cell_csv_label'].notna() &
            (df['cell_csv_label'].str.strip() != '')
        )
        if 'address' not in df.columns:
            df['address'] = ''
        df.loc[_mv2_invalid & _mv2_enriched, 'address'] = df.loc[_mv2_invalid & _mv2_enriched, 'cell_csv_label']

    has_loc='loc_method' in df.columns
    has_csv_lat='cell_lat' in df.columns
    has_csv_dist='csv_district' in df.columns
    has_csv_thana='csv_thana' in df.columns  # may or may not exist
    has_csv_label='cell_csv_label' in df.columns

    ADDR_SIM_THRESHOLD = 0.5   # CSV GPS accepted if sim >= this
    DIST_SUSPECT_KM   = 40.0   # CSV GPS vs text GPS > this → suspicious

    rows_out=[]
    for _,row in df.sort_values('start').iterrows():
        addr=str(row.get('address',''))
        method=str(row.get('loc_method','')) if has_loc else ''
        lat=lon=None; district=thana=''; acc_m=None; gps_method='none'; suspect=False

        # ── PRIORITY 1: CSV LAC+CID match ──
        if method=='cell_exact' and has_csv_lat and _pd.notna(row.get('cell_lat')):
            csv_lat=float(row['cell_lat']); csv_lon=float(row['cell_lon'])

            # Address similarity: CDR addr vs CSV label
            csv_lbl=str(row.get('cell_csv_label','')) if has_csv_label else ''
            sim=addr_sim(addr, csv_lbl) if csv_lbl else 1.0  # no label = trust CSV

            csv_dist_val=str(row.get('csv_district','')).strip().title() if has_csv_dist else ''
            csv_thana_val=str(row.get('csv_thana','')).strip().title() if has_csv_thana else ''

            # Cross-check: text parse for this address
            tg=text_gps(addr)
            if tg and sim>=ADDR_SIM_THRESHOLD:
                text_lat,text_lon=tg[0],tg[1]
                dist_diff=hav(csv_lat,csv_lon,text_lat,text_lon)
                if dist_diff>DIST_SUSPECT_KM:
                    suspect=True  # GPS vs text disagree by >40km

            lat=csv_lat; lon=csv_lon
            district=csv_dist_val; thana=csv_thana_val
            acc_m=1500
            gps_method='csv_exact'

        # ── PRIORITY 2: Text parse (no CSV match) ──
        elif not is_invalid(addr):
            tg=text_gps(addr)
            if tg:
                lat,lon,district,thana,acc_m,gps_method=tg

        if lat is not None and 19<=lat<=27 and 87<=lon<=93:
            # BTS display: valid CDR address → সেটা, নইলে cell_csv_label
            _raw_addr = str(row.get('address','')).strip()
            _csv_lbl_disp = str(row.get('cell_csv_label','') or '').strip() if has_csv_label else ''
            if _raw_addr in ('', 'nan', 'None', 'NaN') or not _is_valid_address(_raw_addr):
                _disp_addr = _csv_lbl_disp if _csv_lbl_disp else _raw_addr
            else:
                _disp_addr = _raw_addr
            rows_out.append({
                'lat':lat,'lon':lon,'district':district,'thana':thana,
                'acc_m':acc_m,'gps_method':gps_method,'suspect':suspect,
                'start':row['start'],'addr':_disp_addr[:70],
            })

    if len(rows_out)<2: return None
    gdf=_pd.DataFrame(rows_out).sort_values('start').reset_index(drop=True)
    gdf['km_raw']=0.0  # placeholder, will compute after home

    # ── HOME: CDR address frequency → GPS from CSV match ──
    # সবচেয়ে বেশি CDR address → সেটাই home
    # GPS: সেই address-এর csv_exact GPS (যদি থাকে), নইলে text parse
    addr_freq = df['address'].value_counts()
    addr_freq = addr_freq[addr_freq.index.str.len() > 5]  # empty address বাদ

    home_lat = home_lon = None
    home_dist_val = ''; home_label = 'Home'

    def _best_home_label(top_addr, df_orig, src_rows):
        """csv_thana → csv_district → address split order-এ home label বানাও।"""
        # src_rows: gdf rows for this address (may have thana/district from text_gps)
        # df_orig: original df (has csv_thana, csv_district from enrichment)
        _thana = ''; _dist = ''
        # CSV enrichment columns থেকে নেওয়ার চেষ্টা
        if 'csv_thana' in df_orig.columns:
            _match = df_orig[df_orig['address'].str.upper().str[:50] == top_addr.upper()[:50]]
            _tv = _match['csv_thana'].dropna()
            if not _tv.empty:
                _thana = str(_tv.mode().iloc[0]).strip().title()
                if _thana in ('Nan', 'None', ''): _thana = ''
        if 'csv_district' in df_orig.columns:
            _match = df_orig[df_orig['address'].str.upper().str[:50] == top_addr.upper()[:50]]
            _dv = _match['csv_district'].dropna()
            if not _dv.empty:
                _dist = str(_dv.mode().iloc[0]).strip().title()
                if _dist in ('Nan', 'None', ''): _dist = ''
        # src_rows district fallback (from gdf text parse)
        if not _dist and src_rows is not None and not src_rows.empty:
            _md = src_rows['district'].dropna()
            if not _md.empty: _dist = str(_md.mode().iloc[0]).strip().title()
        # Priority: csv_thana → csv_district → address first token
        if _thana:
            return _thana, _dist
        if _dist:
            return _dist, _dist
        return top_addr.split(',')[0][:25].split('|')[0].strip().title(), _dist

    for top_addr in addr_freq.index:
        # CDR rows for this address
        addr_rows = gdf[gdf['addr'].str.upper().str[:50] == top_addr.upper()[:50]]
        if addr_rows.empty:
            # Try text-parsed GPS for this address
            tg = text_gps(top_addr)
            if tg:
                home_lat, home_lon = tg[0], tg[1]
                home_dist_val = tg[2]
                home_label, _ = _best_home_label(top_addr, df, None)
                break
            continue
        # Prefer csv_exact rows
        ex = addr_rows[addr_rows['gps_method'] == 'csv_exact']
        src = ex if not ex.empty else addr_rows
        home_lat = float(src['lat'].mean())
        home_lon = float(src['lon'].mean())
        home_label, home_dist_val = _best_home_label(top_addr, df, src)
        if not home_dist_val:
            home_dist_val = str(src['district'].mode().iloc[0]) if not src['district'].empty else ''
        break

    if home_lat is None:
        # Fallback: most frequent GPS cluster
        gdf['lat_r']=gdf['lat'].round(3); gdf['lon_r']=gdf['lon'].round(3)
        freq=gdf.groupby(['lat_r','lon_r','district']).size().reset_index(name='cnt').sort_values('cnt',ascending=False)
        home_lat=float(freq.iloc[0]['lat_r']); home_lon=float(freq.iloc[0]['lon_r'])
        home_dist_val=str(freq.iloc[0]['district'])
        home_rows=gdf[(gdf['lat_r']==freq.iloc[0]['lat_r'])&(gdf['lon_r']==freq.iloc[0]['lon_r'])]
        _hlvc = home_rows['addr'].value_counts()
        _top_fallback = str(_hlvc.index[0] if not _hlvc.empty else '')
        home_label, _fd = _best_home_label(_top_fallback, df, home_rows)
        if not home_label: home_label = 'Home'

    gdf['km']=gdf.apply(lambda r:hav(home_lat,home_lon,r['lat'],r['lon']),axis=1)

    # ── Transit day detection ──
    gdf['date_str']=_pd.to_datetime(gdf['start']).dt.date.astype(str)
    transit_days=set()
    for date,grp in gdf.groupby('date_str'):
        far=grp[grp['km']>30]
        if len(far)>=2 and far['district'].nunique()>1:
            transit_days.add(date)
    gdf['is_transit']=gdf['date_str'].isin(transit_days)
    transit_info={}
    for date in sorted(transit_days):
        grp=gdf[gdf['date_str']==date].sort_values('start')
        dists=list(dict.fromkeys(d for d in grp['district'].tolist() if d))
        transit_info[date]=' → '.join(dists)

    # ── Build steps from non-transit ──
    stay=gdf[~gdf['is_transit']].copy()
    MAX_STEPS=60
    steps=[]; step_rows=[]; prev_lat=prev_lon=None; prev_dist=None
    for _,row in stay.iterrows():
        clat=float(row['lat']); clon=float(row['lon'])
        cur_dist=str(row.get('district','')).strip()
        d2p=hav(prev_lat,prev_lon,clat,clon) if prev_lat else 999
        # Same district within 50km OR any location within 25km → same step
        same_grp=(d2p<25 and step_rows) or (cur_dist and cur_dist==prev_dist and d2p<50 and step_rows)
        if same_grp:
            step_rows.append(row)
        else:
            if step_rows:
                sr=_pd.DataFrame(step_rows)
                ex=sr[sr['gps_method']=='csv_exact']
                rep=ex if not ex.empty else sr
                top_d=sr['district'].value_counts().index[0] if not sr['district'].value_counts().empty else ''
                top_th=sr['thana'].value_counts().index[0] if not sr['thana'].value_counts().empty and sr['thana'].any() else ''
                best_acc=int(sr['acc_m'].min())
                best_m='csv_exact' if not ex.empty else sr['gps_method'].mode().iloc[0]
                any_suspect=bool(sr['suspect'].any())
                steps.append({'lat':float(rep['lat'].mean()),'lon':float(rep['lon'].mean()),
                    'district':top_d,'thana':top_th,'count':len(sr),
                    'km':round(float(sr['km'].max()),1),
                    'start':str(sr['start'].min()),'end':str(sr['start'].max()),
                    'addr':str(sr['addr'].value_counts().index[0])[:65] if not sr['addr'].value_counts().empty else '',
                    'method':best_m,'acc_m':best_acc,'suspect':any_suspect})
            step_rows=[row]
        prev_lat=clat; prev_lon=clon; prev_dist=cur_dist
    if step_rows:
        sr=_pd.DataFrame(step_rows)
        ex=sr[sr['gps_method']=='csv_exact']
        rep=ex if not ex.empty else sr
        top_d=sr['district'].value_counts().index[0] if not sr['district'].value_counts().empty else ''
        top_th=sr['thana'].value_counts().index[0] if not sr['thana'].value_counts().empty and sr['thana'].any() else ''
        best_m='csv_exact' if not ex.empty else sr['gps_method'].mode().iloc[0]
        steps.append({'lat':float(rep['lat'].mean()),'lon':float(rep['lon'].mean()),
            'district':top_d,'thana':top_th,'count':len(sr),
            'km':round(float(sr['km'].max()),1),
            'start':str(sr['start'].min()),'end':str(sr['start'].max()),
            'addr':str(sr['addr'].value_counts().index[0])[:65] if not sr['addr'].value_counts().empty else '',
            'method':best_m,'acc_m':int(sr['acc_m'].min()),'suspect':bool(sr['suspect'].any())})

    # ── Suspicious filter ──
    # Rule 1: 1 record AND >35km from home → unconfirmed
    # Rule 2: suspect flag (CSV vs text GPS disagree >40km) → unconfirmed
    # FIX: mv_data trips থেকে confirmed GPS inject করো — report-এ যে trips আছে
    #       সেগুলো map-এ missing হওয়া ঠেকাতে।

    # mv_data trips দিয়ে একটা lookup তৈরি করো: district → (lat, lon, is_exact)
    _mv_trip_gps = {}
    if mv_data and mv_data.get('trips'):
        for _t in mv_data['trips']:
            _dlabel = str(_t.get('district') or _t.get('upazila') or '').strip().lower()
            _tlat = _t.get('lat'); _tlon = _t.get('lon')
            if _dlabel and _tlat and _tlon:
                _mv_trip_gps[_dlabel] = {
                    'lat': float(_tlat), 'lon': float(_tlon),
                    'is_exact': bool(_t.get('is_exact', False)),
                    'km': float(_t.get('km', 0)),
                    'start': str(_t.get('start_date', '')),
                    'end': str(_t.get('end_date', '')),
                    'district': str(_t.get('district', '')),
                    'thana': str(_t.get('upazila', '') or ''),
                    'addr': str(_t.get('address', ''))[:65],
                }

    main_steps=[]; suspicious=[]
    for s in steps:
        is_sus=False
        if s['count']==1 and s['km']>35:
            is_sus=True
        elif s['suspect'] and s['km']>35:
            is_sus=True

        if is_sus:
            # mv_data-তে এই district-এর confirmed GPS আছে কিনা দেখো
            _s_dist = str(s.get('district', '')).strip().lower()
            _mv_hit = _mv_trip_gps.get(_s_dist)
            if _mv_hit:
                # Report-এর GPS দিয়ে replace করো — confirmed location
                s = dict(s)
                s['lat']     = _mv_hit['lat']
                s['lon']     = _mv_hit['lon']
                s['method']  = 'csv_exact' if _mv_hit['is_exact'] else s.get('method', 'text_district')
                s['suspect'] = False   # confirmed, not suspicious
                main_steps.append(s)
            else:
                suspicious.append(s)
        else:
            # Normal step — তবু mv_data-তে better GPS থাকলে upgrade করো
            _s_dist = str(s.get('district', '')).strip().lower()
            _mv_hit = _mv_trip_gps.get(_s_dist)
            if _mv_hit and _mv_hit.get('is_exact') and s.get('method') != 'csv_exact':
                s = dict(s)
                s['lat']    = _mv_hit['lat']
                s['lon']    = _mv_hit['lon']
                s['method'] = 'csv_exact'
            main_steps.append(s)

    # mv_data trips-এ যে destinations আছে কিন্তু steps-এ নেই → directly add করো
    if mv_data and mv_data.get('trips'):
        _existing_dists = set(str(s.get('district','')).strip().lower() for s in main_steps)
        for _t in mv_data['trips']:
            _dlabel = str(_t.get('district') or _t.get('upazila') or '').strip().lower()
            _tlat = _t.get('lat'); _tlon = _t.get('lon')
            if not _tlat or not _tlon:
                continue
            if _dlabel not in _existing_dists:
                # এই trip map-এ missing — সরাসরি যোগ করো
                main_steps.append({
                    'lat': float(_tlat), 'lon': float(_tlon),
                    'district': str(_t.get('district', '')),
                    'thana': str(_t.get('upazila', '') or ''),
                    'count': 1,
                    'km': float(_t.get('km', 0)),
                    'start': str(_t.get('start_date', '')),
                    'end': str(_t.get('end_date', '')),
                    'addr': str(_t.get('address', ''))[:65],
                    'method': 'csv_exact' if _t.get('is_exact') else 'text_district',
                    'acc_m': 500 if _t.get('is_exact') else 15000,
                    'suspect': False,
                })
                _existing_dists.add(_dlabel)

        # Chronological sort — start date অনুযায়ী
        main_steps.sort(key=lambda s: s.get('start', ''))

    if not main_steps: return None
    steps=main_steps

    # ── Step limit: MAX 60 steps — বেশি হলে same-district consecutive merge ──
    while len(steps) > MAX_STEPS:
        # Find two consecutive steps with same district → merge
        merged=False
        for i in range(len(steps)-1):
            if steps[i]['district']==steps[i+1]['district']:
                s1=steps[i]; s2=steps[i+1]
                merged_step={
                    'lat':(s1['lat']*s1['count']+s2['lat']*s2['count'])/(s1['count']+s2['count']),
                    'lon':(s1['lon']*s1['count']+s2['lon']*s2['count'])/(s1['count']+s2['count']),
                    'district':s1['district'],'thana':s1.get('thana',''),
                    'count':s1['count']+s2['count'],
                    'km':max(s1['km'],s2['km']),
                    'start':s1['start'],'end':s2['end'],
                    'addr':s1['addr'] if s1['count']>=s2['count'] else s2['addr'],
                    'method':s1['method'] if s1['method']=='csv_exact' else s2['method'],
                    'acc_m':min(s1['acc_m'],s2['acc_m']),
                    'suspect':s1.get('suspect',False) or s2.get('suspect',False),
                }
                steps=steps[:i]+[merged_step]+steps[i+2:]
                merged=True; break
        if not merged:
            # No same-district pair → merge closest pair by km difference
            min_diff=float('inf'); min_i=0
            for i in range(len(steps)-1):
                diff=abs(steps[i]['km']-steps[i+1]['km'])
                if diff<min_diff: min_diff=diff; min_i=i
            s1=steps[min_i]; s2=steps[min_i+1]
            merged_step={
                'lat':(s1['lat']+s2['lat'])/2,'lon':(s1['lon']+s2['lon'])/2,
                'district':s1['district'] if s1['count']>=s2['count'] else s2['district'],
                'thana':s1.get('thana',''),'count':s1['count']+s2['count'],
                'km':max(s1['km'],s2['km']),'start':s1['start'],'end':s2['end'],
                'addr':s1['addr'],'method':s1['method'],'acc_m':min(s1['acc_m'],s2['acc_m']),
                'suspect':s1.get('suspect',False) or s2.get('suspect',False),
            }
            steps=steps[:min_i]+[merged_step]+steps[min_i+2:]

    # ── Stats ──
    csv_exact_cnt=int((gdf['gps_method']=='csv_exact').sum())
    text_cnt=int((gdf['gps_method']!='csv_exact').sum())
    period_start=str(_pd.to_datetime(gdf['start'].min()).date())
    period_end=str(_pd.to_datetime(gdf['start'].max()).date())
    total=len(df)

    def acc_badge(method,acc_m,suspect=False):
        s_tag='<span style="color:#dc2626"> ⚠️</span>' if suspect else ''
        if method=='csv_exact':
            return f"<span style='background:#d1fae5;color:#065f46;font-size:9px;padding:1px 5px;border-radius:8px;font-weight:600'>📡 CSV ±{acc_m}m</span>"+s_tag
        elif 'thana' in method:
            return f"<span style='background:#fef3c7;color:#92400e;font-size:9px;padding:1px 5px;border-radius:8px;font-weight:600'>📍 Thana ~{acc_m//1000}km</span>"+s_tag
        return f"<span style='background:#fee2e2;color:#991b1b;font-size:9px;padding:1px 5px;border-radius:8px;font-weight:600'>🌍 District ~{acc_m//1000}km</span>"+s_tag

    # ── Leaflet JS ──
    js=[]
    coords=[[round(s['lat'],6),round(s['lon'],6)] for s in steps]
    js.append("var coords="+_json.dumps(coords)+";")
    js.append("var route=L.polyline.antPath(coords,{color:'#1d4ed8',weight:3,opacity:0.75,delay:600,dashArray:[14,18],pulseColor:'#93c5fd',paused:false,reverse:false}).addTo(map);")

    for i in range(len(steps)-1):
        p1=steps[i]; p2=steps[i+1]
        ml=round((p1['lat']+p2['lat'])/2,6); mlo=round((p1['lon']+p2['lon'])/2,6)
        b=bearing(p1['lat'],p1['lon'],p2['lat'],p2['lon'])
        dk=hav(p1['lat'],p1['lon'],p2['lat'],p2['lon'])
        col=get_color(p2['km'],home_dist_val,p2['district'])
        d1=(p1['district'] or '?').replace('"','')
        d2=(p2['district'] or '?').replace('"','')
        svg=("<svg width='26' height='26' viewBox='0 0 26 26' style='overflow:visible;display:block;margin:-13px 0 0 -13px'>"
            "<defs><marker id='ah{i}' markerWidth='7' markerHeight='7' refX='5' refY='3.5' orient='auto'>"
            "<path d='M0,0.5 L6,3.5 L0,6.5 Z' fill='{col}' stroke='white' stroke-width='0.7'/></marker></defs>"
            "<circle cx='13' cy='13' r='6' fill='white' fill-opacity='0.8' stroke='{col}' stroke-width='1.5'/>"
            "<line x1='4' y1='13' x2='20' y2='13' stroke='{col}' stroke-width='2.8' stroke-linecap='round' "
            "marker-end='url(#ah{i})' transform='rotate({b},13,13)'/></svg>").format(i=i,col=col,b=b)
        tip="{a}&#8594;{b2}: {d1}&#8594;{d2} ({dk}km)".format(a=i+1,b2=i+2,d1=d1,d2=d2,dk=dk)
        js.append("L.marker([{ml},{mlo}],{{icon:L.divIcon({{html:{svg},iconSize:[26,26],iconAnchor:[13,13],className:''}}),zIndexOffset:-50}}).addTo(map).bindTooltip('{tip}',{{sticky:true}});".format(ml=ml,mlo=mlo,svg=_json.dumps(svg),tip=tip))

    for i,s in enumerate(steps):
        num=i+1; la=round(s['lat'],6); lo=round(s['lon'],6)
        km=s['km']; col=get_color(km,home_dist_val,s['district'])
        dist=(s['district'] or '—').replace('"','')
        thana=(s.get('thana','') or '').replace('"','')
        loc_lbl=f"{thana}, {dist}" if thana and thana.lower()!=dist.lower() else dist
        cnt=s['count']; st2=s['start'][:16]; en=s['end'][:16]
        addr=s['addr'].replace('"',' ').replace("'",' ')
        badge=acc_badge(s['method'],s['acc_m'],s.get('suspect',False))
        rad=max(9,min(32,9+cnt//6))
        popup=("<div style='font-family:Arial;min-width:230px'>"
            "<div style='background:{col};color:#fff;padding:7px 12px;border-radius:8px 8px 0 0;font-weight:700;font-size:13px'>Step {num} — {loc_lbl} ({km}km)</div>"
            "<table style='width:100%;font-size:12px;border-collapse:collapse;border:1px solid #e5e7eb;border-top:none'>"
            "<tr style='background:#f9fafb'><td style='padding:4px 8px;color:#6b7280'>Period</td><td style='padding:4px 8px'>{st2}<br>&rarr; {en}</td></tr>"
            "<tr><td style='padding:4px 8px;color:#6b7280'>Records</td><td style='padding:4px 8px;font-weight:600'>{cnt}</td></tr>"
            "<tr style='background:#f9fafb'><td style='padding:4px 8px;color:#6b7280'>From Home</td><td style='padding:4px 8px'>{km} km</td></tr>"
            "<tr><td style='padding:4px 8px;color:#6b7280'>GPS</td><td style='padding:4px 8px;font-family:monospace;font-size:11px'>{la}, {lo}</td></tr>"
            "<tr style='background:#f9fafb'><td style='padding:4px 8px;color:#6b7280'>Accuracy</td><td style='padding:4px 8px'>{badge}</td></tr>"
            "<tr><td style='padding:4px 8px;color:#6b7280'>BTS</td><td style='padding:4px 8px;font-size:11px'>{addr}</td></tr>"
            "</table></div>"
        ).format(col=col,num=num,loc_lbl=loc_lbl,km=km,st2=st2,en=en,cnt=cnt,la=la,lo=lo,badge=badge,addr=addr)
        tip2="Step {num}: {loc} | {s0}&rarr;{e0} | {cnt}rec | {km}km".format(num=num,loc=loc_lbl,s0=s['start'][:10],e0=s['end'][:10],cnt=cnt,km=km)
        js.append("L.circleMarker([{la},{lo}],{{radius:{rad},fillColor:'{col}',color:'#fff',weight:2.5,opacity:1,fillOpacity:0.9}}).addTo(map).bindPopup({pop}).bindTooltip('{tip}',{{sticky:true}});".format(la=la,lo=lo,rad=rad,col=col,pop=_json.dumps(popup),tip=tip2))
        ni="<div style='background:{col};color:#fff;border-radius:50%;width:20px;height:20px;font-size:10px;font-weight:700;display:flex;align-items:center;justify-content:center;box-shadow:0 2px 6px rgba(0,0,0,.35);margin:-10px 0 0 -10px'>{num}</div>".format(col=col,num=num)
        js.append("L.marker([{la},{lo}],{{icon:L.divIcon({{html:{ni},iconSize:[20,20],iconAnchor:[10,10],className:''}}),zIndexOffset:10}}).addTo(map);".format(la=la,lo=lo,ni=_json.dumps(ni)))

    for s in suspicious:
        la=round(s['lat'],6); lo=round(s['lon'],6)
        dist=(s['district'] or '?').replace('"','')
        thana=(s.get('thana','') or '').replace('"','')
        loc_s=f"{thana}, {dist}" if thana and thana.lower()!=dist.lower() else dist
        pop=("<div style='font-family:Arial;padding:10px;min-width:200px'>"
             "<b style='color:#f59e0b'>&#9888; Unconfirmed Location</b><br>"
             "<b>Location:</b> {loc}<br><b>GPS:</b> {la},{lo}<br>"
             "<b>From Home:</b> {km}km<br><b>Records:</b> {cnt}<br>"
             "<b>Date:</b> {dt}</div>").format(loc=loc_s,la=la,lo=lo,km=round(s['km'],1),cnt=s['count'],dt=s['start'][:10])
        js.append("L.circleMarker([{la},{lo}],{{radius:7,fillColor:'#f59e0b',color:'white',weight:1.5,opacity:0.8,fillOpacity:0.35,dashArray:'5,4'}}).addTo(map).bindPopup({pop}).bindTooltip('&#9888; {loc} ({km}km) — {cnt}rec',{{sticky:true}});".format(la=la,lo=lo,pop=_json.dumps(pop),loc=loc_s,km=round(s['km'],1),cnt=s['count']))

    home_pop="<div style='font-family:Arial;padding:10px'><b style='font-size:14px'>&#127968; Home Location</b><br><br><b>Area:</b> {hl}<br><b>District:</b> {hd}<br><b>GPS:</b> {la}, {lo}<br><b>Source:</b> Most frequent BTS location</div>".format(hl=home_label,hd=home_dist_val,la=round(home_lat,6),lo=round(home_lon,6))
    js.append("L.marker([{la},{lo}],{{icon:L.divIcon({{html:\"<div style='font-size:30px;margin:-15px 0 0 -15px'>&#127968;</div>\",iconSize:[30,30],iconAnchor:[15,15],className:''}}),zIndexOffset:1000}}).addTo(map).bindPopup({pop}).bindTooltip('&#127968; Home: {hl}',{{sticky:true,permanent:true,direction:'right',offset:[15,0]}});".format(la=round(home_lat,6),lo=round(home_lon,6),pop=_json.dumps(home_pop),hl=home_label))

    # ── Top 3 Frequent Locations — mv_data থেকে GPS নিয়ে map-এ দেখাও ──────
    if mv_data and mv_data.get('top_locations'):
        _freq_icons  = ['&#127968;', '&#11088;', '&#11088;']  # 🏠 ⭐ ⭐
        _freq_colors = ['#16a34a', '#2563eb', '#7c3aed']
        _freq_labels = ['Home (Most Frequent)', '2nd Frequent Location', '3rd Frequent Location']
        for _fi, _floc in enumerate(mv_data['top_locations'][:3]):
            _fgps = _floc.get('gps', '')
            if not _fgps:
                continue
            try:
                _fla, _flo = [float(x.strip()) for x in _fgps.split(',')]
                if not (19 <= _fla <= 27 and 87 <= _flo <= 93):
                    continue
            except Exception:
                continue
            _fname   = str(_floc.get('name', '')).replace('"', '').replace("'", '')
            _fdist   = str(_floc.get('district', '')).replace('"', '').replace("'", '')
            _fcount  = int(_floc.get('count', 0))
            _faddr   = str(_floc.get('address', ''))[:60].replace('"', '').replace("'", '')
            _fcol    = _freq_colors[_fi]
            _ficon   = _freq_icons[_fi]
            _flabel  = _freq_labels[_fi]
            _fpop = (
                "<div style='font-family:Arial;padding:10px;min-width:220px'>"
                "<div style='background:{col};color:#fff;padding:6px 10px;border-radius:6px 6px 0 0;"
                "font-weight:700;font-size:13px'>{icon} {label}</div>"
                "<table style='width:100%;font-size:12px;border-collapse:collapse;"
                "border:1px solid #e5e7eb;border-top:none'>"
                "<tr style='background:#f9fafb'><td style='padding:4px 8px;color:#6b7280'>Area</td>"
                "<td style='padding:4px 8px;font-weight:600'>{name}</td></tr>"
                "<tr><td style='padding:4px 8px;color:#6b7280'>District</td>"
                "<td style='padding:4px 8px'>{dist}</td></tr>"
                "<tr style='background:#f9fafb'><td style='padding:4px 8px;color:#6b7280'>GPS</td>"
                "<td style='padding:4px 8px;font-family:monospace;font-size:11px'>{la}, {lo}</td></tr>"
                "<tr><td style='padding:4px 8px;color:#6b7280'>Records</td>"
                "<td style='padding:4px 8px;font-weight:600'>{cnt} records</td></tr>"
                "<tr style='background:#f9fafb'><td style='padding:4px 8px;color:#6b7280'>BTS</td>"
                "<td style='padding:4px 8px;font-size:11px'>{addr}</td></tr>"
                "</table></div>"
            ).format(col=_fcol,icon=_ficon,label=_flabel,name=_fname,
                     dist=_fdist,la=round(_fla,6),lo=round(_flo,6),
                     cnt=_fcount,addr=_faddr)
            # Star marker — numbered (1,2,3)
            _fnum = _fi + 1
            _fni = ("<div style='background:{col};color:#fff;border-radius:50%;"
                    "width:24px;height:24px;font-size:11px;font-weight:800;"
                    "display:flex;align-items:center;justify-content:center;"
                    "box-shadow:0 3px 8px rgba(0,0,0,.4);margin:-12px 0 0 -12px;"
                    "border:2px solid white'>F{n}</div>").format(col=_fcol, n=_fnum)
            _ftip = "{icon} {label}: {name} ({dist}) | {cnt} records | {la},{lo}".format(
                icon=_ficon, label=_flabel, name=_fname, dist=_fdist,
                cnt=_fcount, la=round(_fla,6), lo=round(_flo,6))
            js.append(
                "L.circleMarker([{la},{lo}],{{radius:14,fillColor:'{col}',color:'white',"
                "weight:3,opacity:0.9,fillOpacity:0.25,dashArray:'6,3',zIndexOffset:800}})"
                ".addTo(map);".format(la=round(_fla,6),lo=round(_flo,6),col=_fcol))
            js.append(
                "L.marker([{la},{lo}],{{icon:L.divIcon({{html:{ni},iconSize:[24,24],"
                "iconAnchor:[12,12],className:''}}),zIndexOffset:900}})"
                ".addTo(map).bindPopup({pop}).bindTooltip({tip},{{sticky:true}});".format(
                    la=round(_fla,6),lo=round(_flo,6),
                    ni=_json.dumps(_fni),pop=_json.dumps(_fpop),tip=_json.dumps(_ftip)))

    all_bounds=[[round(s['lat'],6),round(s['lon'],6)] for s in steps]+[[round(home_lat,6),round(home_lon,6)]]
    js.append("map.fitBounds("+_json.dumps(all_bounds)+",{padding:[80,80]});")
    all_js='\n'.join(js)

    # ── Timeline ──
    tl_rows=''
    for i,s in enumerate(steps):
        col=get_color(s['km'],home_dist_val,s['district'])
        dist=s['district'] or '—'; thana=s.get('thana','') or ''
        loc_lbl=f"{thana}, {dist}" if thana and thana.lower()!=dist.lower() else dist
        icon='&#127968;' if s['km']<3 else ('&#9992;' if s['km']>150 else ('&#128663;' if s['km']>35 else '&#128205;'))
        mi='&#128249;' if s['method']=='csv_exact' else ('&#128270;' if 'thana' in s['method'] else '&#127758;')
        tl_rows+=("<div class='tl-row' onclick=\"map.setView([{la},{lo}],13)\">"
            "<div class='tl-num' style='background:{col}'>{n}</div>"
            "<div class='tl-info'><span style='font-weight:600;color:{col}'>{icon} {loc}</span>"
            " <span class='tl-km'>{km}km</span> <span title='source'>{mi}</span><br>"
            "<span class='tl-date'>{st} &rarr; {en}</span>"
            " &middot; <span class='tl-cnt'>{cnt}rec</span></div></div>\n"
        ).format(la=round(s['lat'],6),lo=round(s['lon'],6),col=col,n=i+1,
                 icon=icon,loc=loc_lbl,km=s['km'],mi=mi,
                 st=s['start'][:10],en=s['end'][:10],cnt=s['count'])

    if transit_info:
        tl_rows+='<div style="margin-top:8px;padding:6px 8px;background:#fffbeb;border-radius:6px;border-left:3px solid #f59e0b;font-size:10px;color:#92400e"><b>&#128652; Transit days:</b><br>'
        for dt,route in sorted(transit_info.items()):
            tl_rows+=f"&nbsp;{dt}: {route}<br>"
        tl_rows+='</div>'
    if suspicious:
        tl_rows+='<div style="margin-top:6px;padding:6px 8px;background:#fef9f0;border-radius:6px;border-left:3px solid #f59e0b;font-size:10px;color:#b45309"><b>&#9888; Unconfirmed:</b><br>'
        for s in suspicious:
            loc_s=(s.get('thana','') or s['district'] or '?')
            tl_rows+=f"&nbsp;{loc_s} {round(s['km'],1)}km &middot; {s['start'][:10]} ({s['count']}rec)<br>"
        tl_rows+='</div>'

    dot=lambda c:f"<span style='display:inline-block;width:13px;height:13px;border-radius:50%;background:{c};vertical-align:middle'></span> "
    lgd_arrow=("<svg width='22' height='14' style='vertical-align:middle;margin-right:2px'>"
        "<defs><marker id='lgd-a' markerWidth='6' markerHeight='6' refX='4' refY='3' orient='auto'>"
        "<path d='M0,0.5 L6,3 L0,5.5 Z' fill='#1d4ed8'/></marker></defs>"
        "<circle cx='5' cy='7' r='3' fill='white' stroke='#1d4ed8' stroke-width='1.2'/>"
        "<line x1='3' y1='7' x2='18' y2='7' stroke='#1d4ed8' stroke-width='2.2' marker-end='url(#lgd-a)'/></svg>")
    note_bar=(f"<div style='position:fixed;bottom:24px;left:280px;z-index:1000;background:#fffbeb;"
        f"border:1px solid #f59e0b;border-radius:8px;padding:8px 14px;font-family:Arial;font-size:11px;color:#92400e'>"
        f"&#128249; CSV exact: {csv_exact_cnt} | &#128270; Text: {text_cnt} | "
        f"&#128652; {len(transit_info)} transit | &#9888; {len(suspicious)} unconfirmed | "
        f"<span style='color:#16a34a;font-weight:700'>F1</span>/<span style='color:#2563eb;font-weight:700'>F2</span>/<span style='color:#7c3aed;font-weight:700'>F3</span> = Most Frequent Locations</div>")

    parts=[]
    parts.append('<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">')
    parts.append('<meta name="viewport" content="width=device-width,initial-scale=1">')
    parts.append(f'<title>Movement Map &mdash; {phone}</title>')
    parts.append('<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/leaflet@1.9.3/dist/leaflet.css"/>')
    parts.append('<script src="https://cdn.jsdelivr.net/npm/leaflet@1.9.3/dist/leaflet.js"></script>')
    parts.append('<script src="https://cdn.jsdelivr.net/npm/leaflet-ant-path@1.1.2/dist/leaflet-ant-path.min.js"></script>')
    parts.append("""<style>
*{box-sizing:border-box;margin:0;padding:0}html,body{width:100%;height:100%;font-family:Arial,sans-serif}
#map{position:absolute;inset:0;z-index:0}
.panel{position:fixed;z-index:1000;background:#fff;border-radius:12px;box-shadow:0 4px 20px rgba(0,0,0,.18)}
#hdr{top:10px;left:50%;transform:translateX(-50%);padding:10px 24px;text-align:center;white-space:nowrap}
#tl{top:70px;left:10px;width:272px;padding:10px 12px;max-height:calc(100vh - 90px);overflow-y:auto}
#lgd{bottom:90px;right:10px;padding:12px 16px;min-width:192px}
.tl-row{display:flex;gap:8px;align-items:flex-start;padding:5px 4px;border-bottom:1px solid #f3f4f6;cursor:pointer;border-radius:4px}
.tl-row:hover{background:#f0f9ff}
.tl-num{min-width:22px;height:22px;border-radius:50%;color:#fff;font-size:10px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0;margin-top:2px}
.tl-info{font-size:11px;line-height:1.6}.tl-km{color:#9ca3af;font-size:10px}.tl-date{color:#6b7280}.tl-cnt{color:#374151;font-weight:600}
#tl::-webkit-scrollbar{width:4px}#tl::-webkit-scrollbar-thumb{background:#d1d5db;border-radius:2px}
</style>""")
    parts.append('</head><body>')
    parts.append(f'<div id="hdr" class="panel"><b style="font-size:15px;color:#111827">&#128205; Movement Analysis &mdash; {phone}</b><br>'
        f'<span style="font-size:12px;color:#6b7280">{operator} &nbsp;|&nbsp; {period_start} &rarr; {period_end} &nbsp;|&nbsp; {total:,} records &nbsp;|&nbsp; {len(steps)} confirmed steps</span></div>')
    parts.append(f'<div id="tl" class="panel"><div style="font-size:13px;font-weight:700;color:#111827;margin-bottom:6px">&#128203; Timeline <span style="font-size:10px;font-weight:400;color:#9ca3af">click to zoom</span></div>'
        f'<div style="margin-bottom:6px;padding:3px 6px;background:#f9fafb;border-radius:5px;font-size:10px;color:#374151">&#128249;CSV &nbsp; &#128270;Thana &nbsp; &#127758;District</div>'+tl_rows+'</div>')
    parts.append(f'<div id="lgd" class="panel"><b style="font-size:13px">Legend</b><div style="line-height:2.1;font-size:12px;margin-top:6px">'
        f'&#127968; <span style="color:#c62828">Home ({home_label})</span><br>'
        +dot('#1565C0')+str(home_dist_val)+'<br>'
        +dot('#E65100')+'Nearby &lt;50 km<br>'
        +dot('#7B1FA2')+'Far 50&ndash;150 km<br>'
        +dot('#1B5E20')+'Very far &gt;150 km<br>'
        +lgd_arrow+' Direction<br>'
        +dot('#f59e0b')+'&#9888; Unconfirmed'
        +'</div><hr style="margin:6px 0;border-color:#f3f4f6"><span style="font-size:10px;color:#9ca3af">Circle &#8733; records | Click for details</span></div>')
    parts.append(note_bar)
    parts.append('<div id="map"></div><script>')
    parts.append("var map=L.map('map',{center:[23.5,90.3],zoom:7,zoomControl:true});")
    parts.append("L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{attribution:'&copy; <a href=\"https://www.openstreetmap.org/copyright\">OpenStreetMap</a>',subdomains:'abc',maxZoom:19}).addTo(map);")
    parts.append(all_js)
    parts.append('</script></body></html>')
    return '\n'.join(parts).encode('utf-8')


# ═══════════════════════════════════════════════════════════════
# CDR LINK ANALYSIS — Multi-CDR Connection & Co-location Analysis
# ═══════════════════════════════════════════════════════════════

import streamlit as st
import pandas as pd
import re
import json
import math
from itertools import combinations
from collections import defaultdict


# ─────────────────────────────────────────────────────────────
# Helper functions
# ─────────────────────────────────────────────────────────────

def _clean_phone(v):
    d = re.sub(r'[^0-9]', '', str(v))
    if len(d) >= 10:
        if d.startswith('880') and len(d) == 13: return '0' + d[3:]
        if d.startswith('88') and len(d) == 12:  return '0' + d[2:]
        if d.startswith('0')  and len(d) == 11:  return d
        if len(d) == 10:                           return '0' + d
    return d if len(d) >= 8 else None


def _norm_lac_cid(v):
    s = str(v).strip()
    if s.endswith('.0') and s[:-2].isdigit(): s = s[:-2]
    if s.isdigit() and len(s) > 1: s = str(int(s))
    return s


def _is_valid_number(p):
    # Valid BD numbers: 10 digits (1XXXXXXXXX), 11 digits (01XXXXXXXXX), 13 digits (8801XXXXXXXXX)
    if not p: return False
    d = re.sub(r'[^0-9]', '', str(p))
    return len(d) in (10, 11, 13)


def _is_promotional(phone):
    """Detect promotional/service numbers — short codes, non-standard formats."""
    if not phone: return True
    d = re.sub(r'[^0-9]', '', str(phone))
    # Short codes (< 8 digits), or starts with non-BD prefix
    if len(d) < 8: return True
    # Common BD promotional prefixes
    promo_prefixes = ['16', '17600', '17601', '17602', '17603', '01500', '01600',
                      '17700', '17800', '17900', '10', '11', '12', '13', '14', '15']
    for p in promo_prefixes:
        if d.startswith(p) and len(d) < 11: return True
    return False

def _remove_anomalies(df):
    """Remove anomalous records. Works with both original and normalized column names."""
    # Usage type column — try both cases
    ut_col = next((c for c in df.columns if c.lower().replace(' ','_') == 'usage_type'), None)
    if ut_col is None: return df
    valid_types = ['MOC','MTC','SMSMO','SMSMT','SMS-MT','CALL-RCF']
    df = df[df[ut_col].str.upper().isin(valid_types)].copy().reset_index(drop=True)
    # Party B column
    pb_col = next((c for c in df.columns if c.lower().replace(' ','_') == 'party_b'), None)
    if pb_col:
        df = df.reset_index(drop=True)
        df['_pb_clean'] = [_clean_phone(v) for v in df[pb_col]]
        df = df[[bool(_is_valid_number(p)) for p in df['_pb_clean']]].reset_index(drop=True)
        df = df[[not _is_promotional(p) for p in df['_pb_clean']]].reset_index(drop=True)
    return df


def _load_cdr(uploaded_file, label):
    """Load and clean a CDR Excel file — public entry point.
    Reads file bytes from the Streamlit upload object (not hashable),
    then delegates to the @st.cache_data inner function.
    """
    try:
        raw_bytes = uploaded_file.read()
        uploaded_file.seek(0)  # Reset for any later reads
        return _load_cdr_bytes(raw_bytes, label)
    except Exception as e:
        st.error(f"Error loading {label}: {e}")
        return None, None, 0


@st.cache_data(show_spinner=False, max_entries=20)
def _load_cdr_bytes(file_bytes, label):
    """Cached CDR loader — takes bytes so Streamlit can hash the input.
    Re-runs only when file content changes (different hash).
    """
    try:
        xl = pd.ExcelFile(io.BytesIO(file_bytes))
        # Pick sheet with most rows AND has CDR columns
        def _sheet_score(s):
            try:
                tmp = pd.read_excel(xl, sheet_name=s, nrows=3)
                cols = [c.lower() for c in tmp.columns]
                has_cdr = any(k in ' '.join(cols) for k in ['party','usage','lac','cell','operator'])
                return (len(pd.read_excel(xl, sheet_name=s)) if has_cdr else 0)
            except Exception: return 0
        scores = {s: _sheet_score(s) for s in xl.sheet_names}
        best_sheet = max(scores, key=scores.get)
        if scores[best_sheet] == 0:
            best_sheet = xl.sheet_names[0]
        df = pd.read_excel(xl, sheet_name=best_sheet)

        # Normalize columns
        col_map = {
            'start':         ['start', 'start_dttime', 'date', 'datetime'],
            'operator':      ['operator', 'provider_name', 'network'],
            'party_a':       ['party a', 'aparty', 'party_a', 'msisdn', 'a_number'],
            'party_b':       ['party b', 'bparty', 'party_b', 'b_number'],
            'duration':      ['call duration', 'call_duration', 'duration'],
            'usage_type':    ['usage type', 'usage_type', 'call_type', 'type'],
            'cell_type':     ['cell type', 'cell_type', 'network_type', 'technology'],
            'lac':           ['lac id', 'lac_id', 'lac', 'lacstarta', 'mccstarta'],
            'cid':           ['cell id', 'cell_id', 'ci', 'cistarta'],
            'address':       ['bts address', 'address', 'location', 'site_address'],
        }
        df.columns = [c.lower().strip() for c in df.columns]
        rename = {}
        for std, variants in col_map.items():
            for v in variants:
                if v in df.columns and std not in rename.values():
                    rename[v] = std
                    break
        df = df.rename(columns=rename)

        df = df.reset_index(drop=True)  # Fix duplicate index issue
        df['_label'] = label
        if 'party_a' in df.columns:
            df['_phone_a'] = [_clean_phone(v) for v in df['party_a']]
        else:
            df['_phone_a'] = label
        if 'party_b' in df.columns:
            df['_phone_b'] = [_clean_phone(v) for v in df['party_b']]
        else:
            df['_phone_b'] = None

        # Determine subject phone (most frequent Party A)
        df = df.reset_index(drop=True)
        if 'party_a' in df.columns:
            pa_counts = df['party_a'].value_counts()
            subject_raw = pa_counts.index[0] if not pa_counts.empty else label
            subject_phone = _clean_phone(subject_raw) or label
        else:
            subject_phone = label

        df['_subject'] = subject_phone
        df['start'] = pd.to_datetime(df['start'], errors='coerce') if 'start' in df.columns else pd.NaT
        df['lac_n'] = df['lac'].apply(_norm_lac_cid) if 'lac' in df.columns else ''
        df['cid_n'] = df['cid'].apply(_norm_lac_cid) if 'cid' in df.columns else ''

        # Remove anomalies
        before = len(df)
        df = _remove_anomalies(df)
        df = df.reset_index(drop=True)
        after = len(df)

        # ── GPS Enrichment (Link Analysis) — Single CDR Pass1+Pass2 logic হুবহু mirror ──
        # GP, Banglalink, Robi, Teletalk সব operator-এর জন্য district/thana/label সহ।
        # CDR Analysis part-এ কোনো পরিবর্তন নেই — শুধু এই block।
        try:
            import os as _la_os, tempfile as _la_tf, re as _la_re, csv as _la_csv
            _la_CELL_DIR = _la_os.path.join(_la_tf.gettempdir(), "celltower_cache")
            _la_os.makedirs(_la_CELL_DIR, exist_ok=True)

            # ── Step 1: Operator detect ──
            _la_ops = set()
            if 'operator' in df.columns:
                for _ov in df['operator'].dropna().astype(str).unique():
                    _ol = _ov.lower()
                    if 'grameen' in _ol or ' gp' in _ol or _ol.startswith('gp'):   _la_ops.add('gp')
                    elif 'banglalink' in _ol or ' bl' in _ol or _ol.startswith('bl'): _la_ops.add('bl')
                    elif 'robi' in _ol or 'airtel' in _ol:  _la_ops.add('robi')
                    elif 'teletalk' in _ol:                   _la_ops.add('teletalk')
            if not _la_ops: _la_ops = {'gp','bl','robi','teletalk'}

            # ── Step 2: Download CSVs ──
            _ensure_cell_tower_cache(_la_ops)

            # ── Config: Single CDR section-এর HF_FILES_CFG/OP_GEN_FILE হুবহু copy ──
            _la_CFG = {
                "GP_2G.csv":           {"enc":"utf-8",   "lac":"lac",      "cid":"cellid",       "lat":"latitude","lon":"longitude","addr":"address",      "thana":"thana","district":"district"},
                "GP_3G.csv":           {"enc":"utf-8",   "lac":"lac",      "cid":"cellid",       "lat":"latitude","lon":"longitude","addr":"address",      "thana":"thana","district":"district"},
                "GP_4G.csv":           {"enc":"latin-1", "lac":"lac",      "cid":"cell_id",      "lat":"latitude","lon":"longitude","addr":"address",      "thana":"thana","district":"district"},
                "Robi_2G.csv":         {"enc":"latin-1", "lac":"lac",      "cid":"cell_id",      "lat":"latitude","lon":"longitude","addr":"address",      "thana":"thana","district":"district"},
                "Robi_4G.csv":         {"enc":"latin-1", "lac":"enodebid", "cid":"cell_id",      "lat":"latitude","lon":"longitude","addr":"address",      "thana":"thana","district":"district","lac_alt":"tac"},
                "Banglalink_2G3G.csv": {"enc":"latin-1", "lac":"lac",      "cid":"ci",           "lat":"latitude","lon":"longitude","addr":"site address", "thana":"thana","district":"district"},
                "Banglalink_4G.csv":   {"enc":"latin-1", "lac":"tac",      "cid":"eutrancellid", "lat":"latitude","lon":"longitude","addr":"site address", "thana":"thana","district":"district"},
                "Teletalk.csv":        {"enc":"utf-8",   "lac":"lac/ tal", "cid":"ci /tac",      "lat":"latitude","lon":"longitude","addr":"full address"},
            }
            _la_OP_GEN = {
                ("gp","2g"):"GP_2G.csv",  ("gp","3g"):"GP_3G.csv",   ("gp","4g"):"GP_4G.csv",
                ("robi","2g"):"Robi_2G.csv", ("robi","3g"):"Robi_2G.csv", ("robi","4g"):"Robi_4G.csv",
                ("bl","2g"):"Banglalink_2G3G.csv", ("bl","3g"):"Banglalink_2G3G.csv", ("bl","4g"):"Banglalink_4G.csv",
                ("teletalk","2g"):"Teletalk.csv", ("teletalk","3g"):"Teletalk.csv",   ("teletalk","4g"):"Teletalk.csv",
            }
            _la_GPS_THR = 40

            # ── Helpers (Single CDR section-এর _norm_id/_addr_tokens/_op_key/_gen_from_cell_type mirror) ──
            def _la_norm(v):
                s = str(v).strip()
                if s.endswith(".0") and s[:-2].isdigit(): s = s[:-2]
                if s.isdigit() and len(s) > 1: s = str(int(s))
                return s

            _la_STOP = {'house','road','vill','village','post','dist','district','para','area',
                        'ward','block','lane','floor','flat','plot','holding','section','street',
                        'avenue','building','tower','police','station','office','market','bazar',
                        'bazaar','union','upazila','thana','mouza','mouja','mauja','north','south',
                        'east','west','central','new','old','more','moor','ganj','pur','nagar','gram','palli'}
            def _la_toks(s):
                s = _la_re.sub(r'[^a-z0-9 ]', ' ', str(s).lower())
                return set(t for t in s.split() if len(t)>4 and t not in _la_STOP and not t.isdigit())

            def _la_op(op_str):
                o = str(op_str).lower()
                if "grameen" in o or " gp" in o or o.startswith("gp"): return "gp"
                if "banglalink" in o or " bl" in o or o.startswith("bl"): return "bl"
                if "robi" in o or "airtel" in o: return "robi"
                if "teletalk" in o: return "teletalk"
                return None

            def _la_gen(ct):
                ct = str(ct).lower().strip()
                if ct in ("","nan","none","-","n/a"): return None
                if "4g" in ct or "lte" in ct: return "4g"
                if "3g" in ct or "wcdma" in ct or "umts" in ct: return "3g"
                if "2g" in ct or "gsm" in ct: return "2g"
                return None

            # ── BD District coords for Signal 1 (copy from Single CDR section) ──
            _la_DIST_COORDS = {
                'dhaka':(23.8103,90.4125),'chittagong':(22.3569,91.7832),'sylhet':(24.8949,91.8687),
                'rajshahi':(24.3636,88.6241),'khulna':(22.8456,89.5403),'barisal':(22.7010,90.3535),
                'rangpur':(25.7439,89.2752),'mymensingh':(24.7471,90.4203),'comilla':(23.4607,91.1809),
                'narayanganj':(23.6238,90.4996),'gazipur':(24.0022,90.4264),'tangail':(24.2513,89.9167),
                'bogra':(24.8465,89.3773),'sirajganj':(24.4535,89.7006),'pabna':(24.0064,89.2372),
                'dinajpur':(25.6279,88.6338),'cox bazar':(21.4272,92.0058),'feni':(23.0230,91.3960),
                'noakhali':(22.8696,91.0998),'chandpur':(23.2333,90.6518),'brahmanbaria':(23.9570,91.1115),
                'habiganj':(24.3745,91.4153),'moulvibazar':(24.4829,91.7774),'sunamganj':(25.0658,91.3950),
                'kishoreganj':(24.4449,90.7764),'narsingdi':(23.9324,90.7154),'faridpur':(23.6070,89.8429),
                'jessore':(23.1664,89.2080),'kushtia':(23.9014,89.1204),'khagrachhari':(23.1193,91.9847),
                'bandarban':(22.1953,92.2184),'rangamati':(22.6522,92.1615),'gaibandha':(25.3288,89.5287),
                'netrakona':(24.8701,90.7268),'jamalpur':(24.9375,89.9377),'sherpur':(25.0198,90.0172),
                'nilphamari':(25.9310,88.8563),'lalmonirhat':(25.9923,89.2847),'kurigram':(25.8054,89.6363),
                'thakurgaon':(26.0319,88.4616),'panchagarh':(26.3411,88.5551),'joypurhat':(25.1007,89.0227),
                'naogaon':(24.7936,88.9312),'natore':(24.4204,88.9872),'narail':(23.1722,89.5120),
                'satkhira':(22.7185,89.0705),'bagerhat':(22.6602,89.7854),'meherpur':(23.7620,88.6317),
                'chuadanga':(23.6401,88.8416),'jhenaidah':(23.5447,89.1530),'magura':(23.4876,89.4196),
                'gopalganj':(23.0050,89.8267),'madaripur':(23.1641,90.2012),'shariatpur':(23.2423,90.4347),
                'rajbari':(23.7574,89.6441),'munshiganj':(23.5422,90.5305),'manikganj':(23.8630,90.0024),
                'lakshmipur':(22.9449,90.8412),
            }

            def _la_hav(la1,lo1,la2,lo2):
                R=6371.0; dlat=_math_mod.radians(la2-la1); dlon=_math_mod.radians(lo2-lo1)
                a=_math_mod.sin(dlat/2)**2+_math_mod.cos(_math_mod.radians(la1))*_math_mod.cos(_math_mod.radians(la2))*_math_mod.sin(dlon/2)**2
                return R*2*_math_mod.asin(_math_mod.sqrt(max(0,a)))

            # ── Step 3: Load cell files (Single CDR _load_cell_file logic mirror) ──
            def _la_load_file(fname, cfg):
                """exact:{(lac,cid):(lat,lon,toks,addr_str)} multi:{cid:[(lat,lon,toks,addr)]} robi4g_multi:{(dlac,sector):[...]}"""
                fpath = _la_os.path.join(_la_CELL_DIR, fname)
                if not (_la_os.path.isfile(fpath) and _la_os.path.getsize(fpath)>5000):
                    return {},{},{}
                exact={}; multi={}; r4g_multi={}
                try:
                    # encoding fallback
                    _cdf = None
                    for _enc in [cfg["enc"],"latin-1","utf-8"]:
                        try:
                            _cdf = pd.read_csv(fpath, dtype=str, encoding=_enc,
                                               low_memory=False, on_bad_lines='skip')
                            break
                        except Exception: continue
                    if _cdf is None or _cdf.empty: return {},{},{}
                    _cdf.columns = [c.lower().strip() for c in _cdf.columns]

                    # Column detect with fallback (Single CDR pattern)
                    lc=cfg["lac"]; ci=cfg["cid"]; la=cfg["lat"]; lo=cfg["lon"]
                    ac=cfg.get("addr",""); thc=cfg.get("thana",""); dic=cfg.get("district","")
                    if lc not in _cdf.columns:
                        for alt in ["enodebid","enodeb_id","tac","lac","lac_id","enbid"]:
                            if alt in _cdf.columns: lc=alt; break
                    if ci not in _cdf.columns:
                        for alt in ["eutrancellid","cell_id","cellid","ci","cid"]:
                            if alt in _cdf.columns: ci=alt; break
                    if la not in _cdf.columns:
                        for alt in ["latitude","lat","y"]:
                            if alt in _cdf.columns: la=alt; break
                    if lo not in _cdf.columns:
                        for alt in ["longitude","lon","lng","x"]:
                            if alt in _cdf.columns: lo=alt; break
                    if not all(c in _cdf.columns for c in [lc,la,lo]) or ci not in _cdf.columns:
                        return {},{},{}

                    has_addr  = ac  and ac  in _cdf.columns
                    has_thana = thc and thc in _cdf.columns
                    has_dist  = dic and dic in _cdf.columns
                    alt_ci_col= "eutrancellid" if "eutrancellid" in _cdf.columns and "eutrancellid"!=ci else None
                    is_r4g    = (fname=="Robi_4G.csv")
                    is_tt     = (fname=="Teletalk.csv")

                    def _ck(col): return col.replace(" ","_").replace("-","_").replace("/","_")
                    lc_k=_ck(lc); ci_k=_ck(ci); la_k=_ck(la); lo_k=_ck(lo)
                    ac_k=_ck(ac) if ac else ""; thc_k=_ck(thc) if thc else ""; dic_k=_ck(dic) if dic else ""
                    alt_ci_k =_ck(alt_ci_col) if alt_ci_col else ""

                    for row in _cdf.itertuples(index=False):
                        try:
                            rd=row._asdict()
                            lat=float(rd[la_k]); lon=float(rd[lo_k])
                            if not (20<=lat<=27 and 88<=lon<=93): continue
                            # Build addr_str: address + thana + district (same as Single CDR)
                            parts=[]
                            if has_addr:
                                a=str(rd.get(ac_k,rd.get(ac,""))).strip().strip('"')
                                if a and a!='nan': parts.append(a)
                            if has_thana:
                                t=str(rd.get(thc_k,rd.get(thc,""))).strip()
                                if t and t!='nan': parts.append(t)
                            if has_dist:
                                d=str(rd.get(dic_k,rd.get(dic,""))).strip()
                                if d and d!='nan': parts.append(d)
                            addr_str=", ".join(parts)
                            toks=_la_toks(addr_str)
                            lv=_la_norm(rd[lc_k]); cv=_la_norm(rd[ci_k])
                            dist_val=str(rd.get(dic_k,rd.get(dic,""))).strip() if has_dist else ""
                            if dist_val in ('nan','None',''): dist_val=""
                            thana_val=str(rd.get(thc_k,rd.get(thc,""))).strip() if has_thana else ""
                            if thana_val in ('nan','None',''): thana_val=""
                            k=(lv,cv)
                            if k not in exact: exact[k]=(lat,lon,toks,addr_str,thana_val,dist_val)
                            multi.setdefault(cv,[]).append((lat,lon,toks,addr_str))
                            # Robi 4G: enodebid//100 = derived_lac, cell_id last 2 digits = sector
                            if is_r4g:
                                try:
                                    _enb=int(float(str(rd[lc_k])))
                                    _dlac=str(_enb//100); _sec=str(int(float(str(rd[ci_k])))%100).zfill(2)
                                    r4g_multi.setdefault((_dlac,_sec),[]).append((lat,lon,toks,addr_str))
                                except Exception: pass
                            # alt_ci fallback (BL 4G eutrancellid)
                            if alt_ci_k and alt_ci_k in rd:
                                cv2=_la_norm(rd[alt_ci_k]); k2=(lv,cv2)
                                if k2 not in exact: exact[k2]=(lat,lon,toks,addr_str,thana_val,dist_val)
                                multi.setdefault(cv2,[]).append((lat,lon,toks,addr_str))
                        except Exception: continue

                    # Teletalk: CGI/ECGI post-loop (itertuples cannot handle slash in col name)
                    if is_tt:
                        try:
                            _cgi_c=next((c for c in _cdf.columns if 'cgi' in c or 'ecgi' in c),None)
                            _lat_c=next((c for c in _cdf.columns if c=='latitude'),None)
                            _lon_c=next((c for c in _cdf.columns if c=='longitude'),None)
                            _adr_c=next((c for c in _cdf.columns if 'full' in c and 'address' in c),None)
                            if _cgi_c and _lat_c and _lon_c:
                                for _,_tr in _cdf.iterrows():
                                    try:
                                        _clat=float(_tr[_lat_c]); _clon=float(_tr[_lon_c])
                                        if not (20<=_clat<=27 and 88<=_clon<=93): continue
                                        _cgi_v=_la_norm(_tr[_cgi_c])
                                        if not _cgi_v or _cgi_v in ('nan','','0'): continue
                                        _caddr=str(_tr[_adr_c]).strip() if _adr_c else ''
                                        _ctoks=_la_toks(_caddr)
                                        _k_cgi=('0',_cgi_v)
                                        if _k_cgi not in exact:
                                            exact[_k_cgi]=(_clat,_clon,_ctoks,_caddr,'','')
                                    except Exception: continue
                        except Exception:
                            logger.debug('LA Teletalk CGI index error', exc_info=True)
                except Exception:
                    logger.debug('_la_load_file failed: %s', fname, exc_info=True)
                return exact, multi, r4g_multi

            # ── Determine files to load ──
            _la_load_fnames = set()
            for _op in _la_ops:
                _gens = set()
                if 'cell_type' in df.columns:
                    for _ct in df['cell_type'].dropna().unique():
                        _g = _la_gen(_ct)
                        if _g: _gens.add(_g)
                if not _gens: _gens = {"2g","3g","4g"}
                for _g in _gens:
                    _fn = _la_OP_GEN.get((_op,_g))
                    if _fn: _la_load_fnames.add(_fn)
            if not _la_load_fnames: _la_load_fnames = set(_la_CFG.keys())

            _la_files = {}
            for _fn in _la_load_fnames:
                _cfg2 = _la_CFG.get(_fn,{})
                if not _cfg2: continue
                _ex,_mu,_r4g = _la_load_file(_fn, _cfg2)
                _la_files[_fn] = {"exact":_ex,"multi":_mu,"robi4g_multi":_r4g}

            # ── LAC cluster cache (Signal 3) ──
            _la_lac_cache = {}
            for _fn,_fd in _la_files.items():
                for (lv,cv),(lat,lon,*_rest) in _fd["exact"].items():
                    if lv and lv not in ('0','nan'):
                        _la_lac_cache.setdefault(lv,[]).append((lat,lon))
            _la_lac_stats = {}
            for lv, pts in _la_lac_cache.items():
                if len(pts) >= 2:
                    lats=[p[0] for p in pts]; lons=[p[1] for p in pts]
                    clat=_statistics_mod.median(lats); clon=_statistics_mod.median(lons)
                    spread=(_statistics_mod.stdev([_la_hav(clat,clon,lt,ln) for lt,ln in pts])
                            if len(pts)>=3 else 15)
                    _la_lac_stats[lv]=(clat,clon,spread)

            # ── GPS confidence (Signal 1–4, same weights as Single CDR) ──
            def _la_conf(lat, lon, cdr_addr, lv, same_laci, nbrs):
                score=0
                addr_lo=str(cdr_addr).lower() if cdr_addr else ""
                matched=[(dlat,dlon) for dn,(dlat,dlon) in _la_DIST_COORDS.items() if dn in addr_lo]
                if not matched:
                    score+=12
                else:
                    mn=min(_la_hav(lat,lon,dlat,dlon) for dlat,dlon in matched)
                    score+=(25 if mn<=30 else 15 if mn<=60 else 5 if mn<=120 else 0)
                if nbrs:
                    nb_lats=[r[0] for r in nbrs if r[0] is not None]
                    nb_lons=[r[1] for r in nbrs if r[1] is not None]
                    if len(nb_lats)>=2:
                        mdn_lat=_statistics_mod.median(nb_lats); mdn_lon=_statistics_mod.median(nb_lons)
                        d=_la_hav(lat,lon,mdn_lat,mdn_lon)
                        score+=(35 if d<=20 else 20 if d<=60 else 8 if d<=150 else 0)
                    else: score+=15
                else: score+=15
                if lv and lv in _la_lac_stats:
                    clat,clon,spread=_la_lac_stats[lv]; tol=max(spread*3,30)
                    d=_la_hav(lat,lon,clat,clon)
                    score+=(25 if d<=tol else 10 if d<=tol*2 else 0)
                else: score+=12
                if same_laci:
                    close=sum(1 for r in same_laci if r[0] is not None and _la_hav(lat,lon,r[0],r[1])<=25)
                    ratio=close/len(same_laci)
                    score+=(15 if ratio>=0.7 else 8 if ratio>=0.4 else 0)
                else: score+=8
                return score

            # ── Pass 1: GPS candidate per row ──
            if _la_files and 'cell_id' in df.columns:
                _la_lac_col = ('lac' if 'lac' in df.columns else
                               'lac_n' if 'lac_n' in df.columns else None)
                if _la_lac_col:
                    _la_cands = []   # (lat|None, lon|None, addr, dist, thana) per row
                    _la_rlac  = []   # lac key per row
                    _la_rlaci = []   # (lac,cid) per row

                    for _,row in df.iterrows():
                        lv  = _la_norm(row.get(_la_lac_col,''))
                        cv  = _la_norm(row.get('cell_id',''))
                        k   = (lv,cv)
                        cdr_toks = _la_toks(row.get('address',''))
                        op_k  = _la_op(row.get('operator',''))  if 'operator'  in df.columns else None
                        gen_k = _la_gen(row.get('cell_type','')) if 'cell_type' in df.columns else None

                        # Operator+Generation priority order (same as Single CDR)
                        fntry=[]
                        if op_k and gen_k:
                            p=_la_OP_GEN.get((op_k,gen_k))
                            if p and p in _la_files: fntry.append(p)
                            for g in ["4g","3g","2g"]:
                                if g!=gen_k:
                                    fb=_la_OP_GEN.get((op_k,g))
                                    if fb and fb in _la_files and fb not in fntry: fntry.append(fb)
                        elif op_k:
                            for g in ["4g","3g","2g"]:
                                fb=_la_OP_GEN.get((op_k,g))
                                if fb and fb in _la_files and fb not in fntry: fntry.append(fb)
                        if not fntry: fntry=list(_la_files.keys())

                        f_lat=None; f_lon=None; f_addr=''; f_dist=''; f_thana=''; _was_exact=False
                        for fn in fntry:
                            fd=_la_files[fn]; ex=fd["exact"]; mu=fd["multi"]
                            if k in ex:
                                f_lat,f_lon,_,f_addr,f_thana,f_dist=ex[k]; _was_exact=True; break
                            # CID+address token fallback (score ≥ 2)
                            elif cv in mu and cdr_toks:
                                bm=None; bs=0
                                for lt,ln,ct,ca in mu[cv]:
                                    sc=len(cdr_toks&ct) if cdr_toks and ct else 0
                                    if sc>bs: bs=sc; bm=(lt,ln,ca)
                                if bm and bs>=2: f_lat,f_lon,f_addr=bm; break
                            # Robi 2G: cell_id-only fallback (LAC mismatch)
                            elif cv in mu and op_k=="robi" and gen_k in ("2g","3g"):
                                cands=mu[cv]
                                if not cdr_toks and len(cands)==1:
                                    f_lat,f_lon,_,f_addr=cands[0]; break
                                elif cdr_toks:
                                    cdr_lo=str(row.get('address','')).lower()
                                    pk=None
                                    for lt,ln,ct2,ca in cands:
                                        if any(kw in cdr_lo for kw in ['dhaka','chittagong','sylhet','rajshahi','khulna','barisal','rangpur','mymensingh'] if kw in ca.lower()):
                                            pk=(lt,ln,ca); break
                                    if not pk and cands: pk=(cands[0][0],cands[0][1],cands[0][3])
                                    if pk: f_lat,f_lon,f_addr=pk; break

                        # Teletalk: CGI/ECGI direct (address column blank in CDR)
                        if f_lat is None and op_k=="teletalk":
                            tf="Teletalk.csv"
                            if tf in _la_files:
                                kc=('0',cv)
                                if kc in _la_files[tf]["exact"]:
                                    f_lat,f_lon,_,f_addr,f_thana,f_dist=_la_files[tf]["exact"][kc]; _was_exact=True

                        # Robi 4G: enodebid//100 = lac, last 2 digits = sector
                        if f_lat is None and op_k=="robi" and gen_k=="4g":
                            r4fn=_la_OP_GEN.get(("robi","4g"))
                            if r4fn and r4fn in _la_files:
                                r4m=_la_files[r4fn].get("robi4g_multi",{})
                                try:
                                    cv_s=cv.lstrip("0") or "0"
                                    if len(cv_s)>=2:
                                        sec=cv_s[-2:].zfill(2); r4c=r4m.get((lv,sec),[])
                                        if r4c:
                                            if len(r4c)==1: f_lat,f_lon,_,f_addr=r4c[0]
                                            elif cdr_toks:
                                                br=None; sr=-1
                                                for lt,ln,ct3,ca in r4c:
                                                    sc=len(cdr_toks&ct3) if cdr_toks and ct3 else 0
                                                    if sc>sr: sr=sc; br=(lt,ln,ca)
                                                if br: f_lat,f_lon,f_addr=br
                                            else: f_lat,f_lon,_,f_addr=r4c[0]
                                except Exception:
                                    logger.debug('suppressed exception', exc_info=True)

                        # dist/thana fallback from addr_str if still empty
                        if f_lat is not None and not f_dist and ',' in f_addr:
                            f_dist=f_addr.split(',')[-1].strip()

                        _la_cands.append((f_lat,f_lon,f_addr,f_dist,f_thana))
                        _la_rlac.append(lv)
                        _la_rlaci.append((lv,cv))

                    # Signal 4 pre-build
                    _la_laci_gps={}
                    for i,(clat,clon,*_) in enumerate(_la_cands):
                        if clat is not None: _la_laci_gps.setdefault(_la_rlaci[i],[]).append((clat,clon))

                    # ── Pass 2: confidence scoring ──
                    NW=4
                    _la_lats=[]; _la_lons=[]; _la_methods=[]
                    _la_labels=[]; _la_dists=[]; _la_thanas=[]

                    for i,(clat,clon,caddr,cdist,cthana) in enumerate(_la_cands):
                        if clat is None:
                            _la_lats.append(None); _la_lons.append(None)
                            _la_methods.append('none'); _la_labels.append('')
                            _la_dists.append(''); _la_thanas.append('')
                            continue
                        lv=_la_rlac[i]; k4=_la_rlaci[i]
                        lo_i=max(0,i-NW); hi_i=min(len(_la_cands),i+NW+1)
                        nbrs=[(c[0],c[1]) for j,c in enumerate(_la_cands[lo_i:hi_i],lo_i)
                              if j!=i and c[0] is not None]
                        same_laci=[(lt,ln) for lt,ln in _la_laci_gps.get(k4,[])
                                   if not(abs(lt-clat)<1e-9 and abs(ln-clon)<1e-9)]
                        cdr_addr_row=''
                        try: cdr_addr_row=df.iloc[i].get('address','')
                        except Exception: pass
                        conf=_la_conf(clat,clon,cdr_addr_row,lv,same_laci,nbrs)
                        if conf>=_la_GPS_THR:
                            _la_lats.append(clat); _la_lons.append(clon)
                            _la_methods.append('cell_exact'); _la_labels.append(caddr)
                            _la_dists.append(cdist); _la_thanas.append(cthana)
                        else:
                            _la_lats.append(None); _la_lons.append(None)
                            _la_methods.append('none'); _la_labels.append('')
                            _la_dists.append(''); _la_thanas.append('')

                    df['cell_lat']       = _la_lats
                    df['cell_lon']       = _la_lons
                    df['loc_method']     = _la_methods
                    df['cell_csv_label'] = _la_labels
                    df['csv_district']   = _la_dists
                    df['csv_thana']      = _la_thanas

        except Exception:
            logger.debug('Link Analysis GPS enrichment failed', exc_info=True)

        return df, subject_phone, before - after
    except Exception as e:
        st.error(f"Error loading {label}: {e}")
        return None, None, 0


# ── Known BD carrier / service number prefixes ──────────────────────────────
_CARRIER_NUMBERS = {
    # GP
    '01700000000', '01711200200', '01800000000', '01711500500',
    # Banglalink
    '01911100100', '01900000000',
    # Robi / Airtel
    '01600000600', '01800000600',
    # Teletalk
    '01500000500',
    # Common shortcodes (normalized to 11-digit)
    '01600162471', '01600162476', '01600162477',
}
_CARRIER_PREFIXES_SHORT = [
    '162', '163', '164', '165',   # 5-digit BD shortcodes
    '1600', '1700', '1800', '1900',  # operator info lines
]

def _is_carrier_number(num: str) -> bool:
    """True if num looks like a carrier/service/IVR number, not a real subscriber."""
    d = re.sub(r'[^0-9]', '', str(num))
    if d in _CARRIER_NUMBERS: return True
    if len(d) < 8: return True
    if len(set(d)) <= 2 and len(d) >= 8: return True
    for pfx in _CARRIER_PREFIXES_SHORT:
        if d.startswith(pfx) and len(d) < 11: return True
    return False


def _build_connections(dfs, exclude_noise=True):
    """Build connection table from multiple CDRs.
    exclude_noise=True → carrier/service/IVR numbers are dropped before building connections.
    """
    # connections[phone_b] = {subject: {call_out, call_in, sms_out, sms_in}}
    connections = defaultdict(lambda: defaultdict(lambda: {
        'call_out': 0, 'call_in': 0, 'sms_out': 0, 'sms_in': 0,
        'total': 0, 'duration': 0.0
    }))

    for df in dfs:
        subject = df['_subject'].iloc[0]
        ut_col = next((c for c in df.columns if c.lower().replace(' ','_') == 'usage_type'), None)
        if ut_col is None: continue
        pb_col = next((c for c in df.columns if c.lower().replace(' ','_') == 'party_b'), '_phone_b')
        dur_col = next((c for c in df.columns if 'duration' in c.lower()), None)

        # Vectorized: work on whole DataFrame at once
        import pandas as _pd_c
        _tmp = df.copy()
        # Phone B
        if '_phone_b' in _tmp.columns:
            _tmp['_pb'] = _tmp['_phone_b'].fillna('')
        else:
            _tmp['_pb'] = _tmp[pb_col].fillna('').apply(
                lambda x: _clean_phone(str(x)))
        # Valid numbers only
        _tmp = _tmp[_tmp['_pb'].apply(_is_valid_number)]
        if exclude_noise:
            _tmp = _tmp[~_tmp['_pb'].apply(_is_carrier_number)]
            _tmp = _tmp[~_tmp['_pb'].apply(_is_promotional)]
        if _tmp.empty: continue

        # Usage type flags
        # ── Fix 2 (_ci false positive) ───────────────────────────────────────
        # আগের logic: _ut.str.contains('IN') ব্যবহার করা হতো — এটি 'ROAMING_IN',
        # 'LOGIN', 'VPN_IN' ইত্যাদি string-এও match করতো → false positive call-in।
        # নতুন logic: শুধুমাত্র BD CDR-এ স্বীকৃত incoming call type-গুলো exact match।
        _ut = _tmp[ut_col].fillna('').str.upper().str.strip()
        _tmp['_co']  = ((_ut.str.contains('MOC') | _ut.str.contains('OUT'))
                        & ~_ut.str.contains('MTC')).astype(int)
        # MTC, CALL-RCF, CALL_IN, CALLIN — exact known incoming types
        _tmp['_ci']  = (
            _ut.str.contains(r'\bMTC\b',    regex=True) |
            _ut.str.contains(r'\bRCF\b',    regex=True) |
            _ut.str.contains('CALL-RCF')                |
            _ut.str.contains('CALL-IN')                 |
            _ut.str.contains('CALLIN')
        ).astype(int)
        _tmp['_so']  = _ut.str.contains('SMSMO').astype(int)
        _tmp['_si']  = (_ut.str.contains('SMSMT') | _ut.str.contains('SMS-MT')).astype(int)
        _tmp['_tot'] = 1
        _tmp['_dur'] = (_tmp[dur_col].fillna(0).astype(float) / 60
                        if dur_col else 0)

        # Group by phone_b
        _grp = _tmp.groupby('_pb').agg(
            _co=('_co','sum'), _ci=('_ci','sum'),
            _so=('_so','sum'), _si=('_si','sum'),
            _tot=('_tot','sum'), _dur=('_dur','sum')
        )
        # Vectorized: iterate over grouped result as dict (faster than iterrows)
        for pb, row in _grp.to_dict('index').items():
            connections[pb][subject]['call_out']  += int(row['_co'])
            connections[pb][subject]['call_in']   += int(row['_ci'])
            connections[pb][subject]['sms_out']   += int(row['_so'])
            connections[pb][subject]['sms_in']    += int(row['_si'])
            connections[pb][subject]['total']     += int(row['_tot'])
            connections[pb][subject]['duration']  += float(row['_dur'])

    # Filter: remove SMS-only entries PER SUBJECT
    # ── Fix 1 (per-subject SMS filter) ───────────────────────────────────────
    # আগের logic: যেকোনো একটি subject-এ call থাকলে সব subject-এর জন্য number রাখা হতো।
    # ফলে Subject A শুধু SMS করলেও common contact হিসেবে দেখাতো, call count = 0 হওয়া সত্বেও।
    # নতুন logic: প্রতিটি subject-এর জন্য আলাদাভাবে check — call নেই মানে সেই subject-এর
    # entry বাদ। তারপর যদি কোনো subject-ই না থাকে, number টি সম্পূর্ণ বাদ।
    sms_filtered = defaultdict(lambda: defaultdict(lambda: {
        'call_out': 0, 'call_in': 0, 'sms_out': 0, 'sms_in': 0,
        'total': 0, 'duration': 0.0
    }))
    for pb, subj_data in connections.items():
        for sub, sd in subj_data.items():
            # এই subject-এর জন্য অন্তত ১টি call (MOC বা MTC) থাকতে হবে
            if sd['call_out'] + sd['call_in'] > 0:
                sms_filtered[pb][sub] = sd
        # কোনো subject-ই call করেনি → number টি drop
        if sms_filtered[pb]:
            pass  # keep
        else:
            # defaultdict-এ empty entry তৈরি হয়ে গেছে, মুছে দাও
            del sms_filtered[pb]
    return sms_filtered


def _build_first_last_dates(dfs):
    """
    প্রতিটি (subject, contact) pair-এর first ও last contact date বের করে।
    Returns: dict { (subject, phone_b): {'first': date, 'last': date} }
    """
    result = {}
    for df in dfs:
        if 'start' not in df.columns: continue
        subject = df['_subject'].iloc[0]
        pb_col = '_phone_b' if '_phone_b' in df.columns else None
        if pb_col is None:
            pb_col = next((c for c in df.columns if c.lower().replace(' ','_') == 'party_b'), None)
        if pb_col is None: continue

        ut_col = next((c for c in df.columns if c.lower().replace(' ','_') == 'usage_type'), None)
        if ut_col is None: continue

        _tmp = df[['start', pb_col, ut_col]].copy()
        _tmp['_pb'] = _tmp[pb_col].fillna('').apply(lambda x: _clean_phone(str(x)))
        _tmp = _tmp[_tmp['_pb'].apply(_is_valid_number)]
        if _tmp.empty: continue

        # Only call rows
        _ut = _tmp[ut_col].fillna('').str.upper().str.strip()
        _is_call = (
            _ut.str.contains('MOC') | _ut.str.contains('MTC') |
            _ut.str.contains('OUT') | _ut.str.contains(r'\bRCF\b', regex=True)
        )
        _tmp = _tmp[_is_call]
        if _tmp.empty: continue

        _grp = _tmp.groupby('_pb')['start'].agg(['min', 'max'])
        for pb, row in _grp.iterrows():
            key = (subject, pb)
            result[key] = {
                'first': row['min'].strftime('%Y-%m-%d') if pd.notna(row['min']) else '—',
                'last':  row['max'].strftime('%Y-%m-%d') if pd.notna(row['max']) else '—',
            }
    return result


def _build_noise_analysis(dfs, connections):
    """
    Carrier/service numbers যেগুলো common contact হিসেবে দেখাচ্ছে কিন্তু
    আসলে operator IVR/promo — সেগুলো flag করে।
    Returns:
      carrier_list  — list of dicts (number, shared_by, total_calls, reason)
      clean_common  — common contacts with carrier numbers removed
    """
    carrier_list = []
    clean_common = {}

    for pb, subj_dict in connections.items():
        if len(subj_dict) < 2: continue  # শুধু common contacts check করব
        is_carrier  = _is_carrier_number(pb)
        is_promo    = _is_promotional(pb)
        total_calls = sum(d.get('call_out', 0) + d.get('call_in', 0) for d in subj_dict.values())

        if is_carrier or is_promo:
            reason = 'Carrier/IVR' if is_carrier else 'Promotional/Service'
            carrier_list.append({
                'Number':      pb,
                'Shared By':   len(subj_dict),
                'Total Calls': total_calls,
                'Reason':      reason,
            })
        else:
            clean_common[pb] = subj_dict

    return carrier_list, clean_common


def _build_suspicious_patterns(dfs, window_min=30):
    """
    দুই ধরনের suspicious pattern detect করে:

    1. Mirror Call — Subject A → X call করার ±window_min মিনিটের মধ্যে
                     Subject B → same X-কে call করে (বা একই X → B-কে)।
                     মানে: A ও B একই number-এর সাথে প্রায় একই সময়ে যোগাযোগ করেছে।

    2. Relay Pattern — A → X call, তারপর X → B call (±window_min মিনিটের মধ্যে),
                       যেখানে A ও B দুজনেই subject। X একটা intermediary হিসেবে কাজ করছে।

    Returns: (mirror_rows, relay_rows) — দুটো list of dicts
    """
    # Subject phone → DataFrame mapping
    subj_dfs = {}
    for df in dfs:
        if 'start' not in df.columns: continue
        subj = df['_subject'].iloc[0]
        ut_col = next((c for c in df.columns if c.lower().replace(' ','_') == 'usage_type'), None)
        pb_col = '_phone_b' if '_phone_b' in df.columns else next(
            (c for c in df.columns if c.lower().replace(' ','_') == 'party_b'), None)
        if not ut_col or not pb_col: continue

        _tmp = df[['start', pb_col, ut_col]].copy()
        _tmp['_pb'] = _tmp[pb_col].fillna('').apply(lambda x: _clean_phone(str(x)))
        _tmp = _tmp[_tmp['_pb'].apply(_is_valid_number)].copy()
        _ut = _tmp[ut_col].fillna('').str.upper().str.strip()
        _is_call = (
            _ut.str.contains('MOC') | _ut.str.contains('MTC') |
            _ut.str.contains('OUT') | _ut.str.contains(r'\bRCF\b', regex=True)
        )
        _tmp = _tmp[_is_call][['start', '_pb']].copy()
        _tmp = _tmp.sort_values('start').reset_index(drop=True)
        subj_dfs[subj] = _tmp

    subjects = list(subj_dfs.keys())
    window_td = pd.Timedelta(minutes=window_min)

    mirror_rows = []
    relay_rows  = []

    # ── Mirror pattern: subject pairs ──────────────────────────────────────
    for i in range(len(subjects)):
        for j in range(i + 1, len(subjects)):
            sa, sb = subjects[i], subjects[j]
            dfa, dfb = subj_dfs[sa], subj_dfs[sb]

            # Common numbers between A and B
            nums_a = set(dfa['_pb'].unique())
            nums_b = set(dfb['_pb'].unique())
            common_nums = nums_a & nums_b

            for num in common_nums:
                if _is_carrier_number(num) or _is_promotional(num): continue
                times_a = dfa[dfa['_pb'] == num]['start'].sort_values().values
                times_b = dfb[dfb['_pb'] == num]['start'].sort_values().values

                # Find pairs within window
                hits = []
                bi = 0
                # Normalize to int64 nanoseconds safely (handles numpy.datetime64 & pandas.Timestamp)
                def _to_ns(t):
                    ts = pd.Timestamp(t)
                    return ts.value  # always int64 nanoseconds

                _window_ns = int(window_td.total_seconds() * 1e9)
                times_a_ns = [_to_ns(t) for t in times_a]
                times_b_ns = [_to_ns(t) for t in times_b]

                for ta_ns in times_a_ns:
                    while bi < len(times_b_ns) and times_b_ns[bi] < ta_ns - _window_ns:
                        bi += 1
                    for k in range(bi, len(times_b_ns)):
                        tb_ns = times_b_ns[k]
                        diff = abs(tb_ns - ta_ns) / 1e9  # nanoseconds → seconds
                        if diff <= window_min * 60:
                            hits.append({
                                'Subject A': sa, 'Subject B': sb,
                                'Common Number': num,
                                'Time A': pd.Timestamp(ta_ns).strftime('%Y-%m-%d %H:%M'),
                                'Time B': pd.Timestamp(tb_ns).strftime('%Y-%m-%d %H:%M'),
                                'Gap (min)': round(diff / 60, 1),
                            })
                        elif tb_ns > ta_ns + _window_ns:
                            break

                if hits:
                    # Deduplicate: same number-এর অনেক instance থাকলে প্রথম ৩টা দেখাও
                    mirror_rows.extend(hits[:3])

    # ── Relay pattern: A → X → B ────────────────────────────────────────────
    # প্রতিটি subject-এর CDR-এ যেসব number আছে, সেগুলো দিয়ে cross-check
    for i in range(len(subjects)):
        for j in range(len(subjects)):
            if i == j: continue
            sa, sb = subjects[i], subjects[j]
            dfa, dfb = subj_dfs[sa], subj_dfs[sb]

            # X = numbers that appear in A's CDR (A called X)
            # AND in B's CDR (X called B, i.e. B received from X — but we only have B's
            # outgoing/incoming perspective, so X appears as party_b in B's CDR too)
            nums_a = set(dfa['_pb'].unique())
            nums_b = set(dfb['_pb'].unique())
            relay_candidates = nums_a & nums_b  # X appears in both

            # X cannot be a subject itself
            relay_candidates -= set(subjects)

            for x in relay_candidates:
                if _is_carrier_number(x) or _is_promotional(x): continue
                times_ax = dfa[dfa['_pb'] == x]['start'].sort_values().values  # A↔X
                times_xb = dfb[dfb['_pb'] == x]['start'].sort_values().values  # X↔B

                hits = []
                bi = 0
                _to_ns2 = lambda t: pd.Timestamp(t).value
                _window_ns2 = int(window_td.total_seconds() * 1e9)
                times_ax_ns = [_to_ns2(t) for t in times_ax]
                times_xb_ns = [_to_ns2(t) for t in times_xb]

                for ta_ns in times_ax_ns:
                    # Find X-B calls that happen AFTER A-X within window
                    while bi < len(times_xb_ns) and times_xb_ns[bi] < ta_ns:
                        bi += 1
                    for k in range(bi, len(times_xb_ns)):
                        tb_ns = times_xb_ns[k]
                        diff = (tb_ns - ta_ns) / 1e9
                        if 0 <= diff <= window_min * 60:
                            hits.append({
                                'Subject A': sa,
                                'Relay Number (X)': x,
                                'Subject B': sb,
                                'A-X Time': pd.Timestamp(ta_ns).strftime('%Y-%m-%d %H:%M'),
                                'X-B Time': pd.Timestamp(tb_ns).strftime('%Y-%m-%d %H:%M'),
                                'Relay Gap (min)': round(diff / 60, 1),
                            })
                        elif tb_ns > ta_ns + _window_ns2:
                            break

                if hits:
                    relay_rows.extend(hits[:3])

    # Sort by gap ascending (tighter = more suspicious)
    mirror_rows.sort(key=lambda r: r['Gap (min)'])
    relay_rows.sort(key=lambda r: r['Relay Gap (min)'])
    return mirror_rows, relay_rows


def _haversine_km(lat1, lon1, lat2, lon2):
    """Calculate distance in km between two GPS points."""
    R = 6371
    dlat = _math_mod.radians(lat2 - lat1)
    dlon = _math_mod.radians(lon2 - lon1)
    a = _math_mod.sin(dlat/2)**2 + _math_mod.cos(_math_mod.radians(lat1)) * _math_mod.cos(_math_mod.radians(lat2)) * _math_mod.sin(dlon/2)**2
    return R * 2 * _math_mod.asin(_math_mod.sqrt(max(0, a)))


def _ensure_cell_tower_cache(operators=None):
    """
    HuggingFace থেকে cell tower CSV download করে CELL_DIR-এ রাখো।
    CDR Analysis আগে না চালালেও Link Analysis-এ GPS পাওয়া যাবে।
    operators: set of 'gp','bl','robi','teletalk' — None মানে সব
    """
    # Convert mutable set → frozenset so the cached helper can be called with a hashable key
    _ops_key = frozenset(operators) if operators else None
    _ensure_cell_tower_cache_inner(_ops_key)


@st.cache_data(show_spinner=False, ttl=86400, max_entries=5)
def _ensure_cell_tower_cache_inner(operators_frozen=None):
    """
    Actual download logic — cacheable because operators_frozen is a frozenset (hashable).
    ttl=86400 → 24 ঘণ্টা পর re-check করবে, তার আগে re-download হবে না।
    """
    import os as _osc, tempfile as _tfc  # shutil → _shutil_mod
    if _requests_mod is None:
        logger.warning("requests not installed — cell tower CSV download skipped")
        return

    CELL_DIR = _osc.path.join(_tfc.gettempdir(), "celltower_cache")
    _osc.makedirs(CELL_DIR, exist_ok=True)

    HF_REPO = "Faruk131086/Celltower"

    def _hf_tok():
        try:
            import streamlit as _stc
            return _stc.secrets.get("HF_TOKEN", None)
        except Exception:
            return _osc.environ.get("HF_TOKEN", None)

    # Operator → files mapping
    OP_FILES = {
        'gp':       ['GP_2G.csv', 'GP_3G.csv', 'GP_4G.csv'],
        'robi':     ['Robi_2G.csv', 'Robi_4G.csv'],
        'bl':       ['Banglalink_2G3G.csv', 'Banglalink_4G.csv'],
        'teletalk': ['Teletalk.csv'],
    }

    # Which files to download
    files_needed = set()
    if operators_frozen:
        for op in operators_frozen:
            files_needed.update(OP_FILES.get(op, []))
    else:
        for flist in OP_FILES.values():
            files_needed.update(flist)

    token = _hf_tok()
    hdrs = {"User-Agent": "Mozilla/5.0"}
    if token:
        hdrs["Authorization"] = f"Bearer {token}"

    for fname in files_needed:
        local = _osc.path.join(CELL_DIR, fname)
        # Skip if already cached and large enough
        if _osc.path.isfile(local) and _osc.path.getsize(local) > 5000:
            continue
        # Try HuggingFace direct URL
        url = f"https://huggingface.co/datasets/{HF_REPO}/resolve/main/{fname}"
        try:
            r = _requests_mod.get(url, headers=hdrs, stream=True, timeout=120)
            r.raise_for_status()
            with open(local, "wb") as _fw:
                for chunk in r.iter_content(65536):
                    if chunk: _fw.write(chunk)
            if _osc.path.getsize(local) < 1000:
                _osc.remove(local)  # bad download
        except Exception:
            # Try hf_hub_download fallback
            try:
                # hf_hub_download → _hf_hub_download (module-level import)
                path = _hf_hub_download(
                    repo_id=HF_REPO, filename=fname,
                    repo_type="dataset", token=token,
                    local_dir=CELL_DIR
                )
                if path and _osc.path.abspath(path) != _osc.path.abspath(local):
                    _shutil_mod.copy2(path, local)
            except Exception:
                logger.debug('suppressed exception', exc_info=True)


@st.cache_resource(show_spinner=False)
def _load_cell_tower_gps(cell_dir=None):
    """Load GPS coordinates from cell tower CSV files.
    Decorated with st.cache_resource: large dict, shared as singleton across reruns.
    """
    import os, glob
    cell_dict = {}
    # Try common upload locations
    search_dirs = [_UPLOAD_CACHE_DIR, '/mnt/user-data/uploads', '/tmp/celltower_cache']
    if cell_dir: search_dirs.insert(0, cell_dir)

    csv_configs = [
        ('GP_2G.csv',  'lac', 'cellid',      'latitude', 'longitude'),
        ('GP_4G.csv',  'lac', 'cell_id',      'latitude', 'longitude'),
        ('2G.csv',     'lac', 'cellid',        'latitude', 'longitude'),
        ('4G.csv',     'lac', 'cell_id',       'latitude', 'longitude'),
        ('Robi_4G.csv','enodebid','cell_id',   'latitude', 'longitude'),
        ('Robi_2G.csv','lac', 'cell_id',       'latitude', 'longitude'),
        ('Banglalink_4G.csv','tac','eutrancellid','latitude','longitude'),
        ('Banglalink_2G3G.csv','lac','ci',     'lat',      'lon'),
    ]

    for d in search_dirs:
        if not os.path.isdir(d): continue
        for fname, lac_col, cid_col, lat_col, lon_col in csv_configs:
            fpath = os.path.join(d, fname)
            if not os.path.isfile(fpath): continue
            try:
                csv = pd.read_csv(fpath, dtype=str, encoding='latin-1', low_memory=False)
                csv.columns = [c.lower().strip() for c in csv.columns]
                _lc = lac_col if lac_col in csv.columns else next((c for c in csv.columns if 'enodebid' in c or c=='lac'),None)
                _cc = cid_col if cid_col in csv.columns else next((c for c in csv.columns if 'cell_id' in c or c=='cellid' or c=='ci'),None)
                _la = lat_col if lat_col in csv.columns else 'lat'
                _lo = lon_col if lon_col in csv.columns else 'lon'
                if not (_lc and _cc and _la in csv.columns and _lo in csv.columns): continue
                for row in csv.itertuples(index=False):
                    try:
                        rd = row._asdict()
                        lat = float(rd[_la]); lon = float(rd[_lo])
                        if not (19<=lat<=27 and 87<=lon<=93): continue
                        lv = str(rd[_lc]).strip().split('.')[0]
                        cv = str(rd[_cc]).strip().split('.')[0]
                        if lv.isdigit() and len(lv)>1: lv=str(int(lv))
                        if cv.isdigit() and len(cv)>1: cv=str(int(cv))
                        k = (lv, cv)
                        if k not in cell_dict: cell_dict[k] = (lat, lon)
                    except Exception: pass
            except Exception: pass
    return cell_dict


def _build_colocation(dfs, window_min=30, radius_km=3.0):
    """
    Common Location Analysis:
    ২+ subject একই সময়ে (window_min মিনিটের মধ্যে) একই এলাকায় (radius_km km) ছিল কিনা।

    GPS resolution — CDR Analysis-এর build_movement_map-এর হুবহু priority:
      Priority 1: cell_lat/cell_lon (cell tower CSV থেকে, loc_method='cell_exact')
      Priority 2: BTS address text parse → P.S: → thana → THANA_GPS (±3-8 km)
                                         → DIST: → district → DISTRICT_GPS (±10-20 km)

    ভিন্ন operator হলেও GPS haversine দিয়ে তুলনা করা হয়।
    একই operator হলে GPS + same LAC+CID উভয়ই চেক করা হয়।
    """
    # math → _math_mod (module-level import)

    results = []
    if len(dfs) < 2:
        return results

    # ── Thana/District GPS — module-level BD_THANA_GPS / BD_DISTRICT_GPS ──
    # (এক জায়গায় রাখা হয়েছে — movement map ও co-location একই GPS table ব্যবহার করে)
    THANA_GPS    = BD_THANA_GPS
    DISTRICT_GPS = BD_DISTRICT_GPS

    INVALID_PATTERNS = [
        'MOUZA NOT FOUND','NOT FOUND IN AG','CAAB PERMISSION',
        'PERMISSION FOUND','MOUZA-MOUZA',
    ]

    DIST_NORM = {
        'Gaibanda':'Gaibandha','Bogura':'Bogra',
        'Sirajgonj':'Sirajganj','Cumilla':'Comilla',
        'Bogra Sadar South New':'Bogra',
    }

    def _text_gps(addr_str):
        """
        Co-location GPS parse — same conditional logic as build_movement_map.
        Priority:
          2a. P.S: আছে → Thana → District
          2b. P.S: নেই → Keyword scan → District fallback
          3.  শুধু miss হলে → Nominatim (last resort)
        """
        if not addr_str:
            return None
        a = str(addr_str).upper()
        if any(p in a for p in INVALID_PATTERNS):
            return None

        # P.S: detect
        mt = re.search(
            r'P[\.\s]*/?\s*S[\.\:\s\-/]+([A-Z][A-Z\s\-]{2,}?)(?:[,\.\n]|DIST|$)', a)
        has_ps = mt is not None
        thana  = mt.group(1).strip().rstrip('.,- ').title() if mt else ''

        # DIST: → district
        district = ''
        md = re.search(
            r'DIST[\.\:\s]+([A-Z][A-Z\s\-]{2,}?)(?:[,\.\n]|$|\s+BD)', a)
        if md:
            district = md.group(1).strip().rstrip('.,- ').title()

        if has_ps:
            # ── PATH A: P.S: আছে ─────────────────────────────────────────
            if not district:
                parts = [p.strip() for p in re.split(r'[,،]', a) if len(p.strip()) > 2]
                parts = [re.sub(r'\b(BD|BANGLADESH|\d{4,})\b', '', p).strip() for p in parts]
                parts = [p for p in parts if p and not p.isdigit()]
                if parts:
                    last = parts[-1].rstrip('.').title()
                    if last.replace(' ', '').isalpha() and len(last) >= 4:
                        district = last
                    if len(parts) >= 2 and not thana:
                        sl = parts[-2].rstrip('.').title()
                        if len(sl) >= 4: thana = sl

            if thana:
                for tk in [thana, thana.replace(' Sadar', '').strip()]:
                    if tk in THANA_GPS:
                        g = THANA_GPS[tk]
                        d = DIST_NORM.get(district, district) or tk.split()[0]
                        return g[0], g[1], d, thana, 'text_thana'

            if district:
                d = DIST_NORM.get(district, district)
                if d in DISTRICT_GPS:
                    g = DISTRICT_GPS[d]
                    return g[0], g[1], d, thana, 'text_district'
                for k, v in DISTRICT_GPS.items():
                    if k.lower() in d.lower() or d.lower() in k.lower():
                        return v[0], v[1], k, thana, 'text_district'

        else:
            # ── PATH B: P.S: নেই → Keyword scan → District fallback ──────
            res = _keyword_scan_gps(a)
            if res:
                return res[0], res[1], res[2], res[3], res[4]

            # District from comma tokens
            parts = [p.strip() for p in re.split(r'[,،]', a) if len(p.strip()) > 2]
            parts = [re.sub(r'\b(BD|BANGLADESH|\d{4,})\b', '', p).strip() for p in parts]
            parts = [p for p in parts if p and not p.isdigit()]
            if parts:
                last = parts[-1].rstrip('.').title()
                if last.replace(' ', '').isalpha() and len(last) >= 4:
                    district = last
            if district:
                d = DIST_NORM.get(district, district)
                if d in DISTRICT_GPS:
                    g = DISTRICT_GPS[d]
                    return g[0], g[1], d, '', 'text_district'
                for k, v in DISTRICT_GPS.items():
                    if k.lower() in d.lower() or d.lower() in k.lower():
                        return v[0], v[1], k, '', 'text_district'

        # ── LAST RESORT: Nominatim ────────────────────────────────────────
        res2 = _nominatim_geocode(addr_str)
        if res2:
            return res2[0], res2[1], res2[2], res2[3], res2[4]

        return None

    def _resolve_gps(row):
        """
        একটি CDR row-এর GPS বের করো।
        Priority 1: cell_lat/cell_lon (CSV match)
        Priority 2: BTS address text parse
        Returns (lat, lon, district, thana, method) or None
        """
        # Priority 1: CSV exact GPS
        lm = str(row.get('loc_method', '') or '').strip()
        if lm == 'cell_exact':
            try:
                clat = row.get('cell_lat')
                clon = row.get('cell_lon')
                if clat is not None and str(clat) not in ('nan', 'None', ''):
                    clat = float(clat)
                    clon = float(clon)
                    if 19 <= clat <= 27 and 87 <= clon <= 93:
                        dist = str(row.get('csv_district', '') or '').strip().title()
                        dist = '' if dist in ('Nan', 'None', 'nan') else dist
                        thana = str(row.get('csv_thana', '') or '').strip().title()
                        return clat, clon, dist, thana, 'csv_exact'
            except Exception:
                logger.debug('suppressed exception', exc_info=True)

        # Priority 2: BTS address text parse
        addr = str(row.get('address', '') or '').strip()
        if not addr:
            # also try 'Bts Address' column name variation
            addr = str(row.get('Bts Address', '') or
                       row.get('bts_address', '') or '').strip()
        if addr:
            res = _text_gps(addr)
            if res:
                return res[0], res[1], res[2], res[3], res[4]

        return None

    # ── প্রতিটি subject-এর GPS-enriched rows তৈরি ──
    prepared = []
    for df in dfs:
        sub = df['_subject'].iloc[0] if '_subject' in df.columns else 'Unknown'
        op  = str(df['operator'].iloc[0]).lower().strip() \
              if 'operator' in df.columns else ''

        sub_df = df[df['start'].notna()].copy()
        if sub_df.empty:
            continue

        # Vectorized GPS resolve using apply (faster than iterrows)
        def _resolve_row(row):
            res = _resolve_gps(row)
            if res:
                return res[0], res[1], res[2], res[3], res[4]
            return None, None, '', '', 'none'

        import pandas as _pd_co
        _resolved = sub_df.apply(_resolve_row, axis=1, result_type='expand')
        _resolved.columns = ['_lat','_lon','_dist','_thana','_method']

        # lac_key vectorized
        _lv = sub_df['lac_n'].astype(str).str.strip() if 'lac_n' in sub_df.columns else _pd_co.Series([''] * len(sub_df))
        _cv = sub_df['cid_n'].astype(str).str.strip() if 'cid_n' in sub_df.columns else _pd_co.Series([''] * len(sub_df))
        _valid_lk = (_lv != '') & (_lv != 'nan') & (_cv != '') & (_cv != 'nan')
        lac_keys = (_lv + '_' + _cv).where(_valid_lk, '').tolist()

        lats    = _resolved['_lat'].tolist()
        lons    = _resolved['_lon'].tolist()
        dists   = _resolved['_dist'].tolist()
        thanas  = _resolved['_thana'].tolist()
        methods = _resolved['_method'].tolist()

        sub_df = sub_df.copy()
        sub_df['_lat']    = lats;    sub_df['_lon']    = lons
        sub_df['_dist']   = dists;   sub_df['_thana']  = thanas
        sub_df['_method'] = methods; sub_df['_lk']     = lac_keys
        sub_df['_op']     = op;      sub_df['_subject'] = sub

        # GPS পাওয়া গেছে এমন rows রাখো
        sub_df = sub_df[sub_df['_lat'].notna()].copy()
        if not sub_df.empty:
            prepared.append(sub_df)

    if len(prepared) < 2:
        return results

    # ── Fix 4 (Subject label mismatch) ───────────────────────────────────────
    # আগের logic: subj_labels enumerate(dfs) দিয়ে তৈরি হতো — কিন্তু dfs এখানে
    # আসলে `prepared` list (GPS-সহ CDR মাত্র)। যদি মূল CDR list-এর মধ্যে একটি
    # GPS-বিহীন হয়ে বাদ পড়ে, তাহলে index shift হয়ে Subject 2 → "Subject 1" দেখাতো।
    # নতুন logic: label মূল `dfs` parameter থেকে phone→label mapping তৈরি।
    # `prepared` list-এ যে subject-ই থাকুক, মূল index অনুযায়ী সঠিক label পাবে।
    _orig_subject_order = []
    for _df in dfs:
        if '_subject' in _df.columns and not _df.empty:
            _orig_subject_order.append(_df['_subject'].iloc[0])

    subj_labels = {
        phone: f"Subject {_orig_subject_order.index(phone) + 1}"
        if phone in _orig_subject_order
        else f"Subject ?"
        for _pdf in prepared
        if '_subject' in _pdf.columns
        for phone in [_pdf['_subject'].iloc[0]]
    }

    # ── প্রতিটি subject pair তুলনা ──
    for df_a, df_b in combinations(prepared, 2):
        if df_a.empty or df_b.empty:
            continue
        sub_a = df_a['_subject'].iloc[0]
        sub_b = df_b['_subject'].iloc[0]
        op_a  = df_a['_op'].iloc[0]
        op_b  = df_b['_op'].iloc[0]
        same_op = bool(op_a and op_b and op_a == op_b)

        # Vectorized time-window merge instead of nested iterrows
        import pandas as _pd_pair
        import numpy as _np_pair

        # Projection includes all the columns we need post-merge.
        # Note: _method must be included (was missing → KeyError downstream).
        _base_cols = ['start','_lat','_lon','_lk','_thana','_dist',
                       '_subject','_op','_method']
        _opt_addr  = ['address','Bts Address']
        _a_cols = _base_cols + [c for c in _opt_addr if c in df_a.columns]
        _b_cols = _base_cols + [c for c in _opt_addr if c in df_b.columns]
        _a = df_a[_a_cols].copy()
        _b = df_b[_b_cols].copy()
        # Rename B's start to preserve B's actual timestamp post-merge
        # (merge_asof keeps only left's key column; without rename, B's time is lost)
        _b = _b.rename(columns={'start': 'start_b'})

        # Sort for merge_asof
        _a = _a.sort_values('start').reset_index(drop=True)
        _b = _b.sort_values('start_b').reset_index(drop=True)

        # ── merge_asof safety: pandas 2.x strict about dtype/NaT/timezone ──
        # 1) Coerce to datetime64[ns]
        _a['start']   = _pd_pair.to_datetime(_a['start'],   errors='coerce')
        _b['start_b'] = _pd_pair.to_datetime(_b['start_b'], errors='coerce')
        # 2) Strip timezone (consistent naive datetimes)
        try:
            if getattr(_a['start'].dt, 'tz', None) is not None:
                _a['start'] = _a['start'].dt.tz_localize(None)
        except Exception:
            logger.debug('suppressed exception', exc_info=True)
        try:
            if getattr(_b['start_b'].dt, 'tz', None) is not None:
                _b['start_b'] = _b['start_b'].dt.tz_localize(None)
        except Exception:
            logger.debug('suppressed exception', exc_info=True)
        # 3) Drop NaT rows
        _a = _a.dropna(subset=['start'])
        _b = _b.dropna(subset=['start_b'])
        # 4) Numeric coords coerce
        for _c in ('_lat','_lon'):
            _a[_c] = _pd_pair.to_numeric(_a[_c], errors='coerce')
            _b[_c] = _pd_pair.to_numeric(_b[_c], errors='coerce')
        _a = _a.dropna(subset=['_lat','_lon'])
        _b = _b.dropna(subset=['_lat','_lon'])
        # 5) Re-sort with stable order + reset index
        _a = _a.sort_values('start',   kind='mergesort').reset_index(drop=True)
        _b = _b.sort_values('start_b', kind='mergesort').reset_index(drop=True)
        if _a.empty or _b.empty:
            continue

        # merge_asof: find nearest B for each A within window
        # Uses left_on / right_on so B's start_b column also flows through to result.
        _win = _pd_pair.Timedelta(minutes=window_min)
        try:
            _merged = _pd_pair.merge_asof(
                _a, _b,
                left_on='start', right_on='start_b',
                tolerance=_win,
                direction='nearest',
                suffixes=('_a','_b')
            )
        except Exception:
            logger.debug('merge_asof failed for subject pair', exc_info=True)
            continue
        _merged = _merged.dropna(subset=['_lat_b','_lon_b','start_b'])
        if _merged.empty:
            continue

        # ── Helper functions defined early (used in both radius + result build) ──
        def _g(row, col, default=''):
            """Safe getter — column may or may not exist."""
            try:
                v = row[col]
            except (KeyError, IndexError):
                return default
            if _pd_pair.isna(v): return default
            return v

        def _vs(series, default=''):
            """Safe string Series: NaN/None/NaT → default."""
            return (series.fillna(default)
                          .astype(str)
                          .replace({'nan': default, 'None': default, 'NaT': default}))

        # Haversine vectorized
        def _hav_vec(lat1, lon1, lat2, lon2):
            R = 6371
            dlat = _np_pair.radians(lat2 - lat1)
            dlon = _np_pair.radians(lon2 - lon1)
            a = (_np_pair.sin(dlat/2)**2 +
                 _np_pair.cos(_np_pair.radians(lat1)) *
                 _np_pair.cos(_np_pair.radians(lat2)) *
                 _np_pair.sin(dlon/2)**2)
            return R * 2 * _np_pair.arcsin(_np_pair.sqrt(_np_pair.clip(a,0,1)))

        _merged['_dist_km'] = _hav_vec(
            _merged['_lat_a'].values, _merged['_lon_a'].values,
            _merged['_lat_b'].values, _merged['_lon_b'].values)

        # ── Fix 4: Adaptive radius per event ────────────────────────────────
        # GPS accuracy by method:
        #   csv_exact   → ±0.5 km  → user-set radius respected as-is
        #   text_thana  → ±5 km    → effective radius = max(radius_km, 8)
        #   text_district→ ±15 km  → effective radius = max(radius_km, 20)
        # If A and B have different accuracy, use the looser of the two.
        # This prevents false negatives when one side has low-precision GPS.
        _ACCURACY_KM = {
            'csv_exact':     0.5,
            'text_thana':    5.0,
            'text_district': 15.0,
            'none':          99.0,
        }
        _RADIUS_FOR_METHOD = {
            'csv_exact':     radius_km,
            'geocoded':      radius_km,        # ±1km → user radius respected
            'keyword_area':  max(radius_km, 5.0),  # ±3km → slightly looser
            'text_thana':    max(radius_km, 8.0),
            'text_district': max(radius_km, 20.0),
            'none':          0.0,
        }
        # Reset index after dropna to avoid alignment issues
        _merged = _merged.reset_index(drop=True)
        _meth_a_col = _vs(_merged['_method_a'])
        _meth_b_col = _vs(_merged['_method_b'])
        # Effective radius = max of the two sides (looser accuracy wins)
        _eff_radius = _pd_pair.Series([
            max(
                _RADIUS_FOR_METHOD.get(str(ma).strip(), radius_km),
                _RADIUS_FOR_METHOD.get(str(mb).strip(), radius_km),
            )
            for ma, mb in zip(_meth_a_col, _meth_b_col)
        ], index=_merged.index)
        # Filter: distance <= effective radius for each row
        _merged = _merged[_merged['_dist_km'] <= _eff_radius].reset_index(drop=True)
        if _merged.empty:
            continue

        # ── Vectorized result build ────────────────────────────────────────
        m = _merged  # shorthand after filtering + reset_index

        # diff_m — vectorized time difference
        _diff = (m['start'] - m['start_b']).dt.total_seconds().abs() / 60.0
        _diff = _diff.fillna(0.0)

        # same_tower — vectorized boolean
        _lk_a_s = _vs(m['_lk_a']); _lk_b_s = _vs(m['_lk_b'])
        _same_tower_v = (
            same_op
            & _lk_a_s.str.len().gt(0)
            & _lk_b_s.str.len().gt(0)
            & (_lk_a_s == _lk_b_s)
        )

        # location strings — thana → district → raw GPS fallback
        _th_a = _vs(m['_thana_a']); _ds_a = _vs(m['_dist_a'])
        _th_b = _vs(m['_thana_b']); _ds_b = _vs(m['_dist_b'])
        _gps_a = m['_lat_a'].map('{:.6f}'.format) + ', ' + m['_lon_a'].map('{:.6f}'.format)
        _gps_b = m['_lat_b'].map('{:.6f}'.format) + ', ' + m['_lon_b'].map('{:.6f}'.format)
        _loc_a = (_th_a.where(_th_a.str.len() > 0, _ds_a)
                       .where((_th_a.str.len() > 0) | (_ds_a.str.len() > 0), _gps_a))
        _loc_b = (_th_b.where(_th_b.str.len() > 0, _ds_b)
                       .where((_th_b.str.len() > 0) | (_ds_b.str.len() > 0), _gps_b))

        # address display — a-side first, fallback b-side
        _addr_a = _vs(m.get('address_a',   _pd_pair.Series([''] * len(m))))
        _bts_a  = _vs(m.get('Bts Address_a', _pd_pair.Series([''] * len(m))))
        _addr_b = _vs(m.get('address_b',   _pd_pair.Series([''] * len(m))))
        _bts_b  = _vs(m.get('Bts Address_b', _pd_pair.Series([''] * len(m))))
        _addr_v = (_addr_a
                   .where(_addr_a.str.len() > 0, _bts_a)
                   .where((_addr_a.str.len() > 0) | (_bts_a.str.len() > 0), _addr_b)
                   .where((_addr_a.str.len() > 0) | (_bts_a.str.len() > 0)
                          | (_addr_b.str.len() > 0), _bts_b))
        _addr_v = _addr_v.str.strip().str[:70]
        _addr_v = _addr_v.where(_addr_v.str.len() > 0, '—')

        # ── Fix 2: Human-readable GPS Source labels ─────────────────────────
        _GPS_LABEL = {
            'csv_exact':     '🛰️ Cell Tower',
            'text_thana':    '📍 Thana (±5km)',
            'text_district': '🗺️ District (±15km)',
            'keyword_area':  '🏘️ Area (±3km)',
            'geocoded':      '🌐 Geocoded (±1km)',
            'none':          '—',
            '':              '—',
        }
        m = _merged  # shorthand for the filtered, reset-indexed merged DataFrame
        _src_a_raw = _vs(m['_method_a'])
        _src_b_raw = _vs(m['_method_b'])
        _src_a_lbl = _src_a_raw.map(lambda v: _GPS_LABEL.get(v, v))
        _src_b_lbl = _src_b_raw.map(lambda v: _GPS_LABEL.get(v, v))

        # Accuracy mismatch flag:
        # একজনের csv_exact, অন্যজনের text_district → ⚠️ mark করো
        _PREC_RANK = {
            'csv_exact':     0,
            'geocoded':      0,   # Nominatim ±1km ≈ csv_exact level
            'keyword_area':  1,   # ±3km ≈ thana level
            'text_thana':    1,
            'text_district': 2,
            'none':          3,
            '':              3,
        }
        _prec_gap = _src_a_raw.map(lambda v: _PREC_RANK.get(v, 3)) - \
                    _src_b_raw.map(lambda v: _PREC_RANK.get(v, 3))
        _accuracy_flag = _prec_gap.abs().map(
            lambda g: '⚠️ Low' if g >= 2 else ('⚡ Med' if g == 1 else '✅ High')
        )

        # Build result DataFrame in one shot
        _subj_a_lbl = f"{subj_labels.get(sub_a, sub_a)} ({sub_a})"
        _subj_b_lbl = f"{subj_labels.get(sub_b, sub_b)} ({sub_b})"

        _res_df = _pd_pair.DataFrame({
            'Subject A':      _subj_a_lbl,
            'Subject B':      _subj_b_lbl,
            'Time A':         m['start'].astype(str).str[:16],
            'Time B':         m['start_b'].astype(str).str[:16],
            'Diff (min)':     _diff.round(1),
            'Distance (km)':  m['_dist_km'].round(2),
            'Same Tower':     _same_tower_v.map({True: '✅', False: '—'}),
            'Accuracy':       _accuracy_flag,
            'GPS Source A':   _src_a_lbl,
            'GPS Source B':   _src_b_lbl,
            'Location A':     _loc_a,
            'Location B':     _loc_b,
            'GPS A':          _gps_a,
            'GPS B':          _gps_b,
            'BTS Address':    _addr_v,
        })
        results.extend(_res_df.to_dict('records'))

    # Deduplicate
    # ── Fix 3 (dedup key minute-precision) ───────────────────────────────────
    # আগের key: Time A[:13] → hour পর্যন্ত। একই ঘণ্টায় দুটো ভিন্ন co-location
    # ঘটনা থাকলে দ্বিতীয়টি বাদ পড়তো (false duplicate)।
    # নতুন key: Time A[:16] → minute পর্যন্ত। + GPS A[:9] (±0.001° ≈ 100m granularity)।
    seen, unique = set(), []
    for r in results:
        key = (r['Subject A'], r['Subject B'], r['Time A'][:16], r['GPS A'][:9])
        if key not in seen:
            seen.add(key)
            unique.append(r)

    unique.sort(key=lambda x: (x['Diff (min)'], x['Distance (km)']))
    return unique[:1000]


def _build_network_html(dfs, connections, subjects, subj_edge_count=None, subj_meta=None, contact_names=None):
    """Network graph — clean labels, delete nodes, filter by connection count.

    contact_names : dict  {phone_str: name_str}  — optional display names for
                    contact nodes (non-subject). When provided, node labels show
                    "Name\nPhone" instead of phone only.
    """
    # math → _math_mod (module-level import)
    if contact_names is None:
        contact_names = {}

    # Subject 0 = primary (magenta/pink like Image 2), rest = normal palette
    colors_subject = ['#db2777','#1d4ed8','#15803d','#7c3aed','#d97706']

    # SVG telephone icon as base64 data-URI for contact nodes
    _PHONE_SVG = (
        "data:image/svg+xml;base64,"
        + __import__('base64').b64encode(
            b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" '
            b'width="36" height="36">'
            b'<rect width="24" height="24" rx="5" fill="#64748b"/>'
            b'<path fill="#fff" d="M6.6 10.8c1.4 2.8 3.8 5.1 6.6 6.6l2.2-2.2c.3-.3.7-.4 1-.2'
            b' 1.1.4 2.3.6 3.6.6.6 0 1 .4 1 1V20c0 .6-.4 1-1 1C10.6 21 3 13.4 3 4c0-.6.4-1 '
            b'1-1h3.5c.6 0 1 .4 1 1 0 1.3.2 2.5.6 3.6.1.3 0 .7-.2 1L6.6 10.8z"/>'
            b'</svg>'
        ).decode()
    )
    _PHONE_SVG_COMMON = (
        "data:image/svg+xml;base64,"
        + __import__('base64').b64encode(
            b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" '
            b'width="36" height="36">'
            b'<rect width="24" height="24" rx="5" fill="#dc2626"/>'
            b'<path fill="#fff" d="M6.6 10.8c1.4 2.8 3.8 5.1 6.6 6.6l2.2-2.2c.3-.3.7-.4 1-.2'
            b' 1.1.4 2.3.6 3.6.6.6 0 1 .4 1 1V20c0 .6-.4 1-1 1C10.6 21 3 13.4 3 4c0-.6.4-1 '
            b'1-1h3.5c.6 0 1 .4 1 1 0 1.3.2 2.5.6 3.6.1.3 0 .7-.2 1L6.6 10.8z"/>'
            b'</svg>'
        ).decode()
    )
    subj_labels = {sub: f"S{i+1}" for i, sub in enumerate(subjects)}

    nodes = {}
    if subj_edge_count is None:
        subj_edge_count = {sub: 99 for sub in subjects}

    # ── Subject nodes ──
    for i, sub in enumerate(subjects):
        edge_cnt   = subj_edge_count.get(sub, 0)
        is_isolated = edge_cnt < 5
        node_color  = colors_subject[i % len(colors_subject)]
        _meta  = (subj_meta or {}).get(sub, {})
        _sname = _meta.get("name", "").strip()
        _sphoto= _meta.get("photo")

        # Label: show name (if given) and number — NO S1/S2 prefix
        if _sname:
            lbl = f"{_sname}\n{sub}"
        else:
            lbl = sub

        subj_title = (
            f"<div style='font-family:Segoe UI,Arial,sans-serif;font-size:14px;"
            f"padding:10px 14px;min-width:220px;line-height:1.8'>"
            f"<b style='font-size:16px;color:#1e3a8a'>\u2b50 Subject {i+1}</b><br>"
            f"<b style='font-size:15px'>\U0001f4f1 {sub}</b>"
            + (f"<br><b style='color:#1e3a8a'>\U0001f464 {_sname}</b>" if _sname else "") +
            f"<hr style='margin:6px 0;border:none;border-top:1px solid #e2e8f0'>"
            + (f"<span style='color:#f59e0b'>\u26a0\ufe0f Few/no common contacts</span>"
               if is_isolated else
               f"<span style='color:#16a34a'>\u2713 Connections: {edge_cnt}</span>") +
            f"<br><span style='color:#94a3b8;font-size:11px'>"
            f"Right-click or use Delete button to remove</span></div>"
        )
        n = {
            'id': sub, 'label': lbl,
            'color': {'background': node_color,
                      'border': '#fbbf24' if is_isolated else '#ffffff',
                      'highlight': {'background': node_color, 'border': '#f59e0b'}},
            'shape': 'star', 'size': 38,
            'borderWidth': 4 if is_isolated else 2, 'borderWidthSelected': 5,
            'font': {'size': 13, 'color': '#000000', 'bold': True,
                     'strokeWidth': 3, 'strokeColor': '#ffffff'},
            'title': subj_title,
            'group': 'isolated_subject' if is_isolated else 'subject',
            'mass': 6,
            '_total': edge_cnt,
            'x': int(600*_math_mod.cos(2*_math_mod.pi*i/max(len(subjects),1)-_math_mod.pi/2)),
            'y': int(600*_math_mod.sin(2*_math_mod.pi*i/max(len(subjects),1)-_math_mod.pi/2)),
        }
        if _sphoto:
            n.update({'shape':'circularImage','image':_sphoto,'size':45,
                      'color':{'border':node_color,'background':'#fff'},'borderWidth':5})
        nodes[sub] = n

    common = {pb for pb, sd in connections.items() if len(sd) >= 2}

    # ── Contact nodes ──
    # Pre-compute max total for importance-ring threshold (top 10%)
    all_totals = [sum(d['total'] for d in sd.values()) for pb, sd in connections.items() if pb not in subjects]
    _importance_thresh = sorted(all_totals, reverse=True)[max(0, len(all_totals)//10 - 1)] if all_totals else 9999

    for pb, subj_dict in connections.items():
        if pb in nodes: continue
        total = sum(d['total'] for d in subj_dict.values())
        is_common = pb in common
        only_sub  = list(subj_dict.keys())[0] if len(subj_dict)==1 else None
        is_iso_c  = (only_sub and subj_edge_count.get(only_sub,99) < 5)
        # ── Importance ring: top-10% by total interaction count ──
        is_important = (total >= _importance_thresh and total >= 10)

        bg     = '#fca5a5' if is_common else ('#fef3c7' if is_iso_c else '#e2e8f0')
        border = '#dc2626' if is_common else ('#d97706' if is_iso_c else '#94a3b8')
        # Importance ring overrides border color (gold ring)
        if is_important:
            border = '#f59e0b'

        # ── Contact name label ──
        _cname = (contact_names or {}).get(pb, '').strip()
        if _cname:
            node_label = f"{_cname}\n{pb}"
        else:
            node_label = pb

        subj_lines = []
        for s, d in subj_dict.items():
            sl = subj_labels.get(s,s)
            calls = d.get('call_out',0)+d.get('call_in',0)
            sms   = d.get('sms_out',0)+d.get('sms_in',0)
            dur   = round(d.get('duration',0),1)
            subj_lines.append(
                f"&nbsp;&nbsp;<b>{sl}</b>: \U0001f4de{calls} \U0001f4ac{sms} \u23f1{dur}min")
        tooltip = (
            f"<div style='font-family:Segoe UI,Arial,sans-serif;font-size:14px;"
            f"padding:10px 14px;min-width:230px;line-height:1.8'>"
            f"<b style='font-size:16px;color:#1e3a8a'>\U0001f4f1 {pb}</b>"
            + (f"<br><b style='color:#1e3a8a'>\U0001f464 {_cname}</b>" if _cname else "") +
            f"<br><span style='color:#64748b;font-size:12px'>"
            f"Shared: {len(subj_dict)} | Total: {total}"
            + (" | <b style='color:#f59e0b'>⭐ High-frequency</b>" if is_important else "") +
            f"</span><br>"
            f"<hr style='margin:6px 0;border:none;border-top:1px solid #e2e8f0'>"
            + "<br>".join(subj_lines) +
            f"<br><span style='color:#94a3b8;font-size:11px'>"
            f"Click = highlight &nbsp;|&nbsp; Delete btn = remove</span></div>"
        )
        # Importance ring → thicker border + slightly larger
        _bw   = 5 if is_important else 1
        _size = min(10+total, 30) + (5 if is_important else 0)
        # Phone icon: common=red bg, normal=grey bg
        _icon = _PHONE_SVG_COMMON if is_common else _PHONE_SVG
        nodes[pb] = {
            'id': pb, 'label': node_label,
            'color': {'background': bg, 'border': border,
                      'highlight': {'background': bg, 'border': '#2563eb'}},
            'shape': 'circularImage',
            'image': _icon,
            'size': max(18, _size),
            'borderWidth': _bw,
            'borderWidthSelected': _bw + 2,
            'font': {'size': 13, 'color': '#000000', 'bold': True,
                     'strokeWidth': 2, 'strokeColor': '#ffffff'},
            'title': tooltip,
            'group': 'common' if is_common else ('iso_c' if is_iso_c else 'contact'),
            'mass': 1,
            '_total': total,
            '_important': is_important,
            '_cname': _cname,
        }

    # ── Edges — combined (undirected): one edge per pair per type ──
    # Call edges: MOC+MTC combined; SMS edges: sms_out+sms_in combined
    edges = []
    eid = 0
    for pb, subj_dict in connections.items():
        for sub, data in subj_dict.items():
            total = data['total']
            if total == 0: continue
            is_common = pb in common
            dur = round(data['duration'], 1)

            ec_common = '#dc2626'

            # ── Merged Call + SMS edge ──────────────────────────────────────
            # SMS no longer drawn as a separate dotted edge.
            # SMS counts are shown in the tooltip; total on the line = calls + sms.
            call_total  = data['call_out'] + data['call_in']
            sms_total   = data['sms_out']  + data['sms_in']
            grand_total = call_total + sms_total
            if grand_total > 0:
                ec = ec_common if is_common else '#2563eb'
                sms_row = (
                    f"<hr style='margin:6px 0;border:none;border-top:1px solid #e2e8f0'>"
                    f"&nbsp;&nbsp;💬 SMS sent: <b>{data['sms_out']}</b><br>"
                    f"&nbsp;&nbsp;💬 SMS received: <b>{data['sms_in']}</b><br>"
                    f"&nbsp;&nbsp;Total SMS: <b>{sms_total}</b>"
                ) if sms_total > 0 else ""
                # Duration badge — red if suspicious
                _dur_badge_col = '#dc2626' if dur >= 30 else '#1e3a8a'
                _dur_badge_txt = (
                    f"<span style='background:#fee2e2;color:#991b1b;"
                    f"border-radius:4px;padding:1px 6px;font-weight:700;font-size:12px'>"
                    f"⚠️ Long call</span> " if dur >= 30 else ""
                )
                edge_title = (
                    f"<div style='font-family:Segoe UI,Arial,sans-serif;font-size:14px;"
                    f"padding:10px 14px;line-height:1.9;min-width:240px'>"
                    f"<b style='font-size:15px;color:{ec}'>📞 {sub} ↔ {pb}</b><br>"
                    f"<hr style='margin:6px 0;border:none;border-top:1px solid #e2e8f0'>"
                    f"&nbsp;&nbsp;MOC (outgoing calls): <b>{data['call_out']}</b><br>"
                    f"&nbsp;&nbsp;MTC (incoming calls): <b>{data['call_in']}</b><br>"
                    f"&nbsp;&nbsp;Total Calls: <b>{call_total}</b><br>"
                    f"&nbsp;&nbsp;⏱ Total Duration: <b style='color:{_dur_badge_col}'>"
                    f"{dur} min</b> {_dur_badge_txt}<br>"
                    f"&nbsp;&nbsp;📊 Avg per call: <b>"
                    f"{round(dur/call_total,1) if call_total>0 else 0} min</b>"
                    f"{sms_row}</div>"
                )
                # Duration-aware edge styling
                _dur_h = int(dur // 60)
                _dur_m = int(dur % 60)
                _dur_label = (f"{_dur_h}h{_dur_m:02d}m" if _dur_h > 0
                              else f"{_dur_m}m" if _dur_m > 0 else "<1m")

                # Width: count + duration both contribute
                _base_w = max(1, min(5, grand_total // 5 + 1))
                _dur_w  = max(0, min(2, int(dur) // 30))   # +1 per 30 min
                _edge_w = _base_w + _dur_w + (2 if is_common else 0)

                # Color: long calls get darker/redder tone
                if dur >= 60:       # 1h+  → deep red (very suspicious)
                    _ec_col = '#991b1b' if is_common else '#b91c1c'
                elif dur >= 30:     # 30m+ → red-orange
                    _ec_col = '#dc2626' if is_common else '#ea580c'
                else:
                    _ec_col = ec   # normal

                # Label: "count\ndur" two-line
                _edge_label = f"{grand_total}\n{_dur_label}"

                # Font color: red for long calls
                _font_color = '#991b1b' if dur >= 30 else '#1e293b'

                edges.append({
                    'id': eid, 'from': sub, 'to': pb,
                    'label': _edge_label,
                    'arrows': {'to': {'enabled': False}},
                    'color': {'color': _ec_col, 'opacity': 0.88},
                    'width': _edge_w,
                    'font': {'size': 10, 'color': _font_color,
                             'strokeWidth': 2, 'strokeColor': '#ffffff',
                             'align': 'middle', 'multi': False},
                    'smooth': {'type': 'dynamic'},
                    'title': edge_title,
                    '_total': grand_total, '_etype': 'call',
                    '_dur': dur,
                    '_origWidth': _edge_w,
                })
                eid += 1

    nodes_json = json.dumps(list(nodes.values()), ensure_ascii=False)
    edges_json = json.dumps(edges, ensure_ascii=False)

    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<script src="https://cdnjs.cloudflare.com/ajax/libs/vis/4.21.0/vis.min.js"></script>
<link  href="https://cdnjs.cloudflare.com/ajax/libs/vis/4.21.0/vis.min.css" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:"Segoe UI",Arial,sans-serif;background:#f8fafc;padding:6px;overflow-x:hidden}}
#network{{width:100%;height:580px;background:#fff;border:1px solid #e2e8f0;border-radius:10px}}
.bar{{background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:6px 10px;
      margin-bottom:6px;display:flex;align-items:center;gap:8px;flex-wrap:wrap}}
.btn{{background:#1e3a8a;color:#fff;border:none;padding:4px 11px;border-radius:6px;
      cursor:pointer;font-size:12px;white-space:nowrap}}
.btn:hover{{background:#1d4ed8}}
.btn.red{{background:#dc2626}}.btn.red:hover{{background:#b91c1c}}
.btn.grn{{background:#16a34a}}.btn.grn:hover{{background:#15803d}}
.btn.orn{{background:#d97706}}.btn.orn:hover{{background:#b45309}}
.btn.del{{background:#7f1d1d}}.btn.del:hover{{background:#991b1b}}
.btn.teal{{background:#0d9488}}.btn.teal:hover{{background:#0f766e}}
.btn.violet{{background:#7c3aed}}.btn.violet:hover{{background:#6d28d9}}
.sl{{display:flex;align-items:center;gap:5px;font-size:11px;color:#475569}}
input[type=range]{{width:80px;accent-color:#2563eb}}
.lgd{{background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:5px 10px;
      display:flex;gap:12px;flex-wrap:wrap;font-size:11px;margin-bottom:6px}}
.li{{display:flex;align-items:center;gap:4px}}
.dot{{width:11px;height:11px;border-radius:50%;flex-shrink:0}}
.ln{{width:16px;height:3px;flex-shrink:0}}
#panel{{display:none;position:absolute;top:8px;right:8px;z-index:999;
        background:#fff;border:2px solid #2563eb;border-radius:12px;
        padding:12px 16px;min-width:255px;max-width:310px;
        box-shadow:0 6px 24px rgba(0,0,0,.18);font-size:13px;line-height:1.8}}
#panelHdr{{display:flex;justify-content:space-between;align-items:center;margin-bottom:4px}}
#panelTag{{font-size:10px;color:#94a3b8;font-weight:700;text-transform:uppercase;letter-spacing:.5px}}
#panelX{{cursor:pointer;color:#94a3b8;font-size:20px;line-height:1}}
#panelX:hover{{color:#dc2626}}
#wrap{{position:relative}}
/* ── Export modal ── */
#exportModal{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.55);
  z-index:9999;align-items:center;justify-content:center}}
#exportModal.show{{display:flex}}
#exportBox{{background:#fff;border-radius:14px;padding:20px 24px;max-width:92vw;
  width:740px;box-shadow:0 12px 40px rgba(0,0,0,.35);position:relative}}
#exportBox h3{{font-size:15px;color:#1e3a8a;margin-bottom:12px;font-weight:700}}
#exportPreview{{width:100%;border:1px solid #e2e8f0;border-radius:8px;
  display:block;margin-bottom:12px;background:#f1f5f9}}
#exportActions{{display:flex;gap:10px;flex-wrap:wrap;align-items:center}}
#exportClose{{position:absolute;top:12px;right:16px;font-size:22px;cursor:pointer;
  color:#94a3b8;line-height:1;background:none;border:none}}
#exportClose:hover{{color:#dc2626}}
#exportStatus{{font-size:12px;color:#16a34a;font-weight:600;margin-left:auto}}
#captureOverlay{{display:none;position:fixed;inset:0;background:rgba(255,255,255,.7);
  z-index:8888;align-items:center;justify-content:center;font-size:15px;
  color:#1e3a8a;font-weight:600;gap:10px}}
#captureOverlay.show{{display:flex}}
</style>
</head>
<body>
<div class="lgd">
  <div class="li"><div class="dot" style="background:#db2777"></div>Primary Subject</div>
  <div class="li"><div class="dot" style="background:#1d4ed8"></div>Subject</div>
  <div class="li"><div class="dot" style="background:#dc2626"></div>Common Contact</div>
  <div class="li"><div class="dot" style="background:#64748b"></div>Single Contact</div>
  <div class="li"><div class="dot" style="background:#e2e8f0;border:3px solid #f59e0b;width:13px;height:13px;"></div>High-freq ⭐</div>
  <div class="li"><div class="ln" style="background:#2563eb"></div>Connection (calls + SMS total on line)</div>
</div>
<div class="bar">
  <button class="btn" onclick="network.fit()">&#x229F; Fit</button>
  <button class="btn" id="physBtn" onclick="togglePhysics()" title="Freeze: stops auto-layout so you can drag nodes freely. Unfreeze: resumes physics.">&#x23F8; Freeze</button>
  <button class="btn red" onclick="showOnlyCommon()">&#128308; Common</button>
  <button class="btn grn" onclick="showAll()">&#128065; All</button>
  <button class="btn del" id="delBtn" onclick="deleteSelected()">&#x1F5D1; Delete</button>
  <button class="btn orn" onclick="undoDelete()">&#x21BA; Undo</button>
  <button class="btn teal" onclick="exportGraphPNG()">&#x1F4F7; Export PNG</button>
  <button class="btn violet" onclick="copyGraphToClipboard()">&#x1F4CB; Copy Image</button>
</div>
<div class="bar">
  <!-- Search box -->
  <div class="sl" style="flex:1;min-width:180px;">
    <span style="font-weight:600;color:#1e3a8a;">&#x1F50D;</span>
    <input type="text" id="searchBox" placeholder="Search number / name…"
      oninput="searchNodes(this.value)"
      onkeydown="handleSearchKey(event)"
      style="flex:1;font-size:11px;padding:3px 7px;border-radius:5px;
             border:1px solid #cbd5e1;background:#f8fafc;color:#0f172a;outline:none;">
    <span id="searchCount" style="font-size:10px;color:#64748b;white-space:nowrap;padding:0 4px;"></span>
    <button class="btn" style="background:#475569;padding:3px 8px;"
      onclick="document.getElementById('searchBox').value='';searchNodes('');document.getElementById('searchCount').textContent='';">✕</button>
  </div>
  <!-- Edge type filter -->
  <span style="font-size:11px;font-weight:600;color:#1e3a8a;">Edge:</span>
  <button class="btn" id="fAll"  onclick="filterEdgeType('all')"  style="background:#0f172a;">All</button>
  <button class="btn" id="fCall" onclick="filterEdgeType('call')" style="background:#1d4ed8;">Calls</button>
  <div class="sl">
    <span style="font-weight:600;color:#1e3a8a;">&#x25A6; Layout:</span>
    <select id="layoutSel" onchange="applyLayout(this.value)"
      style="font-size:11px;padding:2px 6px;border-radius:5px;border:1px solid #cbd5e1;
             background:#f8fafc;color:#1e3a8a;cursor:pointer;">
      <option value="physics">&#x1F300; Physics (default)</option>
      <option value="hierarchyLR">&#x27A1; Hierarchy L→R</option>
      <option value="hierarchyUD">&#x2B07; Hierarchy U→D</option>
      <option value="bipartite">&#x21C4; Bipartite (Subj left/right)</option>
      <option value="circle">&#x25EF; Circle</option>
      <option value="grid">&#x22EE; Grid</option>
    </select>
  </div>
  <button class="btn" id="impRingBtn" onclick="toggleImportanceRing()" title="High-frequency gold ring">&#11088; Ring: ON</button>
  <div class="sl">
    <span>Min conn:</span>
    <input type="range" id="minConn" min="1" max="20" value="1"
           oninput="filterByConnCount(this.value)">
    <span id="minConnVal">1</span>
  </div>
  <div class="sl">
    <span>Node Lbl:</span>
    <input type="range" id="fontSz" min="0" max="22" value="13"
           oninput="changeFontSize(this.value)">
    <span id="fontVal">13</span>
  </div>
  <div class="sl">
    <span>Edge Lbl:</span>
    <input type="range" id="edgeFontSz" min="0" max="20" value="0"
           oninput="changeEdgeFontSize(this.value)">
    <span id="edgeFontVal">0</span>
  </div>
  <div class="sl">
    <span>Node Size:</span>
    <input type="range" id="nodeSz" min="6" max="40" value="14"
           oninput="changeNodeSize(this.value)">
    <span id="nodeVal">14</span>
  </div>
  <span style="font-size:10px;color:#94a3b8;margin-left:auto">
    Scroll=zoom | Drag=move | Click=info | Del=remove
  </span>
</div>
<div id="wrap">
  <div id="network"></div>
  <div id="panel">
    <div id="panelHdr">
      <span id="panelTag">INFO</span>
      <span id="panelX" onclick="closePanel()">&#xd7;</span>
    </div>
    <div id="panelBody"></div>
  </div>
</div>

<!-- ── Capture overlay ── -->
<div id="captureOverlay">
  <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#1e3a8a" stroke-width="2.5">
    <circle cx="12" cy="12" r="9"/><path d="M12 6v6l4 2"/>
  </svg>
  Capturing graph…
</div>

<!-- ── Export modal ── -->
<div id="exportModal">
  <div id="exportBox">
    <button id="exportClose" onclick="closeExportModal()">&#xd7;</button>
    <h3>&#x1F4F7; Network Graph Export</h3>
    <img id="exportPreview" src="" alt="preview"/>
    <div id="exportActions">
      <button class="btn teal" onclick="downloadExportedPNG()">&#x2B07; Download PNG</button>
      <button class="btn violet" onclick="copyExportedToClipboard()">&#x1F4CB; Copy to Clipboard</button>
      <button class="btn" style="background:#475569" onclick="closeExportModal()">Close</button>
      <span id="exportStatus"></span>
    </div>
    <div style="font-size:11px;color:#94a3b8;margin-top:10px;">
      &#x2139;&#xFE0F; PNG download করুন → Word/PowerPoint-এ Insert → Pictures দিয়ে যোগ করুন।
      অথবা Copy করে সরাসরি Ctrl+V দিয়ে paste করুন।
    </div>
  </div>
</div>
<script>
var nodesData = {nodes_json};
var edgesData = {edges_json};
var allNodes  = new vis.DataSet(nodesData);
var allEdges  = new vis.DataSet(edgesData);
// Delete history for undo
var deletedNodes = [];
var deletedEdges = [];
var selectedNodeId = null;
var selectedEdgeId = null;

var network = new vis.Network(
  document.getElementById('network'),
  {{nodes:allNodes, edges:allEdges}},
  {{
    nodes:{{borderWidth:2,shadow:{{enabled:true,size:4}}}},
    edges:{{
      smooth:{{type:'dynamic'}},shadow:false,font:{{size:10,strokeWidth:2,strokeColor:'#ffffff',align:'middle',multi:false}}
    }},
    physics:{{
      enabled:true,solver:'repulsion',
      stabilization:{{iterations:500,updateInterval:20}},
      repulsion:{{centralGravity:.1,springLength:220,springConstant:.04,
                  nodeDistance:200,damping:.10}}
    }},
    interaction:{{hover:true,tooltipDelay:150,navigationButtons:true,
                  hideEdgesOnDrag:true,keyboard:true,
                  multiselect:true}},
    layout:{{improvedLayout:false}}
  }}
);

// Auto-fit after stabilization
network.once('stabilizationIterationsDone', function(){{
  network.fit({{animation:{{duration:600,easingFunction:'easeInOutQuad'}}}});
  // Auto-freeze: nodes stay put, user can drag freely without physics fighting
  network.setOptions({{physics:{{enabled:false}}}});
  physicsOn = false;
  document.getElementById('physBtn').textContent='\u25B6 Unfreeze';
}});
var physicsOn = true;

setTimeout(function(){{if(network)network.fit();}}, 2500);

// Pin a node after dragging so it stays where the user placed it
network.on('dragEnd', function(params){{
  if(params.nodes.length > 0){{
    params.nodes.forEach(function(nodeId){{
      var pos = network.getPositions([nodeId])[nodeId];
      allNodes.update({{id:nodeId, x:pos.x, y:pos.y, fixed:{{x:true,y:true}}}});
    }});
  }}
}});

// Unpin all on double-click (background) so physics can re-run
network.on('doubleClick', function(params){{
  if(params.nodes.length === 0 && params.edges.length === 0){{
    var unpinned = allNodes.get().map(function(n){{
      return {{id:n.id, fixed:{{x:false,y:false}}}};
    }});
    allNodes.update(unpinned);
  }}
}});

// ── Shared: capture graph canvas to dataURL ──
var _exportDataURL = null;

function _captureGraph(callback){{
  // 1. Temporarily hide the info panel and physics UI clutter
  var panel = document.getElementById('panel');
  var prevPanel = panel.style.display;
  panel.style.display = 'none';

  // 2. Show overlay
  var ov = document.getElementById('captureOverlay');
  ov.classList.add('show');

  // 3. Fit graph to canvas with no animation
  network.fit({{animation:false}});

  // 4. Wait one frame then capture
  setTimeout(function(){{
    var netDiv = document.getElementById('network');
    var canvas = netDiv.querySelector('canvas');
    if(!canvas){{
      ov.classList.remove('show');
      panel.style.display = prevPanel;
      alert('Canvas not found — try again after graph settles.');
      return;
    }}

    // 5. Compose: white background + vis canvas
    var w = canvas.width, h = canvas.height;
    var composed = document.createElement('canvas');
    composed.width  = w;
    composed.height = h;
    var ctx = composed.getContext('2d');

    // white bg
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0,0,w,h);

    // draw vis canvas
    ctx.drawImage(canvas,0,0);

    // optional: add subtle border + timestamp watermark
    ctx.strokeStyle = '#e2e8f0';
    ctx.lineWidth   = 2;
    ctx.strokeRect(1,1,w-2,h-2);

    var ts = new Date().toLocaleString('bn-BD',{{
      year:'numeric',month:'short',day:'numeric',
      hour:'2-digit',minute:'2-digit'}});
    ctx.font = 'bold 13px Segoe UI, Arial, sans-serif';
    ctx.fillStyle = 'rgba(30,58,138,0.55)';
    ctx.textAlign = 'right';
    ctx.fillText('CDR Network Graph  |  ' + ts, w-12, h-10);

    _exportDataURL = composed.toDataURL('image/png',1.0);

    ov.classList.remove('show');
    panel.style.display = prevPanel;
    callback(_exportDataURL);
  }}, 350);
}}

// ── Export PNG → opens preview modal ──
function exportGraphPNG(){{
  _captureGraph(function(dataURL){{
    document.getElementById('exportPreview').src = dataURL;
    document.getElementById('exportStatus').textContent = '';
    document.getElementById('exportModal').classList.add('show');
  }});
}}

// ── Quick copy (no modal) ──
function copyGraphToClipboard(){{
  _captureGraph(function(dataURL){{
    document.getElementById('exportPreview').src = dataURL;
    _doCopy(dataURL, true);
  }});
}}

// ── Download from modal ──
function downloadExportedPNG(){{
  if(!_exportDataURL)return;
  var a = document.createElement('a');
  a.href = _exportDataURL;
  a.download = 'CDR_Network_Graph_' + Date.now() + '.png';
  a.click();
  document.getElementById('exportStatus').textContent = '✅ Downloaded!';
  setTimeout(function(){{document.getElementById('exportStatus').textContent='';}},2500);
}}

// ── Copy from modal ──
function copyExportedToClipboard(){{
  if(!_exportDataURL)return;
  _doCopy(_exportDataURL, false);
}}

function _doCopy(dataURL, quick){{
  // Convert dataURL → Blob → ClipboardItem
  var b64 = dataURL.split(',')[1];
  var byteChars = atob(b64);
  var byteArr = new Uint8Array(byteChars.length);
  for(var i=0;i<byteChars.length;i++) byteArr[i]=byteChars.charCodeAt(i);
  var blob = new Blob([byteArr],{{type:'image/png'}});

  if(navigator.clipboard && window.ClipboardItem){{
    navigator.clipboard.write([new ClipboardItem({{'image/png':blob}})])
      .then(function(){{
        var msg = '✅ Clipboard-এ copy হয়েছে! Ctrl+V দিয়ে Word/PowerPoint-এ paste করুন।';
        if(quick){{ alert(msg); }}
        else{{ document.getElementById('exportStatus').textContent='✅ Copied!';
               setTimeout(function(){{document.getElementById('exportStatus').textContent='';}},2500); }}
      }})
      .catch(function(){{
        // Fallback: open in new tab
        var w=window.open();
        w.document.write('<img src="'+dataURL+'" style="max-width:100%"><br>'
          +'<p style="font-family:sans-serif;color:#1e3a8a">Right-click → Copy Image অথবা Save Image As করুন।</p>');
      }});
  }} else {{
    // Old browser fallback
    var w=window.open();
    w.document.write('<img src="'+dataURL+'" style="max-width:100%"><br>'
      +'<p style="font-family:sans-serif;color:#1e3a8a">Right-click → Copy Image অথবা Save Image As করুন।</p>');
  }}
}}

function closeExportModal(){{
  document.getElementById('exportModal').classList.remove('show');
}}

function togglePhysics(){{
  physicsOn=!physicsOn;
  if(physicsOn){{
    // Unpin all nodes so physics can move them again
    var unpinned = allNodes.get().map(function(n){{
      return {{id:n.id, fixed:{{x:false,y:false}}}};
    }});
    allNodes.update(unpinned);
  }}
  network.setOptions({{physics:{{enabled:physicsOn}}}});
  document.getElementById('physBtn').textContent=physicsOn?'\u23F8 Freeze':'\u25B6 Unfreeze';
}}

// ── Layout switcher (i2-style) ──
function applyLayout(mode){{
  if(mode==='physics'){{
    network.setOptions({{
      layout:{{improvedLayout:false,hierarchical:{{enabled:false}}}},
      physics:{{
        enabled:true,solver:'repulsion',
        stabilization:{{iterations:500,updateInterval:20}},
        repulsion:{{centralGravity:.1,springLength:220,springConstant:.04,
                    nodeDistance:200,damping:.10}}
      }}
    }});
    physicsOn=true;
    document.getElementById('physBtn').textContent='\u23F8 Freeze';
    network.once('stabilizationIterationsDone',function(){{
      network.fit({{animation:{{duration:500,easingFunction:'easeInOutQuad'}}}});
    }});
    return;
  }}
  if(mode==='hierarchyLR'||mode==='hierarchyUD'){{
    var dir=mode==='hierarchyLR'?'LR':'UD';
    network.setOptions({{
      layout:{{
        improvedLayout:true,
        hierarchical:{{
          enabled:true,direction:dir,
          sortMethod:'hubsize',
          nodeSpacing:200,
          levelSeparation:250,
          treeSpacing:250,
          blockShifting:true,
          edgeMinimization:true,
          parentCentralization:true
        }}
      }},
      physics:{{enabled:false}}
    }});
    physicsOn=false;
    document.getElementById('physBtn').textContent='\u25B6 Unfreeze';
    setTimeout(function(){{network.fit({{animation:{{duration:500}}}});}},400);
    return;
  }}
  if(mode==='circle'){{
    network.setOptions({{
      layout:{{improvedLayout:false,hierarchical:{{enabled:false}}}},
      physics:{{enabled:false}}
    }});
    physicsOn=false;
    document.getElementById('physBtn').textContent='\u25B6 Unfreeze';
    var visibleNodes=allNodes.get().filter(function(n){{return !n.hidden;}});
    var n=visibleNodes.length;
    var cx=0,cy=0,r=Math.max(220,n*55);
    var posUpdates=[];
    visibleNodes.forEach(function(nd,i){{
      var angle=(2*Math.PI*i/n)-Math.PI/2;
      posUpdates.push({{id:nd.id,
        x:Math.round(cx+r*Math.cos(angle)),
        y:Math.round(cy+r*Math.sin(angle)),
        fixed:false}});
    }});
    allNodes.update(posUpdates);
    setTimeout(function(){{network.fit({{animation:{{duration:500}}}});}},150);
    return;
  }}
  if(mode==='grid'){{
    network.setOptions({{
      layout:{{improvedLayout:false,hierarchical:{{enabled:false}}}},
      physics:{{enabled:false}}
    }});
    physicsOn=false;
    document.getElementById('physBtn').textContent='\u25B6 Unfreeze';
    var visibleNodes=allNodes.get().filter(function(n){{return !n.hidden;}});
    var cols=Math.ceil(Math.sqrt(visibleNodes.length));
    var spacing=220;
    var posUpdates=[];
    visibleNodes.forEach(function(nd,i){{
      posUpdates.push({{id:nd.id,
        x:(i%cols)*spacing - (cols/2*spacing),
        y:Math.floor(i/cols)*spacing - (Math.ceil(visibleNodes.length/cols)/2*spacing),
        fixed:false}});
    }});
    allNodes.update(posUpdates);
    setTimeout(function(){{network.fit({{animation:{{duration:500}}}});}},150);
    return;
  }}
  if(mode==='bipartite'){{
    // Subjects split left / right, contacts in the middle
    network.setOptions({{
      layout:{{improvedLayout:false,hierarchical:{{enabled:false}}}},
      physics:{{enabled:false}}
    }});
    physicsOn=false;
    document.getElementById('physBtn').textContent='\u25B6 Unfreeze';
    var visibleNodes=allNodes.get().filter(function(n){{return !n.hidden;}});
    var subjNodes=visibleNodes.filter(function(n){{
      return n.group==='subject'||n.group==='isolated_subject';
    }});
    var contactNodes=visibleNodes.filter(function(n){{
      return n.group!=='subject'&&n.group!=='isolated_subject';
    }});
    // Left half subjects on X=-700, right half on X=+700
    var leftSubj=subjNodes.slice(0,Math.ceil(subjNodes.length/2));
    var rightSubj=subjNodes.slice(Math.ceil(subjNodes.length/2));
    var yStep=180;
    var posUpdates2=[];
    leftSubj.forEach(function(n,i){{
      posUpdates2.push({{id:n.id,
        x:-700,
        y:(i-(leftSubj.length-1)/2)*yStep,
        fixed:{{x:true,y:false}}}});
    }});
    rightSubj.forEach(function(n,i){{
      posUpdates2.push({{id:n.id,
        x:700,
        y:(i-(rightSubj.length-1)/2)*yStep,
        fixed:{{x:true,y:false}}}});
    }});
    // Contacts in the middle in a grid
    var cCols=Math.max(1,Math.ceil(Math.sqrt(contactNodes.length*0.6)));
    var cSpacingX=200, cSpacingY=160;
    var totalRows2=Math.ceil(contactNodes.length/cCols);
    contactNodes.forEach(function(n,i){{
      var col=i%cCols, row=Math.floor(i/cCols);
      posUpdates2.push({{id:n.id,
        x:(col-(cCols-1)/2)*cSpacingX,
        y:(row-(totalRows2-1)/2)*cSpacingY,
        fixed:false}});
    }});
    allNodes.update(posUpdates2);
    setTimeout(function(){{network.fit({{animation:{{duration:600}}}});}},200);
    return;
  }}
}}

// ── Importance Ring toggle ──
var _impRingOn = true;
function toggleImportanceRing(){{
  _impRingOn = !_impRingOn;
  var btn = document.getElementById('impRingBtn');
  btn.textContent = _impRingOn ? '\\u2B50 Ring: ON' : '\\u2B50 Ring: OFF';
  btn.style.background = _impRingOn ? '#1e3a8a' : '#64748b';
  // Update borderWidth & border color for important nodes
  var updates = [];
  allNodes.get().forEach(function(n){{
    if(!n._important) return;
    updates.push({{
      id: n.id,
      borderWidth: _impRingOn ? 5 : 1,
      color: Object.assign({{}}, n.color, {{
        border: _impRingOn ? '#f59e0b' : '#94a3b8'
      }})
    }});
  }});
  allNodes.update(updates);
}}

// ── Filter by min connection count ──
function filterByConnCount(val){{
  val=parseInt(val);
  document.getElementById('minConnVal').textContent=val;
  var updates=[];
  nodesData.forEach(function(n){{
    if(n.group==='subject'||n.group==='isolated_subject'){{
      updates.push({{id:n.id,hidden:false}});return;
    }}
    var tot=n._total||0;
    updates.push({{id:n.id,hidden:(tot<val)}});
  }});
  allNodes.update(updates);
  // Re-apply edge filter respecting both hidden nodes + active etype
  var hiddenNodes=new Set();
  allNodes.get().forEach(function(n){{if(n.hidden)hiddenNodes.add(n.id);}});
  var edgeUpdates=[];
  allEdges.get().forEach(function(e){{
    var nodeHidden=hiddenNodes.has(e.to)||hiddenNodes.has(e.from);
    var etypeHidden=false;
    if(_activeEtype!=='all'){{
      if(_activeEtype==='call') etypeHidden=e._etype!=='call';
      else if(_activeEtype==='sms') etypeHidden=e._etype!=='sms';
      else etypeHidden=e._etype!==_activeEtype;
    }}
    edgeUpdates.push({{id:e.id,hidden:nodeHidden||etypeHidden}});
  }});
  allEdges.update(edgeUpdates);
}}

// ── Delete selected node/edge ──
function deleteSelected(){{
  if(selectedNodeId!==null){{
    var node=allNodes.get(selectedNodeId);
    if(!node)return;
    var connEdges=network.getConnectedEdges(selectedNodeId);
    var removedEdges=[];
    connEdges.forEach(function(eid){{
      var e=allEdges.get(eid);
      if(e){{removedEdges.push(e);allEdges.remove(eid);}};
    }});
    deletedNodes.push({{node:node,edges:removedEdges}});
    allNodes.remove(selectedNodeId);
    selectedNodeId=null;
    closePanel();
  }} else if(selectedEdgeId!==null){{
    var edge=allEdges.get(selectedEdgeId);
    if(edge){{
      deletedEdges.push(edge);
      allEdges.remove(selectedEdgeId);
      selectedEdgeId=null;
      closePanel();
    }}
  }}
}}

// ── Undo last delete ──
function undoDelete(){{
  if(deletedNodes.length>0){{
    var last=deletedNodes.pop();
    allNodes.add(last.node);
    last.edges.forEach(function(e){{allEdges.add(e);}});
  }} else if(deletedEdges.length>0){{
    allEdges.add(deletedEdges.pop());
  }}
}}

// ── Keyboard delete ──
document.addEventListener('keydown',function(e){{
  if(e.key==='Delete'||e.key==='Backspace'){{
    if(document.activeElement===document.body||
       document.activeElement===document.getElementById('network')){{
      deleteSelected();
    }}
  }}
}});

function showOnlyCommon(){{
  var keep=nodesData.filter(n=>n.group==='subject'||n.group==='isolated_subject'||n.group==='common').map(n=>n.id);
  allNodes.update(nodesData.map(n=>({{id:n.id,hidden:!keep.includes(n.id)}})));
  allEdges.update(edgesData.map(e=>({{id:e.id,hidden:!keep.includes(e.to)}})));
  network.fit();
}}
function showAll(){{
  allNodes.update(nodesData.map(n=>({{id:n.id,hidden:false}})));
  allEdges.update(edgesData.map(e=>({{id:e.id,hidden:false}})));
  network.fit();
}}

// ── Label / node size sliders ──
function changeFontSize(val){{
  val=parseInt(val);
  document.getElementById('fontVal').textContent=val===0?'off':val;
  allNodes.update(allNodes.get().map(n=>({{id:n.id,
    font:Object.assign({{}},n.font,{{size:val}})}})));
}}
function changeEdgeFontSize(val){{
  val=parseInt(val);
  document.getElementById('edgeFontVal').textContent=val===0?'off':val;
  allEdges.update(allEdges.get().map(e=>({{id:e.id,
    font:Object.assign({{}},e.font,{{size:val}})}})));
}}
function changeNodeSize(val){{
  val=parseInt(val);
  document.getElementById('nodeVal').textContent=val;
  allNodes.update(allNodes.get().map(n=>{{
    if(n.group==='subject'||n.group==='isolated_subject')return{{id:n.id}};
    return{{id:n.id,size:val}};
  }}));
}}

// ── Search nodes by number or name ──
// ── Search state ──────────────────────────────────────────────────────────
var _searchMatches = [];   // ordered list of matched node IDs
var _searchIdx     = -1;   // current focus index (for Enter cycling)

function searchNodes(q){{
  q = q.trim().toLowerCase();
  var countEl = document.getElementById('searchCount');
  if(!q){{
    _searchMatches = []; _searchIdx = -1;
    // Restore — only if not in hover-dim mode
    if(!_hoverActive){{
      allNodes.update(allNodes.get().map(n=>({{id:n.id,opacity:1.0,
        borderWidth:n._origBW||2,color:n._origColor||undefined}})));
      allEdges.update(allEdges.get().map(e=>({{id:e.id,hidden:false,opacity:1.0}})));
    }}
    if(countEl) countEl.textContent='';
    return;
  }}

  var matched = new Set();
  allNodes.get().forEach(function(n){{
    var lbl   = (n.label  ||'').toLowerCase();
    var cname = (n._cname ||'').toLowerCase();
    var id    = (String(n.id)||'').toLowerCase();
    if(lbl.includes(q)||cname.includes(q)||id.includes(q)) matched.add(n.id);
  }});

  _searchMatches = [...matched];
  _searchIdx = _searchMatches.length > 0 ? 0 : -1;

  // Highlight matched (gold ring + full opacity), dim others
  allNodes.update(allNodes.get().map(function(n){{
    if(matched.has(n.id)){{
      return {{id:n.id, opacity:1.0,
               borderWidth:4,
               color:{{border:'#f59e0b',background:n._origBg||n.color&&n.color.background||'#fff'}}}};
    }} else {{
      return {{id:n.id, opacity:0.08,
               borderWidth:n._origBW||2,
               color:n._origColor||undefined}};
    }}
  }}));

  // Show only edges between matched nodes
  allEdges.update(allEdges.get().map(e=>
    ({{id:e.id, hidden:!(matched.has(e.from)&&matched.has(e.to)),
       opacity: matched.has(e.from)&&matched.has(e.to)?1.0:0.0}})));

  // Count badge
  if(countEl) countEl.textContent = matched.size > 0
    ? matched.size+' found'
    : 'No match';

  // Zoom to first match
  if(_searchMatches.length > 0){{
    network.focus(_searchMatches[0],{{scale:1.6,animation:{{duration:400,easingFunction:'easeOutQuad'}}}});
    network.selectNodes([_searchMatches[0]]);
  }}
}}

// Enter key cycles through matches
function handleSearchKey(e){{
  if(e.key !== 'Enter' || _searchMatches.length === 0) return;
  _searchIdx = (_searchIdx + 1) % _searchMatches.length;
  var nodeId = _searchMatches[_searchIdx];
  network.focus(nodeId,{{scale:1.6,animation:{{duration:300,easingFunction:'easeOutQuad'}}}});
  network.selectNodes([nodeId]);
  // Show NODE INFO panel for focused node
  var nodeObj = allNodes.get(nodeId);
  if(nodeObj&&nodeObj.title) showPanel('NODE INFO',nodeObj.title);
}}

// ── Hover dim ─────────────────────────────────────────────────────────────
// State: are we in hover-dim mode?
var _hoverActive   = false;
var _hoveredNodeId = null;

// Store original colors on first hover (once)
var _origColorsStored = false;
function _storeOrigColors(){{
  if(_origColorsStored) return;
  allNodes.update(allNodes.get().map(function(n){{
    return {{id:n.id,
      _origColor: n.color   || null,
      _origBg:    n.color && n.color.background ? n.color.background : null,
      _origBorder:n.color && n.color.border     ? n.color.border     : null,
      _origBW:    n.borderWidth || 2}};
  }}));
  _origColorsStored = true;
}}

network.on('hoverNode',function(params){{
  // Skip if search is active or node is selected
  var q = document.getElementById('searchBox').value.trim();
  if(q) return;
  _storeOrigColors();
  _hoverActive   = true;
  _hoveredNodeId = params.node;

  var hovered   = new Set([params.node]);
  var connected = new Set(network.getConnectedNodes(params.node));
  connected.add(params.node);

  // Hovered node: full + gold ring
  // Connected nodes: full opacity, normal border
  // Others: heavily dimmed
  allNodes.update(allNodes.get().map(function(n){{
    if(n.id === params.node){{
      return {{id:n.id, opacity:1.0, borderWidth:4,
               color:{{border:'#f59e0b',
                       background:n._origBg||undefined}}}};
    }} else if(connected.has(n.id)){{
      return {{id:n.id, opacity:1.0, borderWidth:n._origBW||2,
               color:n._origColor||undefined}};
    }} else {{
      return {{id:n.id, opacity:0.07, borderWidth:1,
               color:n._origColor||undefined}};
    }}
  }}));

  // Edges: bright if connected to hovered, very dim otherwise
  allEdges.update(allEdges.get().map(function(e){{
    var isConn = e.from===params.node||e.to===params.node;
    return {{id:e.id, opacity:isConn?1.0:0.05,
             width: isConn ? Math.max(e._origWidth||e.width||1, 2.5) : (e._origWidth||e.width||1)}};
  }}));
}}

network.on('blurNode',function(params){{
  var q = document.getElementById('searchBox').value.trim();
  if(q) return;  // search is driving — don't reset
  _hoverActive   = false;
  _hoveredNodeId = null;

  // Restore everything
  allNodes.update(allNodes.get().map(function(n){{
    return {{id:n.id, opacity:1.0,
             borderWidth:n._origBW||2,
             color:n._origColor||undefined}};
  }}));
  allEdges.update(allEdges.get().map(function(e){{
    return {{id:e.id, opacity:1.0,
             width:e._origWidth||e.width||1}};
  }}));
}}

// ── Edge type filter ──
var _activeEtype = 'all';
function filterEdgeType(etype){{
  _activeEtype = etype;
  // button highlight
  ['fAll','fCall','fSms'].forEach(function(id){{
    document.getElementById(id).style.opacity='0.5';
  }});
  var activeId = etype==='all'?'fAll':etype==='call'?'fCall':'fSms';
  document.getElementById(activeId).style.opacity='1.0';

  var hiddenNodes=new Set();
  allNodes.get().forEach(function(n){{if(n.hidden)hiddenNodes.add(n.id);}});

  allEdges.update(allEdges.get().map(function(e){{
    if(hiddenNodes.has(e.from)||hiddenNodes.has(e.to)) return{{id:e.id,hidden:true}};
    if(etype==='all') return{{id:e.id,hidden:false}};
    return{{id:e.id,hidden:e._etype!==etype}};
  }}));
}}

// ── Info panel ──
function closePanel(){{document.getElementById('panel').style.display='none';}}
function showPanel(tag,html){{
  document.getElementById('panelTag').textContent=tag;
  document.getElementById('panelBody').innerHTML=html;
  document.getElementById('panel').style.display='block';
}}

// ── Click handler ──
network.on('click',function(params){{
  if(params.nodes.length>0){{
    selectedNodeId=params.nodes[0];
    selectedEdgeId=null;
    var nodeObj=allNodes.get(selectedNodeId);
    var connected=network.getConnectedNodes(selectedNodeId);
    connected.push(selectedNodeId);
    allNodes.update(allNodes.get().map(n=>({{id:n.id,
      opacity:connected.includes(n.id)?1.0:0.12}})));
    if(nodeObj&&nodeObj.title)showPanel('NODE INFO',nodeObj.title);
  }} else if(params.edges.length>0){{
    selectedEdgeId=params.edges[0];
    selectedNodeId=null;
    var edgeObj=allEdges.get(selectedEdgeId);
    if(edgeObj&&edgeObj.title)showPanel('LINK INFO',edgeObj.title);
  }} else{{
    selectedNodeId=null; selectedEdgeId=null;
    _hoverActive=false; _hoveredNodeId=null;
    var q=document.getElementById('searchBox').value.trim();
    if(!q){{
      allNodes.update(allNodes.get().map(n=>({{id:n.id,opacity:1.0,
        borderWidth:n._origBW||2,color:n._origColor||undefined}})));
      allEdges.update(allEdges.get().map(e=>({{id:e.id,opacity:1.0,
        width:e._origWidth||e.width||1}})));
    }}
    closePanel();
  }}
}});

// ── Context menu HTML ─────────────────────────────────────────────────────
var _ctxMenu = (function(){{
  var el = document.createElement('div');
  el.id = 'ctxMenu';
  el.style.cssText = [
    'position:fixed;z-index:9999;background:#fff',
    'border:1px solid #e2e8f0;border-radius:8px',
    'box-shadow:0 8px 24px rgba(0,0,0,.18)',
    'padding:4px 0;min-width:190px',
    'font-family:Segoe UI,Arial,sans-serif;font-size:13px',
    'display:none;user-select:none'
  ].join(';');
  document.body.appendChild(el);

  function item(icon, label, action, danger){{
    var d = document.createElement('div');
    d.style.cssText = [
      'padding:7px 14px;cursor:pointer;display:flex;gap:8px;align-items:center',
      danger ? 'color:#dc2626' : 'color:#0f172a'
    ].join(';');
    d.innerHTML = '<span style="font-size:15px">'+icon+'</span><span>'+label+'</span>';
    d.onmouseenter = function(){{ d.style.background='#f1f5f9'; }};
    d.onmouseleave = function(){{ d.style.background=''; }};
    d.onclick = function(){{ hide(); action(); }};
    return d;
  }}

  function sep(){{
    var hr = document.createElement('hr');
    hr.style.cssText = 'margin:3px 0;border:none;border-top:1px solid #f1f5f9';
    return hr;
  }}

  function show(x, y, items){{
    el.innerHTML = '';
    items.forEach(function(it){{
      if(it === 'sep') el.appendChild(sep());
      else el.appendChild(it);
    }});
    el.style.display = 'block';
    // Prevent overflow off screen
    var rect = el.getBoundingClientRect();
    var vw = window.innerWidth; var vh = window.innerHeight;
    el.style.left = (x + rect.width > vw ? vw - rect.width - 8 : x) + 'px';
    el.style.top  = (y + rect.height > vh ? vh - rect.height - 8 : y) + 'px';
  }}

  function hide(){{ el.style.display='none'; }}
  document.addEventListener('click', hide);
  document.addEventListener('keydown', function(e){{ if(e.key==='Escape') hide(); }});

  return {{ show:show, hide:hide, item:item }};
}})();

// ── Right-click handler ────────────────────────────────────────────────────
network.on('oncontext', function(params){{
  params.event.preventDefault();
  var x = params.event.clientX;
  var y = params.event.clientY;

  if(params.nodes.length > 0){{
    var nid = params.nodes[0];
    selectedNodeId = nid;
    var nodeObj = allNodes.get(nid);
    var label   = nodeObj ? (nodeObj.label||nid) : nid;
    var numOnly = String(nid).replace(/[^0-9+]/g,'');

    _ctxMenu.show(x, y, [
      _ctxMenu.item('📋', 'Copy Number',       function(){{ _copyText(numOnly||String(nid)); }}),
      _ctxMenu.item('📝', 'Copy Full Label',   function(){{ _copyText(label); }}),
      'sep',
      _ctxMenu.item('🔦', 'Highlight Network', function(){{
        var conn = new Set(network.getConnectedNodes(nid)); conn.add(nid);
        allNodes.update(allNodes.get().map(n=>({{id:n.id,opacity:conn.has(n.id)?1.0:0.08}})));
      }}),
      _ctxMenu.item('👁',  'Show Node Info',   function(){{
        if(nodeObj&&nodeObj.title) showPanel('NODE INFO', nodeObj.title);
      }}),
      'sep',
      _ctxMenu.item('🗑', 'Remove Node',       function(){{ deleteSelected(); }}, true),
    ]);

  }} else if(params.edges.length > 0){{
    var eid2 = params.edges[0];
    selectedEdgeId = eid2;
    var edgeObj = allEdges.get(eid2);
    var fromN   = edgeObj ? String(edgeObj.from) : '';
    var toN     = edgeObj ? String(edgeObj.to)   : '';
    var durVal  = edgeObj ? (edgeObj._dur || 0)  : 0;

    _ctxMenu.show(x, y, [
      _ctxMenu.item('📋', 'Copy: '+fromN.slice(-8),  function(){{ _copyText(fromN); }}),
      _ctxMenu.item('📋', 'Copy: '+toN.slice(-8),    function(){{ _copyText(toN); }}),
      'sep',
      _ctxMenu.item('ℹ️',  'Show Edge Info',          function(){{
        if(edgeObj&&edgeObj.title) showPanel('LINK INFO', edgeObj.title);
      }}),
      'sep',
      _ctxMenu.item('🗑', 'Remove Edge',             function(){{ deleteSelected(); }}, true),
    ]);

  }} else {{
    // Background right-click
    _ctxMenu.show(x, y, [
      _ctxMenu.item('🔲', 'Fit All Nodes',    function(){{ network.fit(); }}),
      _ctxMenu.item('▶', 'Unfreeze Physics', function(){{
        physicsOn = true;
        network.setOptions({{physics:{{enabled:true}}}});
        document.getElementById('physBtn').textContent='\u23F8 Freeze';
      }}),
      _ctxMenu.item('⏸', 'Freeze Layout',    function(){{
        physicsOn = false;
        network.setOptions({{physics:{{enabled:false}}}});
        document.getElementById('physBtn').textContent='\u25B6 Unfreeze';
      }}),
      'sep',
      _ctxMenu.item('👁', 'Show All',         function(){{ showAll(); }}),
      _ctxMenu.item('🔴', 'Common Only',       function(){{ showOnlyCommon(); }}),
    ]);
  }}
}});

// ── Copy helper ────────────────────────────────────────────────────────────
function _copyText(text){{
  if(navigator.clipboard && navigator.clipboard.writeText){{
    navigator.clipboard.writeText(text).then(function(){{
      _showCopyToast(text);
    }}).catch(function(){{ _fallbackCopy(text); }});
  }} else {{
    _fallbackCopy(text);
  }}
}}

function _fallbackCopy(text){{
  var ta = document.createElement('textarea');
  ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
  document.body.appendChild(ta); ta.select();
  try {{ document.execCommand('copy'); _showCopyToast(text); }} catch(e){{}}
  document.body.removeChild(ta);
}}

function _showCopyToast(text){{
  var toast = document.createElement('div');
  toast.textContent = '✅ Copied: ' + text;
  toast.style.cssText = [
    'position:fixed;bottom:28px;left:50%;transform:translateX(-50%)',
    'background:#0f172a;color:#fff;padding:8px 18px',
    'border-radius:20px;font-size:13px;font-family:Segoe UI,Arial,sans-serif',
    'z-index:99999;pointer-events:none',
    'box-shadow:0 4px 16px rgba(0,0,0,.35)',
    'opacity:0;transition:opacity .2s'
  ].join(';');
  document.body.appendChild(toast);
  requestAnimationFrame(function(){{ toast.style.opacity='1'; }});
  setTimeout(function(){{
    toast.style.opacity='0';
    setTimeout(function(){{ document.body.removeChild(toast); }}, 300);
  }}, 2000);
}}
</script>
</body>
</html>"""
    return html


def link_analysis_page():
    """Main Link Analysis Page."""
    st.markdown("""
    <div style="background:white;border-radius:12px;padding:1.2rem 1.5rem;margin-bottom:1rem;
                border-left:4px solid #2563eb;box-shadow:0 1px 4px rgba(0,0,0,.06)">
        <div style="font-size:1.2rem;font-weight:800;color:#1e3a8a">🔗 CDR Link Analysis</div>
        <div style="font-size:.85rem;color:#64748b">Upload up to 5 CDR files to analyze connections, common contacts, and co-location events</div>
    </div>
    """, unsafe_allow_html=True)

    # ── Upload Section ──
    st.markdown("### 📂 Upload CDR Files (max 5)")
    cols = st.columns(5)
    uploaded_files = []
    labels = ['Subject A', 'Subject B', 'Subject C', 'Subject D', 'Subject E']

    for i, col in enumerate(cols):
        with col:
            f = st.file_uploader(
                labels[i], type=['xlsx', 'xls', 'csv'],
                key=f'link_cdr_{i}',
                label_visibility='visible'
            )
            if f is not None and not validate_upload(f, kind="excel"):
                f = None
            uploaded_files.append(f)

    active_files = [(f, labels[i]) for i, f in enumerate(uploaded_files) if f is not None]

    if len(active_files) < 2:
        st.info("📌 Upload at least 2 CDR files to start link analysis")
        return

    # Settings
    with st.expander("⚙️ Settings", expanded=False):
        c1, c2, c3 = st.columns(3)
        with c1:
            top_n = st.slider("Top contacts per subject (excl. common contacts)", 5, 20, 5, help="Each subject shows up to this many unique contacts. Common contacts are always shown separately.")
        with c2:
            coloc_window = st.slider("Co-location time window (minutes)", 5, 120, 30)
        with c3:
            radius_km = st.slider("Co-location radius (km)", 1, 20, 5)
        exclude_noise = st.checkbox(
            "🧹 Filter carrier/service numbers (IVR, promo, shortcodes)",
            value=True,
            help="When ON, operator IVR, promotional and shortcode numbers are removed from "
                 "connections, common contacts, and the network graph. "
                 "Turn OFF only if you need to investigate a specific service number."
        )

    # ── Subject Name & Photo ──
    with st.expander("👤 Subject Names & Photos (optional)", expanded=False):
        st.caption("Enter name and upload photo. Photo will replace the star icon in the network graph.")
        _subj_meta = {}
        _meta_cols = st.columns(5)
        for _si, (_mc, _lbl) in enumerate(zip(_meta_cols, labels)):
            with _mc:
                _sname = st.text_input(f"Name", key=f"subj_name_{_si}",
                                       placeholder=f"e.g. John ({_lbl})")
                _sphoto = st.file_uploader(f"Photo", type=["jpg","jpeg","png"],
                                           key=f"subj_photo_{_si}")
                if _sphoto is not None and not validate_upload(_sphoto, kind="image"):
                    _sphoto = None
                _photo_b64 = None
                if _sphoto:
                    import base64 as _b64
                    _photo_b64 = "data:" + _sphoto.type + ";base64," + _b64.b64encode(_sphoto.read()).decode()
                _subj_meta[_lbl] = {"name": _sname.strip(), "photo": _photo_b64}

    # ── Contact Names (optional) ──
    with st.expander("📇 Contact Names (optional)", expanded=False):
        st.caption(
            "পরিচিত নম্বরের নাম দিন। Graph-এ নম্বরের পাশে নাম দেখাবে। "
            "Format: একটি করে লাইনে `880XXXXXXXXXX = নাম`"
        )
        _contact_names_raw = st.text_area(
            "Number = Name (একটি লাইনে একটি)",
            placeholder="8801XXXXXXXXX = Rahim Uddin\n8801YYYYYYYYY = Karim Vai",
            height=140,
            key="contact_names_input"
        )
        # Parse contact names
        _contact_names_dict = {}
        for _line in _contact_names_raw.splitlines():
            _line = _line.strip()
            if '=' in _line:
                _parts = _line.split('=', 1)
                _num = _parts[0].strip()
                _nm  = _parts[1].strip()
                if _num and _nm:
                    _contact_names_dict[_num] = _nm
        if _contact_names_dict:
            st.success(f"✅ {len(_contact_names_dict)} contact name(s) loaded.")

    if st.button("🔗 Run Link Analysis", type="primary", use_container_width=False,
                 key="run_link_analysis"):

        # ── Load CDRs ──
        dfs = []
        subjects = []
        anomaly_counts = []

        with st.spinner("Loading CDR files..."):
            for f, label in active_files:
                df, subject, anomalies = _load_cdr(f, label)
                if df is not None and len(df) > 0:
                    dfs.append(df)
                    subjects.append(subject)
                    anomaly_counts.append(anomalies)

        if len(dfs) < 2:
            st.error("Could not load at least 2 valid CDR files")
            return

        # ── Summary cards ──
        st.markdown("### 📊 Summary")
        summary_cols = st.columns(len(dfs))
        for i, (df, sub, anoms) in enumerate(zip(dfs, subjects, anomaly_counts)):
            with summary_cols[i]:
                st.markdown(f"""
                <div style="background:white;border-top:4px solid {'#2563eb #dc2626 #16a34a #7c3aed #d97706'.split()[i % 5]};
                            border-radius:10px;padding:1rem;box-shadow:0 1px 3px rgba(0,0,0,.06);text-align:center">
                    <div style="font-size:.75rem;color:#94a3b8;text-transform:uppercase">{active_files[i][1]}</div>
                    <div style="font-size:.95rem;font-weight:800;color:#0f172a;margin:.25rem 0">{sub}</div>
                    <div style="font-size:.82rem;color:#64748b">{len(df):,} records</div>
                    <div style="font-size:.78rem;color:#f59e0b">{anoms} anomalies removed</div>
                </div>
                """, unsafe_allow_html=True)

        st.markdown("---")

        # ── Build connections ──
        with st.spinner("Analyzing connections..."):
            connections = _build_connections(dfs, exclude_noise=exclude_noise)

        # Filter top N per subject
        # Sort connections by total across subjects
        conn_sorted = sorted(
            connections.items(),
            key=lambda x: (len(x[1]), sum(d['total'] for d in x[1].values())),
            reverse=True
        )

        # ── Common contacts highlight ──
        common_contacts = [(pb, d) for pb, d in conn_sorted if len(d) >= 2]
        st.markdown(f"### 🔴 Common Contacts ({len(common_contacts)} found)")

        if common_contacts:
            rows = []
            for pb, subj_dict in common_contacts[:50]:
                row = {'Contact Number': pb, 'Shared By': len(subj_dict)}
                total_calls = total_sms = total_dur = 0
                for sub in subjects:
                    d = subj_dict.get(sub, {})
                    out_c = d.get('call_out', 0); in_c = d.get('call_in', 0)
                    out_s = d.get('sms_out', 0); in_s = d.get('sms_in', 0)
                    dur = d.get('duration', 0)
                    row[f'{sub[:8]} Calls'] = f"↑{out_c} ↓{in_c}" if (out_c+in_c) > 0 else '—'
                    row[f'{sub[:8]} SMS'] = f"↑{out_s} ↓{in_s}" if (out_s+in_s) > 0 else '—'
                    total_calls += out_c + in_c
                    total_sms += out_s + in_s
                    total_dur += dur
                row['Total Calls'] = total_calls
                row['Total SMS'] = total_sms
                row['Duration (min)'] = round(total_dur, 1)
                rows.append(row)

            common_df = pd.DataFrame(rows)
            st.dataframe(common_df, use_container_width=True, height=300)
        else:
            st.info("No common contacts found between subjects")

        st.markdown("---")

        # ── Full Connection Table ──
        # Common contacts shown in their own section above.
        # Each subject tab shows top_n contacts EXCLUDING common contacts.
        _common_pbs_table = {pb for pb, sd in connections.items() if len(sd) >= 2}
        st.markdown(f"### 📋 All Connections (Top {top_n} per subject, excl. common contacts)")
        tabs = st.tabs([f"📞 {sub}" for sub in subjects])

        for tab, sub in zip(tabs, subjects):
            with tab:
                sub_conns = [(pb, d[sub]) for pb, d in conn_sorted
                             if sub in d and pb not in _common_pbs_table][:top_n]
                if not sub_conns:
                    st.info(f"No connections found for {sub}")
                    continue
                rows = []
                for pb, d in sub_conns:
                    is_common = len(connections[pb]) >= 2
                    rows.append({
                        'Contact': pb,
                        'Common': '🔴 Yes' if is_common else '—',
                        'Shared Subjects': len(connections[pb]),
                        'Call Out (↑)': d.get('call_out', 0),
                        'Call In (↓)': d.get('call_in', 0),
                        'SMS Out (↑)': d.get('sms_out', 0),
                        'SMS In (↓)': d.get('sms_in', 0),
                        'Total': d.get('total', 0),
                        'Duration (min)': round(d.get('duration', 0), 1),
                    })
                sub_df = pd.DataFrame(rows)
                st.dataframe(sub_df, use_container_width=True, height=350)

        st.markdown("---")

        # ── Network Graph ──
        st.markdown("### 🕸️ Network Graph")
        with st.spinner("Building network graph..."):
            # ── Graph connections: Fair per-subject top_n + shared bonus ──
            #
            # Option A: প্রতি subject থেকে exactly top_n contacts নেওয়া হয়।
            #           ফলে ৪টা subject থাকলে প্রত্যেকের top_n সমান।
            # Option B: top_n slider এখন graph-এও apply হয়।
            # Shared bonus: একাধিক subject-এর সাথে common হলে সে সবসময়
            #               graph-এ থাকবে (top_n limit-এর বাইরেও)।

            top_connections = defaultdict(dict)

            # Pre-compute common numbers (appear in 2+ subjects)
            _common_pbs = {pb for pb, sd in connections.items() if len(sd) >= 2}

            # Step 1 — Per subject: top_n contacts EXCLUDING common contacts
            # Common contacts are added separately (Step 2), so they never
            # consume the per-subject quota.
            for df_s in dfs:
                sub = df_s['_subject'].iloc[0]
                sub_contacts = sorted(
                    [(pb, sd[sub]) for pb, sd in connections.items()
                     if sub in sd and pb not in _common_pbs],   # exclude commons
                    key=lambda x: x[1]['total'], reverse=True
                )
                for pb, data in sub_contacts[:top_n]:
                    if pb not in top_connections:
                        top_connections[pb] = {}
                    top_connections[pb][sub] = data

            # Step 2 — Always include ALL common contacts (shared between 2+ subjects)
            for pb, subj_dict in connections.items():
                if len(subj_dict) >= 2:
                    if pb not in top_connections:
                        top_connections[pb] = {}
                    for sub, data in subj_dict.items():
                        if sub not in top_connections[pb]:
                            top_connections[pb][sub] = data

            # Step 3 — edge count per subject (graph stats-এর জন্য)
            subj_edge_count = {
                sub: sum(1 for sd in top_connections.values() if sub in sd)
                for sub in subjects
            }

            # Build subject meta dict by phone
            _subj_meta_by_phone = {}
            for _si, (df_s, _lbl) in enumerate(zip(dfs, [f[1] for f in active_files])):
                sub = df_s["_subject"].iloc[0]
                meta = _subj_meta.get(_lbl, {}) if "_subj_meta" in dir() else {}
                _subj_meta_by_phone[sub] = meta

            graph_html = _build_network_html(dfs, top_connections, subjects, subj_edge_count, _subj_meta_by_phone, _contact_names_dict)

        st.components.v1.html(graph_html, height=780, scrolling=False)

        # Download graph
        st.download_button(
            "⬇️ Download Network Graph",
            data=graph_html.encode('utf-8'),
            file_name="CDR_Link_Analysis_Network.html",
            mime="text/html",
            key="dl_network_graph"
        )

        st.markdown("---")

        # ── Co-location Analysis ──
        st.markdown(f"### 📍 Co-location Events (±{coloc_window} min, adaptive radius)")
        with st.spinner("Analyzing co-location..."):
            coloc_results = _build_colocation(dfs, coloc_window, radius_km)

        if coloc_results:
            coloc_df = pd.DataFrame(coloc_results)

            # ── Fix 1: Correct same_tower_count ─────────────────────────────
            # Actual value is '✅', not '✅ Yes'
            same_tower_count = int((coloc_df['Same Tower'] == '✅').sum())

            st.success(f"✅ {len(coloc_results)} co-location event(s) found "
                       f"({same_tower_count} at exact same tower)")

            # Summary stats
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.metric("Total Events", len(coloc_results))
            with c2:
                st.metric("Same Tower", same_tower_count)
            with c3:
                avg_dist = coloc_df['Distance (km)'].mean()
                st.metric("Avg Distance", f"{avg_dist:.2f} km")
            with c4:
                high_acc = int((coloc_df['Accuracy'] == '✅ High').sum())
                st.metric("High Accuracy", high_acc)

            # ── Fix 3: GPS accuracy mismatch warning ─────────────────────────
            low_acc  = int((coloc_df['Accuracy'] == '⚠️ Low').sum())
            med_acc  = int((coloc_df['Accuracy'] == '⚡ Med').sum())
            if low_acc > 0 or med_acc > 0:
                with st.expander(
                    f"⚠️ GPS Accuracy Warning — {low_acc} low / {med_acc} medium accuracy events",
                    expanded=(low_acc > 0)
                ):
                    st.markdown("""
**GPS accuracy mismatch detected between subjects:**

| Accuracy | Meaning | Effective Radius Used |
|---|---|---|
| ✅ High | Both subjects: Cell Tower CSV (±0.5 km) | User-set radius |
| ⚡ Med | One: Cell Tower, One: Thana-level (±5 km) | max(user, 8 km) |
| ⚠️ Low | Large gap — e.g. Cell Tower vs District (±15 km) | max(user, 20 km) |

**Recommendation:** ⚠️ Low accuracy events-এ co-location নিশ্চিত নয়।
Cell Tower CSV আপলোড করলে accuracy উন্নত হবে।
                    """)
                    if low_acc > 0:
                        low_df = coloc_df[coloc_df['Accuracy'] == '⚠️ Low'][
                            ['Subject A','Subject B','Time A','Time B',
                             'Distance (km)','GPS Source A','GPS Source B',
                             'Location A','Location B']
                        ].head(10)
                        st.dataframe(low_df, use_container_width=True, height=250)
                        st.caption(f"Showing first 10 of {low_acc} ⚠️ Low accuracy events")

            st.dataframe(coloc_df, use_container_width=True, height=400)
            st.caption(
                "🛰️ Cell Tower = exact LAC+CID match (±0.5 km) · "
                "📍 Thana = BTS address parsed (±5 km) · "
                "🗺️ District = approximate (±15 km) · "
                "Same Tower = exact tower match (same operator only)"
            )

            st.download_button(
                "⬇️ Download Co-location Data",
                data=coloc_df.to_csv(index=False).encode('utf-8'),
                file_name="CDR_CoLocation_Events.csv",
                mime="text/csv",
                key="dl_coloc"
            )
        else:
            st.info("No co-location events found — subjects were not at the same location "
                    f"(within {coloc_window} min, {radius_km} km radius)")

        st.markdown("---")

        # ══════════════════════════════════════════════════════════════════
        # ── Section: Noise Filter Report ──
        # Only shown when carrier/service numbers are actually found
        # ══════════════════════════════════════════════════════════════════
        with st.spinner("Running noise analysis..."):
            raw_connections = _build_connections(dfs, exclude_noise=False)
            carrier_list, _ = _build_noise_analysis(dfs, raw_connections)

        if carrier_list:
            st.markdown("### 🧹 Noise Filter — Carrier & Service Numbers")
            noise_df = pd.DataFrame(carrier_list)
            if exclude_noise:
                st.success(
                    f"✅ {len(carrier_list)} carrier/service number(s) automatically filtered "
                    f"from connections, common contacts, and network graph."
                )
            else:
                st.warning(
                    f"⚠️ Filter is OFF — {len(carrier_list)} carrier/service number(s) "
                    f"are still visible. Enable filter in Settings."
                )
            st.dataframe(noise_df, use_container_width=True, hide_index=True)
            st.markdown("---")

        # ══════════════════════════════════════════════════════════════════
        # ── Section: First / Last Contact Date ──
        # ══════════════════════════════════════════════════════════════════
        st.markdown("### 📅 First & Last Contact Date")
        st.caption("First and last contact date between each subject and their common contacts.")

        with st.spinner("Calculating contact dates..."):
            fl_dates = _build_first_last_dates(dfs)

        if fl_dates and common_contacts:
            fl_rows = []
            # শুধু common contacts-এর জন্য দেখাও (most investigative value)
            for pb, subj_dict in common_contacts[:50]:
                if _is_carrier_number(pb) or _is_promotional(pb): continue
                for sub in subjects:
                    if sub not in subj_dict: continue
                    key = (sub, pb)
                    dates = fl_dates.get(key, {})
                    fl_rows.append({
                        'Subject':       sub,
                        'Contact':       pb,
                        'First Contact': dates.get('first', '—'),
                        'Last Contact':  dates.get('last',  '—'),
                        'Total Calls':   subj_dict[sub].get('call_out', 0) + subj_dict[sub].get('call_in', 0),
                        'Duration (min)': round(subj_dict[sub].get('duration', 0), 1),
                    })

            if fl_rows:
                fl_df = pd.DataFrame(fl_rows).sort_values(['Contact', 'Subject'])

                # Highlight: same contact-এর জন্য subjects-এর first contact date কতটা কাছাকাছি
                st.dataframe(fl_df, use_container_width=True, hide_index=True, height=350)

                # ── Date alignment insight ──
                # একই contact-এ দুই subject-এর first contact date gap বের করো
                align_rows = []
                contact_groups = fl_df.groupby('Contact')
                for contact, grp in contact_groups:
                    if len(grp) < 2: continue
                    dates_valid = grp[grp['First Contact'] != '—']['First Contact'].tolist()
                    if len(dates_valid) < 2: continue
                    try:
                        parsed = sorted([pd.to_datetime(d) for d in dates_valid])
                        gap_days = (parsed[-1] - parsed[0]).days
                        align_rows.append({
                            'Contact':          contact,
                            'Subjects':         ' / '.join(grp['Subject'].tolist()),
                            'Earliest Contact': parsed[0].strftime('%Y-%m-%d'),
                            'Latest Contact':   parsed[-1].strftime('%Y-%m-%d'),
                            'Date Gap (days)':  gap_days,
                            'Alignment':        '🔴 Same Week' if gap_days <= 7
                                                else ('🟡 Same Month' if gap_days <= 30
                                                else '🟢 Different Period'),
                        })
                    except Exception:
                        pass

                if align_rows:
                    align_df = pd.DataFrame(align_rows).sort_values('Date Gap (days)')
                    st.markdown("""<div style="font-weight:700;color:#1e3a8a;margin-top:1rem;margin-bottom:0.4rem;">
                        📊 Contact Date Alignment — When did each subject first contact the same number?
                    </div>""", unsafe_allow_html=True)
                    st.dataframe(align_df, use_container_width=True, hide_index=True)
                    st.caption("🔴 Same Week = highly suspicious alignment · 🟡 Same Month = moderate · 🟢 Different Period = likely coincidental")
            else:
                st.info("Date information unavailable for common contacts.")
        else:
            st.info("Date analysis requires CDR files with a valid timestamp column.")

        st.markdown("---")

        # ══════════════════════════════════════════════════════════════════
        # ── Section: Suspicious Patterns ──
        # ══════════════════════════════════════════════════════════════════
        st.markdown("### 🚨 Suspicious Communication Patterns")

        _sp_window = st.slider(
            "Pattern detection window (minutes)", 5, 120, 30,
            key="sp_window",
            help="Calls to/from the same number within this time window will be flagged as a suspicious pattern."
        )

        with st.spinner("Detecting suspicious patterns..."):
            mirror_rows, relay_rows = _build_suspicious_patterns(dfs, _sp_window)

        sp_tab1, sp_tab2 = st.tabs(["🪞 Mirror Call Pattern", "🔗 Relay Pattern"])

        with sp_tab1:
            st.markdown("""
            <div style="background:#fef2f2;border-left:4px solid #dc2626;border-radius:8px;
                        padding:0.75rem 1rem;font-size:0.83rem;color:#7f1d1d;margin-bottom:0.75rem">
            <b>🪞 What is a Mirror Call Pattern?</b><br>
            Subject A and Subject B both call the same number within <b>±window minutes</b> of each other.
            This indicates they may be <b>coordinated</b> or receiving instructions from the same source.
            </div>
            """, unsafe_allow_html=True)

            if mirror_rows:
                m_df = pd.DataFrame(mirror_rows)
                st.error(f"🚨 {len(mirror_rows)} mirror call instance(s) detected")
                st.dataframe(m_df, use_container_width=True, hide_index=True)

                # Summary: most frequently mirrored numbers
                top_mirror = pd.DataFrame(mirror_rows)['Common Number'].value_counts().head(10)
                if len(top_mirror) > 0:
                    st.markdown("**Top mirrored numbers:**")
                    st.dataframe(
                        top_mirror.reset_index().rename(columns={'Common Number': 'Number', 'count': 'Mirror Instances'}),
                        use_container_width=True, hide_index=True
                    )
            else:
                st.success(f"✅ No mirror call patterns found within ±{_sp_window} min window.")

        with sp_tab2:
            st.markdown("""
            <div style="background:#fff7ed;border-left:4px solid #f59e0b;border-radius:8px;
                        padding:0.75rem 1rem;font-size:0.83rem;color:#78350f;margin-bottom:0.75rem">
            <b>🔗 What is a Relay Pattern?</b><br>
            Subject A calls X, then X calls Subject B within <b>±window minutes</b>.
            X acts as an <b>intermediary</b>, relaying messages or instructions.
            This is an indicator of indirect coordination to avoid direct communication.
            </div>
            """, unsafe_allow_html=True)

            if relay_rows:
                r_df = pd.DataFrame(relay_rows)
                st.warning(f"⚠️ {len(relay_rows)} relay pattern instance(s) detected")
                st.dataframe(r_df, use_container_width=True, hide_index=True)

                # Top relay numbers
                top_relay = pd.DataFrame(relay_rows)['Relay Number (X)'].value_counts().head(10)
                if len(top_relay) > 0:
                    st.markdown("**Top relay numbers (most active intermediaries):**")
                    st.dataframe(
                        top_relay.reset_index().rename(columns={'Relay Number (X)': 'Number', 'count': 'Relay Instances'}),
                        use_container_width=True, hide_index=True
                    )
            else:
                st.success(f"✅ No relay patterns found within {_sp_window} min window.")


def _parse_profile_docs(doc_files):
    """
    NTMC PDF / Image থেকে robust line-based + regex parsing।
    Source tracking: প্রতিটি value কোন document থেকে এসেছে।
    """
    import io as _io, base64 as _b64, re as _re

    FIELDS = ['name','father','mother','spouse','dob','gender',
               'nid','nid_new','passport','tin','tin_new',
               'license_no','vehicle_reg','address_present',
               'address_permanent','profession','mobile',
               'blood_group','nationality','smart_id',
               'license_issue','license_expiry','license_type',
               'vehicle_type','vehicle_color','vehicle_cc',
               'tax_token_expire','fitness_expire',
               'prev_passport','passport_issue','passport_expiry',
               'passport_status','vehicle_reg_date','vehicle_num']
    # store: {field: [(value, doc_label), ...]}
    store      = {k: [] for k in FIELDS}
    photo_b64     = None
    docs_found    = []
    _nid_page_img = None
    _current_doc_label = ['Unknown']  # mutable via list trick

    def _add(field, val, doc_label=None):
        if not val: return
        v = str(val).strip().rstrip('.,;')
        # ── OCR cleanup (image scans থেকে আসা garbage strip করো) ──────────
        # 1) Leading curly/smart quotes এবং junk characters — OCR-এ common
        v = _re.sub(r'^[\'\u2018\u2019\u201C\u201D"`*,;:\-\s]+', '', v).strip()
        v = _re.sub(r'[\u2018\u2019]', "'", v)  # normalize curly to straight
        # 2) Leading label-prefix strip (e.g. "Permanent Address: ASMA")
        #    Common labels that OCR may glue onto next field's value
        _LABEL_PREFIXES = (
            r'Permanent\s*Address', r'Present\s*Address', r'Father\s*(?:\'s)?\s*Name',
            r'Mother\s*(?:\'s)?\s*Name', r'Spouse\s*(?:\'s)?\s*Name', r'Mobile\s*Number',
            r'Date\s*of\s*Birth', r'Passport\s*Number', r'NID\s*Number',
            r'License\s*No', r'License\s*Type', r'Vehicle\s*Classes',
            r'Issuing\s*Authority', r'Issue\s*Date', r'Expiry\s*Date',
            r'Profession', r'Occupation', r'Other\s*Occupation',
            r'Nationality', r'Marital\s*Status', r'Gender',
            r'Emergency\s*Contact[\s\w]*', r'Emergency\s*Mobile\s*Number',
            r'Personal\s*Info', r'License\s*Detail', r'Reference\s*Number',
            r'Applicant\s*Type', r'Apply\s*Date', r'Reference\s*Date',
        )
        for _lp in _LABEL_PREFIXES:
            v = _re.sub(rf'^(?:{_lp})\s*[:\-]?\s*', '', v, flags=_re.I).strip()
        # 3) Strip trailing OCR signature-noise patterns
        v = _re.sub(r'\s+(Seem|See|seer|sees)\s+pete.*$', '', v, flags=_re.I).strip()

        if v.upper() in ('N/A','NA','NONE','880','','N'): return
        if len(v) < 2: return
        # ── Reject obvious label-leak for name fields ──────────────────────
        # PDF column-layout এ ভুলে label ধরা পড়লে এখানে আটকাবে
        if field in ('name','father','mother','spouse'):
            _vu = v.upper().strip()
            # Reject Bengali (NID থেকে বাংলা নাম profile-এ আসবে না)
            if _re.search(r'[\u0980-\u09FF]', v): return
            _NAME_LABEL_BLACKLIST = {
                'PASSPORT STATUS','PASSPORT NUMBER','PASSPORT TYPE',
                'PERMANENT ADDRESS','PRESENT ADDRESS','PROFESSION',
                'SPOUSE NAME',"FATHER'S NAME","MOTHER'S NAME",
                'FATHERS NAME','MOTHERS NAME','FATHER NAME','MOTHER NAME',
                'GOVERNMENT SERVICE','PVT SERVICE','PRIVATE SERVICE',
                'DOCUMENT REVOKED','DOCUMENT REVOKED DUE TO REISSUE',
                'REVOKED DUE TO REISSUE','ACTIVE','REVOKED','OFFICIAL',
                'BIRTH ID','DATE OF BIRTH','DATE OF ISSUE','DATE OF EXPIRY',
                'PREVIOUS PASSPORT','PREVIOUS PASSPORT NO',
                'FIRST NAME','LAST NAME','GENDER','NATIONALITY',
                'BANGLADESH','BANGLADESHI','NID',
            }
            if _vu in _NAME_LABEL_BLACKLIST: return
            # Reject if value contains BOTH a known label keyword AND non-name pattern
            _LABEL_TOKENS = {'PASSPORT','STATUS','PERMANENT','ADDRESS','PRESENT',
                             'PROFESSION','DOCUMENT','REVOKED','REISSUE','SPOUSE',
                             'GOVERNMENT','SERVICE'}
            _toks = set(_vu.split())
            # 2+ label tokens with no other words → reject
            _label_count = len(_toks & _LABEL_TOKENS)
            _other_count = len(_toks - _LABEL_TOKENS - {'OF','THE','TO','DUE','AND','-'})
            if _label_count >= 2 and _other_count == 0: return
            # Single token that IS just a label
            if len(_toks) == 1 and _vu in _LABEL_TOKENS: return
            # Reject if contains digits (names don't have digits)
            if _re.search(r'\d', v): return
            # Reject if looks like an address (3+ commas = district/thana/village format)
            if v.count(',') >= 3: return
            # Reject if contains known address keywords
            _ADDR_KEYWORDS = {'ROAD','LANE','BLOCK','HOUSE','FLAT','VILLAGE','MOUZA',
                              'THANA','DISTRICT','UPAZILA','WARD','UNION','BAZAR',
                              'SADAR','PARA','CHOWK','CANTONMENT','RAJBANDH',
                              'STREET','AVENUE','NAGAR','GRAM','TOLA'}
            _name_words = set(_vu.split())
            if len(_name_words & _ADDR_KEYWORDS) >= 2: return
        # OCR fix for name fields — common tesseract misreads
        if field in ('name','father','mother','spouse'):
            _ocr_pairs = [
                ('Suralya','Suraiya'),('SURALYA','SURAIYA'),
                ('Suratya','Suraiya'),('SURATYA','SURAIYA'),
                ('Shabjahan','Shahjahan'),('SHABJAHAN','SHAHJAHAN'),
            ]
            for _w, _r in _ocr_pairs:
                v = v.replace(_w, _r)
        # Blood group: OCR "At" → "A+"
        if field == 'blood_group':
            v = _re.sub(r'^At$', 'A+', v)
            v = _re.sub(r'^Ot$', 'O+', v)
            v = _re.sub(r'^Bt$', 'B+', v)
        # ── Profession field — reject person-name-like values ─────────────
        # DL OCR সময় profession-এ ব্যক্তির নিজের নাম + signature noise আসতে পারে
        # ('MD. FARHAD HOSSAIN Seem pete' টাইপ)
        if field == 'profession':
            _vu = v.upper().strip()
            # Strip OCR signature noise (single random word at end)
            v = _re.sub(r'\s+\b(Seem|seer|See|sees|pete|peter)\s*\w*\s*$',
                        '', v, flags=_re.I).strip()
            # Reject if starts with person-name prefix (MD./MR./MRS.) + ALL CAPS
            if _re.match(r'^(MD\.?|MR\.?|MRS\.?|MS\.?)\s+[A-Z]', v):
                # If after stripping name pattern nothing meaningful left → reject
                # But keep things like "Md. Government Service" — not real, skip
                return
            # Reject if entire value is a person name (3-5 uppercase tokens, no digits, no role words)
            _toks = v.upper().split()
            _ROLE_WORDS = {'SERVICE','OFFICER','STAFF','MANAGER','TEACHER',
                           'DOCTOR','ENGINEER','BUSINESS','BUSINESSMAN','FARMER',
                           'STUDENT','EMPLOYEE','AGRICULTURE','GOVT','GOVERNMENT',
                           'PRIVATE','PVT','RETIRED','AUTONOMOUS','ORGANIZATION',
                           'BANK','COMPANY','LIMITED','SHOP','PROFESSIONAL'}
            if (2 <= len(_toks) <= 5 and
                    all(_re.match(r'^[A-Z]+$', t) for t in _toks) and
                    not any(t in _ROLE_WORDS for t in _toks)):
                return  # Looks like a person name, not a profession
        # ── Vehicle Type — strip leading nationality leak ────────────────
        # OCR "Nationality\nBangladesh\nVehicle Classes\nLightVehicle..." এর জন্য
        # "Bangladesh LightVehicle" আসত — সেটা ঠিক করি
        if field == 'vehicle_type':
            v = _re.sub(r'^(Bangladesh[i]?|Bangladeshi)\s+', '', v, flags=_re.I).strip()
            v = _re.sub(r'^(Nationality|License\s*Type)\s*[:\-]?\s*',
                        '', v, flags=_re.I).strip()
        # ── License Issue Date — must look like a valid date, not just any value ─
        if field == 'license_issue':
            # Must be actual date format
            if not _re.search(r'\d{4}-\d{2}-\d{2}|\d{2}/\d{2}/\d{4}', v): return
            # Reject if value matches an already-stored DOB (sign of mis-mapping)
            for _ev, _ in store.get('dob', []):
                if _ev.strip() == v.strip(): return
        # ── Address field — reject OCR garbage from failed Bengali OCR ──
        if field in ('address_permanent', 'address_present'):
            _vs = v.strip()
            # Reject if starts with stray punctuation (OCR noise)
            if _re.match(r'^[_=\-\.,]+\s', _vs): return
            # Clean trailing OCR garbage: remove short non-word sequences
            # e.g. "BIYZID CANTONMENT, -, BAYJID BOSTAMI, a eed cai BAYJID..."
            # Anything after the last valid address segment
            # Pattern: lowercase 1-2 char tokens mixed with garbage → strip from that point
            def _clean_addr_tail(s):
                # Find last position where address content is valid
                # Remove sequences of "short garbage tokens" at the end
                parts = _re.split(r',\s*', s)
                good_parts = []
                for pt in parts:
                    pt = pt.strip()
                    if not pt: continue
                    # '-' placeholder (যেমন RAFIQPUR, -, JOMIDAR HAT) — সবসময় রাখো
                    if pt == '-':
                        good_parts.append(pt)
                        continue
                    toks = pt.split()
                    # Garbage detection:
                    # 1) short tokens (<=3 char) with no vowel — e.g. 'cai' NOT counted (has vowel)
                    # 2) BUT lowercase tokens in an otherwise uppercase address = OCR garbage
                    _is_mostly_upper = sum(1 for t in toks if t[0].isupper()) >= len(toks) * 0.5 if toks else False
                    _garbage_toks = sum(1 for t in toks
                                       if t != '-' and (
                                           # no-vowel short tokens
                                           (len(t) <= 3 and not _re.search(r'[aeiouAEIOU]', t)) or
                                           # lowercase token in uppercase address = OCR garbage
                                           (t.islower() and len(t) <= 4)
                                       ))
                    if len(toks) > 0 and _garbage_toks / len(toks) > 0.5 and len(pt) < 20:
                        break  # stop here, this part is garbage
                    good_parts.append(pt)
                return ', '.join(good_parts) if good_parts else s

            v = _clean_addr_tail(_vs)
            if not v or len(v) < 5: return

            # Reject if too many lowercase-only short tokens (Bengali OCR garbage)
            _tokens = [t for t in _re.split(r'[\s,]+', v) if t]
            if _tokens:
                _short_garbage = sum(
                    1 for t in _tokens
                    if 2 <= len(t) <= 4
                    and _re.match(r'^[A-Za-z]+$', t)
                    and not _re.search(r'[aeiouAEIOU]', t)
                )
                if len(_tokens) >= 3 and _short_garbage / len(_tokens) >= 0.4:
                    return
                _has_word = any(
                    len(t) >= 4 and _re.match(r'^[A-Za-z]+$', t)
                    for t in _tokens
                )
                if not _has_word: return
        lbl = doc_label or _current_doc_label[0]
        if field in store:
            # Avoid exact duplicate (same value from same doc)
            if not any(existing_v == v for existing_v, _ in store[field]):
                store[field].append((v, lbl))

    def _get_text_and_words(raw_bytes):
        full_text = ''
        all_words = []
        _ext_photo = None

        # ── pdftotext fallback: Bengali font (SolaimanLipi) সঠিকভাবে পড়তে পারে ──
        _pdftotext_txt = ''
        try:
            import subprocess as _sp
            _res = _sp.run(['pdftotext', '-layout', '-', '-'],
                           input=raw_bytes, capture_output=True, timeout=15)
            if _res.returncode == 0:
                _pdftotext_txt = _res.stdout.decode('utf-8', errors='replace')
        except Exception:
            logger.debug('suppressed exception', exc_info=True)

        try:
            # pdfplumber → _pdfplumber_mod (module-level import)
            with _pdfplumber_mod.open(_io.BytesIO(raw_bytes)) as pdf:
                for _pi, page in enumerate(pdf.pages):
                    pt = page.extract_text() or ''
                    # pdftotext দিয়ে better text পাওয়া গেলে সেটা ব্যবহার করো
                    if _pdftotext_txt.strip():
                        full_text = _pdftotext_txt
                    else:
                        full_text += pt + '\n'
                    try:
                        words = page.extract_words(x_tolerance=3, y_tolerance=3,
                                                   keep_blank_chars=False)
                        all_words.extend(words)
                    except Exception:
                        logger.debug('suppressed exception', exc_info=True)
                    # Image-based page → OCR fallback
                    if not pt.strip():
                        try:
                            # pytesseract → _pytesseract_mod (module-level import)
                            _img = page.to_image(resolution=200).original
                            _ocr = _pytesseract_mod.image_to_string(_img, lang='eng')
                            full_text += _ocr + '\n'
                        except Exception:
                            logger.debug('suppressed exception', exc_info=True)
                    # Person photo extraction from PDF
                    # Only consider images that are portrait-oriented (taller than wide)
                    # and larger than typical logos (> 100px height)
                    if not _ext_photo:
                        try:
                            for _img in (page.images or []):
                                _iw = float(_img.get('width', 0))
                                _ih = float(_img.get('height', 0))
                                # Portrait: height > width, reasonable size
                                if _ih > _iw * 1.1 and _ih > 100 and _iw > 60:
                                    _bbox = (max(0,float(_img['x0'])), max(0,float(_img['top'])),
                                             min(page.width, float(_img['x1'])),
                                             min(page.height, float(_img['bottom'])))
                                    _cr = page.within_bbox(_bbox).to_image(resolution=150)
                                    _pil = _cr.original
                                    if _pil.mode in ('P','RGBA','LA'):
                                        _pil = _pil.convert('RGB')
                                    _buf = _io.BytesIO()
                                    _pil.save(_buf, format='JPEG', quality=85)
                                    if len(_buf.getvalue()) > 5000:  # real photo > 5KB
                                        _ext_photo = ("data:image/jpeg;base64,"
                                                      + _b64.b64encode(_buf.getvalue()).decode())
                                        break
                        except Exception:
                            logger.debug('suppressed exception', exc_info=True)
        except Exception:
            logger.debug('suppressed exception', exc_info=True)
        return full_text, all_words, _ext_photo

    def _detect_doc_type(tu):
        # TIN must come BEFORE passport (TIN PDF contains 'PASSPORT NUMBER' as N/A)
        if 'ASSESSEE NAME' in tu or 'OLD TIN' in tu:
            return 'tin'
        if 'OLD NID NUMBER' in tu or 'NEW NID NUMBER' in tu:
            return 'nid'
        if 'PASSPORT NUMBER' in tu or ('DATE OF ISSUE' in tu and 'PASSPORT' in tu):
            return 'passport'
        if 'VEHICLE REGISTRATION NUMBER' in tu:
            return 'vehicle'
        if ('LICENSE NO' in tu or 'LICENCE NO' in tu or
                'VEHICLE CLASSES' in tu or
                ('ISSUING AUTHORITY' in tu and 'BRTA' in tu)):
            return 'driving_license'
        # ── NID image OCR fallback (no clear label) ──────────────────────
        # NID has: 17-digit NID number + "Permanent Address" + "Present Address"
        # + "Old NID Number" / "New NID Number" (OCR-noisy variant)
        _has_nid17 = bool(_re.search(r'\b19\d{15}\b', tu))
        _has_addr  = 'PERMANENT ADDRESS' in tu and 'PRESENT ADDRESS' in tu
        _has_nid_keywords = any(k in tu for k in [
            'OLD NID', 'NEW NID', 'NATIONAL ID', 'NATIONAL IDENTITY',
            'OLD NO', 'NEW NO', 'NID NUMBER',
        ])
        if _has_nid17 and (_has_addr or _has_nid_keywords):
            return 'nid'
        return 'unknown'

    def _store_has(field):
        """Return True if store has any value for field."""
        return bool(store.get(field))

    def _store_vals(field):
        """Return list of values (without labels) for field."""
        return [v for v, _ in store.get(field, [])]

    # ── PARSER: NID ──────────────────────────────────────────────────────
    def _parse_nid(txt, lines):
        """
        NID থেকে শুধু নিচের তথ্য নেওয়া হবে:
          - NID Number (17-digit)
          - Smart ID / New NID Number
          - Blood Group
        বাকি সব (নাম, ঠিকানা, পিতা/মাতা, DOB ইত্যাদি) NID থেকে নেওয়া হবে না।
        সেগুলো Passport / DL / TIN / Vehicle Reg থেকে আসবে।
        """
        # ── NID Number ──────────────────────────────────────────────────
        for l in lines:
            ls = l.strip()
            m = _re.match(r"^(19\d{15})\s+(\d{7,12})\s*$", ls)
            if m:
                _add("nid", m.group(1))
                _add("nid_new", m.group(2))
                break
        if not _store_has("nid"):
            m = _re.search(r"\b(19\d{15})\b", txt)
            if m: _add("nid", m.group(1))
        if not _store_has("nid_new"):
            m = _re.search(r"\b(19\d{15})\b\s+\b(\d{7,12})\b", txt)
            if m: _add("nid_new", m.group(2))
        if not _store_has("nid_new"):
            # 10-digit fallback — reject phone numbers (starts 880/01) and postal codes (4-6 digit)
            all_nums = _re.findall(r"\b(\d{10})\b", txt)
            for n in all_nums:
                if n.startswith(('8801','01')): continue  # phone number
                if len(set(n)) <= 2: continue              # repeated digits e.g. 0000000000
                _add("nid_new", n); break

        # ── Blood Group ─────────────────────────────────────────────────
        for i, l in enumerate(lines):
            if "Blood Group" in l or "Blood group" in l:
                for j in range(i+1, min(i+6, len(lines))):
                    bg = lines[j].strip()
                    if not bg: continue
                    tok = bg.split()[0] if bg.split() else ""
                    if _re.match(r"^[ABO]{1,3}[+-]$", tok):
                        _add("blood_group", tok); break
                    if _re.match(r"^[ABO]{1,3}[+-]$", bg):
                        _add("blood_group", bg); break
                    if _re.match(r"^(Date|Gender|Occupation|Spouse|Father|Mother)", bg):
                        break
                break
        if not _store_has("blood_group"):
            m = _re.search(r"\b([ABO]{1,2}[+-])\b", txt)
            if m: _add("blood_group", m.group(1))

    # ── PARSER: Driving License ─────────────────────────────────────────
    def _parse_driving_license(txt, lines):
        """
        두 가지 format 지원:
        1. NTMC multi-column: "Father Name  Expiry Date\nMD ABUL KALAM  14/02/2031"
        2. Scanned OCR: "Father Name\nMD ABUL KALAM"
        """
        # Detect format: NTMC has "Personal Info License Detail" header
        _is_ntmc = 'Personal Info' in txt and 'License Detail' in txt

        if _is_ntmc:
            # ── NTMC format: use exact line patterns ──────────────────────
            # Name: standalone CAPS line
            for l in lines:
                l = l.strip()
                if (_re.match(r'^MD\.?\s+[A-Z]+\s+[A-Z]+$', l) and
                        'CENTRE' not in l and 'METRO' not in l and 'BRTA' not in l):
                    _add('name', l); break

            # Line: "22/06/1970 15/02/2021" → DOB + Issue
            for l in lines:
                m = _re.match(r'^(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})$', l.strip())
                if m: _add('dob', m.group(1)); _add('license_issue', m.group(2)); break

            # Line: "MD ABUL KALAM 14/02/2031" → father + expiry
            for l in lines:
                m = _re.match(r'^([A-Z][A-Z ]+?)\s+(\d{2}/\d{2}/\d{4})$', l.strip())
                if m: _add('father', m.group(1).strip()); _add('license_expiry', m.group(2)); break

            # License No: "N/A DK1182865L00001"
            for l in lines:
                m = _re.search(r'N/A\s+([A-Z]{2}\d{7}[A-Z]\d{5})', l)
                if m: _add('license_no', m.group(1)); break

            # License Type: "N/A NON-PROFESSIONAL"
            for l in lines:
                m = _re.search(r'N/A\s+(NON-PROFESSIONAL|PROFESSIONAL|MEDIUM|LIGHT)', l, _re.I)
                if m: _add('license_type', m.group(1)); break

            # Gender: "MALE DHAKA METRO..."
            for l in lines:
                m = _re.match(r'^(MALE|FEMALE)\s+DHAKA', l, _re.I)
                if m: _add('gender', m.group(1).upper()); break

            # Spouse + address: "H-30,...RAMPURA, SPOUSE NAME DD/MM/YYYY"
            # Format: address_part, CAPS_NAME DATE
            for l in lines:
                m = _re.search(
                    r'([A-Z][A-Z .]{3,}?)\s+(\d{2}/\d{2}/\d{4})\s*$', l.strip()
                )
                if m:
                    _spouse_candidate = m.group(1).strip().rstrip(',').strip()
                    # Must look like a name: 2-4 words, no junk keywords
                    _sp_words = _spouse_candidate.split()
                    _JUNK_KW = {'ROAD','LANE','BLOCK','HOUSE','FLOOR','FLAT',
                                'VILLAGE','THANA','DISTRICT','METRO','DHAKA',
                                'CHITTAGONG','RAJSHAHI','NATIONAL','MONITORING'}
                    # Reject if it matches the already-captured father name
                    _captured_father = (store.get('father') or [('','')])[0][0]
                    _is_father = (
                        _captured_father and
                        _spouse_candidate.upper() == _captured_father.upper()
                    )
                    if (2 <= len(_sp_words) <= 4 and
                            not set(_sp_words) & _JUNK_KW and
                            len(_spouse_candidate) < 50 and
                            not _is_father):
                        _add('spouse', _spouse_candidate)
                        addr = l[:m.start()].strip().rstrip(',').strip()
                        if addr and len(addr) > 5:
                            _add('address_permanent', addr)
                    break

            # Present address — flexible matching
            for i, l in enumerate(lines):
                if 'Present Address' in l:
                    for j in range(i+1, min(i+5, len(lines))):
                        a = lines[j].strip()
                        # Accept any address-like line (has comma or road/block/house)
                        if (a and len(a) > 8 and
                                not a.upper().startswith('PERMANENT') and
                                not a.upper().startswith('THANA')):
                            addr = _re.sub(r'\s+Emergency.*$', '', a, flags=_re.I).strip()
                            addr = _re.sub(r',?\s*Khilgaon.*$', '', addr).strip()
                            if len(addr) > 5:
                                _add('address_present', addr); break
                    break

        else:
            # ── Scanned OCR format: label on one line, value on next ──────
            def _next_val(label):
                for i, l in enumerate(lines):
                    if label.lower() in l.lower():
                        for j in range(i+1, min(i+3, len(lines))):
                            v = lines[j].strip()
                            if v and v.upper() not in ('N/A','NA',''):
                                return v
                return ''

            # Name: caps line near top
            for l in lines[:10]:
                if _re.match(r'^MD\.?\s+[A-Z]+\s+[A-Z]+$', l.strip()):
                    _add('name', l.strip()); break

            # Known DL field labels — if _next_val returns one of these, it's a mis-parse
            _DL_LABELS = {
                'emergency contact relation', 'emergency contact name',
                'emergency mobile number', 'nid number', 'spouse name',
                'marital status', 'other occupation', 'mobile number',
                'date of birth', 'father name', 'mother name', 'gender',
                'occupation', 'nationality', 'license type', 'vehicle classes',
            }
            for label, field in [
                ('Gender','gender'), ('Father Name','father'),
                ('Mother Name','mother'), ('Nationality','nationality'),
                ('Occupation','profession'), ('Spouse Name','spouse'),
                ('License Type','license_type'), ('Vehicle Classes','vehicle_type'),
            ]:
                v = _next_val(label)
                if v:
                    # Reject if value is a field label (mis-parse)
                    if v.strip().lower() in _DL_LABELS:
                        continue
                    # Strip trailing comma/semicolon from gender etc.
                    v = v.strip().rstrip(',;').strip()
                    # ── Multi-column OCR leak fix ─────────────────────────────
                    # NTMC DL has 3-column layout. OCR collapses to single col but
                    # values from adjacent columns can leak in.
                    # Strip known left-column values that appear before the actual value:
                    if field == 'license_type':
                        # "Married Medium Vehicle" → "Medium Vehicle"
                        v = _re.sub(r'^(Married|Unmarried|Single|Divorced|Widowed)\s+',
                                    '', v, flags=_re.I).strip()
                    if field == 'vehicle_type':
                        # "Bangladesh LightVehicle, Motorcycle" → "LightVehicle, Motorcycle"
                        v = _re.sub(r'^(Bangladesh[i]?|Nationality)\s+', '', v, flags=_re.I).strip()
                    if field == 'profession':
                        # "N/A 29-Nov-2023" → drop dates leaked from Apply Date column
                        v = _re.sub(r'\s+\d{1,2}[-/]\w{3,}[-/]\d{2,4}\s*$', '', v).strip()
                    if field == 'gender':
                        # "MALE SYLHET, BRTA" → "MALE"
                        m = _re.match(r'^(MALE|FEMALE)\b', v, _re.I)
                        if m: v = m.group(1).upper()
                    if field == 'nationality':
                        # "Bangladesh Reference Number" → "Bangladesh"
                        m = _re.match(r'^(Bangladesh[i]?|Bangladeshi)\b', v, _re.I)
                        if m: v = m.group(1)
                    if field == 'father' or field == 'mother':
                        # Strip trailing date or license-no leaked from right column
                        v = _re.sub(r'\s+\d{4}[-/]\d{2}[-/]\d{2}\s*$', '', v).strip()
                        v = _re.sub(r'\s+SL\d+\w*\s*$', '', v).strip()
                    if field == 'spouse':
                        # Reject if spouse value matches already-captured father/mother name
                        _cap_father = (store.get('father') or [('','')])[0][0]
                        _cap_mother = (store.get('mother') or [('','')])[0][0]
                        if ((_cap_father and v.upper() == _cap_father.upper()) or
                                (_cap_mother and v.upper() == _cap_mother.upper())):
                            continue
                    _add(field, v)

            # Mobile
            v = _next_val('Mobile Number')
            if v:
                m = _re.search(r'(0?1[3-9]\d{8}|880\d{10})', v)
                if m:
                    mob = m.group(1)
                    if mob.startswith('880'): mob = '0' + mob[3:]
                    _add('mobile', mob)

            # NID — only if actual NID present (not N/A)
            v = _next_val('NID Number')
            if v and v.upper() not in ('N/A','NA',''):
                # Only accept full 17-digit NID starting with 19
                m = _re.search(r'\b(19\d{15})\b', v)
                if m: _add('nid', m.group(1))

            # DOB
            v = _next_val('Date of Birth')
            if v:
                m = _re.search(r'(\d{2}[/-]\d{2}[/-]\d{4}|\d{4}-\d{2}-\d{2})', v)
                if m: _add('dob', m.group(1))

            # License No (whole text regex)
            m = _re.search(r'\b([A-Z]{2}\d{7}[A-Z]\d{5})\b', txt)
            if m: _add('license_no', m.group(1))

            # Issue/Expiry dates
            for label, field in [('Issue Date','license_issue'), ('Expiry Date','license_expiry')]:
                v = _next_val(label)
                if v:
                    m = _re.search(r'(\d{4}-\d{2}-\d{2}|\d{2}/\d{2}/\d{4})', v)
                    if m: _add(field, m.group(1))

            # Addresses
            # ── DL Addresses (scanned OCR — NTMC 3-column layout) ─────────
            # OCR দিচ্ছে:
            #   'Permanent Address: ASMA KHATOON'  ← label + spouse mixed
            #   'Thana Diviston Post Code ...'     ← header (skip)
            #   'Satkhira'                          ← Thana
            #   'Sauer KHULNA 9400 ...'            ← Division + postcode mixed
            #   'Ni ...'                            ← noise
            #   'Presest Acitreee ...'             ← "Present Address" (OCR-noisy)
            #   'Thana Diision Post Code'           ← header (skip)
            #   'Sylhet /'                          ← Thana
            #   'Eeeat  SYLHET 3100'               ← Division + postcode

            def _dl_addr_from_thana_block(start_idx, max_lines=8):
                """
                'Permanent/Present Address' label পরে Thana/Division block থেকে
                address তৈরি করে। Format:
                  header line (Thana Division Post Code) → skip
                  thana line  (1-2 words, location name)
                  division line (division_name + postcode noise)
                """
                thana = ''; division = ''; postcode = ''
                _HEADER_RE = _re.compile(
                    r'\b(Thana|Upazila|Division|Post\s*Code|Bnecvency|Contact)\b', _re.I)
                _DIV_NAMES = ('DHAKA','CHITTAGONG','CHATTOGRAM','RAJSHAHI','KHULNA',
                              'BARISAL','BARISHAL','SYLHET','MYMENSINGH','RANGPUR',
                              'COMILLA','CUMILLA','NOAKHALI','GAZIPUR')
                _state = 'find_header'
                for j in range(start_idx, min(start_idx + max_lines, len(lines))):
                    lj = lines[j].strip()
                    if not lj: continue
                    if _state == 'find_header':
                        if _HEADER_RE.search(lj):
                            _state = 'find_thana'
                    elif _state == 'find_thana':
                        # Thana: 1-2 word line with capital letters
                        if _HEADER_RE.search(lj): continue  # next header
                        words = lj.split()
                        # Remove trailing N/A, numbers
                        clean_words = [w for w in words if w.upper() != 'N/A'
                                       and not _re.match(r'^\d+$', w)]
                        if 1 <= len(clean_words) <= 3:
                            thana = ' '.join(clean_words)
                            _state = 'find_division'
                    elif _state == 'find_division':
                        if _HEADER_RE.search(lj): break  # next address block started
                        # Division: line containing a known division name
                        for dname in _DIV_NAMES:
                            if dname in lj.upper():
                                division = dname
                                # Postcode: 4-digit number in this line
                                pc_m = _re.search(r'\b(\d{4})\b', lj)
                                if pc_m: postcode = pc_m.group(1)
                                break
                        if division: break
                if thana and division:
                    # Clean thana — strip trailing OCR noise
                    thana = _re.sub(r'\s*[/\\]+\s*$', '', thana).strip()
                    thana = _re.sub(r'\s*\d+\s*$', '', thana).strip()
                    # Build: "Satkhira Sadar, KHULNA" / "Sylhet, SYLHET"
                    parts = [p for p in [thana, division] if p]
                elif thana:
                    thana = _re.sub(r'\s*[/\\]+\s*$', '', thana).strip()
                    parts = [thana]
                else:
                    parts = []
                return ', '.join(parts) if parts else ''

            # Find "Permanent Address" and "Present Address" in DL OCR
            _ADDR_LABELS = [
                ('Permanent Address', 'address_permanent'),
                ('Present Address',   'address_present'),
            ]
            # Also match OCR-noisy variants: 'Presest Acitreee', 'Presentaddress' etc.
            def _matches_addr_label(line_str, label):
                """Check if line contains the address label (OCR-tolerant)."""
                l = line_str.lower().strip().lstrip("'\"")
                if label.lower() in l: return True
                # OCR noise patterns for "Permanent" and "Present"
                if 'permanent' in label.lower():
                    return _re.search(r'perma\w*\s+addr', l) is not None
                if 'present' in label.lower():
                    return bool(_re.search(r'pre[a-z]*\s+a[a-z]*(?:ress|tress|itre)', l))
                return False

            for lbl, field in _ADDR_LABELS:
                if _store_has(field): continue
                for i, l in enumerate(lines):
                    if _matches_addr_label(l, lbl):
                        addr = _dl_addr_from_thana_block(i+1)
                        if addr:
                            _add(field, addr)
                        break

        # Common: any format
        if not _store_has('license_no'):
            m = _re.search(r'\b([A-Z]{2}\d{7}[A-Z]\d{5})\b', txt)
            if m: _add('license_no', m.group(1))


    # ── PARSER: Vehicle Registration ─────────────────────────────────────
    def _parse_vehicle(txt, lines):
        # Vehicle Reg PDF has 3-column layout — labels and values on separate lines:
        #   "Owner's Name Vehicle Number"  ← labels merged (skip)
        #   "MD. JAHIRUL ISLAM 34-5505"    ← values merged
        # So: find the label line, then look for the NEXT line that is an
        # actual NAME value (starts with MD./MR./MRS. or all-caps words),
        # NOT another label like "Vehicle Number" or "Registration Number".

        _LABEL_WORDS = ('vehicle', 'registration', 'number', 'name', 'detail',
                        'license', 'info', 'owner', 'father', 'mother', 'nid',
                        'mobile', 'dob', 'nationality', 'address', 'joint',
                        'date', 'office', 'tax', 'token', 'route', 'permit',
                        'fitness', 'class', 'type', 'series', 'color', 'colour',
                        'weight', 'capacity', 'axle', 'expire', 'issue', 'cc')

        def _is_label_line(s):
            """A line that's only field labels (no actual name value)."""
            sl = s.lower().strip()
            if not sl: return True
            # If line is mostly label words → it's a label line
            words = [w for w in _re.split(r'[\s,]+', sl) if w]
            if not words: return True
            label_hits = sum(1 for w in words if any(lw in w for lw in _LABEL_WORDS))
            return label_hits >= max(1, len(words) // 2)

        def _looks_like_name(s):
            """Real person name: MD./MR./MRS. prefix or 2+ all-caps words."""
            s = s.strip()
            if len(s) < 5: return False
            if _is_label_line(s): return False
            # Strip trailing vehicle number / digits
            s_clean = _re.sub(r'\s+[\d\-]+\s*$', '', s).strip()
            if _re.match(r'^(MD\.?|MR\.?|MRS\.?|MST\.?|MOST\.?|MOHAMMAD|MOHAMMED)\s+[A-Z]', s_clean, _re.I):
                return True
            # 2+ consecutive all-caps words (e.g. "JAHIRUL ISLAM")
            caps = _re.findall(r'\b[A-Z]{2,}\b', s_clean)
            return len(caps) >= 2

        # ── Owner's Name — find label, then next NAME-looking line ──────────
        _name_found = False
        for i, l in enumerate(lines):
            if "owner's name" in l.lower() or "owners name" in l.lower():
                # Check same line after label (in case value is appended)
                rest = _re.split(r"owner'?s?\s+name", l, flags=_re.I)[-1].strip()
                # rest must be a name, not another label like "Vehicle Number"
                if rest and _looks_like_name(rest):
                    val = _re.sub(r'\s+[\d\-]+\s*$', '', rest).strip()
                    _add('name', val); _name_found = True; break
                # Otherwise scan next few lines for the actual name value
                for j in range(i+1, min(i+6, len(lines))):
                    cand = lines[j].strip()
                    if _looks_like_name(cand):
                        val = _re.sub(r'\s+[\d\-]+\s*$', '', cand).strip()
                        _add('name', val); _name_found = True; break
                if _name_found: break

        # Fallback: "MD. JAHIRUL ISLAM 34-5505" (name + vehicle number)
        if not _store_has('name'):
            for l in lines:
                m = _re.match(r'^(MD\.?\s+[A-Z]+(?:\s+[A-Z]+)+)\s+[\d\-]+\s*$', l.strip())
                if m:
                    _add('name', m.group(1).strip()); break

        # line 11: "MD. ABUL KALAM CAR (SALOON)" → father + vehicle type
        for l in lines:
            m = _re.match(r'^(MD\.?\s+[A-Z ]+?)\s{2,}(CAR|TRUCK|BUS|JEEP|HARD JEEP|VAN|MICROBUS)', l, _re.I)
            if m:
                _add('father', m.group(1).strip())
                _add('vehicle_type', m.group(2).strip()); break
        if not _store_has('father'):
            # Fallback: "Father's Name" label → next name-looking line
            for i, l in enumerate(lines):
                if "father's name" in l.lower() or "father name" in l.lower():
                    rest = _re.split(r"father'?s?\s+name", l, flags=_re.I)[-1].strip()
                    if rest and _looks_like_name(rest):
                        val = _re.sub(r'\s+(CAR|TRUCK|BUS|JEEP|VAN|MICROBUS).*$', '', rest, flags=_re.I).strip()
                        _add('father', val); break
                    for j in range(i+1, min(i+6, len(lines))):
                        cand = lines[j].strip()
                        if _looks_like_name(cand):
                            val = _re.sub(r'\s+(CAR|TRUCK|BUS|JEEP|VAN|MICROBUS).*$', '', cand, flags=_re.I).strip()
                            val = _re.sub(r'\s+[\d\-]+\s*$', '', val).strip()
                            _add('father', val); break
                    break
        # Vehicle Type: label-based extraction (image OCR format)
        if not _store_has('vehicle_type'):
            _VT_TYPES = ('CAR', 'SALOON', 'JEEP', 'HARD JEEP', 'TRUCK', 'BUS',
                         'VAN', 'MICROBUS', 'PICK UP', 'MOTOR CYCLE', 'AUTO RICKSHAW',
                         'AMBULANCE', 'DUMPER', 'TANKER', 'TRACTOR', 'THREE WHEELER',
                         'DOUBLE CABIN', 'WAGON', 'ST. WAGN', 'PVT. PASS')
            for i, l in enumerate(lines):
                if 'Vehicle Type' in l:
                    # Same line or next line
                    rest = l.split('Vehicle Type')[-1].strip()
                    if rest and rest.upper() not in ('N/A', 'NA', ''):
                        _add('vehicle_type', rest.split()[0] if rest.split() else rest); break
                    if i+1 < len(lines):
                        v = lines[i+1].strip()
                        if v and v.upper() not in ('N/A', 'NA', ''):
                            _add('vehicle_type', v); break
                    break
        # Vehicle Class: PVT. PASS. (JEEP/ ST. WAGN)
        if not _store_has('vehicle_type'):
            for i, l in enumerate(lines):
                if 'Vehicle Class' in l:
                    rest = l.split('Vehicle Class')[-1].strip()
                    if rest and rest.upper() not in ('N/A', 'NA', ''):
                        _add('vehicle_type', rest); break
                    if i+1 < len(lines):
                        v = lines[i+1].strip()
                        if v and v.upper() not in ('N/A', 'NA', ''):
                            _add('vehicle_type', v); break
                    break

        # Mobile: line 19 "8801711385207 1500" (PDF multi-column format)
        # Image OCR format: "Mobile Number\n8801711385207"
        if not _store_has('mobile'):
            for i, l in enumerate(lines):
                if 'MOBILE' in l.upper() and ('NUMBER' in l.upper() or 'MOBILE' == l.strip().upper()):
                    for j in range(i+1, min(i+4, len(lines))):
                        mv = lines[j].strip()
                        m = _re.search(r'\b(8801[3-9]\d{8}|01[3-9]\d{8})\b', mv)
                        if m:
                            mob = m.group(1)
                            if mob.startswith('880'): mob = '0' + mob[3:]
                            _add('mobile', mob); break
                    break
        if not _store_has('mobile'):
            for l in lines:
                m = _re.search(r'\b(8801\d{9})\b.*?(\d{3,4})\s*$', l)
                if not m: m = _re.search(r'\b(8801\d{9})\s+(\d+)', l)
                if m:
                    _add('mobile', '0' + m.group(1)[3:])
                    if not _store_has('vehicle_cc'):
                        _add('vehicle_cc', m.group(2))
                    break
        # Vehicle reg: DHAKA METRO anywhere in line (multi-column layout)
        for l in lines:
            m = _re.search(r'(DHAKA[ -]METRO-[A-Z]+-\d+-\d+)', l, _re.I)
            if m: _add('vehicle_reg', m.group(1).upper()); break
        # Fallback: "Vehicle Registration Number" label → next value token
        if not _store_has('vehicle_reg'):
            for i, l in enumerate(lines):
                if 'Vehicle Registration Number' in l:
                    # Check same line after label
                    rest = l.split('Vehicle Registration Number')[-1].strip()
                    m = _re.search(r'(DHAKA[ -]METRO-[A-Z]+-\d+-\d+)', rest, _re.I)
                    if m:
                        _add('vehicle_reg', m.group(1).upper()); break
                    # Or next line
                    if i+1 < len(lines):
                        m = _re.search(r'(DHAKA[ -]METRO-[A-Z]+-\d+-\d+)', lines[i+1], _re.I)
                        if m: _add('vehicle_reg', m.group(1).upper()); break
        # Color: "NO BLACK" → BLACK OR label-based "Vehicle Color\nPEARL"
        if not _store_has('vehicle_color'):
            _VC_COLORS = {'BLACK','WHITE','SILVER','RED','BLUE','GREEN','GREY','GRAY',
                          'PEARL','GOLDEN','YELLOW','ORANGE','MAROON','BROWN','BEIGE',
                          'PURPLE','PINK','INDIGO','CREAM','WINE'}
            for i, l in enumerate(lines):
                if 'Vehicle Color' in l:
                    # In multi-column: take value between "Vehicle Color" and next column
                    # split by 2+ spaces to get columns, find the "Vehicle Color" column value
                    cols = _re.split(r'\s{2,}', l)
                    for ci, col in enumerate(cols):
                        if 'Vehicle Color' in col:
                            # Next column may be on same line or next line
                            if ci+1 < len(cols):
                                cval = cols[ci+1].strip()
                                if cval and cval.upper() not in ('N/A','NA','') and cval.upper() in _VC_COLORS:
                                    _add('vehicle_color', cval.split()[0]); break
                    if _store_has('vehicle_color'): break
                    # Next line — same column position
                    if i+1 < len(lines):
                        cols2 = _re.split(r'\s{2,}', lines[i+1])
                        for ci, col in enumerate(cols):
                            if 'Vehicle Color' in col and ci < len(cols2):
                                cval = cols2[ci].strip()
                                if cval and cval.upper() in _VC_COLORS:
                                    _add('vehicle_color', cval); break
                    if _store_has('vehicle_color'): break
        if not _store_has('vehicle_color'):
            _COLORS = {'BLACK','WHITE','SILVER','RED','BLUE','GREEN','GREY','GRAY',
                       'PEARL','GOLDEN','YELLOW','ORANGE','MAROON','BROWN','BEIGE',
                       'PURPLE','PINK','INDIGO','CREAM','WINE'}
            for l in lines:
                # "NO BLACK N/A" or "NO  BLACK"
                m = _re.search(r'\bNO\s+([A-Z]+)\b', l)
                if m and m.group(1) in _COLORS:
                    _add('vehicle_color', m.group(1)); break

        # Vehicle CC: "Vehicle CC\n1797" label-based
        if not _store_has('vehicle_cc'):
            for i, l in enumerate(lines):
                if 'VEHICLE CC' in l.upper() or 'Vehicle CC' in l:
                    rest = _re.split(r'Vehicle CC', l, flags=_re.I)[-1].strip()
                    if rest and _re.match(r'^\d+', rest):
                        _add('vehicle_cc', rest.split()[0]); break
                    if i+1 < len(lines):
                        v = lines[i+1].strip()
                        if v and _re.match(r'^\d+$', v):
                            _add('vehicle_cc', v); break

        # Tax Token / Fitness: "2026-09-22" date format after label
        for i, l in enumerate(lines):
            if 'Tax Token Expire' in l:
                for j in range(i, min(i+3, len(lines))):
                    m2 = _re.search(r'(\d{4}-\d{2}-\d{2})', lines[j])
                    if m2: _add('tax_token_expire', m2.group(1)); break
                break
        for i, l in enumerate(lines):
            if 'Fitness Expire' in l:
                for j in range(i, min(i+3, len(lines))):
                    m2 = _re.search(r'(\d{4}-\d{2}-\d{2})', lines[j])
                    if m2: _add('fitness_expire', m2.group(1)); break
                break
        # Address: Owner's Address label-based
        _VEHICLE_LABELS = {'vehicle series', 'vehicle number', 'registration',
                           'route permit', 'fitness', 'joint owner', 'license'}
        for i, l in enumerate(lines):
            if "Owner's Address" in l or "OWNER'S ADDRESS" in l.upper():
                # Check same line after label
                rest = _re.split(r"Owner'?s?\s+Address", l, flags=_re.I)[-1].strip()
                # Reject if rest looks like another label (not an address)
                _rest_lower = rest.lower()
                _is_label_rest = any(lk in _rest_lower for lk in _VEHICLE_LABELS)
                if rest and len(rest) > 5 and not _is_label_rest and ',' in rest:
                    addr = _re.sub(r'\s+\d+\s*$', '', rest).strip()
                    if addr: _add('address_present', addr); break
                # Try next 2 lines
                for j in range(i+1, min(i+4, len(lines))):
                    addr_line = lines[j].strip()
                    if not addr_line: continue
                    # Must contain comma (real address)
                    if ',' not in addr_line: continue
                    # Skip if it's another label
                    if any(lk in addr_line.lower() for lk in _VEHICLE_LABELS): continue
                    addr = _re.split(r'\s{3,}', addr_line)[0].strip()
                    addr = _re.sub(r'\s+\d+\s*$', '', addr).strip()
                    if addr and len(addr) > 8:
                        _add('address_present', addr); break
                break
        _add('nationality', 'BANGLADESHI')
        # NID / Smart ID: "Nid Number\nN/A" or "Nid Number\n3282849656" label-based
        # Only accept if NOT N/A AND is a proper NID (10-17 digits, not a phone)
        for i, l in enumerate(lines):
            if 'NID NUMBER' in l.upper() or 'NID Number' in l or 'Nid Number' in l:
                for j in range(i+1, min(i+4, len(lines))):
                    v = lines[j].strip()
                    if not v or v.upper() in ('N/A', 'NA', ''): break  # explicit N/A → skip
                    if _re.match(r'^\d{10,17}$', v):
                        # Reject if looks like a phone number (starts with 8801 or 01)
                        if v.startswith(('8801', '01')):
                            break  # It's a phone number, not NID
                        if len(v) >= 13:
                            _add('nid', v)
                        else:
                            _add('smart_id', v)
                            _add('nid_new', v)
                        break
                break
        # Mobile: "8801711982571" on same line or label-based
        if not _store_has('mobile'):
            for i, l in enumerate(lines):
                if 'MOBILE' in l.upper():
                    for j in range(i+1, min(i+4, len(lines))):
                        v = lines[j].strip()
                        m = _re.search(r'\b(8801[3-9]\d{8}|01[3-9]\d{8})\b', v)
                        if m:
                            mob = m.group(1)
                            if mob.startswith('880'): mob = '0' + mob[3:]
                            _add('mobile', mob); break
                    break
    def _parse_passport(txt, lines):
        # ── Column-aware extraction (handles 3-column passport layout) ────
        # pdftotext -layout output এ Father's Name/Mother's Name/Spouse Name
        # মাঝ ও ডান column-এ থাকে; line-based parser তা ভুল ধরে। নিচের
        # helper character-column position match করে সঠিক value আনে।
        _PASSPORT_STOP_LABELS = {
            "Father's Name", "Mother's Name", "Spouse Name", "Passport Status",
            "Permanent Address", "Present Address", "Profession",
            "Date of Issue", "Date of Expiry", "Passport Number",
            "Previous Passport No", "First Name", "Last Name", "Gender",
            "Date of Birth", "Passport Type", "NID", "Birth ID", "Age",
        }
        def _val_by_col(label, max_search=12, col_tol=10, multi_line=False):
            """Find `label` in lines; return value at same character-column in next lines.
            If multi_line=True, join consecutive non-label lines (for addresses/profession)."""
            for _i, _ln in enumerate(lines):
                _pos = _ln.find(label)
                if _pos < 0:
                    continue
                _col_min = max(0, _pos - col_tol)
                _col_max = _pos + len(label) + col_tol
                _collected = []
                for _j in range(_i + 1, min(_i + max_search, len(lines))):
                    _row = lines[_j]
                    if not _row.strip():
                        if _collected: continue  # skip blank between value lines
                        else: continue
                    # Split row into (col_pos, text) by 2+ space gaps
                    _parts, _pp = [], 0
                    for _m in _re.finditer(r'\s{2,}', _row):
                        if _m.start() > _pp:
                            _parts.append((_pp, _row[_pp:_m.start()].strip()))
                        _pp = _m.end()
                    if _pp < len(_row):
                        _tail = _row[_pp:].strip()
                        if _tail:
                            _parts.append((_pp, _tail))
                    _found_in_row = False
                    for _cp, _ct in _parts:
                        if not _ct:
                            continue
                        if not (_col_min <= _cp <= _col_max):
                            continue
                        if _ct in _PASSPORT_STOP_LABELS:
                            # New label hit — stop collecting
                            if _collected:
                                return ' '.join(_collected)
                            return None
                        _collected.append(_ct)
                        _found_in_row = True
                        if not multi_line:
                            return _ct
                        break  # take first matching part from this row
                    if multi_line and not _found_in_row and _collected:
                        # Row had no matching column → end of value
                        return ' '.join(_collected)
                if _collected:
                    return ' '.join(_collected)
                break  # only first occurrence of label
            return None

        # Run column-aware extraction first; falls back to line-based later
        _fv = _val_by_col("Father's Name")
        if _fv: _add('father', _fv.upper())
        _mv = _val_by_col("Mother's Name")
        if _mv: _add('mother', _mv.upper())
        _sv = _val_by_col("Spouse Name")
        if _sv: _add('spouse', _sv.upper())
        _stv = _val_by_col("Passport Status")
        if _stv: _add('passport_status', _stv)

        # ── Unified Passport Address Parser (PDF + Image OCR) ────────────────
        # NTMC passport-এ address layout (PDF বা image OCR দুই ক্ষেত্রেই):
        #   "Present Address"  বা  "Present Address <garbage>"      ← label line
        #   "<spouse name>"                                          ← skip (right column leak)
        #   "HOUSE-30, ...DHAKA   <other label/garbage>"            ← address (left col)
        #   "Permanent Address   MD ABUL KALAM"                     ← next label + father
        #   "RAFIQPUR, ...NOAKHALI  <garbage>"                       ← address (left col)
        #
        # মূল নিয়ম: address line-এ COMMA থাকে ও যথেষ্ট লম্বা (>=12 char)।
        # Spouse name (FARZANA SHAHID), father name ইত্যাদিতে comma থাকে না → বাদ যায়।
        # Right column-এর label/value এবং OCR garbage strip করা হয়।

        def _addr_looks_valid(s):
            """Real BD address: comma আছে এবং যথেষ্ট লম্বা।
            Spouse/Father name (comma নেই) এতে আটকে যায়।"""
            return (',' in s) and (len(s.strip()) >= 12)

        def _addr_clean(s):
            """Address line থেকে trailing label/OCR-garbage strip করো।"""
            # Right column label OCR-এ glue হয়ে আসে — কেটে দাও
            s = _re.sub(
                r"\s+(etinad|Naa|Father|Mother|Spouse|Passport|Profession|"
                r"Status|Name|Religion|Marital|Citizen|Emergency|DMiayihar|"
                r"'s\s+\w+).*$",
                '', s, flags=_re.I)
            # PDF: 3+ space দিয়ে আলাদা right-column অংশ কেটে দাও
            s = _re.split(r'\s{3,}', s)[0]
            # comma-part ধরে ধরে garbage tail কাটো
            parts = [p.strip() for p in s.split(',')]
            good = []
            for p in parts:
                # garbage part: ছোট (≤2 char, '-' ছাড়া) বা mixed-case junk বা quote/!
                if p and p != '-' and (len(p) <= 2 or _re.search(r"[a-z]{2}[A-Z]|['\"!]", p)):
                    if good:
                        break
                good.append(p)
            return ', '.join(good).strip().rstrip(',').strip()

        def _parse_addr_unified(label_str, field_key):
            """PDF + Image OCR উভয় passport থেকে address parse।
            Multi-line address support: OCR-এ address দুই লাইনে থাকতে পারে।
            """
            _STOP = {'Profession', 'Father', 'Mother', 'Spouse', 'Religion',
                     'Marital', 'Citizen', 'Emergency',
                     'Permanent Address', 'Present Address'}
            for _i, _ln in enumerate(lines):
                if label_str not in _ln:
                    continue
                _parts = []
                # label line-এর পরে যদি same-line-এ address থাকে (left col)
                _rest = _ln.split(label_str, 1)[1].strip()
                _rest = _re.split(r'\s{3,}', _rest)[0].strip() if _rest else ''
                if _addr_looks_valid(_rest):
                    _parts.append(_rest)
                # পরের কয়েকটা line scan করো — multi-line address collect করো
                for _j in range(_i + 1, min(_i + 7, len(lines))):
                    _row = lines[_j]
                    if not _row.strip():
                        continue
                    # PDF: right-column-only line (indent >= 60) skip
                    _indent = len(_row) - len(_row.lstrip())
                    if _indent >= 60:
                        continue
                    _left = _re.split(r'\s{3,}', _row.strip())[0].strip()
                    if not _left:
                        continue
                    # নতুন label এলে থামো
                    if any(_left == s or _left.startswith(s + ' ') or _left.startswith(s) for s in _STOP):
                        break
                    if _addr_looks_valid(_left):
                        _parts.append(_left)
                        # continuation check: পরের line-ও address হতে পারে
                        # (OCR-এ multi-line address — e.g. PABLA COLLEGE CROSS / ROAD 2, ...)
                        # শুধু collect করো, break করো না
                    elif _parts and _left and not _left[0].isdigit():
                        # আগে address পাওয়া গেছে, এই line-এ comma নেই কিন্তু
                        # continuation হতে পারে: e.g. "ROAD 2, DAULATPUR, DAULATPUR, KHULNA"
                        # Check: valid content (uppercase, no label keyword)
                        _is_addr_continuation = (
                            _re.search(r'[A-Z]{3}', _left) and  # uppercase word আছে
                            len(_left) >= 5 and
                            not any(_left.startswith(s) for s in _STOP)
                        )
                        if _is_addr_continuation:
                            _parts.append(_left)
                if _parts:
                    # সব parts জোড়া দাও, trailing comma ঠিক করো
                    _combined = ', '.join(_parts)
                    # যদি first part trailing comma-তে শেষ হয় দ্বিতীয় part-এর সাথে merge
                    # e.g. "NSI OFFICE, -, BANGLA BAZAR, DHAMRAL," + "DHAKA"
                    _combined = _re.sub(r',\s*,', ',', _combined)  # double comma fix
                    _combined = _re.sub(r',\s+([A-Z])', r', \1', _combined)  # spacing normalize
                    _cleaned = _addr_clean(_combined)
                    if len(_cleaned) > 8:
                        _add(field_key, _cleaned)
                return  # label পাওয়া গেছে — duplicate এড়াতে exit

        _parse_addr_unified('Present Address', 'address_present')
        _parse_addr_unified('Permanent Address', 'address_permanent')

        # Profession: LEFT column থেকে নাও (right col-এ mother name leak হতে পারে)
        for _i, _ln in enumerate(lines):
            _ls = _ln.strip()
            if not (_ls == 'Profession' or _ls.startswith('Profession')):
                continue
            _pf_cands = []
            # same line-এ label-এর ঠিক পরে value (left col, indent < 30)
            _after = _ln.split('Profession', 1)[1] if 'Profession' in _ln else ''
            _after_strip = _after.lstrip()
            _after_indent = len(_after) - len(_after_strip)
            if _after_indent < 30 and _after_strip:
                _sv = _re.split(r'\s{3,}', _after_strip)[0].strip()
                if _sv: _pf_cands.append(_sv)
            # পরের non-empty left-column lines — multi-line profession support
            # e.g. "PERMANENT OFFICER/ STAFF OF AUTONOMOUS" + "ORGANIZATION"
            _PF_STOP = {'Father', 'Mother', 'Spouse', 'Address', 'Passport',
                        'Name', 'Religion', 'Marital', 'Present', 'Permanent',
                        'NID', 'Birth', 'Date', 'Gender', 'Spouse', 'Emergency'}
            for _j in range(_i + 1, min(_i + 5, len(lines))):
                _row = lines[_j]
                if not _row.strip(): continue
                _ind = len(_row) - len(_row.lstrip())
                if _ind >= 60: continue  # right col only — skip
                _lf = _re.split(r'\s{3,}', _row.strip())[0].strip()
                if not _lf: continue
                # নতুন label এলে থামো
                if any(_lf == s or _lf.startswith(s) for s in _PF_STOP): break
                _pf_cands.append(_lf)
                # continuation: শুধু যদি line-এ কোনো sentence-ending না থাকে
                # e.g. "AUTONOMOUS" কোনো punctuation ছাড়া → পরের line continuation
                if _lf.endswith(('.', ';', ':')): break
            # name/label নয় এমন candidates একসাথে join করো
            _PF_BAD = ('Father', 'Mother', 'Spouse', 'Address', 'Passport',
                       'Name', 'Religion', 'Marital', 'KHATUN', 'BEGUM')
            _pf_good = []
            for _c in _pf_cands:
                if (_c and _c.upper() not in ('N/A', 'NA', '')
                        and not any(_b in _c for _b in _PF_BAD)):
                    _pf_good.append(_c)
                else:
                    break  # bad candidate এলে থামো
            if _pf_good:
                _add('profession', ' '.join(_pf_good))
            break
        # ── End Unified Address Parser ───────────────────────────────────────

        # NID: dedicated label or 17-digit number
        for i, l in enumerate(lines):
            if l.strip() == 'NID' and i+1 < len(lines):
                v = lines[i+1].strip()
                # Only accept pure numeric NID (17-digit starting with 19)
                # Reject if it looks like an address (contains letters/commas)
                if _re.match(r'^19\d{15}$', v):
                    _add('nid', v); break
        if not _store_has('nid'):
            m = _re.search(r'\b(19\d{15})\b', txt)
            if m: _add('nid', m.group(1))

        # Passport number: OA8012490 format
        # OCR may produce "0A8012490" (zero instead of O) — fix it
        def _fix_passport(v):
            # Clean OCR special chars: ¢→C, 0→O at start if followed by letter
            v = v.replace('¢', 'C').replace('©', 'O')
            v = _re.sub(r'[^A-Za-z0-9]', '', v)  # strip non-alphanumeric
            # Leading digit → 'O' (passport number starts with letter)
            v = _re.sub(r'^0([A-Z]\d)', r'O\1', v.upper())
            return v

        for l in lines:
            # Clean OCR noise then search
            l_clean = l.replace('¢', 'C').replace('©', 'O')
            m = _re.search(r'\b([A-Z0O]{1,2}\d{6,8})\b', l_clean)
            if m:
                pn = _fix_passport(m.group(1))
                if _re.match(r'^[A-Z]{1,2}\d{6,8}$', pn) and len(pn) >= 7:
                    _add('passport', pn); break
        if not _store_has('passport'):
            for l in lines:
                l_clean = l.replace('¢', 'C').replace('©', 'O')
                m = _re.search(r'N/A\s+([A-Z0O]{1,2}\d{6,8})', l_clean)
                if m: _add('passport', _fix_passport(m.group(1))); break

        # Full name: First Name + Last Name combine
        # Case 1: standalone CAPS line "MOHAMMAD AMINUL ISLAM KHAN"
        # ⚠️ But SKIP if the previous line is "Father's Name" / "Mother's Name" / "Spouse Name"
        # (OCR sometimes puts label and value on adjacent lines)
        _name_label_prev = {"father's name", "mother's name", "spouse name",
                            "fathers name", "mothers name",
                            "father name", "mother name"}
        for i, l in enumerate(lines):
            ls = l.strip()
            # Skip if previous non-empty line is a relative-name label
            _prev_label = False
            for _k in range(i-1, max(-1, i-3), -1):
                _pl = lines[_k].strip().lower()
                if _pl:
                    if any(lbl in _pl for lbl in _name_label_prev):
                        _prev_label = True
                    break
            if _prev_label: continue
            if (_re.match(r'^[A-Z][A-Z ]{4,}$', ls) and
                    2 <= len(ls.split()) <= 6 and
                    not any(kw in ls for kw in ('OFFICIAL','ACTIVE','REVOKED',
                        'GOVERNMENT','SERVICE','BANGLADESH','ADDRESS',
                        'PASSPORT','PRESENT','PERMANENT','PROFESSION'))):
                _add('name', ls, 'Passport'); break
        # Case 2: First Name + Last Name separate fields
        if not _store_has('name'):
            fn = ln = ''
            for i, l in enumerate(lines):
                # "First Name" — tolerant OCR variants: Fire Name, Fist Name, Firzt Name
                if (_re.search(r'\bF[iu]r[se][te]\s+Name\b|First\s+Name', l, _re.I)
                        and i+1 < len(lines)):
                    fn = lines[i+1].strip()
                if 'Last Name' in l and i+1 < len(lines):
                    ln = lines[i+1].strip()
            if fn and ln:
                _add('name', (fn + ' ' + ln).upper(), 'Passport')
            elif fn:
                _add('name', fn.upper(), 'Passport')
        # Case 3: OCR line like "MD FARHAD HOSSAIN age" — name + age leak
        if not _store_has('name'):
            for l in lines:
                ls = l.strip()
                m = _re.match(r'^((?:MD|MR|MRS|MOHAMMAD|MOHAMMED)\.?\s+[A-Z]+(?:\s+[A-Z]+){0,4})\s+(?:age|aged|\d{1,3})\s*$', ls, _re.I)
                if m:
                    _add('name', m.group(1).upper().strip(), 'Passport'); break
        # Case 4: "MD MOHIDUL KHAN jioiiie Number Birth" — name + OCR garbage
        # Extract leading name pattern before OCR noise starts
        if not _store_has('name'):
            for l in lines:
                ls = l.strip()
                m = _re.match(r'^((?:MD|MR|MRS|MOHAMMAD|MOHAMMED)\.?\s+[A-Z]+(?:\s+[A-Z]+){1,3})\s+\S', ls)
                if m:
                    name_cand = m.group(1).strip()
                    # Validate: all parts are 2+ caps letters, max 5 words
                    _parts = name_cand.split()
                    if (2 <= len(_parts) <= 5 and
                            all(_re.match(r'^[A-Z.]{2,}$', p) for p in _parts) and
                            len(name_cand) < 50):
                        _add('name', name_cand, 'Passport'); break

        # Gender: "M OFFICIAL" or label-based
        for l in lines:
            m = _re.match(r'^(M|F)\s+OFFICIAL', l.strip())
            if m: _add('gender', 'MALE' if m.group(1)=='M' else 'FEMALE'); break
        if not _store_has('gender'):
            for i, l in enumerate(lines):
                if 'Gender' in l and i+1 < len(lines):
                    v = lines[i+1].strip().upper()
                    if v in ('M','MALE'): _add('gender', 'MALE'); break
                    if v in ('F','FEMALE'): _add('gender', 'FEMALE'); break

        # DOB
        for i, l in enumerate(lines):
            if 'Date of Birth' in l and 'Previous' not in l and i+1 < len(lines):
                nxt = lines[i+1].strip()
                m = _re.match(r'^(\d{2}/\d{2}/\d{4})', nxt)
                if m: _add('dob', m.group(1)); break

        # Spouse
        for i, l in enumerate(lines):
            if 'Spouse Name' in l and i+1 < len(lines):
                sp = lines[i+1].strip()
                if sp and 'Address' not in sp and len(sp) > 3:
                    _add('spouse', sp)
                break

        # Father - passport image OCR may have address merged with father's name
        for i, l in enumerate(lines):
            if "Father" in l and "Name" in l and i+1 < len(lines):
                v = lines[i+1].strip()
                # Skip empty line
                if not v:
                    if i+2 < len(lines): v = lines[i+2].strip()
                    else: break
                # If line looks like address (commas, districts) — extract name part
                if ',' in v and len(v) > 30:
                    # Address + name merged: "RAJBANDH, -, KOYA BAZAR, MD ABDUL JABBAR"
                    # Name is usually at the end after the last comma-group
                    parts = [p.strip() for p in v.split(',')]
                    # Find the part that looks like a name (all caps, no digits, 2-4 words)
                    name_part = ''
                    for p in reversed(parts):
                        p = p.strip()
                        if (p and _re.match(r'^[A-Z ]+$', p) and
                                2 <= len(p.split()) <= 5 and
                                not _re.search(r'\d', p) and
                                p.upper() not in ('N/A','NA','KHAN','BAZAR','DISTRICT')):
                            name_part = p; break
                    if name_part:
                        v = name_part
                    else:
                        break  # Can't extract clean name from this line

                # Check if multi-column data: "MOHAMMAD GOLZAR   ALI KHAN   ..."
                _cols = _re.split(r'\s{2,}', v)
                if len(_cols) >= 2:
                    c1 = _cols[0].strip().upper()
                    c2 = _cols[1].strip().upper()
                    _LABEL_WORDS = {'PASSPORT','STATUS','NUMBER','DATE','TYPE',
                                    'ADDRESS','PROFESSION','GENDER','NAME','N/A'}
                    if (c2 and _re.match(r'^[A-Z][A-Z ]+$', c2) and
                            not _re.search(r'\d', c2) and
                            c2 not in _LABEL_WORDS and
                            len(c2.split()) <= 3):
                        v = c1 + ' ' + c2
                    else:
                        v = c1
                else:
                    v = _re.sub(r'\s+Passport\s+Status.*$', '', v, flags=_re.I).strip()
                    v = _re.sub(r'\s+[A-Z][a-z]+\s+(Status|Name|Number|Type|Address).*$',
                                '', v, flags=_re.I).strip()
                v = v.upper().strip()
                if v and v not in ('N/A','NA','') and _re.search(r'[A-Z]{2,}', v):
                    _add('father', v); break
        if not _store_has('father'):
            m = _re.search(r"Father'?s? Name\s*\n\s*(.+?)(?:\n|$)", txt, _re.I)
            if m:
                v = _re.sub(r'\s+Passport.*$','',m.group(1),flags=_re.I).strip()
                # Reject if address-like
                if ',' not in v or len(v) < 30:
                    _add('father', v.upper())

        # Mother - generic label
        for i, l in enumerate(lines):
            if "Mother" in l and "Name" in l and i+1 < len(lines):
                v = lines[i+1].strip()
                # Strip trailing dates: "Khodeza Begum 01/01/1970 03/24/2014"
                v = _re.sub(r'\s+\d{1,2}/\d{1,2}/\d{4}.*$', '', v).strip()
                v = _re.sub(r'\s+\d{4}-\d{2}-\d{2}.*$', '', v).strip()
                if v and v.upper() not in ('N/A','NA','') and _re.search(r'[A-Za-z]{2,}', v):
                    _add('mother', v); break
        if not _store_has('mother'):
            m = _re.search(r"Mother'?s? Name\s*\n\s*(.+?)(?:\n|$)", txt, _re.I)
            if m:
                mv = _re.sub(r'\s+\d{1,2}/\d{1,2}/\d{4}.*$', '', m.group(1)).strip()
                _add('mother', mv)

        # Permanent / Present address
        # Column layout: OCR may interleave Religion/Marital lines between label and actual address
        # Permanent address fallback: Emergency Address — only if unified parser failed
        if not _store_has('address_permanent'):
            for i, l in enumerate(lines):
                if 'Emergency Address' in l or 'Emergency' in l:
                    _em_parts = []
                    for j in range(i+1, min(i+6, len(lines))):
                        nxt = lines[j].strip()
                        if not nxt: continue
                        if _re.search(r'\b(Father|Mother|Profession|Passport|Marital|Spouse|Name)\b', nxt): break
                        if ',' in nxt or len(nxt) > 10:
                            _em_parts.append(nxt)
                    if _em_parts:
                        _em_full = ' '.join(_em_parts)
                        _em_full = _re.sub(r'\s*,\s*', ', ', _em_full)
                        _em_full = _re.sub(r',\s*$', '', _em_full).strip()
                        _add('address_permanent', _em_full)
                    break

        # Mobile — passport may have "Mobile Number" label (or OCR variant)
        # OR standalone 11-digit number starting with 01
        if not _store_has('mobile'):
            for i, l in enumerate(lines):
                # Label-based (including OCR variants like "jioiiie Number")
                if _re.search(r'\b(Mobile|Mob\.?)\b', l, _re.I) and i+1 < len(lines):
                    for j in range(i+1, min(i+4, len(lines))):
                        mv = lines[j].strip()
                        m = _re.search(r'\b(01[3-9]\d{8})\b', mv)
                        if m: _add('mobile', m.group(1)); break
                    if _store_has('mobile'): break
        if not _store_has('mobile'):
            # Scan all lines for standalone BD mobile number
            for l in lines:
                m = _re.search(r'\b(01[3-9]\d{8})\b', l)
                if m: _add('mobile', m.group(1)); break

        # Profession — multi-line support
        _PROF_SKIP = r'\b(Name|Address|Profession|Passport|Status|Spouse|' \
                     r'Father|Mother|Gender|Permanent|Present|NID|Birth|Age|' \
                     r'Issue|Expiry|Date|Type|Number|Marital|Emergency|Revoked)\b'
        for i, l in enumerate(lines):
            if 'Profession' in l and i+1 < len(lines):
                _prof_lines = []
                for j in range(i+1, min(i+5, len(lines))):
                    p = lines[j].strip()
                    if not p: continue
                    if _re.search(_PROF_SKIP, p): break
                    if 'N/A' in p.upper(): break
                    # Reject if it's a standalone person name (all caps, 1-2 words, no profession keywords)
                    _PROF_WORDS = {'SERVICE','OFFICER','STAFF','TEACHER','DOCTOR',
                                   'ENGINEER','GOVERNMENT','AUTONOMOUS','ORGANIZATION',
                                   'PRIVATE','BUSINESS','RETIRED','PROFESSION'}
                    _pw = p.upper().split()
                    _is_name_only = (
                        1 <= len(_pw) <= 3 and
                        all(_re.match(r'^[A-Z]+$', w) for w in _pw) and
                        not any(w in _PROF_WORDS for w in _pw) and
                        not _re.search(r'\d', p)
                    )
                    if _is_name_only: break
                    _prof_lines.append(p)
                if _prof_lines:
                    _add('profession', ' '.join(_prof_lines))
                break
        _add('nationality', 'BANGLADESHI')

        # Passport number from dedicated "Passport Number" label
        if not _store_has('passport'):
            for i, l in enumerate(lines):
                if 'Passport Number' in l:
                    rest = l.split('Passport Number')[-1].strip()
                    if not rest or rest.upper() in ('N/A','NA'):
                        if i+1 < len(lines): rest = lines[i+1].strip()
                    m = _re.search(r'([A-Z]{1,2}\d{6,8})', rest or '')
                    if m: _add('passport', m.group(1)); break

        # Issue Date
        # TIN PDF-এ "Passport Issue Date" label থাকে কিন্তু value = "01/01/1970" (dummy/N/A)
        # ঐ dummy date reject করতে হবে। Rule: 1970 সালের যেকোনো date + "01/01/YYYY" pattern = dummy
        # Extra guard: TIN doc হলে "Passport Issue Date" label-এর পরে value নেওয়া হবে না
        _DUMMY_ISSUE_DATES = {'01/01/1970', '1/1/1970', '01-01-1970'}
        _is_tin_doc = ('ASSESSEE NAME' in txt.upper() or 'OLD TIN' in txt.upper())
        for i, l in enumerate(lines):
            if 'Date of Issue' in l:
                # Skip if this is a TIN doc (the label here is "Passport Issue Date" col header)
                if _is_tin_doc: break
                m = _re.search(r'(\d{1,2}/\d{2}/\d{4}|\d{2}/\d{2}/\d{4})', l)
                if not m and i+1 < len(lines):
                    m = _re.search(r'(\d{1,2}/\d{2}/\d{4})', lines[i+1])
                if m:
                    _dv = m.group(1)
                    # Reject dummy dates (01/01/1970 or any year before 1980)
                    if _dv in _DUMMY_ISSUE_DATES: break
                    _yr_m = _re.search(r'(\d{4})$', _dv)
                    if _yr_m and int(_yr_m.group(1)) < 1980: break
                    _add('passport_issue', _dv)
                break

        # Expiry Date
        for i, l in enumerate(lines):
            if 'Date of Expiry' in l:
                m = _re.search(r'(\d{1,2}/\d{2}/\d{4}|\d{2}/\d{2}/\d{4})', l)
                if not m and i+1 < len(lines):
                    m = _re.search(r'(\d{1,2}/\d{2}/\d{4})', lines[i+1])
                if m: _add('passport_expiry', m.group(1)); break

        # Previous Passport
        for l in lines:
            if 'Previous Passport' in l or 'Prev' in l and 'Passport' in l:
                m = _re.search(r'([A-Z]{1,2}\d{6,8})', l)
                if m and m.group(1) != profile.get('passport',''):
                    _add('prev_passport', m.group(1)); break

        # Passport Status
        def _clean_passport_status(raw, idx, lines):
            """Strip leading name garbage and join next line if Reissue is cut off."""
            ls = _re.sub(
                r"^['\"]?(?:[A-Z]+\s+){1,5}(?=Document|REVOKED|ACTIVE|revoked|active)",
                '', raw.strip(), flags=_re.I).strip()
            ls = ls.lstrip("'\"").strip()
            # If "Reissue" বা completion word missing, scan ahead 1-4 lines
            if ls and 'reissue' not in ls.lower():
                for _k in range(idx+1, min(idx+5, len(lines))):
                    _nxt = lines[_k].strip()
                    if not _nxt: continue
                    if _re.search(r'\breissue\b', _nxt, _re.I):
                        # Join only the Reissue-containing chunk
                        _m = _re.search(r'(\bReissue\w*)', _nxt, _re.I)
                        if _m:
                            ls = ls.rstrip() + ' ' + _m.group(1)
                        break
            # Trim dangling preposition ("due to" with nothing after)
            ls = _re.sub(r'\s+(due\s+to|to)\s*$', '', ls, flags=_re.I).strip()
            return ls

        for i, l in enumerate(lines):
            if 'Passport Status' in l:
                rest = l.split('Passport Status')[-1].strip()
                if rest and rest.upper() not in ('N/A','NA',''):
                    # Clean garbage names that leaked into status
                    rest = _clean_passport_status(rest, i, lines)
                    if rest: _add('passport_status', rest)
                else:
                    # Image OCR: label on its own line; value on next line(s)
                    # "Document revoked due to" + "Reissue" (split across 2 lines)
                    _status_parts = []
                    for j in range(i+1, min(i+5, len(lines))):
                        nxt = lines[j].strip()
                        if not nxt: continue
                        # Stop যদি new label আসে
                        if _re.match(r'^(Father|Mother|Spouse|Permanent|Present|'
                                     r'Profession|NID|Date)\b', nxt):
                            break
                        _status_parts.append(nxt)
                        if 'reissue' in nxt.lower() or 'active' in nxt.lower():
                            break
                    if _status_parts:
                        _full = ' '.join(_status_parts).strip()
                        _full = _re.sub(r'\s+(due\s+to|to)\s*$', '', _full, flags=_re.I).strip()
                        if _full: _add('passport_status', _full)
                break
        if not _store_has('passport_status'):
            for l in lines:
                ls = l.strip().lstrip("'\"")
                if ls.upper() in ('ACTIVE', 'REVOKED'): _add('passport_status', ls); break
                if 'REVOKED DUE' in ls.upper():         _add('passport_status', ls); break
            for i, l in enumerate(lines):
                if 'revoked' in l.lower() or 'reissue' in l.lower():
                    ls = _clean_passport_status(l, i, lines)
                    if ls: _add('passport_status', ls); break

    # ── PARSER: TIN ──────────────────────────────────────────────────────
    def _parse_tin(txt, lines):
        """
        TIN PDF: 3-column pdftotext -layout format.
        Each data line: "Value_col1          Value_col2          Value_col3"
        We always take col1 (first column) by splitting on 2+ spaces.
        """
        # Detect format: PDF (3-column) vs Image OCR (single-column)
        # PDF: "Old TIN   NID Number   Incorporation Number" (has 2+ spaces between)
        # IMG: "Old TIN" alone, then next line "N/A" or value
        _is_multicol = any(
            _re.search(r'\s{3,}', l) and (
                'TIN' in l or 'Name' in l or 'Number' in l
            )
            for l in lines[:10]
        )

        def _col1(line):
            """First column of multi-column line (split on 2+ spaces)."""
            if _is_multicol:
                parts = _re.split(r'\s{2,}', line.strip())
                v = parts[0].strip() if parts else ''
            else:
                v = line.strip()  # single-column: take whole line
            return v if v.upper() not in ('N/A','NA','NIA','') else ''

        def _label_val(label):
            for i, l in enumerate(lines):
                if label.upper() in l.upper():
                    if i+1 < len(lines):
                        return _col1(lines[i+1])
            return ''

        # ── Old TIN ────────────────────────────────────────────────────
        for i, l in enumerate(lines):
            if 'Old TIN' in l:
                if i+1 < len(lines):
                    v = _col1(lines[i+1])
                    if v and _re.match(r'^\d{7,12}$', v): _add('tin', v)
                break

        # ── New TIN ─────────────────────────────────────────────────────
        # PDF: 3rd column of Father Name row
        # IMG: "New TIN" label, next line = value
        for i, l in enumerate(lines):
            if 'New TIN' in l:
                if i+1 < len(lines):
                    if _is_multicol:
                        parts = _re.split(r'\s{2,}', lines[i+1].strip())
                        v = parts[-1].strip() if len(parts) >= 3 else parts[0].strip() if parts else ''
                    else:
                        v = lines[i+1].strip()
                    if v and v.upper() not in ('N/A','NA','NIA') and _re.match(r'^\d{10,12}$', v):
                        _add('tin_new', v)
                break

        # ── NID / Smart ID ──────────────────────────────────────────────
        if not _store_has('nid'):
            for l in lines:
                m = _re.search(r'\b(19\d{15})\b', l)
                if m: _add('nid', m.group(1)); break
        # Smart ID: 10-digit in "Smart ID" row last column
        for i, l in enumerate(lines):
            if 'Smart ID' in l:
                if i+1 < len(lines):
                    parts = _re.split(r'\s{2,}', lines[i+1].strip())
                    v = parts[-1].strip() if parts else ''
                    if v and _re.match(r'^\d{10}$', v): _add('smart_id', v)
                break

        # ── Assessee Name ───────────────────────────────────────────────
        # Row: "Assessee Name    Contact Telephone    Incorporation Date"
        #      "Md. Jahirul Islam 880               N/A"
        # col1 = name, col2 = phone (880/number), col3 = N/A
        for i, l in enumerate(lines):
            if 'Assessee Name' in l:
                if i+1 < len(lines):
                    v = _col1(lines[i+1])
                    if (v and not _re.match(r'^\d', v)
                            and _re.search(r'[A-Za-z]{2,}', v)
                            and 2 <= len(v.split()) <= 5):
                        _add('name', v.upper(), 'TIN')
                break

        # ── Father Name ─────────────────────────────────────────────────
        # Row: "Father Name    Passport Number    New TIN"
        #      "Md. Abul kalam N/A               149801304174"
        # col1 = father name
        for i, l in enumerate(lines):
            if 'Father Name' in l:
                if i+1 < len(lines):
                    v = _col1(lines[i+1])
                    if v and _re.search(r'[A-Za-z]{2,}', v) and len(v.split()) >= 1:
                        _add('father', v.upper())
                break

        # ── Mother Name ─────────────────────────────────────────────────
        # Row: "Mothers Name    Passport Issue Date    Tin Date"
        #      "Mrs. Khodeza Begum 01/01/1970          03/24/2014"
        # col1 = mother name, col2 = date, col3 = date
        for i, l in enumerate(lines):
            if 'Mothers Name' in l:
                if i+1 < len(lines):
                    v = _col1(lines[i+1])
                    # Strip "Mrs." prefix if present
                    v = _re.sub(r'^Mrs?\.?\s+', '', v).strip()
                    if v and _re.search(r'[A-Za-z]{2,}', v) and len(v.split()) >= 1:
                        _add('mother', v.upper())
                break

        # ── Spouse Name ─────────────────────────────────────────────────
        # Row: "Spouse Name    BOI Date    Contact Telephone Sub"
        #      "N/A            N/A         880"
        # If col1 is N/A → no spouse
        for i, l in enumerate(lines):
            if 'Spouse Name' in l:
                if i+1 < len(lines):
                    v = _col1(lines[i+1])
                    if v and v.upper() not in ('N/A','NA',''):
                        _add('spouse', v)
                break

        # ── Date of Birth ──────────────────────────────────────────────
        for i, l in enumerate(lines):
            if 'DATE OF BIRTH' in l.upper() or 'Date of Birth' in l:
                for j in range(i+1, min(i+4, len(lines))):
                    v = lines[j].strip()
                    m = _re.search(r'(\d{2}/\d{2}/\d{4}|\d{4}-\d{2}-\d{2})', v)
                    if m: _add('dob', m.group(1)); break
                break
        if not _store_has('dob'):
            for l in lines:
                m = _re.match(r'^(\d{2}/\d{2}/\d{4})\s+N/A', l)
                if m: _add('dob', m.group(1)); break

        # ── Gender ─────────────────────────────────────────────────────
        for i, l in enumerate(lines):
            if 'GENDER' in l.upper():
                for j in range(i+1, min(i+4, len(lines))):
                    v = lines[j].strip().upper()
                    if v in ('MALE','FEMALE','M','F'):
                        _add('gender', 'MALE' if v in ('M','MALE') else 'FEMALE'); break
                break
        if not _store_has('gender'):
            for l in lines:
                m = _re.match(r'^(Male|Female|MALE|FEMALE|FeMale)\s+N/A', l)
                if m: _add('gender', m.group(1).upper()); break

        # ── Mobile ─────────────────────────────────────────────────────
        for l in lines:
            m = _re.search(r'\b(8801[3-9]\d{8}|01[3-9]\d{8})\b', l)
            if m:
                mob = m.group(1)
                if mob.startswith('880'): mob = '0' + mob[3:]
                _add('mobile', mob); break

    # ── NID Address Image Crop ──────────────────────────────────────────────
    def _crop_nid_address_images(raw_bytes, is_pdf=True):
        """
        NID PDF/Image থেকে Permanent ও Present Address section crop করো।
        Coordinate-based: word positions দিয়ে সঠিক boundary বের করো।
        Returns: (perm_b64, pres_b64)
        """
        try:
            # PIL → _PILImage (module-level import)
            import io as _io2

            # ── Get page image ──
            if is_pdf:
                try:
                    # pdfplumber → _pdfplumber_mod (module-level import)
                    with _pdfplumber_mod.open(_io.BytesIO(raw_bytes)) as _pdf:
                        pg = _pdf.pages[0]
                        img = pg.to_image(resolution=180).original
                        page_width  = pg.width
                        page_height = pg.height
                        scale = img.size[0] / page_width  # pixel per PDF unit
                        words = pg.extract_words(x_tolerance=3, y_tolerance=3)
                except Exception:
                    return None, None
            else:
                img = _PILImage.open(_io.BytesIO(raw_bytes)).convert('RGB')
                # For image, use pytesseract to find positions
                try:
                    # pytesseract → _pytesseract_mod (module-level import)
                    data = _pytesseract_mod.image_to_data(img, lang='eng',
                                               output_type=_pytesseract_mod.Output.DICT)
                    scale = 1.0
                    words = [{'text': data['text'][i],
                               'top':  data['top'][i],
                               'x0':   data['left'][i]}
                             for i in range(len(data['text']))
                             if (data['text'][i] or '').strip()]
                    page_height = img.size[1]
                    page_width  = img.size[0]
                except Exception:
                    return None, None

            W, H = img.size

            # ── Find key y-positions from words ──
            perm_y = pres_y = old_nid_y = None
            # NID PDF: right column (x > page_width * 0.25) has address sections
            # Left column (x < page_width * 0.25) has personal info

            for w in words:
                txt = (w.get('text') or '').strip()
                top = w.get('top', 0)
                x0  = w.get('x0', 0)

                # "Permanent" heading — must be in address column (right side)
                if txt == 'Permanent' and x0 > page_width * 0.2 and perm_y is None:
                    perm_y = top

                # "Present" heading — must be in address column
                elif txt == 'Present' and x0 > page_width * 0.2 and pres_y is None:
                    pres_y = top

                # "Old" NID label — marks end of present address section
                elif txt == 'Old' and top > (pres_y or 250) + 30 and old_nid_y is None:
                    old_nid_y = top

            if perm_y is None and pres_y is None:
                return None, None

            def _crop_to_b64(y0_pdf, y1_pdf):
                """Crop image using PDF coordinates, convert to b64."""
                y0_px = max(0, int(y0_pdf * scale) - 2)
                y1_px = min(H, int(y1_pdf * scale) + 4)
                if y1_px <= y0_px + 10:
                    return None
                # Only right column: x from ~25% width to edge
                # to avoid left-column personal info bleeding in
                x0_px = int(page_width * 0.22 * scale)
                cropped = img.crop((x0_px, y0_px, W, y1_px))
                buf = _io2.BytesIO()
                cropped.save(buf, format='PNG', optimize=True)
                buf.seek(0)
                return "data:image/png;base64," + _b64.b64encode(buf.read()).decode()

            perm_b64 = None
            pres_b64 = None

            # ── Permanent Address ──
            if perm_y is not None:
                y0 = perm_y - 2
                # End just before Present Address
                y1 = (pres_y - 4) if pres_y else (perm_y + 140)
                perm_b64 = _crop_to_b64(y0, y1)

            # ── Present Address ──
            if pres_y is not None:
                y0 = pres_y - 2
                # Find NID card thumbnails — they appear BEFORE Old NID label
                # Look for image-like gaps: search for last word in address section
                # NID thumbnails start around old_nid_y - 65 (empirically)
                if old_nid_y is not None:
                    # Find last word in right column before thumbnails
                    _last_addr_y = pres_y
                    for _w2 in words:
                        _wy = _w2.get('top', 0)
                        _wx = _w2.get('x0', 0)
                        # Right column word, after pres_y, before old_nid_y
                        if (_wx > page_width * 0.2 and
                                _wy > pres_y + 5 and
                                _wy < old_nid_y - 30):
                            if _wy > _last_addr_y:
                                _last_addr_y = _wy
                    # Add small padding below last word
                    y1 = _last_addr_y + 18
                else:
                    y1 = pres_y + 120
                pres_b64 = _crop_to_b64(y0, y1)

            return perm_b64, pres_b64

        except Exception:
            return None, None


    # ── Image OCR parser (for BRTA DL / Passport / NID images) ──────────────
    def _parse_image_ocr(raw_bytes, fname):
        """OCR করে image থেকে text বের করে parse করো।"""
        _ocr_txt = ''
        try:
            # PIL → _PILImage (module-level import)
            # pytesseract → _pytesseract_mod (module-level import)
            _img = _PILImage.open(_io.BytesIO(raw_bytes))
            if _img.mode in ('P', 'RGBA', 'LA'):
                _img = _img.convert('RGB')
            _ocr_txt = _pytesseract_mod.image_to_string(_img, lang='eng')
        except Exception:
            logger.debug('suppressed exception', exc_info=True)
        return _ocr_txt

    # ── Process each file ─────────────────────────────────────────────────
    _nid_addr_imgs = {}  # NID address crop images (built before result dict)
    for f in (doc_files or []):
        fname = f.name.lower()
        raw   = f.read(); f.seek(0)

        # ── Image files — OCR করে parse করো ──
        if any(fname.endswith(x) for x in ['.jpg','.jpeg','.png']):

            # OCR করে document type detect করো (photo assign-এর আগে)
            _ocr_txt = _parse_image_ocr(raw, fname)
            _is_doc_img = False

            if _ocr_txt.strip() and len(_ocr_txt.strip()) > 30:
                _tu = _ocr_txt.upper()
                _doc_type_img = _detect_doc_type(_tu)
                _lines_img = [l.strip() for l in _ocr_txt.split('\n') if l.strip()]

                if _doc_type_img == 'driving_license' or ('LICENSE' in _tu and 'BRTA' in _tu):
                    _is_doc_img = True
                    _current_doc_label[0] = 'Driving License'
                    _parse_driving_license(_ocr_txt, _lines_img)
                    if 'Driving License' not in docs_found:
                        docs_found.append('Driving License')
                    # DL image থেকে photo crop (left ~28%)
                    if not photo_b64:
                        try:
                            # PIL → _PILImage (module-level import)
                            _pi = _PILImage.open(_io.BytesIO(raw))
                            _pw, _ph2 = _pi.size
                            _pc = _pi.crop((0, 0, int(_pw*0.28), int(_ph2*0.65)))
                            _pb = _io.BytesIO(); _pc.save(_pb, format='JPEG', quality=88)
                            photo_b64 = "data:image/jpeg;base64," + _b64.b64encode(_pb.getvalue()).decode()
                        except Exception: pass

                elif _doc_type_img == 'vehicle':
                    _is_doc_img = True
                    _current_doc_label[0] = 'Vehicle Reg'
                    _parse_vehicle(_ocr_txt, _lines_img)
                    if 'Vehicle Registration' not in docs_found:
                        docs_found.append('Vehicle Registration')

                elif _doc_type_img == 'passport':
                    _is_doc_img = True
                    _current_doc_label[0] = 'Passport'
                    _parse_passport(_ocr_txt, _lines_img)
                    if 'Passport' not in docs_found:
                        docs_found.append('Passport')
                    # Passport image থেকে photo crop (left panel ~30%)
                    if not photo_b64:
                        try:
                            # PIL → _PILImage (module-level import)
                            _pi = _PILImage.open(_io.BytesIO(raw))
                            _pw, _ph2 = _pi.size
                            # Passport photo is top-left
                            _pc = _pi.crop((0, 0, int(_pw*0.30), int(_ph2*0.55)))
                            _pb = _io.BytesIO(); _pc.save(_pb, format='JPEG', quality=88)
                            photo_b64 = "data:image/jpeg;base64," + _b64.b64encode(_pb.getvalue()).decode()
                        except Exception: pass

                elif _doc_type_img == 'nid':
                    _is_doc_img = True
                    _current_doc_label[0] = 'NID'
                    _parse_nid(_ocr_txt, _lines_img)
                    if 'NID' not in docs_found:
                        docs_found.append('NID')
                    # NID Image থেকে address section crop করো
                    # Full NID image = the uploaded image itself
                    if not _nid_page_img:
                        try:
                            _nid_page_img = ("data:image/jpeg;base64,"
                                           + _b64.b64encode(raw).decode())
                        except Exception:
                            logger.debug('suppressed exception', exc_info=True)
                    _perm_img, _pres_img = _crop_nid_address_images(raw, is_pdf=False)
                    if _perm_img: _nid_addr_imgs['address_permanent_img'] = _perm_img
                    if _pres_img: _nid_addr_imgs['address_present_img']   = _pres_img
                    # NID photo: left panel — face only, exclude name/Bengali text
                    if not photo_b64:
                        try:
                            # PIL → _PILImage (module-level import)
                            _pi = _PILImage.open(_io.BytesIO(raw))
                            _pw, _ph2 = _pi.size
                            # Face box: left panel ~22% width, top ~38% height
                            # Stop before name/Bengali text (around 40% height)
                            # Small left/top margin to avoid edge artifacts
                            _face = _pi.crop((
                                int(_pw * 0.02),
                                int(_ph2 * 0.02),
                                int(_pw * 0.22),
                                int(_ph2 * 0.38),   # 38% — name text starts ~40%
                            ))
                            _pb = _io.BytesIO()
                            _face.save(_pb, format='JPEG', quality=88)
                            photo_b64 = "data:image/jpeg;base64," + _b64.b64encode(_pb.getvalue()).decode()
                        except Exception: pass

                elif _doc_type_img == 'tin':
                    _is_doc_img = True
                    _current_doc_label[0] = 'TIN'
                    _parse_tin(_ocr_txt, _lines_img)
                    if 'TIN' not in docs_found:
                        docs_found.append('TIN')

            # Document image নয় → subject photo হিসেবে রাখো
            if not _is_doc_img and not photo_b64:
                mime = 'image/jpeg' if fname.endswith(('.jpg','.jpeg')) else 'image/png'
                photo_b64 = "data:" + mime + ";base64," + _b64.b64encode(raw).decode()
            continue

        if not fname.endswith('.pdf'):
            continue

        txt, words, _ext_photo = _get_text_and_words(raw)
        if not photo_b64 and _ext_photo:
            photo_b64 = _ext_photo
        lines = [l for l in txt.split('\n')]
        tu = txt.upper()

        doc_type = _detect_doc_type(tu)

        if not txt.strip() or len(txt.strip()) < 30:
            # Image-based PDF → try OCR via pdfplumber page render
            _ocr_txt = ''
            try:
                # pdfplumber/pytesseract → module-level imports (_pdfplumber_mod/_pytesseract_mod)
                with _pdfplumber_mod.open(_io.BytesIO(raw)) as _pdf2:
                    for _pg in _pdf2.pages:
                        _img = _pg.to_image(resolution=200).original
                        _ocr_txt += _pytesseract_mod.image_to_string(_img, lang='eng') + '\n'
            except Exception:
                logger.debug('suppressed exception', exc_info=True)
            if _ocr_txt.strip():
                txt   = _ocr_txt
                lines = txt.split('\n')
                doc_type = _detect_doc_type(txt.upper())
                if doc_type == 'driving_license' or 'LICENSE' in txt.upper():
                    _current_doc_label[0] = 'Driving License'
                    _parse_driving_license(txt, lines)
                    if 'Driving License' not in docs_found:
                        docs_found.append('Driving License')
                elif doc_type == 'tin':
                    _current_doc_label[0] = 'TIN'
                    _parse_tin(txt, lines)
                    if 'TIN' not in docs_found: docs_found.append('TIN')
                elif doc_type == 'nid':
                    _current_doc_label[0] = 'NID'
                    _parse_nid(txt, lines)
                    if 'NID' not in docs_found: docs_found.append('NID')
                elif doc_type == 'passport':
                    _current_doc_label[0] = 'Passport'
                    _parse_passport(txt, lines)
                    if 'Passport' not in docs_found: docs_found.append('Passport')
                elif doc_type == 'vehicle':
                    _current_doc_label[0] = 'Vehicle Reg'
                    _parse_vehicle(txt, lines)
                    if 'Vehicle Registration' not in docs_found: docs_found.append('Vehicle Registration')
                elif doc_type == 'unknown':
                    # Unknown: run all parsers but in safe order
                    # Each parser only adds fields if not already present (store dedup handles it)
                    _current_doc_label[0] = 'Document'
                    for _fn in [_parse_tin, _parse_nid, _parse_passport,
                                 _parse_driving_license, _parse_vehicle]:
                        _fn(txt, lines)
                    docs_found.append('Document (Scanned)')
            else:
                docs_found.append('Driving License (Scanned — OCR failed)')
            continue

        if doc_type == 'nid':
            _current_doc_label[0] = 'NID'
            _parse_nid(txt, lines)
            if 'NID' not in docs_found: docs_found.append('NID')
            # NID PDF → full page image + face photo crop
            # Strategy: PyMuPDF দিয়ে embedded image সরাসরি extract করা
            # (pdfplumber page render-এ crop coordinate ভুল হওয়ার সম্ভাবনা থাকে)
            try:
                import fitz as _fitz_nid
                _nid_doc = _fitz_nid.open(stream=_io.BytesIO(raw), filetype="pdf")
                _nid_pg  = _nid_doc[0]

                # ── Full NID page render (nid_page_img) ──────────────────────
                _nid_pix = _nid_pg.get_pixmap(dpi=120)
                _nid_buf = _io.BytesIO(_nid_pix.tobytes('jpeg'))
                _nid_page_img = ("data:image/jpeg;base64,"
                                + _b64.b64encode(_nid_buf.getvalue()).decode())

                # ── Face photo: best portrait image from embedded images ──────
                # NID PDF-এ person-এর photo embedded থাকে portrait orientation-এ
                # সবচেয়ে বড় portrait image-টাই face
                if not photo_b64:
                    _best_face = None
                    _best_area = 0
                    for _xref_info in _nid_pg.get_images(full=True):
                        _xref  = _xref_info[0]
                        _bimg  = _nid_doc.extract_image(_xref)
                        _bpil  = _PILImage.open(_io.BytesIO(_bimg['image']))
                        _bw, _bh = _bpil.size
                        _ratio = _bh / max(_bw, 1)
                        # Portrait + reasonable size + larger than logos
                        if _ratio >= 1.1 and _bh >= 100 and _bw >= 70:
                            _area = _bw * _bh
                            if _area > _best_area:
                                _best_area = _area
                                _best_face = _bpil
                    if _best_face:
                        _fb = _io.BytesIO()
                        if _best_face.mode in ('P','RGBA','LA'):
                            _best_face = _best_face.convert('RGB')
                        _best_face.thumbnail((200, 260), _PILImage.LANCZOS)
                        _best_face.save(_fb, format='JPEG', quality=88)
                        photo_b64 = ("data:image/jpeg;base64,"
                                     + _b64.b64encode(_fb.getvalue()).decode())
                _nid_doc.close()
            except Exception:
                # PyMuPDF না থাকলে pdfplumber fallback
                try:
                    with _pdfplumber_mod.open(_io.BytesIO(raw)) as _pdf2:
                        _nid_pg2  = _pdf2.pages[0]
                        _nid_pil2 = _nid_pg2.to_image(resolution=150).original
                        _nid_buf2 = _io.BytesIO()
                        _nid_pil2.save(_nid_buf2, format='JPEG', quality=80)
                        _nid_page_img = ("data:image/jpeg;base64,"
                                        + _b64.b64encode(_nid_buf2.getvalue()).decode())
                        if not photo_b64:
                            _npw, _nph = _nid_pil2.size
                            _face2 = _nid_pil2.crop((
                                int(_npw * 0.005), int(_nph * 0.008),
                                int(_npw * 0.22),  int(_nph * 0.40),
                            ))
                            _fb2 = _io.BytesIO()
                            _face2.save(_fb2, format='JPEG', quality=88)
                            photo_b64 = ("data:image/jpeg;base64,"
                                         + _b64.b64encode(_fb2.getvalue()).decode())
                except Exception:
                    logger.debug('suppressed exception', exc_info=True)
            _perm_img, _pres_img = _crop_nid_address_images(raw, is_pdf=True)
            if _perm_img: _nid_addr_imgs['address_permanent_img'] = _perm_img
            if _pres_img: _nid_addr_imgs['address_present_img']   = _pres_img
        elif doc_type == 'driving_license':
            _current_doc_label[0] = 'Driving License'
            _parse_driving_license(txt, lines)
            if 'Driving License' not in docs_found: docs_found.append('Driving License')
        elif doc_type == 'vehicle':
            _current_doc_label[0] = 'Vehicle Reg'
            _parse_vehicle(txt, lines)
            if 'Vehicle Registration' not in docs_found: docs_found.append('Vehicle Registration')
        elif doc_type == 'passport':
            _current_doc_label[0] = 'Passport'
            _parse_passport(txt, lines)
            if 'Passport' not in docs_found: docs_found.append('Passport')
        elif doc_type == 'tin':
            _current_doc_label[0] = 'TIN'
            _parse_tin(txt, lines)
            if 'TIN' not in docs_found: docs_found.append('TIN')
        else:
            # Unknown: try all
            _current_doc_label[0] = 'Document'
            for fn in [_parse_nid, _parse_driving_license, _parse_vehicle,
                       _parse_passport, _parse_tin]:
                fn(txt, lines)

    # ── Build result ──────────────────────────────────────────────────────
    # store = {field: [(val, doc_label), ...]}
    # Normalize DOB
    # NTMC document format: MM/DD/YYYY (passport, TIN) or YYYY-MM-DD (NID)
    if _store_has('dob'):
        def _nd(v, lbl):
            import datetime as _dt
            v = v.strip()
            # If already ISO format
            if _re.match(r'^\d{4}-\d{2}-\d{2}$', v):
                return (v, lbl)
            # DD/MM/YYYY vs MM/DD/YYYY — use context:
            # If first component > 12 → must be day (so DD/MM/YYYY)
            # If second component > 12 → must be month? no that's impossible
            # NTMC uses MM/DD/YYYY (based on "05/25/1970" for May 25)
            m = _re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})$', v)
            if m:
                mm, dd, yy = int(m.group(1)), int(m.group(2)), int(m.group(3))
                # If first part > 12 → it's a day, format is DD/MM/YYYY
                if mm > 12:
                    try: return (_dt.date(yy, dd, mm).strftime('%Y-%m-%d'), lbl)
                    except Exception: pass
                # Otherwise assume MM/DD/YYYY (NTMC standard)
                try: return (_dt.date(yy, mm, dd).strftime('%Y-%m-%d'), lbl)
                except Exception: pass
                # Fallback: try DD/MM/YYYY
                try: return (_dt.date(yy, dd, mm).strftime('%Y-%m-%d'), lbl)
                except Exception: pass
            # Other formats
            for _fmt in ('%Y/%m/%d','%d-%m-%Y','%Y.%m.%d'):
                try: return (_dt.datetime.strptime(v, _fmt).strftime('%Y-%m-%d'), lbl)
                except Exception: pass
            return (v, lbl)
        store['dob'] = list(dict.fromkeys(_nd(v, lbl) for v, lbl in store['dob'] if v))
    # Normalize name
    if _store_has('name'):
        store['name'] = [(
            _re.sub(r'^MD\.?\s+', 'MD. ', v.strip()), lbl
        ) for v, lbl in store['name'] if v]
    # Normalize mother
    if _store_has('mother'):
        store['mother'] = [(
            _re.sub(r'(?i)^(?:Profession|Mrs\.?)\s+', '', v).strip(), lbl
        ) for v, lbl in store['mother'] if v]

    # ── Post-process passport_status: drop dangling "due to" ──────────
    # OCR-এ "Document revoked due to Reissue" multi-line হলে কিছু source-এ
    # শুধু "Document revoked due to" আসে। Reissue লাগানো বা trim করি।
    if _store_has('passport_status'):
        _ps_new = []
        for _v, _lbl in store['passport_status']:
            _vc = _v.strip()
            # If ends with "due to" / "to" — try to find Reissue from other passport sources
            if _re.search(r'(due\s+to|to)\s*$', _vc, _re.I):
                _has_reissue = any(_re.search(r'reissue', other_v, _re.I)
                                    for other_v, _ in store['passport_status']
                                    if other_v != _v)
                if _has_reissue:
                    _vc = _vc.rstrip() + ' Reissue'
                else:
                    _vc = _re.sub(r'\s+(due\s+to|to)\s*$', '', _vc, flags=_re.I).strip()
            _ps_new.append((_vc, _lbl))
        store['passport_status'] = _ps_new

    # ── Post-process name fields: "Late X" / "X" dedup ─────────────────
    # NBR TIN-এ ব্যক্তির বাবা-মা মৃত থাকলে "Late" prefix থাকে। অন্য doc-এ
    # সেটা নেই। দুজনকে আলাদা mismatch হিসেবে না দেখিয়ে, "Late" version-কে
    # parent value হিসেবেই রাখি কিন্তু দুটো display-এ আসবে (a/b)। তবে fuzzy
    # match-এ "Late" + base form এক হিসেবে গণ্য হবে — এটা mismatch detection-এ
    # হ্যান্ডল করা যাবে। এখানে আমরা শুধু dedup করি যদি base form same হয় ও
    # source same হয় (clean repeated extraction)।
    for _nf in ('father','mother','spouse'):
        if not _store_has(_nf): continue
        _seen_base = {}
        _new_list = []
        for _v, _lbl in store[_nf]:
            _base = _re.sub(r'(?i)^late\s+', '', _v).strip().upper()
            _key = (_base, _lbl)
            if _key in _seen_base: continue
            _seen_base[_key] = True
            _new_list.append((_v, _lbl))
        store[_nf] = _new_list

    # result: primary value (first found) for each field
    result = {k: (store[k][0][0] if store[k] else '') for k in FIELDS}
    # Also store with source for table display: {field}_src = [(val, label), ...]
    for _k in FIELDS:
        result[f'{_k}_src'] = store.get(_k, [])
    # Legacy _all compatibility
    for _k in FIELDS:
        result[f'{_k}_all'] = [v for v, _ in store.get(_k, [])]

    # ── NID name priority block removed ──────────────────────────────────
    # NID parser no longer extracts names — only NID number & blood group.
    # Name comes from Passport / DL / TIN only.

    if not result['tin'] and result.get('tin_new'):
        result['tin'] = result['tin_new']
        if not result.get('tin_src'):
            result['tin_src'] = result.get('tin_new_src', [])
    if not result['nid'] and result.get('nid_new'):
        result['nid'] = result['nid_new']
        if not result.get('nid_src'):
            result['nid_src'] = result.get('nid_new_src', [])
    # smart_id → nid_new fallback (sync _src as well)
    if not result['nid_new'] and result.get('smart_id'):
        result['nid_new'] = result['smart_id']
        if not result.get('nid_new_src'):
            result['nid_new_src'] = result.get('smart_id_src', [])

    # ── Address primary selection — সবচেয়ে informative value বেছে নাও ─────
    # ক্রাইটেরিয়া: সম্পূর্ণ address (comma বেশি, character বেশি)
    def _best_addr(field_key):
        candidates = [(v, lbl) for v, lbl in store.get(field_key, [])
                      if v and len(v) >= 8 and 'N/A' not in v.upper()]
        if not candidates: return ''
        # Score: comma count × 3 + len (longer = more complete)
        def _score(v): return v.count(',') * 3 + len(v)
        return max(candidates, key=lambda x: _score(x[0]))[0]

    _best_perm = _best_addr('address_permanent')
    _best_pres = _best_addr('address_present')
    if _best_perm: result['address_permanent'] = _best_perm
    if _best_pres: result['address_present']   = _best_pres
    result['photo_b64']    = photo_b64
    result['nid_page_img'] = _nid_page_img
    result['address_permanent_img'] = _nid_addr_imgs.get('address_permanent_img')
    result['address_present_img']   = _nid_addr_imgs.get('address_present_img')
    _df_clean = []
    _seen_dl = False
    for _d in docs_found:
        if 'Driving License' in _d:
            if not _seen_dl:
                _df_clean.append('Driving License')
                _seen_dl = True
        else:
            _df_clean.append(_d)
    result['docs_found'] = list(dict.fromkeys(_df_clean))

    # ── Mismatch detection — source-aware ──────────────────────────────────
    mismatches = []

    CHECK_FIELDS = [
        ('nid',      'NID Number'),
        ('nid_new',  'New NID Number'),
        ('name',     'Name'),
        ('father',   "Father's Name"),
        ('mother',   "Mother's Name"),
        ('dob',      'Date of Birth'),
        ('profession','Profession'),
        ('address_permanent', 'Permanent Address'),
    ]

    def _norm_val(field, v):
        v = str(v).strip().upper()
        if not v or v in ('N/A','NA','NONE',''): return ''
        if field == 'dob':
            import datetime as _dtt
            for _fmt in ('%d/%m/%Y','%m/%d/%Y','%Y-%m-%d','%d-%m-%Y'):
                try: return _dtt.datetime.strptime(v, _fmt).strftime('%Y-%m-%d')
                except Exception: pass
        if field in ('name','father','mother'):
            # Strip trailing dates: "Khodeza Begum 01/01/1970" → "Khodeza Begum"
            v = _re.sub(r'\s+\d{1,2}/\d{1,2}/\d{4}.*$', '', v).strip()
            v = _re.sub(r'\s+\d{4}-\d{2}-\d{2}.*$', '', v).strip()
            v = _re.sub(r'^MD\.?\s+', 'MD. ', v)
            v = _re.sub(r'\s+', ' ', v).strip()
            _NAME_REJECT = {'PASSPORT STATUS','REVOKED','ACTIVE','OFFICIAL',
                            'GOVERNMENT SERVICE','PVT SERVICE','BANGLADESH',
                            'N/A','NA','NONE','GOVERNMENT','SERVICE'}
            if v.upper() in _NAME_REJECT or len(v) < 4: return ''
            if _re.search(r'[\u0980-\u09FF]', v): return ''  # reject Bengali
            if len(v) > 60: return ''
            # Reject if contains digits or N/A tokens (garbage from multi-column)
            if _re.search(r'\b(N/A|NA|880|\d{4,})\b', v): return ''
            # Reject if looks like a label/keyword combination (VEHICLE NUMBER REGISTRATION)
            _LABEL_WORDS = {'VEHICLE','NUMBER','REGISTRATION','PERMIT','ROUTE','FITNESS',
                            'OFFICE','RMO','BPO','POST','THANA','UPAZILA','DIVISION',
                            'DISTRICT','UNION','WARD','MOUZA','REGION','VILLAGE',
                            'EXPIRE','ISSUE','TOKEN','TAX','JOINT','OWNER',
                            'LADEN','UNLADEN','WEIGHT','AXLE','SERIES','CAPACITY',
                            'TYPE','CLASS','COLOR','CC','CONTACT','TELEPHONE','BOI',
                            'VISA','INCORPORATION','RESPONSE','SMART','GENDER','BLOOD',
                            'MONITORING','CENTRE','NATIONAL','TELECOMMUNICATION',
                            'POLICE','STATION','PERMANENT','PRESENT','ADDRESS'}
            _vwords = set(v.upper().split())
            # Reject if majority of tokens are infrastructure/label keywords
            _INFRA_WORDS = {'VEHICLE','NUMBER','REGISTRATION','PERMIT','ROUTE','FITNESS',
                            'OFFICE','RMO','BPO','POST','THANA','UPAZILA','DIVISION',
                            'DISTRICT','UNION','WARD','MOUZA','REGION','VILLAGE',
                            'EXPIRE','ISSUE','TOKEN','TAX','JOINT','OWNER',
                            'LADEN','WEIGHT','AXLE','SERIES','CAPACITY',
                            'TYPE','CLASS','COLOR','CC','CONTACT','TELEPHONE',
                            'MONITORING','CENTRE','NATIONAL','TELECOMMUNICATION',
                            'POLICE','STATION','PERMANENT','PRESENT','ADDRESS',
                            'PASSPORT','STATUS','PROFESSION','DOCUMENT','REVOKED',
                            'REISSUE','SPOUSE','GOVERNMENT','SERVICE','OFFICIAL','ACTIVE',
                            'FATHER','MOTHER','NAME','DATE','BIRTH','GENDER',
                            'NATIONALITY','BANGLADESH','BANGLADESHI'}
            _bad_count = len(_vwords & _INFRA_WORDS)
            if len(_vwords) > 0 and _bad_count / len(_vwords) >= 0.5:
                return ''
            if len(_vwords) >= 2 and _vwords.issubset(_INFRA_WORDS | {'OF','THE','AND','IN'}):
                return ''
            if not _re.search(r'[A-Z]{2,}', v): return ''
            # Reject if contains Bengali characters (OCR garbage)
            if _re.search(r'[ঀ-৿]', v): return ''
            # Reject if looks like address (contains comma+digits or road keywords)
            if _re.search(r'[ঀ-৿]|ROAD-\d|BLOCK-\w|WARD|KHILGAON|RAMPURA', v): return ''
            # Max reasonable name length
            if len(v) > 60: return ''
        if 'nid' in field:
            v = _re.sub(r'[^0-9]', '', v)
            if len(v) < 10: return ''
        if 'address' in field:
            if len(v) < 10: return ''
            if _re.search(r'[ঀ-৿]', v): return ''
            # Reject passport status, revocation notices etc
            _BAD_ADDR = ('DOCUMENT REVOKED','REISSUE','PASSPORT STATUS',
                         'REVOKED DUE','CANCELLED','ACTIVE','OFFICIAL',
                         'GOVERNMENT','MONITORING CENTRE')
            if any(b in v for b in _BAD_ADDR): return ''
            tokens = v.split()
            na_count = sum(1 for t in tokens if t in ('N/A','NA','N'))
            if na_count > len(tokens) * 0.3: return ''
            digit_count = sum(1 for c in v if c.isdigit())
            if digit_count > len(v) * 0.5: return ''
        return v

    # Mismatch: different normalized values from different sources
    for field, label in CHECK_FIELDS:
        src_list = store.get(field, [])  # [(val, doc_label), ...]
        if not src_list: continue
        # Normalize each value, keep source
        norm_with_src = []
        for v, lbl in src_list:
            nv = _norm_val(field, v)
            if nv:
                norm_with_src.append((nv, lbl))
        # Group by normalized value → {norm_val: first_source}
        seen = {}
        for nv, lbl in norm_with_src:
            if nv not in seen:
                seen[nv] = lbl
        if len(seen) > 1:
            # NID value first as primary reference
            parts = []
            for val, lbl in seen.items():
                if lbl == 'NID':
                    parts.insert(0, f"NID: {val}")
                else:
                    parts.append(f"{lbl}: {val}")
            mismatches.append(f"{label} — " + " | ".join(parts))

    result['mismatches'] = mismatches
    result['mismatches_detail'] = mismatches
    return result
def _profile_html_section(profile):
    """
    Profile Analysis section — 3-column table design per specification.
    NID: only English data (name, old NID, new NID).
    Address: NID excluded; other sources only.
    Mismatch: highlighted at bottom.
    """
    if not profile:
        return ''
    has_data = any(profile.get(k) for k in [
        'name','nid','passport','license_no','vehicle_reg',
        'tin','dob','father','photo_b64','docs_found'
    ])
    if not has_data:
        return ''

    import re as _re_ph

    # ── Helpers ──────────────────────────────────────────────────────────
    def _src_badge(lbl):
        colors = {
            'NID':          ('#dbeafe','#1e40af'),
            'Passport':     ('#fce7f3','#9d174d'),
            'TIN':          ('#fef9c3','#854d0e'),
            'Driving License': ('#dcfce7','#166534'),
            'Vehicle Reg':  ('#ede9fe','#5b21b6'),
            'Vehicle Registration': ('#ede9fe','#5b21b6'),
        }
        bg, fg = colors.get(lbl, ('#f1f5f9','#475569'))
        return (f"<span style='background:{bg};color:{fg};padding:1px 7px;"
                f"border-radius:10px;font-size:10px;font-weight:700;"
                f"margin-left:5px;white-space:nowrap'>{lbl}</span>")

    def _multi(field_key, formatter=None):
        """All values for a field across sources — comma-separated or a/b list."""
        # ── Garbage / label-leak detection for person-name fields ────────
        _NAME_FIELDS = {'name','father','mother','spouse'}

        # Comprehensive OCR label / infrastructure keyword blacklist
        _BAD_TOKENS = {
            # NID structural labels
            'POST','OFFICE','RMO','BPO','THANA','UPAZILA','DIVISION','DISTRICT',
            'UNION','WARD','MOUZA','REGION','VILLAGE','ROAD','BLOCK','HOUSE','FLAT',
            'POSTAL','CODE','ADDRESS','PERMANENT','PRESENT',
            # Passport / DL labels
            'PASSPORT','STATUS','PROFESSION','DOCUMENT','REVOKED','REISSUE',
            'SPOUSE','GOVERNMENT','SERVICE','OFFICIAL','ACTIVE',
            'FATHER','MOTHER','NAME','NUMBER','DATE','BIRTH','GENDER',
            # TIN / misc
            'NATIONALITY','BANGLADESH','BANGLADESHI','N/A', 'NA',
            # OCR infrastructure noise
            'MONITORING','CENTRE','NATIONAL','TELECOMMUNICATION','POLICE','STATION',
            'PERMIT','FITNESS','TOKEN','SERIES','AXLE','CAPACITY','WEIGHT',
        }

        # Exact bad-value set (full strings)
        _BAD_EXACT = {
            'POST OFFICE RMO', 'DIVISION POST OFFICE', 'PERMANENT ADDRESS',
            'PRESENT ADDRESS', 'PASSPORT STATUS', 'PASSPORT NUMBER',
            'PASSPORT TYPE', 'SPOUSE NAME', "FATHER'S NAME", "MOTHER'S NAME",
            'FATHERS NAME', 'MOTHERS NAME', 'DOCUMENT REVOKED',
            'DOCUMENT REVOKED DUE TO REISSUE', 'REVOKED DUE TO REISSUE',
            'GOVERNMENT SERVICE', 'PVT SERVICE', 'ACTIVE', 'REVOKED',
            'OFFICIAL', 'BANGLADESH', 'BANGLADESHI', 'BIRTH ID',
            'DATE OF BIRTH', 'PASSPORT', 'NATIONAL ID',
        }

        def _is_bad_name_val(s):
            """True if value looks like OCR garbage / label leak — not a real person name."""
            if field_key not in _NAME_FIELDS: return False
            su = str(s).upper().strip()
            if not su: return True
            if su in _BAD_EXACT: return True
            toks = su.split()
            if not toks: return True
            # Reject if every token is in infrastructure blacklist
            if all(t in _BAD_TOKENS for t in toks): return True
            # Reject if majority (≥ 60%) of tokens are bad
            bad_count = sum(1 for t in toks if t in _BAD_TOKENS)
            if len(toks) > 0 and bad_count / len(toks) >= 0.6: return True
            # Reject if no token looks like a real name word
            # Real name words: ≥4 chars, has vowel, not in infra blacklist
            def _looks_like_name_word(w):
                return (len(w) >= 4
                        and bool(_re_ph.search(r'[AEIOU]', w))
                        and w not in _BAD_TOKENS)
            if not any(_looks_like_name_word(t) for t in toks):
                return True
            return False

        src_list = profile.get(f'{field_key}_src', [])
        if not src_list:
            v = profile.get(field_key, '')
            if not v or str(v).upper() in ('N/A','NA',''): return ''
            if _re_ph.search(r'[\u0980-\u09FF]', str(v)): return ''
            if _is_bad_name_val(v): return ''
            return str(v)
        seen = {}
        for val, lbl in src_list:
            v = (val or '').strip()
            if not v or v.upper() in ('N/A','NA',''): continue
            if _re_ph.search(r'[\u0980-\u09FF]', v): continue  # skip Bengali
            if _is_bad_name_val(v): continue  # skip label leaks / garbage
            if v not in seen:
                seen[v] = lbl
        if not seen: return ''
        items = list(seen.items())
        if len(items) == 1:
            v, lbl = items[0]
            return f"{formatter(v) if formatter else v}{_src_badge(lbl)}"
        # Multiple values → a) b) c) list
        parts = []
        for i, (v, lbl) in enumerate(items):
            letter = chr(ord('a') + i)
            parts.append(
                f"<div style='margin:2px 0'><strong>{letter})</strong> "
                f"{formatter(v) if formatter else v}{_src_badge(lbl)}</div>")
        return ''.join(parts)

    def _single(field_key, formatter=None):
        """Best single value — NID preferred."""
        src_list = profile.get(f'{field_key}_src', [])
        # NID first
        for val, lbl in src_list:
            if lbl == 'NID' and val and val.upper() not in ('N/A','NA',''):
                if not _re_ph.search(r'[\u0980-\u09FF]', val):
                    return formatter(val) if formatter else val
        # Then any
        for val, lbl in src_list:
            if val and val.upper() not in ('N/A','NA',''):
                if not _re_ph.search(r'[\u0980-\u09FF]', val):
                    return formatter(val) if formatter else val
        v = profile.get(field_key, '')
        if v and str(v).upper() not in ('N/A','NA',''):
            # Bengali check on fallback path too
            if _re_ph.search(r'[\u0980-\u09FF]', str(v)): return ''
            return formatter(v) if formatter else v
        return ''

    def _fmt_mobile(v):
        """Format mobile: 880XXXXXXXXXX → +880-XX-XXXX-XXXX / 01XXXXXXXXX → 0XXX-XXXXXX"""
        v = _re_ph.sub(r'[^0-9]', '', str(v))
        if v.startswith('880') and len(v) == 13:
            return f"+{v[:3]}-{v[3:5]}-{v[5:9]}-{v[9:]}"
        if v.startswith('0') and len(v) == 11:
            return f"{v[:4]}-{v[4:7]}-{v[7:]}"
        return v

    def _eng_addr_multi(field_key):
        """Address: Passport/DL/other sources preferred; NID as fallback if no other source."""
        src_list = profile.get(f'{field_key}_src', [])
        seen_non_nid = {}
        seen_nid = {}
        for val, lbl in src_list:
            v = (val or '').strip()
            if not v or v.upper() in ('N/A','NA',''): continue
            if _re_ph.search(r'[\u0980-\u09FF]', v): continue
            # Reject garbage
            if _re_ph.search(r'\b(Route|Permit|Fitness|Series|Token|NIA|N/A N/A)\b', v): continue
            # Reject passport status leaks (revoked/reissue/active/document revoked)
            if _re_ph.search(r'\b(REVOKED|REISSUE|ACTIVE|Document\s+revoked)\b', v, _re_ph.I): continue
            if len(v) < 8: continue
            if lbl == 'NID':
                if v not in seen_nid: seen_nid[v] = lbl
            else:
                if v not in seen_non_nid: seen_non_nid[v] = lbl
        # Non-NID sources preferred; fallback to NID if nothing else available
        seen = seen_non_nid if seen_non_nid else seen_nid
        if not seen: return ''
        items = list(seen.items())
        if len(items) == 1:
            return f"{items[0][0]}{_src_badge(items[0][1])}"
        parts = []
        for i, (v, lbl) in enumerate(items):
            letter = chr(ord('a') + i)
            parts.append(f"<div style='margin:2px 0'><strong>{letter})</strong> {v}{_src_badge(lbl)}</div>")
        return ''.join(parts)

    # ── NID image (full page) ──────────────────────────────────────────
    nid_img_html = ''
    if profile.get('nid_page_img'):
        nid_img_html = (f"<img src='{profile['nid_page_img']}' "
                        f"style='max-width:100%;height:auto;border-radius:6px;"
                        f"border:1px solid #e2e8f0'/>")

    # ── Photo ──────────────────────────────────────────────────────────
    photo_html = ''
    if profile.get('photo_b64'):
        # NID no longer provides names — profile['name'] comes from Passport/DL/TIN only
        import re as _re_pn
        _photo_name = (profile.get('manual_name') or profile.get('name') or '').strip()
        # Reject Bengali just in case
        if _re_pn.search(r'[\u0980-\u09FF]', _photo_name):
            _photo_name = ''

        photo_html = (f"<div style='text-align:center;margin-bottom:8px'>"
                      f"<img src='{profile['photo_b64']}' "
                      f"style='width:120px;height:150px;object-fit:cover;"
                      f"border-radius:8px;border:3px solid #e2e8f0'/>"
                      f"<div style='margin-top:6px;font-size:14px;font-weight:700;"
                      f"color:#0f172a'>{_photo_name}</div>"
                      f"</div>")

    # ── Docs badges ──────────────────────────────────────────────────
    docs_badges = ''.join(_src_badge(d) for d in (profile.get('docs_found') or []))

    # ── TIN combined ──────────────────────────────────────────────────
    tin_parts = []
    if profile.get('tin'):     tin_parts.append(f"<div>Old TIN: <strong>{profile['tin']}</strong></div>")
    if profile.get('tin_new'): tin_parts.append(f"<div>New TIN: <strong>{profile['tin_new']}</strong></div>")
    tin_str = ''.join(tin_parts)

    # ── Passport details ──────────────────────────────────────────────
    passport_status = ''
    for val, lbl in profile.get('passport_status_src', []):
        if val and val.upper() not in ('N/A','NA',''):
            passport_status = val; break
    if not passport_status: passport_status = profile.get('passport_status','')

    # ── Table rows ─────────────────────────────────────────────────────
    def tr(field, detail):
        """Return (field, detail) — only included if detail has content."""
        if not detail or str(detail).strip() in ('', 'N/A', 'NA'):
            return None
        return (field, detail)

    def _build_table(row_defs):
        """Build table with sequential serial numbers, skipping empty rows."""
        rows_html = ''
        serial = 1
        for item in row_defs:
            if item is None:
                continue
            field, detail = item
            rows_html += (
                f"<tr>"
                f"<td style='padding:6px 8px;text-align:center;color:#94a3b8;"
                f"font-size:11px;border:1px solid #e2e8f0;width:36px'>{serial}</td>"
                f"<td style='padding:6px 10px;font-weight:600;font-size:12px;"
                f"color:#374151;border:1px solid #e2e8f0;white-space:nowrap;"
                f"background:#f8fafc;min-width:160px'>{field}</td>"
                f"<td style='padding:6px 10px;font-size:12px;color:#0f172a;"
                f"border:1px solid #e2e8f0;line-height:1.7'>{detail}</td>"
                f"</tr>"
            )
            serial += 1
        return rows_html

    # ── Name row: Passport/DL/TIN sources only (NID not involved) ────────
    def _name_row_html():
        import re as _re_nr
        seen = {}
        for val, lbl in profile.get('name_src', []):
            v = (val or '').strip()
            if not v or v.upper() in ('N/A', 'NA', ''): continue
            if _re_nr.search(r'[\u0980-\u09FF]', v): continue
            if len(v) < 4: continue
            if v.upper() not in {k.upper() for k in seen}:
                seen[v] = lbl
        if not seen:
            nm = (profile.get('name') or '').strip()
            if nm and not _re_nr.search(r'[\u0980-\u09FF]', nm):
                seen[nm] = ''
        if not seen: return ''
        items = list(seen.items())
        if len(items) == 1:
            v, lbl = items[0]
            return f"{v}{_src_badge(lbl) if lbl else ''}"
        parts = []
        for i, (v, lbl) in enumerate(items):
            letter = chr(ord('a') + i)
            parts.append(f"<div style='margin:2px 0'><strong>{letter})</strong> "
                         f"{v}{_src_badge(lbl) if lbl else ''}</div>")
        return ''.join(parts)

    rows = [
        tr('Name', _name_row_html()),
        tr('NID Info', nid_img_html or ''),
        tr('Date of Birth', _single('dob')),
        tr("Father's Name", _multi('father')),
        tr("Mother's Name", _multi('mother')),
        tr('Spouse Name', _multi('spouse')),
        tr('Gender', _single('gender')),
        tr('Blood Group', _single('blood_group')),
        tr('Occupation / Profession', _multi('profession')),
        tr('Mobile Number', _multi('mobile', _fmt_mobile)),
        tr('Nationality', _single('nationality')),
        tr('Old NID Number', _single('nid')),
        tr('New NID / Smart ID', _single('nid_new')),
        tr('TIN Number', tin_str),
        tr('Passport Number', _single('passport')),
        tr('Old / Previous Passport No', _single('prev_passport')),
        tr('Passport Issue Date', _single('passport_issue')),
        tr('Passport Expiry Date', _single('passport_expiry')),
        tr('Passport Status', passport_status),
        tr('Permanent Address', _eng_addr_multi('address_permanent')),
        tr('Present Address', _eng_addr_multi('address_present')),
        tr('Driving License Number', _single('license_no')),
        tr('DL Issue Date', _single('license_issue')),
        tr('DL Expiry Date', _single('license_expiry')),
        tr('Vehicle Number', _single('vehicle_reg')),
        tr('Vehicle Type', _single('vehicle_type')),
        tr('Vehicle Registration Date', _single('vehicle_reg_date')),
        tr('Tax Token Expiry Date', _single('tax_token_expire')),
        tr('Fitness Expiry Date', _single('fitness_expire')),
    ]
    table_html = '<table style="border-collapse:collapse;width:100%">' + _build_table(rows) + '</table>'

    # ── Mismatch section ──────────────────────────────────────────────
    mismatch_html = ''
    mm_list = profile.get('mismatches', [])
    if mm_list:
        mm_items = []
        for mm in mm_list:
            if ' — ' in mm:
                fld, rest = mm.split(' — ', 1)
                parts = rest.split(' | ')
                parts_html = ''.join(
                    f"<div style='margin:3px 0 3px 16px;color:#374151'>"
                    f"<span style='color:#6b7280;font-size:11px'>{p.split(':')[0].strip()} →</span> "
                    f"<strong>{p.split(':',1)[1].strip() if ':' in p else p}</strong></div>"
                    for p in parts
                )
                mm_items.append(
                    f"<div style='margin-bottom:10px'>"
                    f"<div style='font-weight:700;color:#dc2626;font-size:12px'>"
                    f"⚠ {fld}</div>{parts_html}</div>"
                )
        if mm_items:
            mismatch_html = (
                f"<div style='background:#fef2f2;border:1px solid #fecaca;"
                f"border-radius:8px;padding:14px 16px;margin-top:14px'>"
                f"<div style='font-weight:700;color:#dc2626;font-size:13px;"
                f"margin-bottom:10px;border-bottom:1px solid #fecaca;padding-bottom:6px'>"
                f"⚠️ Mismatch / Inconsistencies Found</div>"
                + ''.join(mm_items)
                + "</div>"
            )

    # ── Layout ────────────────────────────────────────────────────────
    return f"""
<div style="background:white;border-radius:12px;border:2px solid #2563eb;
            padding:20px 24px;margin-bottom:24px;
            box-shadow:0 4px 16px rgba(37,99,235,0.10)">
  <div style="font-size:1.1rem;font-weight:800;color:#1e3a8a;
              margin-bottom:14px;border-bottom:2px solid #dbeafe;
              padding-bottom:8px;display:flex;align-items:center;gap:8px">
    👤 Profile Analysis
    <span style="font-size:0.75rem;font-weight:400;color:#64748b">{docs_badges}</span>
  </div>
  <div style="display:flex;gap:20px;align-items:flex-start">
    <div style="min-width:130px;max-width:150px">{photo_html}</div>
    <div style="flex:1;overflow-x:auto">{table_html}</div>
  </div>
  {mismatch_html}
</div>
"""


def main():
    # ── Top App Header ──
    current_user = st.session_state.get("current_user", "user")
    st.markdown(f"""
    <div class="app-header">
        <div class="app-header-left">
            <div class="app-header-logo">📊</div>
            <div class="app-header-title">CDR Analysis Platform</div>
        </div>
        <div class="app-header-nav" id="main-nav">
            <span class="nav-link" id="nav-cdr">📈 CDR Analysis</span>
            <span class="nav-link" id="nav-link">🔗 Link Analysis</span>
            <span style="color:#64748b;font-size:0.82rem;">👤 {current_user}</span>
            <span style="color:#94a3b8;font-size:0.75rem;margin-left:0.5rem;">⏱ {max(0, 30 - int(_idle_secs//60))}m left</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Page Navigation via session state ──
    if "current_page" not in st.session_state:
        st.session_state["current_page"] = "cdr"

    page = st.session_state["current_page"]

    # ── Actual Navigation Buttons (hidden but functional) ──
    # ── Page Navigation + Logout ──
    _nav_col1, _nav_col2, _nav_spacer, _logout_col = st.columns([1, 1, 7, 1])
    with _nav_col1:
        if st.button("📈 CDR Analysis", key="nav_cdr_btn",
                     type="primary" if page == "cdr" else "secondary",
                     use_container_width=True):
            st.session_state["current_page"] = "cdr"
            st.rerun()
    with _nav_col2:
        if st.button("🔗 Link Analysis", key="nav_link_btn",
                     type="primary" if page == "link" else "secondary",
                     use_container_width=True):
            st.session_state["current_page"] = "link"
            st.rerun()
    with _logout_col:
        if st.button("🚪 Logout", key="logout_btn",
                     use_container_width=True):
            _logout()

    if page == "link":
        link_analysis_page()
        return

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
            type=['xlsx', 'xls', 'csv'],
            help="CDR Excel file (any operator)",
            label_visibility="collapsed"
        )
        st.caption("Maximum file size: 200MB")
        # ── security: validate size + magic bytes (no-op on valid files) ──
        if uploaded is not None and not validate_upload(uploaded, kind="excel"):
            uploaded = None

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
            label_visibility="collapsed",
            key="target_number_input"
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

        # ── Target Location ──
        st.markdown("""
        <div class="card-title" style="margin-top:1rem;">
            <div class="card-icon-blue">📍</div>
            <div>
                <div>Target Location <span style="color:#94a3b8; font-weight:500; font-size:0.85rem;">(Optional)</span></div>
                <div class="card-subtitle" style="font-weight:400;">Enter district or upazila name to find visit dates.</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        target_location = st.text_input(
            label=" ",
            placeholder="e.g. Rampal, Bagerhat, Keraniganj",
            key="target_location_input",
            label_visibility="collapsed"
        )
        target_location = target_location.strip() if target_location else None

        if target_location:
            st.markdown(f"""
            <div style="background:#f0fdf4; border-radius:10px; padding:0.6rem 1rem; margin-top:0.4rem; display:flex; gap:0.5rem; align-items:center;">
                <div style="color:#16a34a; font-size:1rem;">📍</div>
                <div style="color:#15803d; font-size:0.83rem;">Will search for activity near: <strong>{target_location}</strong></div>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div style="background:#f0fdf4; border-radius:10px; padding:0.6rem 1rem; margin-top:0.4rem; display:flex; gap:0.5rem; align-items:flex-start;">
                <div style="color:#16a34a; font-size:1rem;">📍</div>
                <div style="color:#15803d; font-size:0.83rem; line-height:1.4;">Shows all dates the subscriber visited a specific district/upazila.</div>
            </div>
            """, unsafe_allow_html=True)

    # ── Profile & Documents Upload (Optional) ──
    with st.expander("👤 Profile & Documents (Optional — NID, Passport, DL, TIN, Vehicle)", expanded=False):
        st.caption(
            "পরিচয়পত্র (NID, Passport, Driving License, TIN, Vehicle Registration) "
            "PDF বা ছবি আকারে আপলোড করুন। CDR রিপোর্টের উপরে Profile Analysis অংশে যুক্ত হবে।"
        )

        _prof_col1, _prof_col2 = st.columns([1, 2])
        with _prof_col1:
            _prof_photo = st.file_uploader(
                "Subject Photo (JPG/PNG)",
                type=["jpg","jpeg","png"],
                key="prof_photo_upload"
            )
            if _prof_photo is not None and not validate_upload(_prof_photo, kind="image"):
                _prof_photo = None
            _manual_name = st.text_input(
                "Subject Name (optional — NID name used if blank)",
                value=st.session_state.get('_manual_name_val',''),
                key="manual_name_input",
                placeholder="e.g. MD. JAHIRUL ISLAM"
            )
            if _manual_name.strip():
                st.session_state['_manual_name_val'] = _manual_name.strip().upper()
            else:
                st.session_state['_manual_name_val'] = ''
        with _prof_col2:
            _prof_docs = st.file_uploader(
                "Identity Documents (PDF or image — multiple allowed)",
                type=["pdf","jpg","jpeg","png"],
                key="prof_docs_upload",
                accept_multiple_files=True
            )
            if _prof_docs:
                _prof_docs = [d for d in _prof_docs if validate_upload(d, kind="any_doc")]

        # Build combined doc list (photo first if given)
        _all_docs = []
        if _prof_photo:
            _all_docs.append(_prof_photo)
        if _prof_docs:
            _all_docs.extend(_prof_docs)

        # Store uploaded files — parsing happens on Run Forensic
        if _all_docs:
            # Save file bytes now (before widget rerender loses them)
            _doc_bytes = []
            for _d in _all_docs:
                _doc_bytes.append({'name': _d.name, 'data': _d.read()})
                try: _d.seek(0)
                except Exception: pass
            st.session_state['_profile_doc_bytes'] = _doc_bytes
            st.info(f"📎 {len(_all_docs)} document(s) ready — will be parsed on Run Forensic Analysis")
        else:
            # No new docs — keep existing parsed data
            pass

    # Store profile_data — expander-এর বাইরে session_state থেকে পড়ো
    # manual_name শুধু দেওয়া থাকলে (doc ছাড়া) empty dict তৈরি করে রাখো
    # যাতে photo/name নিচে সঠিকভাবে render হয়
    _mn_val = st.session_state.get('_manual_name_val', '')
    if _mn_val and not st.session_state.get('_profile_data'):
        st.session_state['_profile_data'] = {'manual_name': _mn_val}
    elif _mn_val and st.session_state.get('_profile_data') is not None:
        st.session_state['_profile_data']['manual_name'] = _mn_val
    profile_data = st.session_state.get('_profile_data', None)

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
    # ── Session State Cache: একই file দিলে re-analysis বন্ধ ──
    # hashlib → _hashlib_mod (module-level import)
    _file_bytes_raw = uploaded.read()
    _file_hash = _hashlib_mod.md5(_file_bytes_raw).hexdigest()

    # ── Run Analysis Button ──
    # File upload হলেই analysis শুরু না করে, button click করলে শুরু হবে
    _run_key = f"run_analysis_{_file_hash}"
    if _run_key not in st.session_state:
        st.session_state[_run_key] = False

    # Link Analysis থেকে ফিরলে _run_key=True কিন্তু cache নাও থাকতে পারে
    # সেক্ষেত্রে "Re-run Analysis" button দেখাও
    _prof_hash_check = ''
    _pdata_check = st.session_state.get('_profile_data', None)
    if _pdata_check:
        # hashlib → _hashlib_mod (module-level import)
        _ps = str(sorted((k,v) for k,v in _pdata_check.items()
                         if k not in ('photo_b64',) and v))
        _prof_hash_check = _hashlib_mod.md5(_ps.encode()).hexdigest()[:8]
    _cache_key_check = f"cdr_result_{_file_hash}_{_prof_hash_check}"
    _cache_exists = _cache_key_check in st.session_state

    # ── Parse profile docs BEFORE Run button check ──
    # যাতে profile hash সঠিক থাকে এবং Re-run loop না হয়
    _doc_bytes_pre = st.session_state.get('_profile_doc_bytes', [])
    if _doc_bytes_pre:
        _docs_hash = str(hash(str([(d['name'], len(d['data'])) for d in _doc_bytes_pre])))
        if st.session_state.get('_profile_docs_hash') != _docs_hash:
            import io as _io_pre
            class _PreFile:
                def __init__(self, d):
                    self.name = d['name']
                    self._buf = _io_pre.BytesIO(d['data'])
                def read(self): return self._buf.read()
                def seek(self, p): self._buf.seek(p)
            _pre_files = [_PreFile(d) for d in _doc_bytes_pre]
            _parsed_pre = _parse_profile_docs(_pre_files)
            st.session_state['_profile_data'] = _parsed_pre
            st.session_state['_profile_docs_hash'] = _docs_hash
        # Apply manual name override if set
        _mn = st.session_state.get('_manual_name_val','')
        if _mn and st.session_state.get('_profile_data'):
            st.session_state['_profile_data']['manual_name'] = _mn

    # ── Recalculate cache key with updated profile ──
    _prof_hash_check = ''
    _pdata_check2 = st.session_state.get('_profile_data', None)
    if _pdata_check2:
        # hashlib → _hashlib_mod (module-level import)
        _ps2 = str(sorted((k,v) for k,v in _pdata_check2.items()
                          if k not in ('photo_b64',) and v))
        _prof_hash_check = _hashlib_mod.md5(_ps2.encode()).hexdigest()[:8]
    _cache_key_check = f"cdr_result_{_file_hash}_{_prof_hash_check}"
    _cache_exists = _cache_key_check in st.session_state

    # ── Run / Re-run button logic ──
    # শুধু তখনই button দেখাও যখন analysis একদমই হয়নি (run_key=False)
    # run_key=True কিন্তু cache নেই → analysis চলুক (button দেখাবে না)
    if not st.session_state[_run_key]:
        st.markdown(f"""
        <div style="background:#f0fdf4;border:1px solid #86efac;border-radius:12px;
                    padding:1rem 1.5rem;margin:1rem 0;display:flex;align-items:center;gap:1rem">
            <div style="font-size:1.5rem">📂</div>
            <div>
                <div style="font-weight:600;color:#166534">CDR File Ready</div>
                <div style="font-size:0.85rem;color:#15803d">Click Run Analysis to start</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # ── Nominatim toggle ──────────────────────────────────────────────
        with st.expander("⚙️ Advanced GPS Options", expanded=False):
            _nom_enabled = st.checkbox(
                "🌐 Enable Geocoding (Nominatim / OpenStreetMap)",
                value=st.session_state.get('_nominatim_enabled', False),
                key=f"nom_toggle_{_file_hash}",
                help=(
                    "P.S: এবং area keyword উভয়ই GPS দিতে ব্যর্থ হলে "
                    "Nominatim API দিয়ে address geocode করবে।\n\n"
                    "⚠️ Rate limit: ১ req/sec — বড় CDR-এ (৫০০০+ row) "
                    "কয়েক মিনিট বেশি সময় লাগতে পারে।\n"
                    "Cell Tower CSV থাকলে এটা on করার দরকার নেই।"
                )
            )
            st.session_state['_nominatim_enabled'] = _nom_enabled
            if _nom_enabled:
                st.info(
                    "🌐 Geocoding চালু — শুধু P.S: এবং keyword scan "
                    "উভয়ই fail করলে Nominatim call হবে। "
                    "Cache করা থাকে, একই address দ্বিতীয়বার query হবে না।"
                )
            else:
                st.caption("Geocoding বন্ধ — Cell Tower CSV ও text parse দিয়ে GPS নেওয়া হবে।")

        # Nominatim global enable/disable — _nominatim_geocode() এই flag check করে
        _NOMINATIM_CACHE['__enabled__'] = st.session_state.get('_nominatim_enabled', False)

        if st.button("▶️ Run Analysis", type="primary", use_container_width=False,
                     key=f"run_btn_{_file_hash}"):
            st.session_state[_run_key] = True
            st.rerun()
        return  # Analysis will not start without clicking the button

    # run_key=True কিন্তু cache নেই → profile/target পরিবর্তন হয়েছে → Re-run option দেখাও
    # কিন্তু analysis চলতে দাও — return করো না
    if not _cache_exists:
        _rerun_col1, _rerun_col2 = st.columns([3, 1])
        with _rerun_col2:
            if st.button("🔄 Re-run Analysis", type="secondary",
                         key=f"rerun_btn_{_file_hash}"):
                # Clear old cache for this file (all profile variants)
                _keys_to_del = [k for k in st.session_state
                                if k.startswith(f"cdr_result_{_file_hash}")]
                for _k in _keys_to_del:
                    del st.session_state[_k]
                st.rerun()

    # Cache key: file hash + profile data hash
    _prof_hash = ''
    _pdata_for_hash = st.session_state.get('_profile_data', None)
    if _pdata_for_hash:
        # hashlib → _hashlib_mod (module-level import)
        _prof_str = str(sorted((k,v) for k,v in _pdata_for_hash.items()
                                if k not in ('photo_b64',) and v))
        _prof_hash = _hashlib_mod.md5(_prof_str.encode()).hexdigest()[:8]
    _cache_key = f"cdr_result_{_file_hash}_{_prof_hash}"

    # ── Apply Nominatim toggle state ──────────────────────────────────────
    # Analysis শুরুর আগে user-এর choice অনুযায়ী geocoding on/off করো
    _NOMINATIM_CACHE['__enabled__'] = st.session_state.get('_nominatim_enabled', False)

    # target_number এবং target_location আলাদা session_state-এ রাখি
    if "target_number_val" not in st.session_state:
        st.session_state["target_number_val"] = ""
    if "target_location_val" not in st.session_state:
        st.session_state["target_location_val"] = ""

    _from_cache = False
    # যদি cache-এ আছে এবং inputs same → cached result দেখাও
    if (_cache_key in st.session_state
            and st.session_state.get(_run_key, False)):
        _cached = st.session_state[_cache_key]

        # Profile inject into cached HTML
        # cached report may have been generated without profile — add it now if needed
        _cur_profile = st.session_state.get('_profile_data', None)
        _has_prof = bool(_cur_profile and any(
            _cur_profile.get(k) for k in ['name','nid','passport','license_no',
                                           'vehicle_reg','docs_found','photo_b64']
        ))
        _html_bytes = _cached["html_bytes"]
        _html_str = _html_bytes.decode('utf-8') if isinstance(_html_bytes, bytes) else _html_bytes

        # build_html() already includes profile — no extra inject needed
        _html_bytes_final = _html_bytes

        st.success("✅ Reports ready (cached)")
        _base = _cached["base_name"]
        dl1, dl2, dl3 = st.columns(3)
        with dl1:
            st.download_button("⬇️ Download HTML Report",
                data=_html_bytes_final, file_name=f"{_base}_Report.html",
                mime="text/html", use_container_width=True, key="dl_html_cached")
        with dl2:
            st.download_button("⬇️ Download Word Report",
                data=_cached["docx_bytes"], file_name=f"{_base}_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True, key="dl_docx_cached")
        with dl3:
            if _cached.get("map_bytes"):
                st.download_button("🗺️ Download Movement Map",
                    data=_cached["map_bytes"], file_name=f"{_base}_Movement_Map.html",
                    mime="text/html", use_container_width=True, key="dl_map_cached")
        # Reload df from pickle and re-render all analysis sections
        if _cached.get("sections") and _cached.get("df_pickle"):
            try:
                df = _safe_pickle_loads(_cached["df_pickle"])
                if df is None:
                    raise ValueError("Cached dataframe rejected (signature mismatch or corrupted)")
                phone          = _cached.get("phone", "")
                operator       = _cached.get("operator", "")
                date_range     = _cached.get("date_range", "")
                total_raw      = _cached.get("total_raw", 0)
                anomaly_count  = _cached.get("anomaly_count", 0)
                target_number  = _cached.get("target_number", "")
                target_location= _cached.get("target_location", "")
                profile_data   = st.session_state.get('_profile_data', None)
                _from_cache = True
            except Exception:
                return
        else:
            return

    class _DummyProgress:
        def progress(self, *a, **k): pass
        def empty(self, *a, **k): pass

    if not _from_cache:
        progress = st.progress(0, text="📥 Reading file...")
    else:
        progress = _DummyProgress()

    try:
        if not _from_cache:
            file_bytes = _file_bytes_raw
            progress.progress(15, text="🔍 Analyzing data structure...")
            df, col_map, total_raw, anomaly_count, sheet = load_and_clean(file_bytes)
        else:
            # df, phone, operator etc already loaded from cache pickle above
            col_map = {}
            sheet = ""

        progress.progress(15, text="🔍 Analyzing data structure...")

        # ── Cell Tower GPS Enrichment ─────────────────────────────────────
        # সব operator-এর CSV থেকে LAC+CID → exact GPS
        # LAC mismatch থাকলে CID+address token দিয়ে smart fallback
        cell_match_count = 0
        if not _from_cache:
          try:
            # hf_hub_download → _hf_hub_download (module-level import)
            import tempfile, os as _os, re as _re

            HF_REPO = "Faruk131086/Celltower"
            CELL_DIR = _os.path.join(tempfile.gettempdir(), "celltower_cache")
            _os.makedirs(CELL_DIR, exist_ok=True)



            # ── Local file fallback: HF_FILES_CFG hf নাম অনুযায়ী uploads ফোল্ডারে খোঁজো ──
            # যদি user uploads ফোল্ডারে Banglalink_4G.csv থাকে সেটা সরাসরি CELL_DIR-এ copy করো
            _UPLOADS_DIR = "/mnt/user-data/uploads"
            _LOCAL_FILE_MAP = {
                "Banglalink_4G.csv":    "Banglalink_4G.csv",
                "Banglalink_2G3G.csv":  "Banglalink_2G3G.csv",
                "GP_2G.csv":            "GP_2G.csv",
                "GP_3G.csv":            "GP_3G.csv",
                "GP_4G.csv":            "GP_4G.csv",
                "Robi_2G.csv":          "Robi_2G.csv",
                "Robi_2G-1.csv":        "Robi_2G.csv",
                "Robi_4G.csv":          "Robi_4G.csv",
                "Teletalk.csv":         "Teletalk.csv",
            }
            # shutil → _shutil_mod (module-level import)
            for _hf_name, _internal_name in _LOCAL_FILE_MAP.items():
                _src_path  = _os.path.join(_UPLOADS_DIR, _hf_name)
                _dest_path = _os.path.join(CELL_DIR, _internal_name)
                if _os.path.isfile(_src_path) and _os.path.getsize(_src_path) > 1000:
                    if not (_os.path.isfile(_dest_path) and _os.path.getsize(_dest_path) > 1000):
                        _shutil_mod.copy2(_src_path, _dest_path)

            # ── HF file config: all operators, all generations ──
            # key: lac/tac col, cid col, lat col, lon col, addr col
            # ── HF Repo: Faruk131086/Celltower ──
            # File names exactly as stored in HF dataset
            HF_FILES_CFG = {
                # ── Grameenphone ──────────────────────────────────────────
                "GP_2G.csv":  {"hf": "GP_2G.csv",  "enc": "utf-8",   "lac": "lac", "cid": "cellid",       "lat": "latitude", "lon": "longitude", "addr": "address",      "thana": "thana", "district": "district"},
                "GP_3G.csv":  {"hf": "GP_3G.csv",  "enc": "utf-8",   "lac": "lac", "cid": "cellid",       "lat": "latitude", "lon": "longitude", "addr": "address",      "thana": "thana", "district": "district"},
                "GP_4G.csv":  {"hf": "GP_4G.csv",  "enc": "latin-1", "lac": "lac", "cid": "cell_id",      "lat": "latitude", "lon": "longitude", "addr": "address",      "thana": "thana", "district": "district"},
                # ── Robi ──────────────────────────────────────────────────
                "Robi_2G.csv": {"hf": "Robi_2G.csv", "enc": "latin-1", "lac": "lac", "cid": "cell_id",   "lat": "latitude", "lon": "longitude", "addr": "address",      "thana": "thana", "district": "district"},
                # Robi_3G.csv নেই HF-এ → Robi_2G.csv fallback
                "Robi_4G.csv": {"hf": "Robi_4G.csv", "enc": "latin-1", "lac": "enodebid", "cid": "cell_id", "lat": "latitude", "lon": "longitude", "addr": "address", "thana": "thana", "district": "district", "lac_alt": "tac"},
                # ── Banglalink ────────────────────────────────────────────
                "Banglalink_2G3G.csv": {"hf": "Banglalink_2G3G.csv", "enc": "latin-1", "lac": "lac", "cid": "ci",           "lat": "latitude", "lon": "longitude", "addr": "site address", "thana": "thana", "district": "district"},
                "Banglalink_4G.csv":   {"hf": "Banglalink_4G.csv",   "enc": "latin-1", "lac": "tac", "cid": "eutrancellid", "lat": "latitude", "lon": "longitude", "addr": "site address", "thana": "thana", "district": "district"},
                # ── Teletalk ──────────────────────────────────────────────
                "Teletalk.csv": {"hf": "Teletalk.csv", "enc": "utf-8", "lac": "lac/ tal", "cid": "ci /tac", "lat": "latitude", "lon": "longitude", "addr": "full address"},
            }

            def _hf_token():
                try: return st.secrets["HF_TOKEN"]
                except Exception: return _os.environ.get("HF_TOKEN", None)

            def _norm_id(v):
                s = str(v).strip()
                # .0 suffix remove (float→int)
                if s.endswith(".0") and s[:-2].isdigit():
                    s = s[:-2]
                # Leading zeros strip for numeric IDs (e.g. "0130646" → "130646")
                # This ensures CDR "0130646" matches CSV "130646"
                if s.isdigit() and len(s) > 1:
                    s = str(int(s))
                return s

            def _addr_tokens(s):
                """Extract meaningful tokens from address for fuzzy matching.
                Filters out generic words that appear in almost every address
                (house, road, vill, post, dist, etc.) to prevent false matches.
                """
                _ADDR_STOPWORDS = {
                    'house','road','vill','village','post','dist','district',
                    'para','area','ward','block','lane','floor','flat','plot',
                    'holding','section','street','avenue','building','tower',
                    'police','station','office','market','bazar','bazaar',
                    'union','upazila','thana','mouza','mouja','mauja',
                    'north','south','east','west','central','new','old',
                    'more','moor','ganj','pur','nagar','gram','palli',
                }
                s = re.sub(r'[^a-z0-9 ]', ' ', str(s).lower())
                return set(t for t in s.split()
                           if len(t) > 4 and t not in _ADDR_STOPWORDS and not t.isdigit())

            def _haversine_km(la1, lo1, la2, lo2):
                """Fast Haversine distance in km between two GPS points."""
                # math → _math_mod (module-level import)
                R = 6371.0
                dlat = _math_mod.radians(la2 - la1)
                dlon = _math_mod.radians(lo2 - lo1)
                a = _math_mod.sin(dlat/2)**2 + _math_mod.cos(_math_mod.radians(la1)) * _math_mod.cos(_math_mod.radians(la2)) * _math_mod.sin(dlon/2)**2
                return R * 2 * _math_mod.asin(_math_mod.sqrt(a))

            # Bangladesh district approximate center coordinates for CDR address cross-validation
            BD_DISTRICT_COORDS = {
                'dhaka': (23.8103, 90.4125), 'chittagong': (22.3569, 91.7832),
                'sylhet': (24.8949, 91.8687), 'rajshahi': (24.3636, 88.6241),
                'khulna': (22.8456, 89.5403), 'barisal': (22.7010, 90.3535),
                'rangpur': (25.7439, 89.2752), 'mymensingh': (24.7471, 90.4203),
                'comilla': (23.4607, 91.1809), 'narayanganj': (23.6238, 90.4996),
                'gazipur': (24.0022, 90.4264), 'tangail': (24.2513, 89.9167),
                'manikganj': (23.8630, 90.0024), 'munshiganj': (23.5422, 90.5305),
                'narsingdi': (23.9324, 90.7154), 'kishoreganj': (24.4449, 90.7764),
                'netrakona': (24.8701, 90.7268), 'jamalpur': (24.9375, 89.9377),
                'sherpur': (25.0198, 90.0172), 'faridpur': (23.6070, 89.8429),
                'gopalganj': (23.0050, 89.8267), 'madaripur': (23.1641, 90.2012),
                'shariatpur': (23.2423, 90.4347), 'rajbari': (23.7574, 89.6441),
                'jessore': (23.1664, 89.2080), 'satkhira': (22.7185, 89.0705),
                'khulna': (22.8456, 89.5403), 'bagerhat': (22.6602, 89.7854),
                'narail': (23.1722, 89.5120), 'magura': (23.4876, 89.4196),
                'jhenaidah': (23.5447, 89.1530), 'kushtia': (23.9014, 89.1204),
                'chuadanga': (23.6401, 88.8416), 'meherpur': (23.7620, 88.6317),
                'bogra': (24.8465, 89.3773), 'sirajganj': (24.4535, 89.7006),
                'pabna': (24.0064, 89.2372), 'natore': (24.4204, 88.9872),
                'naogaon': (24.7936, 88.9312), 'chapainawabganj': (24.5965, 88.2785),
                'joypurhat': (25.1007, 89.0227), 'dinajpur': (25.6279, 88.6338),
                'thakurgaon': (26.0319, 88.4616), 'panchagarh': (26.3411, 88.5551),
                'nilphamari': (25.9310, 88.8563), 'lalmonirhat': (25.9923, 89.2847),
                'kurigram': (25.8054, 89.6363), 'gaibandha': (25.3288, 89.5287),
                'cox bazar': (21.4272, 92.0058), "cox's bazar": (21.4272, 92.0058),
                'bandarban': (22.1953, 92.2184), 'rangamati': (22.6522, 92.1615),
                'khagrachhari': (23.1193, 91.9847), 'feni': (23.0230, 91.3960),
                'noakhali': (22.8696, 91.0998), 'lakshmipur': (22.9449, 90.8412),
                'chandpur': (23.2333, 90.6518), 'brahmanbaria': (23.9570, 91.1115),
                'habiganj': (24.3745, 91.4153), 'moulvibazar': (24.4829, 91.7774),
                'sunamganj': (25.0658, 91.3950), 'gazipur': (24.0022, 90.4264),
                'savar': (23.8580, 90.2670), 'ashulia': (23.9481, 90.2895),
                'narayanganj': (23.6238, 90.4996),
            }

            GPS_CONFIDENCE_THRESHOLD = 40  # accept GPS if confidence score >= 40/100

            # ── Pre-build LAC-level GPS cluster stats (for Signal 3) ──
            # Populated after all CSV files are loaded, before row-by-row enrichment
            _lac_cluster_cache = {}  # lac → (median_lat, median_lon, std_km)

            def _build_lac_clusters(loaded_files):
                """
                For each LAC, compute the median GPS and spread (std_km) of all towers.
                Used to detect outlier GPS entries within a LAC.
                """
                # statistics → _statistics_mod (module-level import)
                lac_points = {}  # lac → [(lat, lon)]
                for fn, fdata in loaded_files.items():
                    for (lv, cv), (lat, lon, toks, addr) in fdata["exact"].items():
                        if lv not in lac_points:
                            lac_points[lv] = []
                        lac_points[lv].append((lat, lon))
                clusters = {}
                for lv, pts in lac_points.items():
                    if len(pts) < 2:
                        continue
                    lats = [p[0] for p in pts]
                    lons = [p[1] for p in pts]
                    med_lat = _statistics_mod.median(lats)
                    med_lon = _statistics_mod.median(lons)
                    # Compute spread: median distance from median point
                    dists = [_haversine_km(med_lat, med_lon, la, lo) for la, lo in pts]
                    spread = _statistics_mod.median(dists)
                    clusters[lv] = (med_lat, med_lon, spread)
                return clusters

            def _gps_confidence(lat, lon, cdr_addr_str, lac_key,
                                same_lac_ci_rows, neighbor_rows):
                """
                Multi-source GPS confidence score (0–100).
                Combines 4 independent signals — CDR address is just one of them.

                Signal 1 – CDR Address Match (0–25 pts)
                  GPS distance vs district mentioned in CDR address.
                  Low weight because CDR address can also be wrong.

                Signal 2 – Neighbor Consistency (0–35 pts)
                  If surrounding records (±3 rows) all cluster near the candidate GPS,
                  this GPS is likely correct even if CDR address disagrees.

                Signal 3 – LAC Cluster Outlier (0–25 pts)
                  All towers in this LAC normally sit within a tight geographic cluster.
                  An outlier GPS far from the LAC median gets penalised.

                Signal 4 – Same LAC+CI Majority Vote (0–15 pts)
                  Other rows with the identical LAC+CI: what GPS do they end up with
                  after previous signals? If >70% agree with this GPS → bonus.

                Total >= GPS_CONFIDENCE_THRESHOLD (40) → accept.
                """
                # statistics → _statistics_mod (module-level import)
                score = 0
                reasons = []

                # ── Signal 1: CDR Address Match (max 25 pts) ──
                addr_lower = str(cdr_addr_str).lower() if cdr_addr_str else ""
                matched_districts = []
                for dn, (dlat, dlon) in BD_DISTRICT_COORDS.items():
                    if dn in addr_lower:
                        matched_districts.append((dlat, dlon, dn))
                if not matched_districts:
                    # CDR address mentions no known district → neutral (12 pts, half credit)
                    score += 12
                    reasons.append("addr:neutral(12)")
                else:
                    min_dist = min(_haversine_km(lat, lon, dlat, dlon)
                                   for dlat, dlon, _ in matched_districts)
                    if min_dist <= 30:
                        score += 25; reasons.append(f"addr:match({min_dist:.0f}km,25)")
                    elif min_dist <= 60:
                        score += 15; reasons.append(f"addr:near({min_dist:.0f}km,15)")
                    elif min_dist <= 120:
                        score += 5;  reasons.append(f"addr:far({min_dist:.0f}km,5)")
                    else:
                        score += 0;  reasons.append(f"addr:mismatch({min_dist:.0f}km,0)")

                # ── Signal 2: Neighbor Consistency (max 35 pts) ──
                if neighbor_rows:
                    neighbor_lats = [r[0] for r in neighbor_rows if r[0] is not None]
                    neighbor_lons = [r[1] for r in neighbor_rows if r[1] is not None]
                    if len(neighbor_lats) >= 2:
                        med_nlat = _statistics_mod.median(neighbor_lats)
                        med_nlon = _statistics_mod.median(neighbor_lons)
                        dist_to_neighbors = _haversine_km(lat, lon, med_nlat, med_nlon)
                        if dist_to_neighbors <= 20:
                            score += 35; reasons.append(f"neighbors:close({dist_to_neighbors:.0f}km,35)")
                        elif dist_to_neighbors <= 60:
                            score += 20; reasons.append(f"neighbors:near({dist_to_neighbors:.0f}km,20)")
                        elif dist_to_neighbors <= 150:
                            score += 8;  reasons.append(f"neighbors:far({dist_to_neighbors:.0f}km,8)")
                        else:
                            score += 0;  reasons.append(f"neighbors:outlier({dist_to_neighbors:.0f}km,0)")
                    else:
                        score += 15; reasons.append("neighbors:insufficient(15)")
                else:
                    score += 15; reasons.append("neighbors:none(15)")

                # ── Signal 3: LAC Cluster Outlier Check (max 25 pts) ──
                if lac_key and lac_key in _lac_cluster_cache:
                    clat, clon, spread = _lac_cluster_cache[lac_key]
                    dist_from_cluster = _haversine_km(lat, lon, clat, clon)
                    # Allow up to 3× the LAC spread, minimum 30km tolerance
                    tolerance = max(spread * 3, 30)
                    if dist_from_cluster <= tolerance:
                        score += 25; reasons.append(f"lac_cluster:ok({dist_from_cluster:.0f}km,25)")
                    elif dist_from_cluster <= tolerance * 2:
                        score += 10; reasons.append(f"lac_cluster:borderline({dist_from_cluster:.0f}km,10)")
                    else:
                        score += 0;  reasons.append(f"lac_cluster:outlier({dist_from_cluster:.0f}km,0)")
                else:
                    score += 12; reasons.append("lac_cluster:unknown(12)")

                # ── Signal 4: Same LAC+CI Majority Vote (max 15 pts) ──
                if same_lac_ci_rows:
                    close = sum(1 for r in same_lac_ci_rows
                                if r[0] is not None and _haversine_km(lat, lon, r[0], r[1]) <= 25)
                    ratio = close / len(same_lac_ci_rows)
                    if ratio >= 0.7:
                        score += 15; reasons.append(f"majority:{ratio:.0%}(15)")
                    elif ratio >= 0.4:
                        score += 8;  reasons.append(f"majority:{ratio:.0%}(8)")
                    else:
                        score += 0;  reasons.append(f"majority:{ratio:.0%}(0)")
                else:
                    score += 8; reasons.append("majority:unknown(8)")

                return score, reasons

            def _load_cell_file(fname, cfg):
                """
                Load one cell tower CSV from HF (public dataset, direct URL).
                Returns:
                  cell_exact: dict (lac,cid) -> (lat, lon, addr_tokens)
                  cid_multi:  dict cid -> [(lat, lon, addr_tokens)]  [for fallback]
                """
                local = _os.path.join(CELL_DIR, fname)
                if not (_os.path.isfile(local) and _os.path.getsize(local) > 5000):
                    # requests → _requests_mod (module-level import)
                    # Try multiple URL formats for public HF datasets
                    urls_to_try = [
                        f"https://huggingface.co/datasets/{HF_REPO}/resolve/main/{cfg['hf']}",
                        f"https://huggingface.co/datasets/{HF_REPO}/resolve/refs%2Fconvert%2Fparquet/default/train/0000.parquet",
                        f"https://datasets-server.huggingface.co/rows?dataset={HF_REPO}&config=default&split=train",
                    ]
                    downloaded = False
                    last_err = ""
                    for url in urls_to_try:  # try all URLs in order
                        try:
                            token = _hf_token()
                            hdrs = {
                                "User-Agent": "Mozilla/5.0",
                                "Cache-Control": "no-cache",
                            }
                            if token:
                                hdrs["Authorization"] = f"Bearer {token}"
                            r = _requests_mod.get(url, headers=hdrs, stream=True, timeout=180)
                            r.raise_for_status()
                            with open(local, "wb") as _f:
                                for chunk in r.iter_content(65536):
                                    if chunk: _f.write(chunk)
                            downloaded = True
                            break
                        except Exception as e:
                            last_err = str(e)
                            # Try hf_hub_download as fallback
                            try:
                                path = _hf_hub_download(
                                    repo_id=HF_REPO, filename=cfg["hf"],
                                    repo_type="dataset", token=_hf_token(),
                                    local_dir=CELL_DIR
                                )
                                # shutil → _shutil_mod (module-level import)
                                if path and _os.path.abspath(path) != _os.path.abspath(local):
                                    _shutil_mod.copy2(path, local)
                                downloaded = True
                                break
                            except Exception as e2:
                                last_err = f"Direct: {e} | HF lib: {e2}"

                    if not downloaded:
                        # Token missing হলে friendly info, অন্যথায় warning
                        _no_token = ("401" in last_err or "Unauthorized" in last_err
                                     or "Repository Not Found" in last_err
                                     or "Invalid username" in last_err)
                        if _no_token:
                            st.info(
                                f"ℹ️ Cell tower GPS CSV ({cfg['hf']}) লোড হয়নি — "
                                f"HuggingFace token সেট করা নেই। "
                                f"Streamlit Secrets-এ `HF_TOKEN` যোগ করুন। "
                                f"GPS ছাড়া text-based location ব্যবহার হবে।"
                            )
                        else:
                            st.warning(f"⚠️ Could not download {cfg['hf']}: {last_err}")
                        return {}, {}, {}

                if not (_os.path.isfile(local) and _os.path.getsize(local) > 5000):
                    st.warning(f"⚠️ Downloaded file too small or missing: {fname}")
                    return {}, {}, {}

                cell_exact = {}
                cid_multi  = {}
                try:
                    if local.endswith(".xlsx"):
                        cdf2 = pd.read_excel(local, dtype=str)
                    else:
                        # Try configured encoding first, fallback to latin-1 then utf-8
                        for _enc in [cfg["enc"], "latin-1", "utf-8"]:
                            try:
                                cdf2 = pd.read_csv(local, dtype=str, encoding=_enc,
                                                   low_memory=False, on_bad_lines="skip")
                                break
                            except Exception:
                                continue
                    cdf2.columns = [c.lower().strip() for c in cdf2.columns]
                    lc = cfg["lac"]; ci = cfg["cid"]
                    la = cfg["lat"]; lo = cfg["lon"]
                    ac = cfg.get("addr", "")

                    # ── LAC/TAC column fallback ──
                    # BL 4G config uses 'tac' but some CSV versions use 'lac' instead
                    if lc not in cdf2.columns:
                        for alt_lc in ["enodebid", "enodeb_id", "enodeb id", "tac", "lac", "lac_id", "lac id", "enbid"]:
                            if alt_lc in cdf2.columns and alt_lc != lc:
                                lc = alt_lc
                                break

                    # ── CID column fallback ──
                    if ci not in cdf2.columns:
                        for alt_ci in ["eutrancellid", "eutrancell_id", "eutran_cell_id",
                                       "cell_id", "cell id", "cellid", "ci", "cid"]:
                            if alt_ci in cdf2.columns and alt_ci != ci:
                                ci = alt_ci
                                break

                    # ── LAT/LON column fallback ──
                    if la not in cdf2.columns:
                        for alt_la in ["latitude", "lat", "y"]:
                            if alt_la in cdf2.columns:
                                la = alt_la; break
                    if lo not in cdf2.columns:
                        for alt_lo in ["longitude", "lon", "lng", "x"]:
                            if alt_lo in cdf2.columns:
                                lo = alt_lo; break

                    if not all(c in cdf2.columns for c in [lc, la, lo]):
                        st.warning(f"⚠️ {fname}: Required columns not found. Available: {list(cdf2.columns[:10])}")
                        return {}, {}, {}
                    if ci not in cdf2.columns:
                        st.warning(f"⚠️ {fname}: Cell ID column '{ci}' not found. Available: {list(cdf2.columns[:10])}")
                        return {}, {}, {}

                    has_addr   = ac and ac in cdf2.columns
                    thana_col  = cfg.get("thana", "")
                    dist_col   = cfg.get("district", "")
                    has_thana  = thana_col and thana_col in cdf2.columns
                    has_dist   = dist_col  and dist_col  in cdf2.columns
                    # alt_ci_col: only useful if it's DIFFERENT from ci
                    alt_ci_col = None
                    if "eutrancellid" in cdf2.columns and "eutrancellid" != ci:
                        alt_ci_col = "eutrancellid"
                    # itertuples()._asdict() converts spaces→underscores in col names
                    def _col_key(col): return col.replace(" ","_").replace("-","_")
                    ac_key    = _col_key(ac)        if ac        else ""
                    thana_key = _col_key(thana_col) if thana_col else ""
                    dist_key  = _col_key(dist_col)  if dist_col  else ""
                    ci_key    = _col_key(ci)
                    lc_key    = _col_key(lc)
                    la_key    = _col_key(la)
                    lo_key    = _col_key(lo)
                    for row in cdf2.itertuples(index=False):
                        try:
                            rd = row._asdict()
                            lat = float(rd[la_key]); lon = float(rd[lo_key])
                            if not (20 <= lat <= 27 and 88 <= lon <= 93): continue
                            lv  = _norm_id(rd[lc_key]); cv = _norm_id(rd[ci_key])
                            # Build full address: address, Thana, District
                            parts = []
                            if has_addr:
                                a = str(rd.get(ac_key, rd.get(ac, ""))).strip().strip('"')
                                if a and a != 'nan': parts.append(a)
                            if has_thana:
                                t = str(rd.get(thana_key, rd.get(thana_col, ""))).strip()
                                if t and t != 'nan': parts.append(t)
                            if has_dist:
                                d = str(rd.get(dist_key, rd.get(dist_col, ""))).strip()
                                if d and d != 'nan': parts.append(d)
                            addr_str = ", ".join(parts)
                            # ── token: address + thana + district সব মিলিয়ে ──
                            # Thana/District token matching CDR address-এ থানার নাম থাকলে সঠিক tower বাছাই করতে সাহায্য করে
                            toks = _addr_tokens(addr_str)
                            k = (lv, cv)
                            # district value for trip filtering
                            _dist_val = str(rd[dist_col]) if has_dist else ""
                            if k not in cell_exact:
                                cell_exact[k] = (lat, lon, toks, addr_str)
                            if cv not in cid_multi:
                                cid_multi[cv] = []
                            cid_multi[cv].append((lat, lon, toks, addr_str))  # addr_str for token matching
                            # ── Robi 4G special: ENODEBID//100 = CDR LAC_ID ──────────
                            # CDR LAC_ID = ENODEBID ÷ 100 (integer division)
                            # CDR Cell ID শেষ ২ digit = CSV CELL_ID
                            # Ambiguity: একই (derived_lac, sector) অনেক tower — address token দিয়ে best নাও
                            # তাই cell_exact-এ first-wins না রেখে robi4g_multi-তে সব রাখো
                            if fname == "Robi_4G.csv":
                                try:
                                    _enb_int = int(float(str(rd[lc_key])))
                                    _derived_lac = str(_enb_int // 100)
                                    _k_robi4g = (_derived_lac, cv)
                                    # robi4g_multi: all candidates for address token matching
                                    if "robi4g_multi" not in locals():
                                        robi4g_multi = {}
                                    if _k_robi4g not in robi4g_multi:
                                        robi4g_multi[_k_robi4g] = []
                                    robi4g_multi[_k_robi4g].append((lat, lon, toks, addr_str))
                                except Exception:
                                    logger.debug('Suppressed exception', exc_info=True)
                            # Also index by eutrancid if present (BL 4G CDR may use it)
                            if alt_ci_col:
                                cv2 = _norm_id(rd[alt_ci_col])
                                k2 = (lv, cv2)
                                if k2 not in cell_exact:
                                    cell_exact[k2] = (lat, lon, toks, addr_str)
                                if cv2 not in cid_multi:
                                    cid_multi[cv2] = []
                                cid_multi[cv2].append((lat, lon, toks, addr_str))
                            # (Teletalk CGI/ECGI index built separately below — after loop)
                        except Exception:
                            continue
                    # ── Teletalk: CGI/ECGI index (post-loop) ─────────────────
                    # itertuples() renames 'CGI/ECGI' → '_8' (slash/space not allowed).
                    # Fix: loop-এর পরে pandas iterrows() দিয়ে CGI/ECGI column পড়ো।
                    if fname == 'Teletalk.csv':
                        try:
                            _cgi_c = next((c for c in cdf2.columns
                                           if 'cgi' in c.lower() or 'ecgi' in c.lower()), None)
                            _lat_c = next((c for c in cdf2.columns if c.lower()=='latitude'), None)
                            _lon_c = next((c for c in cdf2.columns if c.lower()=='longitude'), None)
                            _adr_c = next((c for c in cdf2.columns
                                           if 'full' in c.lower() and 'address' in c.lower()), None)
                            if _cgi_c and _lat_c and _lon_c:
                                for _ti, _tr in cdf2.iterrows():
                                    try:
                                        _clat = float(_tr[_lat_c])
                                        _clon = float(_tr[_lon_c])
                                        if not (20 <= _clat <= 27 and 88 <= _clon <= 93): continue
                                        _cgi_v = _norm_id(_tr[_cgi_c])
                                        if not _cgi_v or _cgi_v in ('nan','','0'): continue
                                        _caddr = str(_tr[_adr_c]).strip() if _adr_c else ''
                                        _ctoks = _addr_tokens(_caddr)
                                        _k_cgi = ('0', _cgi_v)
                                        if _k_cgi not in cell_exact:
                                            cell_exact[_k_cgi] = (_clat, _clon, _ctoks, _caddr)
                                    except Exception:
                                        continue
                        except Exception:
                            logger.debug('Teletalk CGI index error', exc_info=True)
                except Exception:
                    logger.debug('suppressed exception', exc_info=True)
                _r4g_multi_out = locals().get("robi4g_multi", {})
                return cell_exact, cid_multi, _r4g_multi_out

            # ── Detect operators used in this CDR ──
            ops_in_cdr = set()
            if "operator" in df.columns:
                for op_val in df["operator"].dropna().unique():
                    ov = str(op_val).lower()
                    if "grameen" in ov or "gp" in ov:    ops_in_cdr.add("gp")
                    elif "banglalink" in ov or "bl" in ov: ops_in_cdr.add("bl")
                    elif "robi" in ov or "airtel" in ov:  ops_in_cdr.add("robi")
                    elif "teletalk" in ov:                 ops_in_cdr.add("teletalk")
            # If single operator detected earlier
            _op_early = get_operator(df)
            op_key_main = _op_early.lower() if _op_early else ""
            if "grameen" in op_key_main or "gp" in op_key_main:   ops_in_cdr.add("gp")
            if "banglalink" in op_key_main or "bl" in op_key_main: ops_in_cdr.add("bl")
            if "robi" in op_key_main or "airtel" in op_key_main:  ops_in_cdr.add("robi")
            if "teletalk" in op_key_main:                          ops_in_cdr.add("teletalk")
            if not ops_in_cdr:
                ops_in_cdr = {"gp", "bl", "robi", "teletalk"}  # load all if unknown

            # ── operator+generation → CSV file mapping ──
            # Key: (operator_key, generation) → internal fname
            OP_GEN_FILE = {
                ("gp",       "2g"): "GP_2G.csv",
                ("gp",       "3g"): "GP_3G.csv",
                ("gp",       "4g"): "GP_4G.csv",
                ("robi",     "2g"): "Robi_2G.csv",
                ("robi",     "3g"): "Robi_2G.csv",       # Robi_3G.csv নেই HF-এ
                ("robi",     "4g"): "Robi_4G.csv",
                ("bl",       "2g"): "Banglalink_2G3G.csv",
                ("bl",       "3g"): "Banglalink_2G3G.csv",
                ("bl",       "4g"): "Banglalink_4G.csv",
                ("teletalk", "2g"): "Teletalk.csv",
                ("teletalk", "3g"): "Teletalk.csv",
                ("teletalk", "4g"): "Teletalk.csv",
            }

            # Determine which files to load based on CDR operators + generations present
            load_files = set()
            for op in ops_in_cdr:
                gens_present = set()
                if "cell_type" in df.columns:
                    for ct in df["cell_type"].dropna().unique():
                        ct_l = str(ct).lower()
                        if "2g" in ct_l or "gsm" in ct_l or "wcdma" not in ct_l and "lte" not in ct_l and "4g" not in ct_l and "3g" not in ct_l:
                            gens_present.add("2g")
                        if "3g" in ct_l or "wcdma" in ct_l or "umts" in ct_l:
                            gens_present.add("3g")
                        if "4g" in ct_l or "lte" in ct_l:
                            gens_present.add("4g")
                if not gens_present:
                    gens_present = {"2g", "3g", "4g"}  # load all if unknown
                for gen in gens_present:
                    fn = OP_GEN_FILE.get((op, gen))
                    if fn:
                        load_files.add(fn)

            # ── Load cell tower CSV files (cached locally) ──
            # file_key → {exact: {(lac,cid):(lat,lon,toks,addr)}, multi: {cid:[...]}}
            loaded_files = {}
            progress.progress(25, text="📡 Loading cell tower GPS data...")
            for fn in load_files:
                cfg = HF_FILES_CFG.get(fn, {})
                if not cfg:
                    continue
                ex, mu, r4g_multi = _load_cell_file(fn, cfg)
                loaded_files[fn] = {"exact": ex, "multi": mu, "addr_list": [], "robi4g_multi": r4g_multi}

            # ── CSV CDR: HF download পরে re-enrich ───────────────────────────
            # load_and_clean() CSV path-এ _apply_bts_enrichment() early call হয়েছিল
            # কিন্তু তখন HuggingFace থেকে Teletalk.csv download হয়নি।
            # এখন download শেষ — Teletalk CDR হলে cell_lat empty rows re-enrich করো।
            try:
                _is_csv_teletalk = (
                    'operator' in df.columns and
                    any('teletalk' in str(v).lower()
                        for v in df['operator'].dropna().unique()) and
                    'cell_id' in df.columns and
                    not getattr(df, '_csv_enriched_post_hf', False)
                )
                if _is_csv_teletalk:
                    _needs_enrich = (
                        'cell_lat' not in df.columns or
                        df['cell_lat'].notna().sum() == 0
                    )
                    if _needs_enrich:
                        df = _apply_bts_enrichment(df)
                        df._csv_enriched_post_hf = True
            except Exception:
                logger.debug('Suppressed exception', exc_info=True)

            # ── Robi 4G: addr_list তৈরি করো address token matching-এর জন্য ──
            for fn in list(loaded_files.keys()) if loaded_files else []:
                if fn == "Robi_4G.csv":
                    _r4g_cfg = HF_FILES_CFG.get(fn, {})
                    _addr_list = []
                    try:
                        _r4g_path = _os.path.join(CELL_DIR, fn)
                        if _os.path.exists(_r4g_path):
                            import csv as _csv_mod
                            with open(_r4g_path, 'r', encoding='latin-1', errors='replace') as _f:
                                _reader = _csv_mod.DictReader(_f)
                                _r4g_noise = {'and','the','road','ward','house','floor','building','thana','area','block','dhaka'}
                                for _row in _reader:
                                    try:
                                        _lat = float(_row.get('latitude','').strip())
                                        _lon = float(_row.get('longitude','').strip())
                                        if not(19<=_lat<=27 and 87<=_lon<=93): continue
                                        _addr_str = str(_row.get('address','')).strip()
                                        _toks = set(t for t in __import__('re').sub(r'[^a-z0-9]',' ',_addr_str.lower()).split() if len(t)>=4 and t not in _r4g_noise)
                                        if not _toks: continue
                                        _addr_list.append({
                                            'lat': _lat, 'lon': _lon, 'toks': _toks,
                                            'district': str(_row.get('district','')).strip().title(),
                                            'thana': str(_row.get('thana','')).strip().title(),
                                            'addr': _addr_str,
                                        })
                                    except Exception: pass
                    except Exception: pass
                    loaded_files[fn]['addr_list'] = _addr_list

            # Helper: detect generation from cell_type string
            def _gen_from_cell_type(ct):
                ct_l = str(ct).lower().strip()
                if ct_l in ("", "nan", "none", "-", "n/a"): return None
                if "4g" in ct_l or "lte" in ct_l:   return "4g"
                if "3g" in ct_l or "wcdma" in ct_l or "umts" in ct_l: return "3g"
                if "2g" in ct_l or "gsm" in ct_l:   return "2g"
                return None

            # Helper: detect operator key from operator string
            def _op_key(op_str):
                o = str(op_str).lower()
                if "grameen" in o or " gp" in o or o.startswith("gp"): return "gp"
                if "banglalink" in o or " bl" in o or o.startswith("bl"): return "bl"
                if "robi" in o or "airtel" in o: return "robi"
                if "teletalk" in o: return "teletalk"
                return None

            # ── Enrich df row by row using operator + generation ──
            if loaded_files and "cell_id" in df.columns and ("lac" in df.columns or "lac_id" in df.columns):
                lac_col = "lac" if "lac" in df.columns else "lac_id"
                lats_col, lons_col, methods_col, csv_lbl_col, dist_col = [], [], [], [], []

                # Debug: show loaded file stats
                for _fn, _fd in loaded_files.items():
                    _ex_cnt = len(_fd["exact"]); _mu_cnt = len(_fd["multi"])
                    if _ex_cnt == 0 and _mu_cnt == 0:
                        # Teletalk CSV-এ CGI/ECGI-based index থাকে — ('0',cgi) keys
                        # তাই normal (lac,cid) count = 0 হলেও Teletalk কাজ করে
                        if "teletalk" not in _fn.lower():
                            st.warning(f"⚠️ {_fn}: Loaded but 0 towers found — check CSV column names")

                # ── Pre-compute LAC cluster stats (Signal 3) ──
                _lac_cluster_cache = _build_lac_clusters(loaded_files)

                # ── Pass 1: collect raw GPS candidates for every row (no validation yet) ──
                # We need neighbor context → do a first pass to get candidate GPS per row
                _candidates = []   # list of (lat|None, lon|None, addr_str, dist_val) per row
                _row_lac    = []   # lac key per row (for Signal 3)
                _row_lacid  = []   # (lac, ci) string per row (for Signal 4)

                df_list = list(df.iterrows())
                for _, row in df_list:
                    lv  = _norm_id(row.get(lac_col, ""))
                    cv  = _norm_id(row.get("cell_id", ""))
                    k   = (lv, cv)
                    cdr_addr_toks = _addr_tokens(row.get("address", ""))
                    op_k  = _op_key(row.get("operator", "")) if "operator" in df.columns else None
                    gen_k = _gen_from_cell_type(row.get("cell_type", "")) if "cell_type" in df.columns else None

                    # Operator+Network strict: CDR-এর operator ও 2G/3G/4G অনুযায়ী
                    # শুধু সেই CSV-এ LAC+CID match করো — অন্য operator বা অন্য generation-এ যাবে না
                    fnames_to_try = []
                    if op_k and gen_k:
                        # Priority 1: exact operator+generation match
                        primary = OP_GEN_FILE.get((op_k, gen_k))
                        if primary and primary in loaded_files:
                            fnames_to_try.append(primary)
                        # Priority 2: same operator, other generations
                        for g in ["4g", "3g", "2g"]:
                            if g != gen_k:
                                fb = OP_GEN_FILE.get((op_k, g))
                                if fb and fb in loaded_files and fb not in fnames_to_try:
                                    fnames_to_try.append(fb)
                    elif op_k and not gen_k:
                        # Cell Type NaN/unknown → same operator-এর সব generation CSV try
                        for g in ["4g", "3g", "2g"]:
                            fb = OP_GEN_FILE.get((op_k, g))
                            if fb and fb in loaded_files and fb not in fnames_to_try:
                                fnames_to_try.append(fb)
                    # অন্য operator-এর CSV দেখবে না
                    if not fnames_to_try:
                        fnames_to_try = list(loaded_files.keys())

                    found_lat = None; found_lon = None
                    found_addr = ""; found_dist_val = ""
                    for fn in fnames_to_try:
                        fdata = loaded_files[fn]
                        exact = fdata["exact"]; multi = fdata["multi"]
                        if k in exact:
                            found_lat, found_lon, _, found_addr = exact[k]
                            found_dist_val = found_addr.split(",")[-1].strip() if "," in found_addr else found_addr[:20]
                            break
                        # CID+Address fallback: LAC mismatch কিন্তু CID same থাকলে
                        # GP 4G-তে একই tower একাধিক LAC-এ থাকতে পারে (TAC reassignment)
                        # CDR BTS address vs CSV address token similarity — score >= 2 হলেই accept
                        # score=1 হলে না নেওয়াই ভালো (Chandpur-style false positive এড়াতে)
                        elif cv in multi and cdr_addr_toks:
                            _best_m = None; _best_sc = 0; _best_addr = ""; _best_dv = ""
                            for _lt, _ln, _ctoks, _caddr in multi[cv]:
                                _sc = len(cdr_addr_toks & _ctoks) if cdr_addr_toks and _ctoks else 0
                                if _sc > _best_sc:
                                    _best_sc = _sc; _best_m = (_lt, _ln); _best_addr = _caddr
                                    _best_dv = _caddr.split(",")[-1].strip() if "," in _caddr else _caddr[:20]
                            if _best_m and _best_sc >= 2:
                                found_lat, found_lon = _best_m; found_addr = _best_addr; found_dist_val = _best_dv
                                break
                        # ── Robi 2G: cell_id-only fallback (LAC mismatch) ────────────
                        # Robi 2G CDR-এ কিছু records-এ LAC mismatch হয় —
                        # একই cell_id ভিন্ন LAC-এ store হয়।
                        # Fallback: cell_id দিয়ে match, address/district দিয়ে disambiguate।
                        elif cv in multi and not cdr_addr_toks and op_k == "robi" and gen_k in ("2g", "3g"):
                            # Address token নেই — শুধু first candidate নাও
                            _cands = multi[cv]
                            if len(_cands) == 1:
                                found_lat, found_lon, _, found_addr = _cands[0]
                                found_dist_val = found_addr.split(",")[-1].strip() if "," in found_addr else found_addr[:20]
                                break
                        elif cv in multi and op_k == "robi" and gen_k in ("2g", "3g"):
                            # Address আছে কিন্তু score < 2 — district match দিয়ে try করি
                            _cdr_addr_lo_r2 = str(row.get("address","")).lower()
                            _cands_r2 = multi[cv]
                            _pick_r2 = None
                            for _lt, _ln, _ctoks, _caddr in _cands_r2:
                                _caddr_lo = _caddr.lower()
                                if any(kw in _cdr_addr_lo_r2 for kw in ['dhaka','chittagong','sylhet','rajshahi','khulna','barisal','rangpur','mymensingh'] if kw in _caddr_lo):
                                    _pick_r2 = (_lt, _ln, _caddr)
                                    break
                            if not _pick_r2 and _cands_r2:
                                _pick_r2 = (_cands_r2[0][0], _cands_r2[0][1], _cands_r2[0][3])
                            if _pick_r2:
                                found_lat, found_lon, found_addr = _pick_r2
                                found_dist_val = found_addr.split(",")[-1].strip() if "," in found_addr else found_addr[:20]
                                break

                    # ── Teletalk special: CGI/ECGI direct match ──────────────────
                    # Teletalk CDR-এ Cell ID = full CGI/ECGI (e.g. 470040122737532)
                    # LAC/CID আলাদা নেই, তাই normal (lv,cv) match ব্যর্থ হয়।
                    # Fix: cell_id-কে CGI/ECGI হিসেবে ('0', cgi_val) key দিয়ে lookup করো।
                    if found_lat is None and op_k == "teletalk":
                        _tel_fname = "Teletalk.csv"
                        if _tel_fname in loaded_files:
                            _tel_exact = loaded_files[_tel_fname]["exact"]
                            # CDR Cell ID = CGI/ECGI value → key = ('0', cv)
                            _k_cgi = ('0', cv)
                            if _k_cgi in _tel_exact:
                                found_lat, found_lon, _, found_addr = _tel_exact[_k_cgi]
                                found_dist_val = found_addr.split(",")[-1].strip() if "," in found_addr else found_addr[:20]

                    # ── Robi 4G special: ENODEBID//100 = CDR LAC_ID ──────────────
                    # CDR LAC_ID = ENODEBID ÷ 100, CDR Cell ID শেষ ২ digit = CSV CELL_ID
                    # Ambiguity: একই (derived_lac, sector) অনেক tower → address token দিয়ে best নাও
                    if found_lat is None and op_k == "robi" and gen_k == "4g":
                        robi_4g_fname = OP_GEN_FILE.get(("robi","4g"))
                        if robi_4g_fname and robi_4g_fname in loaded_files:
                            _r4g_multi = loaded_files[robi_4g_fname].get("robi4g_multi", {})
                            try:
                                _cv_str = cv.lstrip("0") or "0"
                                if len(_cv_str) >= 2:
                                    _sector_cv = _cv_str[-2:]   # শেষ ২ digit
                                    _k_r4g = (lv, _sector_cv)
                                    _r4g_cands = _r4g_multi.get(_k_r4g, [])
                                    if _r4g_cands:
                                        if len(_r4g_cands) == 1:
                                            found_lat, found_lon, _, found_addr = _r4g_cands[0]
                                            found_dist_val = found_addr.split(",")[-1].strip() if "," in found_addr else found_addr[:20]
                                        elif cdr_addr_toks:
                                            # address token matching — best score জিতবে
                                            _best_r4g = None; _best_sc_r4g = -1
                                            for _lt, _ln, _ctoks, _caddr in _r4g_cands:
                                                _sc = len(cdr_addr_toks & _ctoks) if cdr_addr_toks and _ctoks else 0
                                                if _sc > _best_sc_r4g:
                                                    _best_sc_r4g = _sc; _best_r4g = (_lt, _ln, _caddr)
                                            if _best_r4g:
                                                found_lat, found_lon, found_addr = _best_r4g
                                                found_dist_val = found_addr.split(",")[-1].strip() if "," in found_addr else found_addr[:20]
                                        else:
                                            # address নেই → first entry
                                            found_lat, found_lon, _, found_addr = _r4g_cands[0]
                                            found_dist_val = found_addr.split(",")[-1].strip() if "," in found_addr else found_addr[:20]
                            except Exception:
                                logger.debug('Suppressed exception', exc_info=True)

                    # ── Robi 4G address token fallback (if still not found) ──
                    # CDR BTS address ↔ CSV address token similarity দিয়ে GPS নেওয়া হয়
                    if found_lat is None and op_k == "robi" and gen_k == "4g" and cdr_addr_toks:
                        robi_4g_fname = OP_GEN_FILE.get(("robi","4g"))
                        if robi_4g_fname and robi_4g_fname in loaded_files:
                            _r4g_data = loaded_files[robi_4g_fname]
                            _r4g_addr_list = _r4g_data.get("addr_list", [])
                            if _r4g_addr_list:
                                best_r4g = None; best_r4g_sc = 0
                                for _entry in _r4g_addr_list:
                                    sc = len(cdr_addr_toks & _entry["toks"])
                                    if sc > best_r4g_sc:
                                        best_r4g_sc = sc; best_r4g = _entry
                                if best_r4g and best_r4g_sc >= 3:
                                    found_lat = best_r4g["lat"]; found_lon = best_r4g["lon"]
                                    found_addr = best_r4g["addr"]; found_dist_val = best_r4g["district"]

                    _candidates.append((found_lat, found_lon, found_addr, found_dist_val))
                    _row_lac.append(lv)
                    _row_lacid.append((lv, cv))

                # ── Pre-build Signal 4 lookup: (lac,ci) → list of candidate GPS from Pass 1 ──
                _lacid_gps = {}  # (lv,cv) → [(lat,lon), ...]
                for i, (clat, clon, _, _) in enumerate(_candidates):
                    k4 = _row_lacid[i]
                    if clat is not None:
                        _lacid_gps.setdefault(k4, []).append((clat, clon))

                # ── Pass 2: validate each candidate using confidence score ──
                NEIGHBOR_WINDOW = 4  # look ±4 rows for neighbor context
                rejected_count = 0

                for i, (_, row) in enumerate(df_list):
                    clat, clon, caddr, cdist = _candidates[i]
                    if clat is None:
                        lats_col.append(None); lons_col.append(None)
                        methods_col.append("text_based"); csv_lbl_col.append(""); dist_col.append("")
                        continue

                    cdr_addr = row.get("address", "")
                    lv = _row_lac[i]
                    k4 = _row_lacid[i]

                    # Neighbor GPS: ±NEIGHBOR_WINDOW rows that have a GPS candidate
                    lo_i = max(0, i - NEIGHBOR_WINDOW)
                    hi_i = min(len(_candidates), i + NEIGHBOR_WINDOW + 1)
                    neighbor_gps = [(c[0], c[1]) for j, c in enumerate(_candidates[lo_i:hi_i], lo_i)
                                    if j != i and c[0] is not None]

                    # Signal 4: other rows with same LAC+CI
                    same_laci_others = [(lt, ln) for lt, ln in _lacid_gps.get(k4, [])
                                        if not (abs(lt - clat) < 1e-9 and abs(ln - clon) < 1e-9)]

                    conf, reasons = _gps_confidence(
                        clat, clon,
                        cdr_addr_str   = cdr_addr,
                        lac_key        = lv,
                        same_lac_ci_rows = same_laci_others,
                        neighbor_rows  = neighbor_gps,
                    )

                    # CID-only matches পেলে থ্রেশহোল্ড বাড়াই
                    effective_threshold = GPS_CONFIDENCE_THRESHOLD
                    if clat is not None and k4 not in loaded_files.get(fnames_to_try[0] if fnames_to_try else "", {}).get("exact", {}):
                        # This was a CID-only match — stricter threshold
                        effective_threshold = 55

                    if conf >= effective_threshold:
                        lats_col.append(clat); lons_col.append(clon)
                        methods_col.append("cell_exact")
                        csv_lbl_col.append(caddr); dist_col.append(cdist)
                        cell_match_count += 1
                    else:
                        # Low confidence → fall back to text_based
                        lats_col.append(None); lons_col.append(None)
                        methods_col.append("text_based"); csv_lbl_col.append(""); dist_col.append("")
                        rejected_count += 1

                df["cell_lat"]       = lats_col
                df["cell_lon"]       = lons_col
                df["loc_method"]     = methods_col
                df["cell_csv_label"] = csv_lbl_col
                df["csv_district"]   = dist_col

          except Exception as _cell_err:
              st.warning(f"⚠️ Cell tower GPS enrichment failed: {_cell_err}")

        gps_match_pct = round(cell_match_count / max(len(df), 1) * 100, 1)
        if not _from_cache:
            df_clean = cdf(df)
            progress.progress(35, text="📊 Generating report...")
            phone      = get_phone(df)
            operator   = get_operator(df)
            date_range = get_date_range(df)
            base_name  = os.path.splitext(uploaded.name)[0]
        else:
            df_clean = cdf(df)
            base_name  = _cached.get("base_name", "CDR")

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

        # ── GPS Accuracy Banner ──
        if cell_match_count > 0:
            st.markdown(f"""
            <div style="background:#ecfdf5; border-left:4px solid #10b981; border-radius:10px;
                        padding:0.75rem 1.25rem; margin-bottom:0.85rem; display:flex;
                        align-items:center; gap:0.75rem;">
                <div style="font-size:1.3rem;">📡</div>
                <div>
                    <strong style="color:#065f46;">GPS Cell Tower Match: {cell_match_count:,}/{len(df):,} ({gps_match_pct}%)</strong>
                    <div style="color:#047857; font-size:0.85rem; margin-top:0.1rem;">
                        Exact GPS coordinates matched from cell tower database · Accuracy: ±0.5–2 km
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



        html_content = build_html(df, phone, operator, date_range, total_raw, anomaly_count, target_number, target_location, profile_data=profile_data)
        # Inject Profile Analysis section after <h1>
        if profile_data and any(profile_data.get(k) for k in ['name','nid','passport','docs_found']):
            _prof_html = _profile_html_section(profile_data)
            if _prof_html:
                _marker = '<h1>📞 CDR Analysis Report</h1>'
                if _marker in html_content:
                    html_content = html_content.replace(
                        _marker, _marker + '\n' + _prof_html, 1)
        html_bytes = html_content.encode('utf-8')

        progress.progress(80, text="📝 Generating Word report...")
        docx_bytes = build_docx(df, phone, operator, date_range, total_raw, anomaly_count, target_number, target_location, profile_data=profile_data)

        # ── Movement Map ──
        progress.progress(90, text="🗺️ Generating movement map...")
        _mv_for_map = movement_pattern_analysis(df)
        map_bytes = build_movement_map(df, phone, operator, mv_data=_mv_for_map)
        progress.progress(100, text="✅ Complete!")

        # ── Save to session_state cache ──
        st.session_state[_cache_key] = {
            "html_bytes": html_bytes,
            "docx_bytes": docx_bytes,
            "map_bytes": map_bytes,
            "base_name": base_name,
            "df_pickle": _safe_pickle_dumps(df),
            "phone": phone,
            "operator": operator,
            "date_range": date_range,
            "total_raw": total_raw,
            "anomaly_count": anomaly_count,
            "target_number": target_number,
            "target_location": target_location,
            "sections": True,
        }

        # ── Success + Download ──
        st.markdown("""
        <div style="background:#ecfdf5; border:1px solid #10b981; border-radius:10px;
                    padding:1rem 1.5rem; margin:1rem 0; color:#065f46;">
            ✅ <strong>Reports generated successfully!</strong>
            Click the buttons below to download.
        </div>
        """, unsafe_allow_html=True)

        dl1, dl2, dl3 = st.columns(3)
        with dl1:
            st.download_button(
                label="⬇️ Download HTML Report",
                data=html_bytes,
                file_name=f"{base_name}_Report.html",
                mime="text/html",
                use_container_width=True,
                key="dl_html_main"
            )
        with dl2:
            st.download_button(
                label="⬇️ Download Word Report",
                data=docx_bytes,
                file_name=f"{base_name}_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key="dl_docx_main"
            )
        with dl3:
            if map_bytes:
                st.download_button(
                    label="🗺️ Download Movement Map",
                    data=map_bytes,
                    file_name=f"{base_name}_Movement_Map.html",
                    mime="text/html",
                    use_container_width=True,
                    key="dl_map_main"
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

        # ── Profile Analysis (Dashboard-এ শুধু docs badge) ──
        _dash_profile = st.session_state.get('_profile_data', None)
        if _dash_profile and _dash_profile.get('docs_found'):
            _docs = _dash_profile.get('docs_found', [])
            _SRC_COLORS = {
                'NID':                  ('#dbeafe', '#1e40af'),
                'Passport':             ('#fce7f3', '#9d174d'),
                'TIN':                  ('#fef9c3', '#854d0e'),
                'Driving License':      ('#dcfce7', '#166534'),
                'Vehicle Registration': ('#f3e8ff', '#6b21a8'),
                'Vehicle Reg':          ('#f3e8ff', '#6b21a8'),
            }
            _badges_html = ''.join(
                f"<span style='background:{_SRC_COLORS.get(d,('#f1f5f9','#475569'))[0]};"
                f"color:{_SRC_COLORS.get(d,('#f1f5f9','#475569'))[1]};"
                f"padding:3px 10px;border-radius:8px;font-size:12px;"
                f"font-weight:700;margin:3px 4px;display:inline-block'>{d}</span>"
                for d in _docs if d
            )
            _name_disp = (_dash_profile.get('name') or '').strip()
            st.markdown(
                f"<div style='background:#f8fafc;border:1px solid #e2e8f0;"
                f"border-radius:10px;padding:10px 14px;margin-bottom:8px'>"
                f"<span style='font-size:12px;color:#64748b;font-weight:600'>👤 Profile Documents: </span>"
                f"{_badges_html}"
                f"{'<span style=\"margin-left:12px;font-size:13px;font-weight:700;color:#0f172a\">'+_name_disp+'</span>' if _name_disp else ''}"
                f"</div>",
                unsafe_allow_html=True
            )

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
                with mv3:
                    _work_d = mv.get("work_district") or mv.get("frequent_districts", [None])[1] if len(mv.get("frequent_districts", [])) > 1 else None
                    st.markdown(f'''<div class="stat-card" style="border-left-color:#0891b2;">
                        <div class="label">Work Location</div>
                        <div class="value" style="font-size:1rem;">{_work_d or "N/A"}</div>
                        <div style="font-size:0.75rem;color:#94a3b8;">2nd most frequent</div>
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
                    trip_rows = []
                    for t in mv['trips']:
                        lat  = t.get('lat')
                        lon  = t.get('lon')
                        gps  = f"{lat}, {lon}" if lat and lon else "—"
                        dist = f"{t['km']} km" if t.get('km') else "—"
                        # Location label: upazila + district from CSV (cell_csv_label preferred)
                        loc_label = ""
                        if t.get('upazila') and str(t['upazila']).strip() not in ("", "nan"):
                            loc_label = str(t['upazila'])
                        if t.get('district') and str(t['district']).strip() not in ("", "nan", loc_label):
                            loc_label = (loc_label + ", " + str(t['district'])).strip(", ")
                        if not loc_label:
                            loc_label = t.get('district') or "—"
                        trip_rows.append({
                            'Upazila / District': loc_label,

                            'Distance (km)':      dist,
                            'Start Date':         t['start_date'],
                            'End Date':           t['end_date'],
                            'Days':               t['days'],
                            'Cell Location (CSV)': str(t.get('address', ''))[:100],
                        })
                    trip_df = pd.DataFrame(trip_rows)
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

        # ── 8. International / Foreign Voice Call Analysis ──
        with st.expander("🌐 International Voice Call Analysis", expanded=True):
            intl_ui = international_call_analysis(df)
            if intl_ui:
                ic1, ic2, ic3, ic4 = st.columns(4)
                with ic1:
                    st.markdown(f'''<div class="stat-card" style="border-left-color:#2563eb;">
                        <div class="label">Total Foreign Calls</div>
                        <div class="value">{intl_ui["total_calls"]}</div>
                        <div style="font-size:0.75rem;color:#94a3b8;">MOC + MTC only</div>
                    </div>''', unsafe_allow_html=True)
                with ic2:
                    st.markdown(f'''<div class="stat-card" style="border-left-color:#7c3aed;">
                        <div class="label">Countries</div>
                        <div class="value">{intl_ui["unique_countries"]}</div>
                    </div>''', unsafe_allow_html=True)
                with ic3:
                    st.markdown(f'''<div class="stat-card" style="border-left-color:#0891b2;">
                        <div class="label">Unique Numbers</div>
                        <div class="value">{intl_ui["unique_numbers"]}</div>
                    </div>''', unsafe_allow_html=True)
                with ic4:
                    st.markdown(f'''<div class="stat-card" style="border-left-color:#16a34a;">
                        <div class="label">Total Duration</div>
                        <div class="value" style="font-size:1.1rem;">{intl_ui["total_duration"]} min</div>
                    </div>''', unsafe_allow_html=True)
                st.markdown("")
                st.markdown("""<div style="background:#eff6ff; border-left:4px solid #2563eb;
                    border-radius:8px; padding:0.6rem 1rem; margin-bottom:0.75rem;
                    font-size:0.85rem; color:#1e40af;">
                    ℹ️ SMS বাদ — শুধুমাত্র Voice Call (MOC/MTC)। Raw E.164 format সাপোর্ট।
                </div>""", unsafe_allow_html=True)
                st.dataframe(intl_ui['table'], use_container_width=True, hide_index=True)
            else:
                st.success("✅ No foreign voice calls detected in this CDR.")

        # ── 9. Specific Number Analysis ──
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

        # ── 8. Target Location Analysis ──
        if target_location:
            with st.expander(f"📍 Target Location — {target_location}", expanded=True):
                loc_results = target_location_analysis(df, target_location)
                if not loc_results:
                    st.warning(f"No CDR activity found near '{target_location}'. Try a different spelling or nearby area name.")
                else:
                    total_rec = sum(r['Records'] for r in loc_results)
                    st.success(f"Found activity on **{len(loc_results)}** day(s) near **{target_location}** — {total_rec} total records.")
                    loc_df = pd.DataFrame(loc_results)
                    st.dataframe(loc_df, use_container_width=True, hide_index=True)


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
