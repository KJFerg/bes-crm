"""
Berkshire Executive Search — Contact CRM
=========================================
Streamlit app for browsing, searching, and editing contacts.
Reads/writes from Google Sheets (consolidated contact database).

Deploy to Streamlit Community Cloud for access from any device.

Ken Ferguson | Berkshire Executive Search | May 29, 2026
"""

import streamlit as st
import pandas as pd
import io
import re
import base64
import uuid
from datetime import datetime
from streamlit_gsheets import GSheetsConnection
from streamlit_paste_button import paste_image_button as paste_button

# Distribution tags available for contacts
AVAILABLE_TAGS = [
    # ── Audience / role ──
    "Big Data",
    "CEO President Mid",
    "HR",
    "HR Managers",
    "IT Execs for Newsletter",
    "IT Security",
    # ── Pipeline / action (from temp-sheet color coding) ──
    "Sent Strategy Link",
    "Sent Strategy Link (2 Options)",
    "Intro Summit/Parliament",
    "Sent CC Proposal",
    "PersonaWise Review",
    "Paid Consult",
    "Pro Bono Session",
    "30-Minute Free",
    "Looking",
    # ── Cohort / source ──
    "We-Connect Campaign A",
    "Inbound (outside We-Connect)",
]

# Source = how a contact came in (dropdown on the Add Contact form)
AVAILABLE_SOURCES = [
    "We-Connect Campaign A",
    "We-Connect Campaign B",
    "LinkedIn inbound (they messaged me)",
    "LinkedIn comment/post engagement",
    "Referral",
    "Website inquiry",
    "Strategy page booking",
    "Parliament / Summit (Mark Blanke)",
    "Past client / existing network",
    "Conference / event",
    "Search assignment (recruiting)",
]

# ══════════════════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="BES Contacts",
    page_icon="📇",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ══════════════════════════════════════════════════════════════════════════
# CUSTOM CSS — clean, mobile-friendly
# ══════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
    /* Hide Streamlit header bar and footer */
    header[data-testid="stHeader"] { display: none !important; }
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }

    /* Tighten padding for mobile */
    .block-container { padding-top: 1rem; padding-bottom: 0; }

    /* Contact card styling */
    .contact-card {
        border: 1px solid #ddd;
        border-radius: 8px;
        padding: 12px 16px;
        margin-bottom: 8px;
        background: white;
        transition: border-color 0.2s;
    }
    .contact-card:hover { border-color: #1a73e8; }

    .contact-name {
        font-size: 1.1rem;
        font-weight: 600;
        color: #1a1a1a;
        margin-bottom: 2px;
    }
    .contact-title {
        font-size: 0.9rem;
        color: #555;
        margin-bottom: 4px;
    }
    .contact-meta {
        font-size: 0.8rem;
        color: #888;
    }
    .tag-badge {
        display: inline-block;
        background: #e8f0fe;
        color: #1a73e8;
        padding: 2px 8px;
        border-radius: 12px;
        font-size: 0.75rem;
        margin-right: 4px;
        margin-top: 4px;
    }
    .source-badge {
        display: inline-block;
        background: #f0f0f0;
        color: #666;
        padding: 2px 6px;
        border-radius: 4px;
        font-size: 0.7rem;
        margin-right: 3px;
    }
    .stat-box {
        text-align: center;
        padding: 8px;
    }
    .stat-number {
        font-size: 1.5rem;
        font-weight: 700;
        color: #1a73e8;
    }
    .stat-label {
        font-size: 0.75rem;
        color: #888;
    }

    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    /* Mobile adjustments */
    @media (max-width: 768px) {
        .block-container { padding-left: 0.5rem; padding-right: 0.5rem; }
    }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
# DATA LOADING
# ══════════════════════════════════════════════════════════════════════════

@st.cache_data(ttl=300)  # Cache for 5 minutes
def load_data():
    """Load contacts from Google Sheets."""
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        df = conn.read(worksheet="Contacts", ttl=300)
        if df is None or df.empty:
            return pd.DataFrame()
        # Force text columns to string to prevent float64 inference on sparse columns
        text_cols = ["Notes", "Phone1", "Phone2", "Email1", "Email2", "Email3",
                     "FirstName", "LastName", "Positions", "Industry", "Location",
                     "City", "State", "Sources", "DistributionTags", "Education",
                     "LinkedInURL"]
        for col in text_cols:
            if col in df.columns:
                df[col] = df[col].astype(str).replace({"nan": "", "None": ""})
        # Fill remaining NaN with empty strings for display
        df = df.fillna("")
        return df
    except Exception as e:
        st.error(f"Could not connect to Google Sheets: {e}")
        return pd.DataFrame()


def _prep_df_for_write(df):
    """Convert ALL columns to string before writing back to Sheets.
    This prevents dtype errors when sparse columns are inferred as float64."""
    for col in df.columns:
        df[col] = df[col].astype(str).replace({"nan": "", "None": "", "NaN": ""})
    return df


# Columns that must be present in a healthy read. Used to abort writes (and to
# warn on load) if a partial/corrupt read would otherwise clobber good data.
REQUIRED_COLUMNS = ["FirstName", "LastName", "Email1", "Notes"]


def _read_for_write(conn):
    """Read the sheet for a write operation, aborting if the read looks corrupt
    or partial — protects against overwriting 40K good rows with a bad read."""
    df = conn.read(worksheet="Contacts", ttl=0)
    if df is None or df.empty:
        raise RuntimeError("Sheet read returned no data — write skipped to protect your data. Please try again.")
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise RuntimeError(f"Sheet read was incomplete (missing {missing}) — write skipped to protect your data. Please try again.")
    return df


@st.cache_resource
def _contacts_ws():
    """The Contacts worksheet via gspread — for fast TARGETED cell/row ops."""
    return _gspread_book().worksheet("Contacts")


@st.cache_data(ttl=600)
def _contacts_header():
    return _contacts_ws().row_values(1)


def _col_num(field_name):
    """1-based column number for a field, or None if absent."""
    header = _contacts_header()
    return header.index(field_name) + 1 if field_name in header else None


# row_index is the 0-based position from load_data(); sheet row = row_index + 2
# (row 1 = header). All writes below are TARGETED single-cell/row gspread ops, so
# we never read or rewrite the whole 40K sheet — fast, and can't clobber the data.

def save_field(row_index, field_name, new_value):
    """Update a single cell for a contact."""
    try:
        ws = _contacts_ws()
        col = _col_num(field_name)
        if col is None:  # field not yet a column — add it at the end
            col = len(_contacts_header()) + 1
            ws.update_cell(1, col, field_name)
            _contacts_header.clear()
        ws.update_cell(row_index + 2, col, str(new_value))
        if "df" in st.session_state and row_index in st.session_state.df.index:
            if field_name not in st.session_state.df.columns:
                st.session_state.df[field_name] = ""
            st.session_state.df.at[row_index, field_name] = str(new_value)
        return True
    except Exception as e:
        st.error(f"Could not save {field_name}: {e}")
        return False


def save_note(row_index, new_note):
    """Append a note via a single-cell read + write (no whole-sheet rewrite)."""
    try:
        ws = _contacts_ws()
        col = _col_num("Notes")
        rownum = row_index + 2
        existing = ws.cell(rownum, col).value or ""
        # Idempotency: if the most recent note is the same text, treat as already
        # saved (prevents accidental double-entry from a missed visual confirm).
        if existing:
            last = existing.split(" || ")[-1].strip()
            last_text = last.split("]", 1)[1].strip() if last.startswith("[") and "]" in last else last
            if last_text == new_note.strip():
                if "df" in st.session_state and row_index in st.session_state.df.index:
                    st.session_state.df.at[row_index, "Notes"] = existing
                return True
        entry = f"[{datetime.now().strftime('%Y-%m-%d')}] {new_note}"
        combined = f"{existing} || {entry}" if existing else entry
        ws.update_cell(rownum, col, combined)
        if "df" in st.session_state and row_index in st.session_state.df.index:
            st.session_state.df.at[row_index, "Notes"] = combined
        return True
    except Exception as e:
        st.error(f"Could not save note: {e}")
        return False


def delete_contact(row_index):
    """Delete a single row via gspread, then reload (row numbers shift)."""
    try:
        ws = _contacts_ws()
        ws.delete_rows(row_index + 2)
        st.cache_data.clear()
        st.session_state.df = load_data()
        return True
    except Exception as e:
        st.error(f"Could not delete contact: {e}")
        return False


def add_contact(new_data):
    """Append a new contact row via gspread, then reload."""
    try:
        ws = _contacts_ws()
        header = list(_contacts_header())
        added = False
        for k in new_data:
            if k not in header:
                header.append(k)
                ws.update_cell(1, len(header), k)
                added = True
        if added:
            _contacts_header.clear()
        ws.append_row([str(new_data.get(c, "")) for c in header], value_input_option="RAW")
        st.cache_data.clear()
        st.session_state.df = load_data()
        return True
    except Exception as e:
        st.error(f"Could not add contact: {e}")
        return False


# ══════════════════════════════════════════════════════════════════════════
# WE-CONNECT ENRICH — fill the blank Positions/Education/Location/City/State/
# Industry fields on EXISTING CRM records, matched by LinkedIn slug, sourced live
# from the We-Connect API (mirrors enrich_new_contacts.py — no CSV export needed).
# Fills blanks ONLY; never overwrites your data, your note, or your photo. Also
# reports who's accepted but isn't in the CRM yet (so you add them by hand with a
# photo + note). API key lives in st.secrets['weconnect']['api_key'] (NOT in repo).
# ══════════════════════════════════════════════════════════════════════════

WECONNECT_BASE = "https://api-us-1.we-connect.io"
_IT_EXEC_KW = (
    "cio", "cto", "ciso", "cdo", "chief information", "chief technology",
    "chief digital", "chief data", "chief ai", "vp of it", "vp information",
    "vice president, information", "head of it", "head of technology",
    "director of it", "it director", "information technology",
)


def _slug_from_url(url):
    """Lowercase LinkedIn slug from a profile URL (or a bare slug)."""
    if not url:
        return ""
    s = str(url).strip()
    m = re.search(r"/in/([^/?#]+)", s)
    slug = m.group(1) if m else s
    return slug.strip().lower().strip("/")


def _wc_get_connections(api_key, page):
    """One page of the We-Connect connections endpoint. Returns a list."""
    import requests
    r = requests.get(
        f"{WECONNECT_BASE}/api/v1/connections",
        params={"api_key": api_key, "page": page},
        headers={"Accept": "application/json"},
        timeout=30,
    )
    r.raise_for_status()
    body = r.json()
    if isinstance(body, dict):
        return body.get("data") or body.get("connections") or []
    return body or []


def _wc_enrich_fields(c):
    """Extract the six enrich-only fields from a We-Connect connection object.
    Returns (slug, {Positions, Education, Location, City, State, Industry})."""
    slug = (c.get("linkedin") or _slug_from_url(c.get("linkedin_profile_url", ""))).strip().lower()
    title = (c.get("title") or "").strip()
    company = (c.get("company") or "").strip()
    exp = c.get("experience") or []
    if isinstance(exp, list) and exp:
        roles = []
        for e in exp:
            if isinstance(e, dict) and (e.get("title") or e.get("name")):
                roles.append(f"{(e.get('title') or '').strip()} at {(e.get('name') or '').strip()}".strip(" at "))
        positions = "; ".join(roles)
    else:
        positions = f"{title} at {company}" if (title and company) else (title or company)
    location = (c.get("location") or "").strip()
    parts = [p.strip() for p in location.split(",")] if location else []
    return slug, {
        "Positions": positions,
        "Education": (c.get("education") or "").strip(),
        "Location": location,
        "City": parts[0] if len(parts) >= 1 else "",
        "State": parts[1] if len(parts) >= 2 else "",
        "Industry": (c.get("industry") or "").strip(),
    }


def enrich_from_weconnect(max_pages=400, recent_days=30):
    """Fill ONLY the blank Positions/Education/Location/City/State/Industry fields on
    EXISTING CRM records, matched by LinkedIn slug — mirrors enrich_new_contacts.py,
    sourced live from the We-Connect API. Never touches notes, photo, email, phone, or
    any field already filled. The 'missing' list is limited to people who connected in
    the last `recent_days` days (the API returns ALL connections, so without this the
    list floods with old/deleted ones). Returns (enriched_count, cells_filled, missing)."""
    import gspread, time
    cutoff = time.time() - recent_days * 86400
    FIELDS = ["Positions", "Education", "Location", "City", "State", "Industry"]
    try:
        api_key = st.secrets["weconnect"]["api_key"]
    except Exception:
        api_key = ""
    if not api_key:
        st.error('We-Connect API key not found. Add it to Streamlit secrets as:\n\n[weconnect]\napi_key = "your-key"')
        return 0, 0, []
    ws = _contacts_ws()
    vals = ws.get_all_values()
    header = vals[0]
    if "LinkedInURL" not in header:
        st.error("No LinkedInURL column in the sheet.")
        return 0, 0, []
    iURL = header.index("LinkedInURL")
    for f in FIELDS:  # ensure the enrich columns exist
        if f not in header:
            header.append(f)
            ws.update_cell(1, len(header), f)
    header = ws.row_values(1)
    col_idx = {f: header.index(f) for f in FIELDS}
    slug_rows = {}
    for i, row in enumerate(vals[1:], start=2):
        g = _slug_from_url(row[iURL]) if len(row) > iURL else ""
        if g:
            slug_rows.setdefault(g, i)

    def _cell(rownum, idx):
        r = vals[rownum - 1] if 0 <= rownum - 1 < len(vals) else []
        return r[idx].strip() if len(r) > idx else ""

    # Enrich based on what the CRM NEEDS, not on connection recency. Build the set of
    # CRM records that have a slug AND at least one blank field, then scan We-Connect
    # until every one of them is filled (stop early when the need-set empties). This
    # catches reconnects/backlog adds wherever they sit in the (unsortable) list.
    need = {}  # slug -> rownum, only records with a blank in the 6 fields
    for _slug, _rn in slug_rows.items():
        if any(not _cell(_rn, col_idx[f]) for f in FIELDS):
            need[_slug] = _rn
    need_total = len(need)
    st.session_state["wc_debug"] = {"crm_records_needing_enrichment": need_total}

    batch, enriched, missing = [], set(), []
    page, sampled = 1, False
    while page <= max_pages:
        try:
            rows = _wc_get_connections(api_key, page)
        except Exception as e:
            st.error(f"We-Connect API error on page {page}: {e}")
            break
        if not rows:
            break
        if not sampled:
            st.session_state["wc_debug"] = {"crm_records_needing_enrichment": need_total,
                                            "page1_row_count": len(rows), "sample": rows[:2]}
            sampled = True
        for c in rows:
            slug, fields = _wc_enrich_fields(c)
            rownum = need.get(slug)
            if rownum:  # a CRM record that needs filling
                for f in FIELDS:
                    val = (fields.get(f) or "").strip()
                    if val and not _cell(rownum, col_idx[f]):
                        batch.append({"range": gspread.utils.rowcol_to_a1(rownum, col_idx[f] + 1), "values": [[val]]})
                        enriched.add(rownum)
                need.pop(slug, None)  # filled — stop tracking
            elif slug and slug not in slug_rows:  # not in CRM at all
                ts = c.get("timestamp_connected_at")
                try:
                    is_recent = bool(ts) and float(ts) >= cutoff
                except (TypeError, ValueError):
                    is_recent = False
                if is_recent:  # only flag RECENT acceptances as "to add"
                    nm = (c.get("name") or f"{c.get('first_name','')} {c.get('last_name','')}").strip()
                    url = (c.get("linkedin_profile_url") or (f"https://www.linkedin.com/in/{slug}/" if slug else "")).strip()
                    missing.append(f"{nm or slug or '(unknown)'} — {c.get('connected_at','')} — {url}")
        page += 1
        if not need:  # every blank CRM record has been filled — done early
            break
    st.session_state["wc_debug"]["records_filled"] = len(enriched)
    st.session_state["wc_debug"]["still_unmatched_in_weconnect"] = len(need)
    if batch:
        try:
            for k in range(0, len(batch), 200):
                ws.batch_update(batch[k:k + 200], value_input_option="USER_ENTERED")
                if k + 200 < len(batch):
                    time.sleep(1.1)  # stay under the Sheets 60-writes/min quota
            st.cache_data.clear()
            st.session_state.df = load_data()
        except Exception as e:
            st.error(f"Pulled enrichment but could not write it: {e}")
            return 0, 0, missing
    return len(enriched), len(batch), missing


# ══════════════════════════════════════════════════════════════════════════
# PHOTOS — stored in a separate "Photos" worksheet (PhotoKey | Name | Base64)
# to keep the big 40K Contacts sheet lean. Accessed directly via gspread so we
# fetch one photo at a time instead of loading thousands of base64 strings.
# ══════════════════════════════════════════════════════════════════════════

@st.cache_resource
def _gspread_book():
    """Open the spreadsheet with gspread using the same service-account creds
    st-gsheets uses (stored in st.secrets['connections']['gsheets'])."""
    import gspread
    from google.oauth2.service_account import Credentials
    cfg = dict(st.secrets["connections"]["gsheets"])
    spreadsheet = cfg.get("spreadsheet", "")
    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(cfg, scopes=scopes)
    gc = gspread.authorize(creds)
    if str(spreadsheet).startswith("http"):
        return gc.open_by_url(spreadsheet)
    return gc.open_by_key(spreadsheet)


def _photos_ws(create=False):
    book = _gspread_book()
    try:
        return book.worksheet("Photos")
    except Exception:
        if not create:
            return None
        ws = book.add_worksheet(title="Photos", rows=100, cols=3)
        ws.update("A1:C1", [["PhotoKey", "Name", "Base64"]])
        return ws


def photo_key_for_row(row):
    """A contact's photo key: explicit PhotoKey, else CrelateId."""
    for col in ("PhotoKey", "CrelateId"):
        v = str(row.get(col, "")).strip()
        if v and v not in ("nan", "None"):
            return v
    return ""


@st.cache_data(ttl=600)
def get_photo_b64(key):
    """Fetch one contact's base64 photo from the Photos tab (cached)."""
    if not key:
        return None
    try:
        ws = _photos_ws()
        if ws is None:
            return None
        cell = ws.find(str(key), in_column=1)
        if not cell:
            return None
        return ws.cell(cell.row, 3).value
    except Exception:
        return None


def pil_thumb_b64(im, max_dim=200, quality=72):
    """Resize a PIL image to a small JPEG thumbnail, return base64."""
    im = im.convert("RGB")
    im.thumbnail((max_dim, max_dim))
    buf = io.BytesIO()
    im.save(buf, format="JPEG", quality=quality)
    return base64.b64encode(buf.getvalue()).decode("ascii")


def set_photo(key, b64, name=""):
    """Upsert a photo row in the Photos tab keyed by PhotoKey."""
    try:
        ws = _photos_ws(create=True)
        cell = ws.find(str(key), in_column=1)
        if cell:
            ws.update_cell(cell.row, 3, b64)
            if name:
                ws.update_cell(cell.row, 2, name)
        else:
            ws.append_row([str(key), name, b64], value_input_option="RAW")
        get_photo_b64.clear()
        return True
    except Exception as e:
        st.error(f"Could not save photo: {e}")
        return False


def parse_current_position(positions_str):
    """Extract the most recent position for display."""
    if not positions_str:
        return "", ""
    # Last position is most recent (We-Connect or LinkedIn latest)
    parts = str(positions_str).split(" | ")
    latest = parts[-1] if parts else ""
    # Strip any trailing source/date tag like "(Crelate)" or "(Updated 2026-06-08)"
    latest = latest.strip()
    if latest.endswith(")") and " (" in latest:
        latest = latest[:latest.rfind(" (")].strip()
    if " @ " in latest:
        title, company = latest.split(" @ ", 1)
        return title.strip(), company.strip()
    return latest.strip(), ""


_SEARCH_FIELDS = ["FirstName", "LastName", "Positions", "Email1", "Email2",
                  "Notes", "Industry", "Location", "City", "State",
                  "DistributionTags", "Education"]


def search_contacts(df, query):
    """Token-based search: every word in the query must appear somewhere across
    the searchable fields. So 'michael morgan' matches a row with FirstName
    Michael AND LastName Morgan, and 'morgan new york' narrows further."""
    tokens = [t for t in str(query).lower().split() if t]
    if not tokens:
        return df
    present = [f for f in _SEARCH_FIELDS if f in df.columns]
    blob = df[present[0]].astype(str).str.lower()
    for f in present[1:]:
        blob = blob.str.cat(df[f].astype(str).str.lower(), sep=" ")
    mask = pd.Series(True, index=df.index)
    for tok in tokens:
        mask &= blob.str.contains(tok, na=False, regex=False)
    return df[mask]


def _clean(v):
    v = str(v).strip()
    return "" if v in ("nan", "None") else v


def _union_list(a, b, sep):
    out = []
    for x in [p.strip() for p in (a.split(sep) + b.split(sep)) if p.strip()]:
        if x not in out:
            out.append(x)
    return out


def merge_records(keeper_idx, dup_idx):
    """Merge the duplicate row into the keeper row (union of data), then delete
    the duplicate. Keeper survives with combined notes/tags/emails; the dup's
    positions are appended so its (newer) job becomes the current one."""
    try:
        df = st.session_state.df
        k, d = df.loc[keeper_idx], df.loc[dup_idx]
        updates = {}
        updates["Notes"] = " || ".join(_union_list(_clean(k.get("Notes", "")), _clean(d.get("Notes", "")), " || "))
        updates["DistributionTags"] = "; ".join(_union_list(_clean(k.get("DistributionTags", "")), _clean(d.get("DistributionTags", "")), ";"))
        # KEEPER is authoritative for the current role: its positions stay last
        # (= current). The duplicate's non-overlapping positions go in front as
        # older history, so an OLD record can't overwrite the current job.
        kp = [p.strip() for p in _clean(k.get("Positions", "")).split(" | ") if p.strip()]
        dp = [p.strip() for p in _clean(d.get("Positions", "")).split(" | ") if p.strip()]
        updates["Positions"] = " | ".join([p for p in dp if p not in kp] + kp)
        updates["Education"] = " | ".join(_union_list(_clean(k.get("Education", "")), _clean(d.get("Education", "")), " | "))
        # emails (3 slots) and phones (2 slots): unique, keeper first
        emails = []
        for col in ["Email1", "Email2", "Email3"]:
            for r in (k, d):
                e = _clean(r.get(col, ""))
                if e and e.lower() not in [x.lower() for x in emails]:
                    emails.append(e)
        for i, col in enumerate(["Email1", "Email2", "Email3"]):
            updates[col] = emails[i] if i < len(emails) else ""
        phones = []
        for col in ["Phone1", "Phone2"]:
            for r in (k, d):
                p = _clean(r.get(col, ""))
                if p and p not in phones:
                    phones.append(p)
        for i, col in enumerate(["Phone1", "Phone2"]):
            updates[col] = phones[i] if i < len(phones) else ""
        # single-value fields: keep keeper's if present, else take dup's
        for col in ["LinkedInURL", "City", "State", "Location", "Industry",
                    "Website", "Twitter", "Sources", "SourceDetail"]:
            updates[col] = _clean(k.get(col, "")) or _clean(d.get(col, ""))
        # PHOTO: the keeper's photo wins. photo_key_for_row uses PhotoKey, then
        # CrelateId. Keep the keeper's; only fall back to the duplicate's photo if
        # the keeper has no photo at all.
        k_pk, k_cid = _clean(k.get("PhotoKey", "")), _clean(k.get("CrelateId", ""))
        updates["CrelateId"] = k_cid or _clean(d.get("CrelateId", ""))
        if k_pk:
            updates["PhotoKey"] = k_pk                 # keeper has an explicit photo key
        elif k_cid:
            updates["PhotoKey"] = ""                   # keeper shows photo via its CrelateId
        else:
            updates["PhotoKey"] = _clean(d.get("PhotoKey", ""))   # keeper had none → take dup's
        # write merged fields to keeper, then delete the duplicate row
        for col, val in updates.items():
            save_field(keeper_idx, col, val)
        delete_contact(dup_idx)
        return True
    except Exception as e:
        st.error(f"Merge failed: {e}")
        return False


# ══════════════════════════════════════════════════════════════════════════
# LEADS — pipeline view (at-a-glance grid of everyone with an action tag)
# ══════════════════════════════════════════════════════════════════════════
PIPELINE_TAGS = [
    "Looking",
    "Sent Strategy Link",
    "Sent Strategy Link (2 Options)",
    "30-Minute Free",
    "Intro Summit/Parliament",
    "PersonaWise Review",
    "Paid Consult",
    "Sent CC Proposal",
    "Pro Bono Session",
]


def _last_note_date(notes):
    dates = re.findall(r"\[(\d{4}-\d{2}-\d{2})\]", str(notes))
    return max(dates) if dates else ""


def build_leads(df, tag_filter="All pipeline stages"):
    rows = []
    for _idx, r in df.iterrows():
        tags = [t.strip() for t in str(r.get("DistributionTags", "")).split(";") if t.strip()]
        pipe = [t for t in tags if t in PIPELINE_TAGS]
        if not pipe:
            continue
        if tag_filter != "All pipeline stages" and tag_filter not in pipe:
            continue
        title, company = parse_current_position(r.get("Positions", ""))
        name = f"{r.get('FirstName','')} {r.get('LastName','')}".strip()
        loc = ", ".join([x for x in [str(r.get("City", "")).strip(), str(r.get("State", "")).strip()]
                         if x and x not in ("nan", "None")])
        url = str(r.get("LinkedInURL", ""))
        rows.append({
            "Name": name,
            "Title": title,
            "Company": company,
            "Status": ", ".join(pipe),
            "Last Note": _last_note_date(r.get("Notes", "")),
            "Location": loc,
            "LinkedIn": url if url not in ("nan", "None") else "",
        })
    leads = pd.DataFrame(rows)
    if not leads.empty:
        leads = leads.sort_values("Last Note", ascending=False).reset_index(drop=True)
    return leads


# ══════════════════════════════════════════════════════════════════════════
# MAIN APP
# ══════════════════════════════════════════════════════════════════════════

# Load data
# Held in session so targeted single-cell edits don't trigger a 40K-row reload.
if "df" not in st.session_state:
    st.session_state.df = load_data()
df = st.session_state.df

# Guard: a momentary Google Sheets glitch can return rows without the expected
# columns. Show a friendly Reload instead of a raw KeyError crash.
if not df.empty and any(c not in df.columns for c in REQUIRED_COLUMNS):
    st.warning("The contact data didn't load completely — a momentary Google Sheets glitch.")
    if st.button("🔄 Reload"):
        st.cache_data.clear()
        st.session_state.df = load_data()
        st.rerun()
    st.stop()

if df.empty:
    st.warning("No contacts loaded. Check your Google Sheets connection in `.streamlit/secrets.toml`.")
    st.info("""
    **Setup steps:**
    1. Upload `consolidated_contacts.csv` to a Google Sheet named "BES Contacts"
    2. Rename the sheet tab to "Contacts"
    3. Create a Google Cloud service account and share the sheet with it
    4. Add credentials to `.streamlit/secrets.toml`

    See `DEPLOYMENT.md` for full instructions.
    """)
    st.stop()

# ── Header with stats ──
col_title, col_stats = st.columns([2, 3])
with col_title:
    st.markdown("## 📇 BES Contacts")

with col_stats:
    s1, s2, s3, s4 = st.columns(4)
    with s1:
        st.markdown(f'<div class="stat-box"><div class="stat-number">{len(df):,}</div><div class="stat-label">Total</div></div>', unsafe_allow_html=True)
    with s2:
        has_email = (df["Email1"] != "").sum()
        st.markdown(f'<div class="stat-box"><div class="stat-number">{has_email:,}</div><div class="stat-label">With Email</div></div>', unsafe_allow_html=True)
    with s3:
        has_resume = (df.get("HasResume", pd.Series(dtype=str)) == "Yes").sum()
        st.markdown(f'<div class="stat-box"><div class="stat-number">{has_resume:,}</div><div class="stat-label">Resumes</div></div>', unsafe_allow_html=True)
    with s4:
        has_notes = (df["Notes"] != "").sum()
        st.markdown(f'<div class="stat-box"><div class="stat-number">{has_notes:,}</div><div class="stat-label">With Notes</div></div>', unsafe_allow_html=True)

st.divider()

# ── Add Contact + Enrich open as pop-up dialogs from the compact top row ──
@st.dialog("➕ Add Contact", width="large")
def _add_contact_dialog():
    st.caption("Photo (optional): copy from LinkedIn (right-click → Copy image), then Paste.")
    _add_paste = paste_button("📋 Paste photo", key="add_paste_dlg")
    if _add_paste.image_data is not None:
        st.session_state["add_photo_img"] = _add_paste.image_data
    if st.session_state.get("add_photo_img") is not None:
        st.image(st.session_state["add_photo_img"], width=96, caption="Will attach to the new contact")
    with st.form("add_contact_form_dlg", clear_on_submit=False):
        ac1, ac2 = st.columns(2)
        with ac1:
            f_first = st.text_input("First name")
            f_last = st.text_input("Last name")
            f_url = st.text_input("LinkedIn URL")
            f_pos = st.text_input("Title / Company", placeholder="e.g. CIO at Acme")
            f_city = st.text_input("City")
            f_state = st.text_input("State")
        with ac2:
            f_email = st.text_input("Email")
            f_phone = st.text_input("Phone")
            f_source = st.selectbox("Source", AVAILABLE_SOURCES, index=0)
            f_source_custom = st.text_input("Custom source", placeholder="only if not in the list above")
            f_tags = st.multiselect("Tags", options=AVAILABLE_TAGS)
            f_note = st.text_area("Note", placeholder="What to remember about this person / interaction")
        submitted = st.form_submit_button("Add Contact", type="primary")
        if submitted:
            if not (f_first.strip() or f_last.strip() or f_url.strip()):
                st.warning("Enter at least a name or a LinkedIn URL.")
            else:
                _today = datetime.now().strftime("%Y-%m-%d")
                _add_img = st.session_state.get("add_photo_img")
                _photo_key = uuid.uuid4().hex if _add_img is not None else ""
                _new = {
                    "FirstName": f_first.strip(), "LastName": f_last.strip(),
                    "LinkedInURL": f_url.strip(), "Positions": f_pos.strip(),
                    "City": f_city.strip(), "State": f_state.strip(),
                    "Email1": f_email.strip(), "Phone1": f_phone.strip(),
                    "Sources": (f_source_custom.strip() or f_source),
                    "SourceDetail": (f_source_custom.strip() or f_source),
                    "DistributionTags": "; ".join(f_tags),
                    "Notes": (f"[{_today}] {f_note.strip()}" if f_note.strip() else ""),
                    "CreatedDate": _today,
                    "PhotoKey": _photo_key,
                }
                if add_contact(_new):
                    if _add_img is not None:
                        set_photo(_photo_key, pil_thumb_b64(_add_img), f"{f_first} {f_last}".strip())
                    st.session_state.pop("add_photo_img", None)
                    st.session_state["add_done_msg"] = f"Added {f_first} {f_last}.".strip()
                    st.rerun()


@st.dialog("🔄 Enrich from We-Connect", width="large")
def _enrich_dialog():
    st.caption(
        "Fills the blank Title/Work-history, Education, Location, City, State, and "
        "Industry fields on contacts you've already added — matched by LinkedIn slug, "
        "pulled live from We-Connect (no export needed). Fills blanks only; never "
        "touches your notes, photo, or anything you typed. Also lists anyone who's "
        "accepted but isn't in the CRM yet."
    )
    if st.button("Enrich now", key="wc_enrich_btn", type="primary"):
        with st.spinner("Pulling from We-Connect and filling blanks…"):
            st.session_state["wc_result"] = enrich_from_weconnect()
    res = st.session_state.get("wc_result")
    if res:
        _e, _c, _miss = res
        st.success(f"Enriched {_e} record(s) · {_c} field(s) filled.")
        if _miss:
            st.warning(f"{len(_miss)} recently-accepted connection(s) not yet in your CRM — add by hand (photo + note):")
            st.download_button(
                "📥 Download this list (CSV)",
                "Name,ConnectedAt,LinkedInURL\n" + "\n".join('"' + m.replace(" — ", '","') + '"' for m in _miss),
                f"to_add_{datetime.now().strftime('%Y%m%d')}.csv", "text/csv", key="wc_miss_dl",
            )
            for m in _miss[:40]:
                st.write(f"- {m}")
            if len(_miss) > 40:
                st.caption(f"…and {len(_miss) - 40} more (all in the CSV).")
        else:
            st.info("No recent acceptances missing from the CRM. 🎉")
    if st.checkbox("🐞 Show raw We-Connect data (debug)", key="wc_dbg"):
        _dbg = st.session_state.get("wc_debug")
        if _dbg is None:
            st.write("Click **Enrich now** first, then check this box.")
        else:
            st.json(_dbg)


# ── Compact top row: view toggle + action buttons (search becomes the first full line) ──
_tc1, _tc2, _tc3 = st.columns([2.2, 1, 1.7])
with _tc1:
    view = st.radio("View", ["📇 Contacts", "📋 Leads"], horizontal=True,
                    label_visibility="collapsed", key="view_mode")
with _tc2:
    if st.button("➕ Add Contact", use_container_width=True):
        _add_contact_dialog()
with _tc3:
    if st.button("🔄 Enrich from We-Connect", use_container_width=True):
        _enrich_dialog()

if st.session_state.get("add_done_msg"):
    st.success(st.session_state.pop("add_done_msg"))

if view == "📋 Leads":
    f1, f2 = st.columns([2, 1])
    with f1:
        lead_tag = st.selectbox("Pipeline stage", ["All pipeline stages"] + PIPELINE_TAGS,
                                label_visibility="collapsed")
    leads_df = build_leads(df, lead_tag)
    with f2:
        st.markdown(f"<div style='text-align:right;padding-top:6px;'><b>{len(leads_df):,}</b> leads</div>",
                    unsafe_allow_html=True)
    if leads_df.empty:
        st.info("No leads with a pipeline tag yet. Tag a contact (Looking, Sent Strategy Link, "
                "Pro Bono Session, etc.) in the Contacts view and they'll show up here.")
    else:
        st.dataframe(
            leads_df, use_container_width=True, hide_index=True, height=560,
            column_config={
                "LinkedIn": st.column_config.LinkColumn("LinkedIn", display_text="Profile"),
            },
        )
        st.download_button(
            "📥 Download leads CSV",
            leads_df.to_csv(index=False).encode("utf-8"),
            f"bes_leads_{datetime.now().strftime('%Y%m%d')}.csv",
            "text/csv",
        )
    if st.button("🔄 Reload data", key="leads_reload"):
        st.cache_data.clear()
        st.session_state.df = load_data()
        st.rerun()
    st.stop()

# (Add Contact and Enrich now open as pop-up dialogs from the top button row above.)


# ── Search and Filters ──
search_col, filter_col1, filter_col2 = st.columns([3, 1.5, 1.5])

with search_col:
    query = st.text_input("🔍 Search contacts", placeholder="Name, company, title, email, notes...", label_visibility="collapsed")

with filter_col1:
    # Tag filter
    all_tags = set()
    for tags in df["DistributionTags"]:
        if tags:
            for t in str(tags).split("; "):
                if t.strip():
                    all_tags.add(t.strip())
    tag_filter = st.selectbox("Tag", ["All Tags"] + sorted(all_tags), label_visibility="collapsed")

with filter_col2:
    # Source filter
    all_sources = set()
    for sources in df["Sources"]:
        if sources:
            for s in str(sources).split("; "):
                if s.strip():
                    all_sources.add(s.strip())
    source_filter = st.selectbox("Source", ["All Sources"] + sorted(all_sources), label_visibility="collapsed")

# Apply filters
filtered = search_contacts(df, query)

if tag_filter != "All Tags":
    filtered = filtered[filtered["DistributionTags"].str.contains(tag_filter, na=False)]

if source_filter != "All Sources":
    filtered = filtered[filtered["Sources"].str.contains(source_filter, na=False)]

# Sort options
sort_col, count_col = st.columns([2, 1])
with sort_col:
    sort_by = st.selectbox(
        "Sort",
        ["Last Name A→Z", "Last Name Z→A", "Most Recent Activity", "Most Emails", "Newest Added"],
        label_visibility="collapsed"
    )
with count_col:
    st.markdown(f"**{len(filtered):,}** contacts")

# Apply sort
if sort_by == "Last Name A→Z":
    filtered = filtered.sort_values("LastName", ascending=True)
elif sort_by == "Last Name Z→A":
    filtered = filtered.sort_values("LastName", ascending=False)
elif sort_by == "Most Recent Activity":
    filtered = filtered.sort_values("LastActivityDate", ascending=False)
elif sort_by == "Most Emails":
    filtered["_email_count_sort"] = pd.to_numeric(filtered["EmailCount"], errors="coerce").fillna(0)
    filtered = filtered.sort_values("_email_count_sort", ascending=False)
    filtered = filtered.drop(columns=["_email_count_sort"])
elif sort_by == "Newest Added":
    filtered = filtered.sort_values("CreatedDate", ascending=False)

# ── Pagination ──
CONTACTS_PER_PAGE = 25
total_pages = max(1, (len(filtered) - 1) // CONTACTS_PER_PAGE + 1)

if "page" not in st.session_state:
    st.session_state.page = 1
# Reset to page 1 when search/filter changes
if "last_query" not in st.session_state:
    st.session_state.last_query = ""
if query != st.session_state.last_query:
    st.session_state.page = 1
    st.session_state.last_query = query

page = st.session_state.page
start_idx = (page - 1) * CONTACTS_PER_PAGE
end_idx = start_idx + CONTACTS_PER_PAGE
page_df = filtered.iloc[start_idx:end_idx]

# ── Session state for selected contact ──
if "selected_contact" not in st.session_state:
    st.session_state.selected_contact = None

# ── Two-Panel Layout: List on Left, Detail on Right ──
list_col, detail_col = st.columns([1, 4])

with list_col:
    for idx, row in page_df.iterrows():
        _, company = parse_current_position(row.get("Positions", ""))
        name = f"{row.get('FirstName', '')} {row.get('LastName', '')}".strip()
        label = f"{name} — {company}" if company else name
        is_selected = st.session_state.selected_contact == idx

        if st.button(
            label,
            key=f"contact_{idx}",
            use_container_width=True,
            type="primary" if is_selected else "secondary"
        ):
            st.session_state.selected_contact = idx
            st.rerun()

with detail_col:
    sel_idx = st.session_state.selected_contact
    if sel_idx is not None and sel_idx in page_df.index:
        row = page_df.loc[sel_idx]
        title, company = parse_current_position(row.get("Positions", ""))
        display_title = f"{title}" if title else ""
        if company:
            display_title += f" @ {company}" if display_title else company
        name = f"{row.get('FirstName', '')} {row.get('LastName', '')}".strip()

        # Source badges
        source_badges = ""
        for s in str(row.get("Sources", "")).split("; "):
            if s.strip():
                source_badges += f'<span class="source-badge">{s.strip()}</span>'

        # ── Split into main area and right sidebar from the top ──
        main_area, right_area = st.columns([3, 2])

        with right_area:
            def _save_on_change(field_key, field_name, row_idx):
                """Callback: auto-save when user presses Enter or tabs away."""
                new_val = st.session_state[field_key]
                save_field(row_idx, field_name, new_val)

            # ── Contact Info (above Tags) ──
            st.markdown("**Contact Info**")

            def _clean_val(v):
                v = str(v).lstrip("'")
                return "" if v in ("nan", "None") else v

            # Email 1 + Email 2 always editable (so a second email can be added)
            st.text_input("Email 1", value=_clean_val(row.get("Email1", "")),
                          key=f"email1_{sel_idx}",
                          on_change=_save_on_change, args=(f"email1_{sel_idx}", "Email1", sel_idx))
            st.text_input("Email 2", value=_clean_val(row.get("Email2", "")),
                          key=f"email2_{sel_idx}", placeholder="second email…",
                          on_change=_save_on_change, args=(f"email2_{sel_idx}", "Email2", sel_idx))
            # Email 3 only if already populated (keeps the panel tidy)
            email3_val = _clean_val(row.get("Email3", ""))
            if email3_val:
                st.text_input("Email 3", value=email3_val, key=f"email3_{sel_idx}",
                              on_change=_save_on_change, args=(f"email3_{sel_idx}", "Email3", sel_idx))
            # Phone 1 + Phone 2 always editable
            st.text_input("Phone 1", value=_clean_val(row.get("Phone1", "")),
                          key=f"phone1_{sel_idx}", placeholder="phone…",
                          on_change=_save_on_change, args=(f"phone1_{sel_idx}", "Phone1", sel_idx))
            st.text_input("Phone 2", value=_clean_val(row.get("Phone2", "")),
                          key=f"phone2_{sel_idx}", placeholder="second phone…",
                          on_change=_save_on_change, args=(f"phone2_{sel_idx}", "Phone2", sel_idx))

            # ── Tags (below Contact Info) ──
            st.markdown("**Tags**")
            current_tags_str = str(row.get("DistributionTags", ""))
            if current_tags_str in ("", "nan", "None"):
                current_tags_list = []
            else:
                current_tags_list = [t.strip() for t in current_tags_str.split(";")
                                     if t.strip() and t.strip() in AVAILABLE_TAGS]

            def _save_tags(tag_key, row_idx):
                """Save selected tags back to sheet as semicolon-separated."""
                selected = st.session_state[tag_key]
                save_field(row_idx, "DistributionTags", "; ".join(selected))

            st.multiselect(
                "Select tags",
                options=AVAILABLE_TAGS,
                default=current_tags_list,
                key=f"tags_{sel_idx}",
                label_visibility="collapsed",
                on_change=_save_tags,
                args=(f"tags_{sel_idx}", sel_idx),
            )

            # ── Source (editable) ──
            st.markdown("**Source**")
            cur_src = str(row.get("Sources", "")).strip()
            if cur_src in ("nan", "None"):
                cur_src = ""
            src_opts = list(AVAILABLE_SOURCES)
            if cur_src and cur_src not in src_opts:
                src_opts = [cur_src] + src_opts  # keep the current value selectable

            def _save_source(skey, row_idx):
                save_field(row_idx, "Sources", st.session_state[skey])

            st.selectbox(
                "Source", src_opts,
                index=(src_opts.index(cur_src) if cur_src in src_opts else 0),
                key=f"src_{sel_idx}", label_visibility="collapsed",
                on_change=_save_source, args=(f"src_{sel_idx}", sel_idx),
            )

        with main_area:
            # ── Photo ──
            pkey = photo_key_for_row(row)
            existing_b64 = get_photo_b64(pkey) if pkey else None
            if existing_b64:
                st.image(base64.b64decode(existing_b64), width=96)
            with st.popover("📷 Add / update photo"):
                st.caption("Copy a photo (right-click → Copy image), then click Paste.")
                # Nonce in the key remounts the paste button after a save, so it
                # stops re-handing the same image (was looping "running photo…").
                _pnonce = st.session_state.get(f"paste_nonce_{sel_idx}", 0)
                paste_res = paste_button("📋 Paste photo", key=f"paste_{sel_idx}_{_pnonce}")
                if paste_res.image_data is not None:
                    key = pkey or uuid.uuid4().hex
                    if not pkey:
                        save_field(sel_idx, "PhotoKey", key)
                    if set_photo(key, pil_thumb_b64(paste_res.image_data), name):
                        st.session_state[f"paste_nonce_{sel_idx}"] = _pnonce + 1
                        st.success("Photo saved!")
                        st.rerun()

            # ── Header: Name, Title, Location ──
            loc_parts = []
            if row.get("City"):
                loc_parts.append(str(row["City"]))
            if row.get("State"):
                loc_parts.append(str(row["State"]))
            if row.get("Location") and not loc_parts:
                loc_parts.append(str(row["Location"]))
            loc_str = ', '.join(loc_parts)

            header_line = f"### {name}"
            if loc_str:
                header_line += f"&emsp;📍 {loc_str}"
            st.markdown(header_line, unsafe_allow_html=True)
            if display_title:
                st.markdown(f"**{display_title}**")

            # Position history & Education popovers inline
            positions = str(row.get("Positions", ""))
            if positions:
                with st.popover("📋 Position History"):
                    for pos in reversed(positions.split(" | ")):
                        st.markdown(f"- {pos}")

            # Update current role — appends a new current position (keeps history)
            with st.popover("✏️ Update current role"):
                ur_title = st.text_input("New title", key=f"urtitle_{sel_idx}")
                ur_co = st.text_input("New company", key=f"urco_{sel_idx}")
                if st.button("Save current role", key=f"ursave_{sel_idx}"):
                    if ur_title.strip() or ur_co.strip():
                        _new_pos = f"{ur_title.strip()} @ {ur_co.strip()}".strip(" @")
                        _today = datetime.now().strftime("%Y-%m-%d")
                        _entry = f"{_new_pos} (Updated {_today})"
                        _existing = str(row.get("Positions", "")).strip()
                        _combined = f"{_existing} | {_entry}" if _existing else _entry
                        if save_field(sel_idx, "Positions", _combined):
                            st.success("Updated.")
                            st.rerun()
                    else:
                        st.warning("Enter a title and/or company.")

            education = str(row.get("Education", ""))
            if education:
                with st.popover("🎓 Education"):
                    for edu in education.split(" | "):
                        st.markdown(f"- {edu}")

            # LinkedIn
            if row.get("LinkedInURL"):
                st.markdown(f"🔗 [LinkedIn Profile]({row['LinkedInURL']})")

            # Sources
            if source_badges:
                st.markdown(source_badges, unsafe_allow_html=True)

            # Stats line
            stats_parts = []
            email_count = row.get("EmailCount", "")
            if email_count and str(email_count) not in ("", "nan", "None", "0"):
                try:
                    stats_parts.append(f"{int(float(email_count))} emails")
                except (ValueError, TypeError):
                    stats_parts.append(f"{email_count} emails")
            if row.get("LastActivityDate"):
                stats_parts.append(f"Last active: {row['LastActivityDate']}")
            if row.get("ConnectionDate"):
                stats_parts.append(f"Connected: {row['ConnectionDate']}")
            if stats_parts:
                st.caption("  ·  ".join(stats_parts))

        # ── Notes (full width below both columns) ──
        notes_raw = str(row.get("Notes", ""))
        all_notes = []
        if notes_raw and notes_raw not in ("", "nan", "None"):
            all_notes = [n.strip() for n in notes_raw.split(" || ") if n.strip() and n.strip() not in ("nan", "None")]

        # Explicit sort: newest first by date prefix [YYYY-MM-DD], pinned notes on top
        def _note_sort_key(note):
            if note.startswith("[Crelate-pinned]"):
                return "9999-99-99"
            if note.startswith("[") and len(note) > 11 and note[5] == "-":
                return note[1:11]
            return "0000-00-00"
        # Reverse first so that within the same date (notes only carry a date,
        # not a time) the most-recently-added note sorts to the top.
        all_notes = sorted(list(reversed(all_notes)), key=_note_sort_key, reverse=True)

        note_count = len(all_notes)
        stats_line = f"**Notes** ({note_count})"
        if email_count and str(email_count) not in ("", "nan", "None", "0"):
            try:
                stats_line += f"&emsp;|&emsp;{int(float(email_count))} emails"
            except (ValueError, TypeError):
                stats_line += f"&emsp;|&emsp;{email_count} emails"
        st.markdown(stats_line, unsafe_allow_html=True)

        # Add new note. A per-contact nonce gives the input a fresh key after each
        # save, so it clears and CANNOT re-fire on rerun (was creating duplicates).
        _nnonce = st.session_state.get(f"note_nonce_{sel_idx}", 0)
        new_note = st.text_input("Add note", key=f"note_{sel_idx}_{_nnonce}",
                                 placeholder="Type a note and press Enter...")
        if new_note:
            if save_note(sel_idx, new_note):
                st.session_state[f"note_nonce_{sel_idx}"] = _nnonce + 1
                st.success("Note saved!")
                st.rerun()

        # Show all notes, most recent first, one per line. Key includes the note
        # nonce so the box remounts (and shows the new note) right after a save —
        # a keyed text_area otherwise caches its first value and ignores updates.
        if all_notes:
            notes_display = "\n".join(all_notes)
            st.text_area("Previous notes", value=notes_display, height=200,
                         disabled=True, key=f"notes_display_{sel_idx}_{_nnonce}",
                         label_visibility="collapsed")

        # ── Record actions: Merge + Delete side by side ──
        st.markdown("")  # spacer
        merge_col, delete_col, _spacer = st.columns([1.4, 1, 2])
        with merge_col:
            with st.popover("🔗 Merge duplicate"):
                st.caption("Keep THIS record as the survivor. Find the duplicate; its "
                           "notes, tags, emails, phones, and newer job are combined here, "
                           "then the duplicate row is deleted.")
                dq = st.text_input("Find the duplicate", key=f"mergeq_{sel_idx}",
                                   placeholder="name…")
                if dq:
                    cands = search_contacts(df, dq)
                    cands = cands[cands.index != sel_idx]
                    opts = {}
                    for cidx, crow in cands.head(15).iterrows():
                        _t, _co = parse_current_position(crow.get("Positions", ""))
                        cnm = f"{crow.get('FirstName','')} {crow.get('LastName','')}".strip()
                        opts[f"{cnm} — {_co or _t}  ·  row {cidx}"] = cidx
                    if opts:
                        pick = st.selectbox("Duplicate record", list(opts.keys()),
                                            key=f"mergepick_{sel_idx}")
                        if st.button("Merge & delete duplicate", key=f"mergego_{sel_idx}",
                                     type="primary"):
                            if merge_records(sel_idx, opts[pick]):
                                st.session_state.selected_contact = None
                                st.success("Merged — reselect the record to view the result.")
                                st.rerun()
                    else:
                        st.write("No other matching records.")
        with delete_col:
            if not st.session_state.get(f"confirm_delete_{sel_idx}"):
                if st.button("🗑️ Delete", key=f"delete_{sel_idx}"):
                    st.session_state[f"confirm_delete_{sel_idx}"] = True
                    st.rerun()

        # Delete confirmation (full width, below the action row)
        if st.session_state.get(f"confirm_delete_{sel_idx}"):
            st.warning(f"Permanently delete **{name}**? This cannot be undone.")
            confirm_col, cancel_col, _ = st.columns([1, 1, 3])
            with confirm_col:
                if st.button("Yes, delete", key=f"do_delete_{sel_idx}", type="primary"):
                    if delete_contact(sel_idx):
                        st.session_state.selected_contact = None
                        st.session_state.pop(f"confirm_delete_{sel_idx}", None)
                        st.rerun()
            with cancel_col:
                if st.button("Cancel", key=f"cancel_delete_{sel_idx}"):
                    st.session_state.pop(f"confirm_delete_{sel_idx}", None)
                    st.rerun()

    else:
        st.markdown("*Select a contact from the list to view details.*")

# ── Pagination Controls ──
if total_pages > 1:
    st.divider()
    p1, p2, p3, p4, p5 = st.columns([1, 1, 2, 1, 1])
    with p1:
        if st.button("⏮ First", disabled=(page == 1)):
            st.session_state.page = 1
            st.rerun()
    with p2:
        if st.button("◀ Prev", disabled=(page == 1)):
            st.session_state.page = page - 1
            st.rerun()
    with p3:
        st.markdown(f"<div style='text-align:center; padding-top:8px;'>Page {page} of {total_pages}</div>", unsafe_allow_html=True)
    with p4:
        if st.button("Next ▶", disabled=(page >= total_pages)):
            st.session_state.page = page + 1
            st.rerun()
    with p5:
        if st.button("Last ⏭", disabled=(page >= total_pages)):
            st.session_state.page = total_pages
            st.rerun()

# ── Sidebar: Export ──
with st.sidebar:
    if st.button("🔄 Reload data"):
        st.cache_data.clear()
        st.session_state.df = load_data()
        st.rerun()

    st.markdown("### Export")
    if not filtered.empty:
        csv = filtered.to_csv(index=False).encode("utf-8")
        st.download_button(
            "📥 Download filtered CSV",
            csv,
            f"bes_contacts_export_{datetime.now().strftime('%Y%m%d')}.csv",
            "text/csv",
        )

    st.divider()
    st.markdown("### Quick Stats")
    st.markdown(f"**Total contacts:** {len(df):,}")

    # Source breakdown
    source_counts = {}
    for sources in df["Sources"]:
        for s in str(sources).split("; "):
            s = s.strip()
            if s:
                source_counts[s] = source_counts.get(s, 0) + 1
    for src, count in sorted(source_counts.items(), key=lambda x: -x[1]):
        st.markdown(f"- {src}: {count:,}")

    st.divider()
    st.caption("Berkshire Executive Search")
    st.caption(f"Last refreshed: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
