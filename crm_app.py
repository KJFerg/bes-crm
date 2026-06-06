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
from datetime import datetime
from streamlit_gsheets import GSheetsConnection

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


def save_field(row_index, field_name, new_value):
    """Update a single field for a contact in Google Sheets."""
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        df = conn.read(worksheet="Contacts", ttl=0)
        df = _prep_df_for_write(df)
        df.at[row_index, field_name] = str(new_value)
        conn.update(worksheet="Contacts", data=df)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Could not save {field_name}: {e}")
        return False


def save_note(row_index, new_note):
    """Append a note to a contact's Notes field in Google Sheets."""
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        df = conn.read(worksheet="Contacts", ttl=0)
        df = _prep_df_for_write(df)
        timestamp = datetime.now().strftime("%Y-%m-%d")
        note_entry = f"[{timestamp}] {new_note}"
        existing = df.at[row_index, "Notes"]
        df.at[row_index, "Notes"] = f"{existing} || {note_entry}" if existing else note_entry
        conn.update(worksheet="Contacts", data=df)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Could not save note: {e}")
        return False


def parse_current_position(positions_str):
    """Extract the most recent position for display."""
    if not positions_str:
        return "", ""
    # Last position is most recent (We-Connect or LinkedIn latest)
    parts = str(positions_str).split(" | ")
    latest = parts[-1] if parts else ""
    # Strip source tag
    for tag in ["(Crelate)", "(LinkedIn export Mar 2026)", "(We-Connect)", "(Interim)"]:
        latest = latest.replace(tag, "").strip()
    if " @ " in latest:
        title, company = latest.split(" @ ", 1)
        return title.strip(), company.strip()
    return latest.strip(), ""


def search_contacts(df, query):
    """Search across name, company, title, email, notes, location."""
    if not query:
        return df
    q = query.lower()
    mask = (
        df["FirstName"].str.lower().str.contains(q, na=False) |
        df["LastName"].str.lower().str.contains(q, na=False) |
        df["Positions"].str.lower().str.contains(q, na=False) |
        df["Email1"].str.lower().str.contains(q, na=False) |
        df["Email2"].str.lower().str.contains(q, na=False) |
        df["Notes"].str.lower().str.contains(q, na=False) |
        df["Industry"].str.lower().str.contains(q, na=False) |
        df["Location"].str.lower().str.contains(q, na=False) |
        df["City"].str.lower().str.contains(q, na=False) |
        df["State"].str.lower().str.contains(q, na=False) |
        df["DistributionTags"].str.lower().str.contains(q, na=False) |
        df["Education"].str.lower().str.contains(q, na=False)
    )
    return df[mask]


# ══════════════════════════════════════════════════════════════════════════
# MAIN APP
# ══════════════════════════════════════════════════════════════════════════

# Load data
df = load_data()

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

        # Tag badges
        tag_badges = ""
        for t in str(row.get("DistributionTags", "")).split("; "):
            if t.strip():
                tag_badges += f'<span class="tag-badge">{t.strip()}</span>'

        # ── Header: Name, Title, Location on one block ──
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

        # ── Two sub-columns: info + contact ──
        info_col, contact_col = st.columns([1, 1])

        with info_col:
            # Position history
            positions = str(row.get("Positions", ""))
            if positions:
                with st.popover("📋 Position History"):
                    for pos in reversed(positions.split(" | ")):
                        st.markdown(f"- {pos}")

            # Education
            education = str(row.get("Education", ""))
            if education:
                with st.popover("🎓 Education"):
                    for edu in education.split(" | "):
                        st.markdown(f"- {edu}")

            # LinkedIn
            if row.get("LinkedInURL"):
                st.markdown(f"🔗 [LinkedIn Profile]({row['LinkedInURL']})")

            if row.get("Industry"):
                st.markdown(f"🏢 {row['Industry']}")

            # Tags
            if tag_badges:
                st.markdown(tag_badges, unsafe_allow_html=True)

            # Sources
            if source_badges:
                st.markdown(source_badges, unsafe_allow_html=True)

            # Stats
            if row.get("EmailCount"):
                st.caption(f"{row['EmailCount']} email exchanges")
            if row.get("LastActivityDate"):
                st.caption(f"Last active: {row['LastActivityDate']}")
            if row.get("ConnectionDate"):
                st.caption(f"Connected: {row['ConnectionDate']}")

        with contact_col:
            # ── Editable Contact Fields (auto-save on Enter) ──
            st.markdown("**Edit Contact Info**")

            def _save_on_change(field_key, field_name, row_idx):
                """Callback: auto-save when user presses Enter or tabs away."""
                new_val = st.session_state[field_key]
                save_field(row_idx, field_name, new_val)

            st.text_input("Email 1", value=row.get("Email1", ""), key=f"email1_{sel_idx}",
                          on_change=_save_on_change, args=(f"email1_{sel_idx}", "Email1", sel_idx))
            st.text_input("Email 2", value=row.get("Email2", ""), key=f"email2_{sel_idx}",
                          on_change=_save_on_change, args=(f"email2_{sel_idx}", "Email2", sel_idx))
            st.text_input("Phone", value=row.get("Phone1", ""), key=f"phone1_{sel_idx}",
                          on_change=_save_on_change, args=(f"phone1_{sel_idx}", "Phone1", sel_idx))

        # ── Notes (full width below) ──
        st.markdown("---")
        notes = str(row.get("Notes", ""))
        if notes and notes not in ("", "nan", "None"):
            st.markdown("**Notes:**")
            for note in notes.split(" || "):
                clean_note = note.strip()
                if clean_note and clean_note not in ("nan", "None"):
                    st.markdown(f"> {clean_note}")

        new_note = st.text_input("Add note", key=f"note_{sel_idx}", placeholder="Type a note and press Enter...")
        if new_note:
            if save_note(sel_idx, new_note):
                st.success("Note saved!")
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
