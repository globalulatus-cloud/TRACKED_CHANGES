"""
Tracked Changes Analyzer for Word Documents (.docx)
Supports Asian (character-level) and Latin (word-level) counting.
"""

import streamlit as st
import zipfile
import re
import io
import unicodedata
from datetime import datetime
import pandas as pd

# ─── Page Config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Tracked Changes Analyzer",
    page_icon="📝",
    layout="wide",
)

# ─── Custom CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
}

.stApp {
    background: #0f1117;
    color: #e8e8e8;
}

h1, h2, h3 {
    font-family: 'IBM Plex Mono', monospace !important;
    letter-spacing: -0.03em;
}

.main-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 2.4rem;
    font-weight: 600;
    color: #f0f0f0;
    letter-spacing: -0.04em;
    margin-bottom: 0.2rem;
    border-left: 4px solid #4ade80;
    padding-left: 1rem;
}

.sub-title {
    font-family: 'IBM Plex Sans', sans-serif;
    color: #888;
    font-size: 0.95rem;
    padding-left: 1.25rem;
    margin-bottom: 2rem;
}

.metric-card {
    background: #1a1d27;
    border: 1px solid #2a2d3a;
    border-radius: 8px;
    padding: 1.5rem;
    text-align: center;
    transition: border-color 0.2s;
}

.metric-card:hover {
    border-color: #4ade80;
}

.metric-value {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 2.8rem;
    font-weight: 600;
    line-height: 1;
    margin-bottom: 0.4rem;
}

.metric-label {
    font-size: 0.8rem;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    color: #888;
}

.ins-color { color: #4ade80; }
.del-color { color: #f87171; }
.total-color { color: #60a5fa; }
.net-color { color: #fbbf24; }

.section-header {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.75rem;
    text-transform: uppercase;
    letter-spacing: 0.15em;
    color: #555;
    border-bottom: 1px solid #2a2d3a;
    padding-bottom: 0.5rem;
    margin: 2rem 0 1rem 0;
}

.change-row-ins {
    background: rgba(74, 222, 128, 0.05);
    border-left: 3px solid #4ade80;
    padding: 0.6rem 1rem;
    margin: 0.3rem 0;
    border-radius: 0 4px 4px 0;
    font-family: 'IBM Plex Sans', sans-serif;
    font-size: 0.9rem;
}

.change-row-del {
    background: rgba(248, 113, 113, 0.05);
    border-left: 3px solid #f87171;
    padding: 0.6rem 1rem;
    margin: 0.3rem 0;
    border-radius: 0 4px 4px 0;
    font-family: 'IBM Plex Sans', sans-serif;
    font-size: 0.9rem;
}

.badge-ins {
    background: rgba(74, 222, 128, 0.2);
    color: #4ade80;
    padding: 0.15rem 0.5rem;
    border-radius: 4px;
    font-size: 0.75rem;
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 600;
}

.badge-del {
    background: rgba(248, 113, 113, 0.2);
    color: #f87171;
    padding: 0.15rem 0.5rem;
    border-radius: 4px;
    font-size: 0.75rem;
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 600;
}

div[data-testid="stFileUploader"] {
    background: #1a1d27;
    border: 2px dashed #2a2d3a;
    border-radius: 8px;
    padding: 1rem;
}

div[data-testid="stFileUploader"]:hover {
    border-color: #4ade80;
}

.stRadio label {
    color: #ccc !important;
}

.stButton > button {
    background: #4ade80;
    color: #0f1117;
    border: none;
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 600;
    letter-spacing: 0.05em;
    border-radius: 6px;
    padding: 0.5rem 1.5rem;
    transition: all 0.2s;
}

.stButton > button:hover {
    background: #22c55e;
    color: #0f1117;
}

.info-box {
    background: #1a1d27;
    border: 1px solid #2a2d3a;
    border-radius: 8px;
    padding: 1.2rem 1.5rem;
    font-size: 0.88rem;
    color: #aaa;
    line-height: 1.7;
}

.stDataFrame {
    background: #1a1d27 !important;
}

/* Scrollable changes list */
.changes-container {
    max-height: 480px;
    overflow-y: auto;
    padding-right: 0.5rem;
}
</style>
""", unsafe_allow_html=True)

# ─── Core Parsing Logic ───────────────────────────────────────────────────────

def is_asian_char(ch: str) -> bool:
    """Return True if character belongs to CJK or other Asian script blocks."""
    try:
        name = unicodedata.name(ch, "")
    except Exception:
        return False
    asian_prefixes = (
        "CJK", "HIRAGANA", "KATAKANA", "HANGUL",
        "BOPOMOFO", "THAI", "TIBETAN", "MYANMAR",
        "KHMER", "MONGOLIAN",
    )
    return any(name.startswith(p) for p in asian_prefixes)


def count_units(text: str, mode: str) -> int:
    """Count characters (Asian) or words (Latin) in text."""
    if not text or not text.strip():
        return 0
    if mode == "asian":
        return sum(1 for ch in text if not ch.isspace())
    else:
        # Word count for Latin/mixed
        words = re.findall(r'\S+', text)
        return len(words)


def extract_text_from_block(xml_block: str, tag: str = "w:t") -> str:
    """Extract inner text from w:t or w:delText tags."""
    pattern = rf'<{tag}(?:\s[^>]*)?>([^<]*)</{tag}>'
    parts = re.findall(pattern, xml_block)
    return "".join(parts)


def parse_tracked_changes(docx_bytes: bytes):
    """
    Parse a .docx file and extract all insertions and deletions
    with their text, author, date, and paragraph context.
    """
    try:
        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
            xml = z.read("word/document.xml").decode("utf-8", errors="replace")
    except Exception as e:
        return None, str(e)

    # Remove XML namespace prefixes for easier regex
    # Keep w: prefix as-is; we'll match directly

    changes = []

    # ── Insertions ──
    ins_blocks = re.finditer(
        r'<w:ins\s([^>]*)>(.*?)</w:ins>',
        xml, re.DOTALL
    )
    for m in ins_blocks:
        attrs = m.group(1)
        body  = m.group(2)
        text  = extract_text_from_block(body, "w:t")
        if not text.strip():
            continue
        author = re.search(r'w:author="([^"]*)"', attrs)
        date   = re.search(r'w:date="([^"]*)"', attrs)
        changes.append({
            "type":   "insertion",
            "text":   text,
            "author": author.group(1) if author else "Unknown",
            "date":   date.group(1)[:10] if date else "—",
        })

    # ── Deletions ──
    del_blocks = re.finditer(
        r'<w:del\s([^>]*)>(.*?)</w:del>',
        xml, re.DOTALL
    )
    for m in del_blocks:
        attrs = m.group(1)
        body  = m.group(2)
        text  = extract_text_from_block(body, "w:delText")
        if not text.strip():
            continue
        author = re.search(r'w:author="([^"]*)"', attrs)
        date   = re.search(r'w:date="([^"]*)"', attrs)
        changes.append({
            "type":   "deletion",
            "text":   text,
            "author": author.group(1) if author else "Unknown",
            "date":   date.group(1)[:10] if date else "—",
        })

    return changes, None


def build_report_csv(changes: list, mode: str) -> bytes:
    """Build a downloadable CSV report."""
    unit_label = "Characters" if mode == "asian" else "Words"
    rows = []
    for i, c in enumerate(changes, 1):
        rows.append({
            "No.": i,
            "Type": c["type"].capitalize(),
            "Author": c["author"],
            "Date": c["date"],
            "Text": c["text"],
            unit_label: count_units(c["text"], mode),
        })
    df = pd.DataFrame(rows)
    return df.to_csv(index=False).encode("utf-8-sig")  # utf-8-sig for Excel compatibility


def build_report_excel(changes: list, mode: str) -> bytes:
    """Build a downloadable Excel report with formatting."""
    unit_label = "Characters" if mode == "asian" else "Words"

    ins_rows, del_rows = [], []
    for i, c in enumerate(changes, 1):
        row = {
            "No.": i,
            "Author": c["author"],
            "Date": c["date"],
            "Text": c["text"],
            unit_label: count_units(c["text"], mode),
        }
        if c["type"] == "insertion":
            ins_rows.append(row)
        else:
            del_rows.append(row)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Summary sheet
        ins_units = sum(r[unit_label] for r in ins_rows)
        del_units = sum(r[unit_label] for r in del_rows)
        summary_df = pd.DataFrame([
            {"Metric": f"Total Insertions ({unit_label})", "Value": ins_units},
            {"Metric": f"Total Deletions ({unit_label})", "Value": del_units},
            {"Metric": f"Net Change ({unit_label})", "Value": ins_units - del_units},
            {"Metric": "Insertion Segments", "Value": len(ins_rows)},
            {"Metric": "Deletion Segments", "Value": len(del_rows)},
            {"Metric": "Report Generated", "Value": datetime.now().strftime("%Y-%m-%d %H:%M")},
        ])
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

        if ins_rows:
            pd.DataFrame(ins_rows).to_excel(writer, sheet_name="Insertions", index=False)
        if del_rows:
            pd.DataFrame(del_rows).to_excel(writer, sheet_name="Deletions", index=False)

        # All changes together
        all_rows = []
        for i, c in enumerate(changes, 1):
            all_rows.append({
                "No.": i,
                "Type": c["type"].capitalize(),
                "Author": c["author"],
                "Date": c["date"],
                "Text": c["text"],
                unit_label: count_units(c["text"], mode),
            })
        pd.DataFrame(all_rows).to_excel(writer, sheet_name="All Changes", index=False)

    return output.getvalue()


# ─── UI Layout ────────────────────────────────────────────────────────────────

st.markdown('<div class="main-title">Tracked Changes Analyzer</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Extract & quantify insertions and deletions from Word (.docx) comparison files</div>', unsafe_allow_html=True)

# Step 1 — Upload
st.markdown('<div class="section-header">① Upload Document</div>', unsafe_allow_html=True)
uploaded_file = st.file_uploader(
    "Drop your tracked-changes .docx file here",
    type=["docx"],
    help="Supports Compare Merge / tracked changes Word documents",
    label_visibility="collapsed",
)

# Step 2 — Language mode
st.markdown('<div class="section-header">② Select Language Mode</div>', unsafe_allow_html=True)
col_lang1, col_lang2 = st.columns([1, 3])
with col_lang1:
    mode = st.radio(
        "Language",
        options=["asian", "latin"],
        format_func=lambda x: "🀄 Asian (character count)" if x == "asian" else "🔤 Latin / English (word count)",
        label_visibility="collapsed",
    )

with col_lang2:
    unit_label = "characters" if mode == "asian" else "words"
    st.markdown(f"""
    <div class="info-box">
    <b>{'Asian mode' if mode == 'asian' else 'Latin / English mode'}:</b>
    counts {'every non-whitespace <b>character</b> (CJK, Hiragana, Katakana, Hangul, etc.)' if mode == 'asian' else 'space-delimited <b>words</b> in inserted and deleted runs'}.
    <br>Switch mode if your document contains the opposite script.
    </div>
    """, unsafe_allow_html=True)

# ─── Analysis ─────────────────────────────────────────────────────────────────
if uploaded_file:
    raw_bytes = uploaded_file.read()
    changes, error = parse_tracked_changes(raw_bytes)

    if error:
        st.error(f"❌ Could not parse document: {error}")
        st.stop()

    if not changes:
        st.warning("No tracked changes (insertions or deletions) were found in this document.")
        st.stop()

    insertions = [c for c in changes if c["type"] == "insertion"]
    deletions  = [c for c in changes if c["type"] == "deletion"]

    ins_units = sum(count_units(c["text"], mode) for c in insertions)
    del_units = sum(count_units(c["text"], mode) for c in deletions)
    total_units = ins_units + del_units
    net_units = ins_units - del_units
    unit_cap = "Characters" if mode == "asian" else "Words"

    # Step 3 — Summary metrics
    st.markdown('<div class="section-header">③ Analysis Results</div>', unsafe_allow_html=True)

    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value ins-color">{ins_units:,}</div>
            <div class="metric-label">Inserted {unit_cap}</div>
            <div style="color:#555;font-size:0.75rem;margin-top:0.3rem">{len(insertions)} segments</div>
        </div>""", unsafe_allow_html=True)
    with m2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value del-color">{del_units:,}</div>
            <div class="metric-label">Deleted {unit_cap}</div>
            <div style="color:#555;font-size:0.75rem;margin-top:0.3rem">{len(deletions)} segments</div>
        </div>""", unsafe_allow_html=True)
    with m3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value total-color">{total_units:,}</div>
            <div class="metric-label">Total Changes</div>
            <div style="color:#555;font-size:0.75rem;margin-top:0.3rem">{len(changes)} segments</div>
        </div>""", unsafe_allow_html=True)
    with m4:
        sign = "+" if net_units >= 0 else ""
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value net-color">{sign}{net_units:,}</div>
            <div class="metric-label">Net Change</div>
            <div style="color:#555;font-size:0.75rem;margin-top:0.3rem">ins − del</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Detailed change list ──────────────────────────────────────────────────
    tab_all, tab_ins, tab_del = st.tabs([
        f"All Changes ({len(changes)})",
        f"✅ Insertions ({len(insertions)})",
        f"🗑️ Deletions ({len(deletions)})",
    ])

    def render_changes(change_list, filter_type=None):
        display = [c for c in change_list if filter_type is None or c["type"] == filter_type]
        if not display:
            st.info("No entries to display.")
            return

        # Build DataFrame for display
        rows = []
        for c in display:
            u = count_units(c["text"], mode)
            rows.append({
                "Type": "↑ INS" if c["type"] == "insertion" else "↓ DEL",
                unit_cap: u,
                "Author": c["author"],
                "Date": c["date"],
                "Text": c["text"][:120] + ("…" if len(c["text"]) > 120 else ""),
            })
        df = pd.DataFrame(rows)

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Type": st.column_config.TextColumn("Type", width=80),
                unit_cap: st.column_config.NumberColumn(unit_cap, width=90),
                "Author": st.column_config.TextColumn("Author", width=120),
                "Date": st.column_config.TextColumn("Date", width=100),
                "Text": st.column_config.TextColumn("Text (preview)"),
            }
        )

    with tab_all:
        render_changes(changes)
    with tab_ins:
        render_changes(changes, "insertion")
    with tab_del:
        render_changes(changes, "deletion")

    # ── Download ──────────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">④ Download Report</div>', unsafe_allow_html=True)

    fname_base = uploaded_file.name.replace(".docx", "")
    d1, d2 = st.columns(2)

    with d1:
        excel_bytes = build_report_excel(changes, mode)
        st.download_button(
            label="📥 Download Excel Report (.xlsx)",
            data=excel_bytes,
            file_name=f"{fname_base}_tracked_changes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with d2:
        csv_bytes = build_report_csv(changes, mode)
        st.download_button(
            label="📄 Download CSV Report (.csv)",
            data=csv_bytes,
            file_name=f"{fname_base}_tracked_changes.csv",
            mime="text/csv",
            use_container_width=True,
        )

    st.markdown("""
    <div style="color:#444;font-size:0.78rem;margin-top:1rem;">
    Excel report contains 4 sheets: Summary · Insertions · Deletions · All Changes
    </div>
    """, unsafe_allow_html=True)

else:
    # Empty state
    st.markdown("""
    <div style="text-align:center; padding: 4rem 2rem; color: #444;">
        <div style="font-size:3rem;margin-bottom:1rem;">📄</div>
        <div style="font-family:'IBM Plex Mono',monospace;font-size:1rem;">
            Upload a tracked-changes .docx file to begin
        </div>
        <div style="font-size:0.82rem;margin-top:0.5rem;color:#333;">
            Works with Compare Merge documents generated by Microsoft Word
        </div>
    </div>
    """, unsafe_allow_html=True)
