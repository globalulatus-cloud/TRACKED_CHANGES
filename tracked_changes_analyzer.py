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

st.set_page_config(page_title="Tracked Changes Analyzer", page_icon="📝", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Fraunces:wght@300;600;700&family=DM+Sans:wght@300;400;500;600&display=swap');

:root {
    --cream: #faf8f4; --white: #ffffff; --border: #e4e0d8;
    --muted: #9e9589; --body: #3a3530; --heading: #1e1a17;
    --accent: #b5813a;
    --ins: #1a6e3f; --ins-bg: #eef8f3; --ins-bd: #9ed4b8;
    --del: #b03030; --del-bg: #fdf1f1; --del-bd: #edaaaa;
    --total: #1a52a0; --net-color: #b5813a;
}

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif !important; color: var(--body); }
.stApp { background: var(--cream); }
.block-container { padding-top: 2.5rem !important; padding-bottom: 3rem !important; max-width: 1080px !important; }

.page-header { padding-bottom: 1.4rem; margin-bottom: 0.2rem; border-bottom: 2px solid var(--heading); }
.page-title { font-family: 'Fraunces', serif; font-size: 2.5rem; font-weight: 700; color: var(--heading); line-height: 1; letter-spacing: -0.02em; margin: 0 0 0.35rem 0; }
.page-title em { font-style: italic; color: var(--accent); }
.page-subtitle { font-size: 0.88rem; color: var(--muted); margin: 0; }

.step-label { display: flex; align-items: center; gap: 0.55rem; font-size: 0.68rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.2em; color: var(--muted); margin: 2.2rem 0 0.9rem 0; }
.step-num { display: inline-flex; align-items: center; justify-content: center; width: 19px; height: 19px; background: var(--heading); color: var(--cream); border-radius: 50%; font-size: 0.62rem; font-weight: 700; flex-shrink: 0; }

div[data-testid="stFileUploader"] { background: var(--white); border: 1.5px dashed var(--border); border-radius: 10px; padding: 0.5rem; transition: border-color 0.2s; }
div[data-testid="stFileUploader"]:hover { border-color: var(--accent); }

.stRadio label { color: var(--body) !important; font-size: 0.92rem !important; }

.info-box { background: var(--white); border: 1px solid var(--border); border-radius: 10px; padding: 1rem 1.3rem; font-size: 0.86rem; color: var(--body); line-height: 1.75; }
.info-box b { color: var(--heading); }

.metrics-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 1rem; margin: 0.5rem 0 1.5rem 0; }
.metric-card { background: var(--white); border: 1px solid var(--border); border-radius: 12px; padding: 1.4rem 1.2rem; text-align: center; transition: box-shadow 0.2s, border-color 0.2s; }
.metric-card:hover { box-shadow: 0 4px 16px rgba(0,0,0,0.07); border-color: var(--accent); }
.metric-card.ins  { border-top: 3px solid var(--ins); }
.metric-card.del  { border-top: 3px solid var(--del); }
.metric-card.tot  { border-top: 3px solid var(--total); }
.metric-card.net  { border-top: 3px solid var(--accent); }
.metric-value { font-family: 'Fraunces', serif; font-size: 2.6rem; font-weight: 700; line-height: 1; margin-bottom: 0.35rem; }
.metric-value.ins   { color: var(--ins); }
.metric-value.del   { color: var(--del); }
.metric-value.total { color: var(--total); }
.metric-value.net   { color: var(--net-color); }
.metric-label { font-size: 0.72rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.12em; color: var(--muted); }
.metric-sub { font-size: 0.74rem; color: #bbb; margin-top: 0.3rem; }

.stTabs [data-baseweb="tab-list"] { background: transparent; border-bottom: 2px solid var(--border); gap: 0; }
.stTabs [data-baseweb="tab"] { background: transparent !important; color: var(--muted) !important; font-size: 0.85rem !important; font-weight: 500 !important; padding: 0.6rem 1.2rem !important; border-bottom: 2px solid transparent !important; margin-bottom: -2px; }
.stTabs [aria-selected="true"] { color: var(--heading) !important; border-bottom-color: var(--heading) !important; font-weight: 600 !important; }

.stDownloadButton > button { background: var(--heading) !important; color: var(--cream) !important; border: none !important; border-radius: 8px !important; font-family: 'DM Sans', sans-serif !important; font-size: 0.88rem !important; font-weight: 600 !important; padding: 0.6rem 1.2rem !important; transition: background 0.2s !important; width: 100%; }
.stDownloadButton > button:hover { background: var(--accent) !important; color: white !important; }

.empty-state { text-align: center; padding: 5rem 2rem; }
.empty-icon { font-size: 3rem; margin-bottom: 1rem; opacity: 0.35; }
.empty-text { font-family: 'Fraunces', serif; font-size: 1.1rem; font-weight: 600; color: var(--muted); margin-bottom: 0.4rem; }
.empty-sub { font-size: 0.83rem; color: #bbb; }
.dl-hint { font-size: 0.78rem; color: var(--muted); margin-top: 0.6rem; }
</style>
""", unsafe_allow_html=True)


# ─── Logic ────────────────────────────────────────────────────────────────────

def count_units(text, mode):
    if not text or not text.strip():
        return 0
    if mode == "asian":
        return sum(1 for ch in text if not ch.isspace())
    return len(re.findall(r'\S+', text))

def extract_text(block, tag):
    return "".join(re.findall(rf'<{tag}(?:\s[^>]*)?>([^<]*)</{tag}>', block))

def parse_tracked_changes(docx_bytes):
    try:
        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
            xml = z.read("word/document.xml").decode("utf-8", errors="replace")
    except Exception as e:
        return None, str(e)

    changes = []
    for m in re.finditer(r'<w:ins\s([^>]*)>(.*?)</w:ins>', xml, re.DOTALL):
        text = extract_text(m.group(2), "w:t")
        if not text.strip(): continue
        a = re.search(r'w:author="([^"]*)"', m.group(1))
        d = re.search(r'w:date="([^"]*)"', m.group(1))
        changes.append({"type": "insertion", "text": text,
                         "author": a.group(1) if a else "Unknown",
                         "date": d.group(1)[:10] if d else "—"})

    for m in re.finditer(r'<w:del\s([^>]*)>(.*?)</w:del>', xml, re.DOTALL):
        text = extract_text(m.group(2), "w:delText")
        if not text.strip(): continue
        a = re.search(r'w:author="([^"]*)"', m.group(1))
        d = re.search(r'w:date="([^"]*)"', m.group(1))
        changes.append({"type": "deletion", "text": text,
                         "author": a.group(1) if a else "Unknown",
                         "date": d.group(1)[:10] if d else "—"})
    return changes, None

def build_csv(changes, mode):
    ul = "Characters" if mode == "asian" else "Words"
    rows = [{"No.": i, "Type": c["type"].capitalize(), "Author": c["author"],
             "Date": c["date"], "Text": c["text"], ul: count_units(c["text"], mode)}
            for i, c in enumerate(changes, 1)]
    return pd.DataFrame(rows).to_csv(index=False).encode("utf-8-sig")

def build_excel(changes, mode):
    ul = "Characters" if mode == "asian" else "Words"
    ins_rows = [{"No.": i, "Author": c["author"], "Date": c["date"], "Text": c["text"],
                 ul: count_units(c["text"], mode)}
                for i, c in enumerate(changes, 1) if c["type"] == "insertion"]
    del_rows = [{"No.": i, "Author": c["author"], "Date": c["date"], "Text": c["text"],
                 ul: count_units(c["text"], mode)}
                for i, c in enumerate(changes, 1) if c["type"] == "deletion"]
    iu = sum(r[ul] for r in ins_rows)
    du = sum(r[ul] for r in del_rows)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame([
            {"Metric": f"Total Insertions ({ul})", "Value": iu},
            {"Metric": f"Total Deletions ({ul})",  "Value": du},
            {"Metric": f"Net Change ({ul})",        "Value": iu - du},
            {"Metric": "Insertion Segments", "Value": len(ins_rows)},
            {"Metric": "Deletion Segments",  "Value": len(del_rows)},
            {"Metric": "Generated",          "Value": datetime.now().strftime("%Y-%m-%d %H:%M")},
        ]).to_excel(w, sheet_name="Summary", index=False)
        if ins_rows: pd.DataFrame(ins_rows).to_excel(w, sheet_name="Insertions", index=False)
        if del_rows: pd.DataFrame(del_rows).to_excel(w, sheet_name="Deletions",  index=False)
        pd.DataFrame([{"No.": i, "Type": c["type"].capitalize(), "Author": c["author"],
                        "Date": c["date"], "Text": c["text"], ul: count_units(c["text"], mode)}
                       for i, c in enumerate(changes, 1)]
        ).to_excel(w, sheet_name="All Changes", index=False)
    return buf.getvalue()


# ─── UI ───────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="page-header">
  <div class="page-title">Tracked Changes <em>Analyzer</em></div>
  <p class="page-subtitle">Extract &amp; quantify insertions and deletions from Word (.docx) comparison files</p>
</div>
""", unsafe_allow_html=True)

# Step 1
st.markdown('<div class="step-label"><span class="step-num">1</span> Upload Document</div>', unsafe_allow_html=True)
uploaded_file = st.file_uploader("Upload", type=["docx"], label_visibility="collapsed",
                                  help="Supports Compare Merge / tracked changes Word documents")

# Step 2
st.markdown('<div class="step-label"><span class="step-num">2</span> Select Language Mode</div>', unsafe_allow_html=True)
cr, ci = st.columns([1, 2])
with cr:
    mode = st.radio("lang", ["asian", "latin"], label_visibility="collapsed",
                    format_func=lambda x: "🀄  Asian  (character count)" if x == "asian"
                                          else "🔤  Latin / English  (word count)")
with ci:
    desc = ("Counts every <b>non-whitespace character</b> — CJK, Hiragana, Katakana, Hangul, and more."
            if mode == "asian" else
            "Counts <b>space-delimited words</b> in each inserted or deleted text run.")
    st.markdown(f'<div class="info-box">{desc}</div>', unsafe_allow_html=True)

# Analysis
if uploaded_file:
    changes, err = parse_tracked_changes(uploaded_file.read())
    if err:
        st.error(f"Could not parse document: {err}"); st.stop()
    if not changes:
        st.warning("No tracked changes found in this document."); st.stop()

    ins  = [c for c in changes if c["type"] == "insertion"]
    dels = [c for c in changes if c["type"] == "deletion"]
    iu   = sum(count_units(c["text"], mode) for c in ins)
    du   = sum(count_units(c["text"], mode) for c in dels)
    tu   = iu + du
    nu   = iu - du
    uc   = "Characters" if mode == "asian" else "Words"
    sgn  = "+" if nu >= 0 else ""

    # Step 3
    st.markdown('<div class="step-label"><span class="step-num">3</span> Analysis Results</div>', unsafe_allow_html=True)
    st.markdown(f"""
    <div class="metrics-row">
      <div class="metric-card ins">
        <div class="metric-value ins">{iu:,}</div>
        <div class="metric-label">Inserted {uc}</div>
        <div class="metric-sub">{len(ins):,} segments</div>
      </div>
      <div class="metric-card del">
        <div class="metric-value del">{du:,}</div>
        <div class="metric-label">Deleted {uc}</div>
        <div class="metric-sub">{len(dels):,} segments</div>
      </div>
      <div class="metric-card tot">
        <div class="metric-value total">{tu:,}</div>
        <div class="metric-label">Total Changes</div>
        <div class="metric-sub">{len(changes):,} segments</div>
      </div>
      <div class="metric-card net">
        <div class="metric-value net">{sgn}{nu:,}</div>
        <div class="metric-label">Net Change</div>
        <div class="metric-sub">insertions − deletions</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    tab_all, tab_ins, tab_del = st.tabs([
        f"All Changes  ({len(changes)})",
        f"✦ Insertions  ({len(ins)})",
        f"✦ Deletions  ({len(dels)})",
    ])

    def render(items, ft=None):
        display = [c for c in items if ft is None or c["type"] == ft]
        if not display: st.info("No entries to display."); return
        rows = [{"Type": "↑ INS" if c["type"] == "insertion" else "↓ DEL",
                  uc: count_units(c["text"], mode), "Author": c["author"], "Date": c["date"],
                  "Text": c["text"][:130] + ("…" if len(c["text"]) > 130 else "")}
                for c in display]
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True,
                     column_config={"Type": st.column_config.TextColumn("Type", width=80),
                                    uc: st.column_config.NumberColumn(uc, width=100),
                                    "Author": st.column_config.TextColumn("Author", width=130),
                                    "Date": st.column_config.TextColumn("Date", width=100),
                                    "Text": st.column_config.TextColumn("Text preview")})

    with tab_all: render(changes)
    with tab_ins: render(changes, "insertion")
    with tab_del: render(changes, "deletion")

    # Step 4
    st.markdown('<div class="step-label"><span class="step-num">4</span> Download Report</div>', unsafe_allow_html=True)
    fname = uploaded_file.name.replace(".docx", "")
    d1, d2 = st.columns(2)
    with d1:
        st.download_button("📥  Download Excel Report (.xlsx)", data=build_excel(changes, mode),
                           file_name=f"{fname}_tracked_changes.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
    with d2:
        st.download_button("📄  Download CSV Report (.csv)", data=build_csv(changes, mode),
                           file_name=f"{fname}_tracked_changes.csv", mime="text/csv",
                           use_container_width=True)
    st.markdown('<p class="dl-hint">Excel contains 4 sheets: Summary · Insertions · Deletions · All Changes</p>',
                unsafe_allow_html=True)

else:
    st.markdown("""
    <div class="empty-state">
      <div class="empty-icon">📄</div>
      <div class="empty-text">Upload a .docx file to begin</div>
      <div class="empty-sub">Works with Compare Merge documents generated by Microsoft Word</div>
    </div>
    """, unsafe_allow_html=True)
