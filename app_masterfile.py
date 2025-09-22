# app_masterfile.py

import io
import json
import re
from difflib import SequenceMatcher

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="Masterfile Automation", page_icon="📦", layout="wide")

# ---------------- Helpers ----------------
def norm(s: str) -> str:
    if s is None:
        return ""
    x = str(s).replace("\xa0", " ").strip().lower()  # kill NBSPs too
    x = re.sub(r"\s*-\s*en\s*[-_ ]\s*us\s*$", "", x)
    x = x.replace("–", "-").replace("—", "-").replace("−", "-")
    x = re.sub(r"[._/\\-]+", " ", x)
    x = re.sub(r"[^0-9a-z\s]+", " ", x)
    return re.sub(r"\s+", " ", x).strip()

def top_matches(query, candidates, k=3):
    q = norm(query)
    scored = [(SequenceMatcher(None, q, norm(c)).ratio(), c) for c in candidates]
    scored.sort(key=lambda t: t[0], reverse=True)
    return scored[:k]

def worksheet_used_cols(ws, header_rows=(1,), hard_cap=512, empty_streak_stop=8):
    max_try = min(ws.max_column, hard_cap)
    last_nonempty, streak = 0, 0
    for c in range(1, max_try + 1):
        any_val = False
        for r in header_rows:
            v = ws.cell(row=r, column=c).value
            if v not in (None, ""):
                any_val = True
                break
        if any_val:
            last_nonempty = c
            streak = 0
        else:
            streak += 1
            if streak >= empty_streak_stop:
                break
    return max(last_nonempty, 1)

def clean_header_row(raw_row):
    """Return list of header strings (trim NBSP etc.)."""
    headers = []
    for v in raw_row.tolist():
        if v is None:
            headers.append("")
        else:
            headers.append(str(v).replace("\xa0", " ").strip())
    return headers

def build_on_df_from_raw(raw_df, header_row_index):
    """
    - Take pandas DataFrame with no header (header=None).
    - header_row_index is 0-based index pointing to the header row.
    - Clean headers: trim, remove blanks, drop duplicates (by normalized key).
    - Return a new DataFrame with these cleaned headers and only kept columns.
    """
    # Extract & clean candidate headers
    headers_raw = clean_header_row(raw_df.iloc[header_row_index])
    keep_indices = []
    keep_names = []
    seen_norm = set()

    for idx, h in enumerate(headers_raw):
        hn = norm(h)
        if not hn:                      # drop truly blank
            continue
        if hn in seen_norm:             # drop duplicates by normalized name
            continue
        seen_norm.add(hn)
        keep_indices.append(idx)
        keep_names.append(h if h else f"col_{idx+1}")

    # Build DF with only kept columns; data is everything after the header row
    df = raw_df.iloc[header_row_index + 1:, keep_indices].copy()
    df.columns = keep_names
    df = df.fillna("")
    return df, keep_names

# sentinel for “Listing Action”
SENTINEL_LISTING_ACTION = object()

# ---------------- UI ----------------
st.title("📦 Masterfile Automation")
st.caption("Auto-detect or choose onboarding header row. Cleans headers (drops blanks/duplicates) and writes only rows that contain data.")

with st.expander("ℹ️ Quick guide", expanded=True):
    st.markdown("""
- **Masterfile**: Row **1** = labels, Row **2** = keys, data starts **Row 3** (styles preserved).
- **Onboarding**: Choose **Auto-detect** or **Pick row** for the header row.  
  The tool **cleans** that row (removes blank/duplicate headers) and uses it for mapping.  
  Data is **below** the selected header row.
- **Mapping JSON**: keys = master labels; values = list of onboarding aliases (priority order).
    """)

st.divider()
colA, colB = st.columns([1, 1])
with colA:
    masterfile_file = st.file_uploader("📄 Masterfile Template (.xlsx)", type=["xlsx"])
with colB:
    onboarding_file = st.file_uploader("🧾 Onboarding Sheet (.xlsx)", type=["xlsx"])

st.markdown("#### 🔗 Mapping JSON")
tab1, tab2 = st.tabs(["Paste JSON", "Upload JSON"])
mapping_text = ""
mapping_file = None
with tab1:
    mapping_text = st.text_area(
        "Paste mapping JSON",
        height=220,
        placeholder='{\n  "Partner SKU": ["Seller SKU", "item_sku"]\n}',
    )
with tab2:
    mapping_file = st.file_uploader("mapping.json", type=["json"])

st.markdown("#### 🧭 Header Row for Onboarding")
header_mode = st.radio(
    "How should the app find your header row?",
    ["Auto-detect", "Pick a row number"],
    horizontal=True,
)
header_row_manual = None
if header_mode == "Pick a row number":
    header_row_manual = st.number_input(
        "Header row number (1-based)", min_value=1, value=1, step=1
    )

st.divider()
go = st.button("🚀 Generate Final Masterfile", type="primary")
log_area = st.container()
download_area = st.container()

# ---------------- Main ----------------
if go:
    with log_area:
        st.markdown("### 📝 Log")
        log = st.empty()
        def slog(msg): log.markdown(msg)

        if not masterfile_file or not onboarding_file:
            st.error("Please upload both **Masterfile** and **Onboarding** files.")
            st.stop()

        # mapping
        try:
            mapping_raw = json.loads(mapping_text) if mapping_text.strip() else (
                json.load(mapping_file) if mapping_file else None
            )
        except Exception as e:
            st.error(f"Mapping JSON could not be parsed. Error: {e}")
            st.stop()
        if mapping_raw is None:
            st.error("Please provide mapping JSON (paste or upload).")
            st.stop()

        MAPPING = {norm(k): (v[:] if isinstance(v, list) else [v])
                   for k, v in mapping_raw.items()}

        # master
        slog("⏳ Reading master (preserving styles)…")
        try:
            master_wb = load_workbook(masterfile_file, keep_links=False)
            master_ws = master_wb.active
        except Exception as e:
            st.error(f"Could not read **Masterfile**: {e}")
            st.stop()
        used_cols = worksheet_used_cols(master_ws, header_rows=(1, 2))
        master_displays = [master_ws.cell(row=1, column=c).value or ""
                           for c in range(1, used_cols + 1)]

        # onboarding raw (no header), then pick/auto-detect header row
        slog("⏳ Reading onboarding…")
        try:
            raw_df = pd.read_excel(onboarding_file, header=None, dtype=str).fillna("")
        except Exception as e:
            st.error(f"Could not read **Onboarding**: {e}")
            st.stop()

        def build_mapping(df):
            on_headers = list(df.columns)
            series_by_alias = {norm(h): df[h] for h in on_headers}

            master_to_source = {}
            chosen_alias = {}
            unmatched = []
            report_lines = []

            resolved = 0
            for c, m_disp in enumerate(master_displays, start=1):
                disp_norm = norm(m_disp)
                aliases = MAPPING.get(disp_norm, [])
                if m_disp:
                    aliases += [m_disp]

                resolved_series = None
                resolved_alias = None
                for a in aliases:
                    a_norm = norm(a)
                    if a_norm in series_by_alias:
                        resolved_series = series_by_alias[a_norm]
                        resolved_alias = a
                        break

                if resolved_series is not None:
                    master_to_source[c] = resolved_series
                    chosen_alias[c] = resolved_alias
                    resolved += 1
                    report_lines.append(f"- ✅ **{m_disp}** ← `{resolved_alias}`")
                else:
                    if disp_norm == norm("Listing Action (List or Unlist)"):
                        master_to_source[c] = SENTINEL_LISTING_ACTION
                        report_lines.append(f"- 🟨 **{m_disp}** ← (will fill `'List'`)")
                    else:
                        unmatched.append(m_disp)
                        suggestions = top_matches(m_disp, on_headers, 3)
                        sug = ", ".join(f"`{n}` ({round(sc*100,1)}%)" for sc, n in suggestions) if suggestions else "*none*"
                        report_lines.append(f"- ❌ **{m_disp}** ← *no match*. Suggestions: {sug}")
            return resolved, master_to_source, chosen_alias, unmatched, report_lines

        if header_mode == "Pick a row number":
            h0 = int(header_row_manual) - 1
            if h0 < 0 or h0 >= len(raw_df):
                st.error("Header row number is out of range for this file.")
                st.stop()
            on_df, kept_headers = build_on_df_from_raw(raw_df, h0)
            resolved, master_to_source, chosen_alias, unmatched, report = build_mapping(on_df)
            detected_header_row = h0 + 1
        else:
            # Auto-detect among first ~10 rows, pick the one that yields most resolved columns
            best = None
            max_try = min(10, len(raw_df) - 1)
            for h0 in range(0, max_try + 1):
                try:
                    candidate_df, kept_headers = build_on_df_from_raw(raw_df, h0)
                except Exception:
                    continue
                resolved, m2s, chosen, unmatch, rep = build_mapping(candidate_df)
                score = (resolved, len(kept_headers))
                if best is None or score > best[0]:
                    best = (score, h0, candidate_df, m2s, chosen, unmatch, rep, kept_headers)

            if best is None:
                st.error("Could not auto-detect a usable header row.")
                st.stop()

            (_, _), h0, on_df, master_to_source, chosen_alias, unmatched, report, kept_headers = best
            detected_header_row = h0 + 1

        st.info(
            f"Header row used for onboarding: **Row {detected_header_row}**. "
            f"Columns kept after cleaning: **{len(on_df.columns)}**"
        )

        st.markdown("#### 🔎 Mapping Summary (Master → Onboarding)")
        st.markdown("\n".join(report))

        # -------- Write only rows that have any mapped data --------
        slog("🛠️ Writing data…")
        out_row_start = 3

        def row_has_any_data(i: int) -> bool:
            for src in master_to_source.values():
                if isinstance(src, pd.Series):
                    val = src.iloc[i] if i < len(src) else ""
                    if pd.notna(val) and str(val).strip() != "":
                        return True
            return False

        blank_streak_limit = 50
        blanks = 0
        written = 0

        for i in range(len(on_df)):
            if not row_has_any_data(i):
                blanks += 1
                if blanks >= blank_streak_limit:
                    break
                continue
            blanks = 0
            target_row = out_row_start + written
            for c in range(1, used_cols + 1):
                src = master_to_source.get(c)
                if src is SENTINEL_LISTING_ACTION:
                    master_ws.cell(row=target_row, column=c, value="List")
                elif isinstance(src, pd.Series) and i < len(src):
                    master_ws.cell(row=target_row, column=c, value=src.iloc[i])
            written += 1

        # save
        slog("💾 Saving…")
        bio = io.BytesIO()
        master_wb.save(bio)
        bio.seek(0)

        with download_area:
            st.success(f"✅ Final masterfile is ready! Rows written: **{written}**")
            st.download_button(
                "⬇️ Download Final Masterfile",
                data=bio.getvalue(),
                file_name="final_masterfile_real.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            if unmatched:
                st.info("Some master columns had no match and were left blank:\n\n- " + "\n- ".join(unmatched))
