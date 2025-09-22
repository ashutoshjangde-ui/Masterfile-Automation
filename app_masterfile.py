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
    x = str(s).strip().lower()
    x = re.sub(r"\s*-\s*en\s*[-_ ]\s*us\s*$", "", x)
    x = x.replace("–", "-").replace("—", "-").replace("−", "-")
    x = re.sub(r"[._/\\-]+", " ", x)           # separators -> space
    x = re.sub(r"[^0-9a-z\s]+", " ", x)        # keep alnum/space
    return re.sub(r"\s+", " ", x).strip()      # collapse spaces

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

def uniquify_headers(headers):
    seen = {}
    out = []
    for h in headers:
        h = "" if h is None else str(h)
        if h in seen:
            seen[h] += 1
            out.append(f"{h}__{seen[h]}")
        else:
            seen[h] = 0
            out.append(h)
    return out

# sentinel for “Listing Action”
SENTINEL_LISTING_ACTION = object()

# ---------------- UI ----------------
st.title("📦 Masterfile Automation")
st.caption("Auto-detect or choose onboarding header row. Only writes rows that actually contain data.")

with st.expander("ℹ️ Quick guide", expanded=True):
    st.markdown("""
- **Masterfile**: Row 1 = labels, Row 2 = keys, data starts Row 3 (styles preserved).
- **Onboarding**: You can **Auto-detect** the header row or **pick it manually** (1-based). Data is below that header row.
- **Mapping JSON**: keys = master labels; values = list of onboarding aliases in order of priority.
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

        # onboarding raw (no header), then pick header row
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

        # get onboarding df according to header mode
        if header_mode == "Pick a row number":
            h0 = int(header_row_manual) - 1  # 0-based
            if h0 < 0 or h0 >= len(raw_df):
                st.error("Header row number is out of range for this file.")
                st.stop()
            headers = uniquify_headers(list(raw_df.iloc[h0].astype(str)))
            on_df = raw_df.iloc[h0 + 1:].copy()
            on_df.columns = headers
            on_df = on_df.fillna("")
            detected_header_row = h0 + 1
            resolved, master_to_source, chosen_alias, unmatched, report = build_mapping(on_df)
        else:
            # auto-detect among first up to 10 rows
            best = None
            max_try = min(10, len(raw_df) - 1)
            for h0 in range(0, max_try + 1):
                headers = uniquify_headers(list(raw_df.iloc[h0].astype(str)))
                candidate = raw_df.iloc[h0 + 1:].copy()
                candidate.columns = headers
                candidate = candidate.fillna("")

                resolved, m2s, chosen, unmatch, rep = build_mapping(candidate)
                nonempty_headers = sum(1 for hh in headers if str(hh).strip())
                score = (resolved, nonempty_headers)
                if best is None or score > best[0]:
                    best = ((resolved, nonempty_headers), h0, candidate, m2s, chosen, unmatch, rep)

            (_, _), h0, on_df, master_to_source, chosen_alias, unmatched, report = best
            detected_header_row = h0 + 1

        st.info(f"Header row used for onboarding: **Row {detected_header_row}** (1-based). "
                f"Resolved **{sum(isinstance(v, pd.Series) for v in master_to_source.values())}** columns.")

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

        blank_streak_limit = 50  # hard stop if we see many consecutive blanks
        blanks = 0
        written = 0

        for i in range(len(on_df)):
            if not row_has_any_data(i):
                blanks += 1
                if blanks >= blank_streak_limit:
                    break
                continue
            blanks = 0  # reset once we hit a row with data
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
