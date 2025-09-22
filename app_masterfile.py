# app_masterfile.py

import io
import json
import re
from difflib import SequenceMatcher

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="Masterfile Automation", page_icon="📦", layout="wide")

# =========================
# Helpers
# =========================
def norm(s: str) -> str:
    """Normalize header strings for robust matching."""
    if s is None:
        return ""
    x = str(s).strip().lower()
    # strip "- en-us" & variants
    x = re.sub(r"\s*-\s*en\s*[-_ ]\s*us\s*$", "", x)
    # normalize dashes
    x = x.replace("–", "-").replace("—", "-").replace("−", "-")
    # replace common separators with a space (keep '-' last in class)
    x = re.sub(r"[._/\\-]+", " ", x)
    # drop anything not alnum or space
    x = re.sub(r"[^0-9a-z\s]+", " ", x)
    # collapse spaces
    return re.sub(r"\s+", " ", x).strip()


def top_matches(query, candidates, k=3):
    q = norm(query)
    scored = [(SequenceMatcher(None, q, norm(c)).ratio(), c) for c in candidates]
    scored.sort(key=lambda t: t[0], reverse=True)
    return scored[:k]


def worksheet_used_cols(ws, header_rows=(1,), hard_cap=512, empty_streak_stop=8):
    """Heuristically detect meaningful column span by scanning header rows."""
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
    """Make duplicate headers unique by adding suffixes."""
    seen = {}
    out = []
    for h in headers:
        h = "" if h is None else str(h)
        key = h
        if key in seen:
            seen[key] += 1
            out.append(f"{h}__{seen[key]}")
        else:
            seen[key] = 0
            out.append(h)
    return out


# Unique sentinel for the special "Listing Action" fill
SENTINEL_LISTING_ACTION = object()

# =========================
# UI
# =========================
st.title("📦 Masterfile Automation")
st.caption("Auto-detects onboarding header row. Map onboarding columns to master template headers and generate a ready-to-upload masterfile.")

with st.expander("ℹ️ Quick guide", expanded=True):
    st.markdown("""
- **Masterfile template (.xlsx)**  
  - Row **1** = display labels  
  - Row **2** = internal keys/helper labels  
  - Data is written starting at **Row 3** (template styles are preserved)

- **Onboarding sheet (.xlsx)**  
  - The app **auto-detects the header row** (it tries the first 10 rows and picks the one that maps best).  
  - Data is everything **after** the detected header row.

- **Mapping JSON**  
  - Keys = **Master display headers** (Row 1 in master)  
  - Values = **list of onboarding header aliases** in priority order.  
  - The **first alias that exists** is used.  
    """)

st.divider()

colA, colB = st.columns([1, 1])
with colA:
    masterfile_file = st.file_uploader("📄 Upload Masterfile Template (.xlsx)", type=["xlsx"])
with colB:
    onboarding_file = st.file_uploader("🧾 Upload Onboarding Sheet (.xlsx)", type=["xlsx"])

st.markdown("#### 🔗 Mapping JSON")
mapping_tab1, mapping_tab2 = st.tabs(["Paste JSON", "Upload JSON file"])
mapping_json_text = ""
mapping_json_file = None
with mapping_tab1:
    mapping_json_text = st.text_area(
        "Paste mapping JSON here",
        height=220,
        placeholder='{\n  "Partner SKU": ["Seller SKU", "item_sku"]\n}',
    )
with mapping_tab2:
    mapping_json_file = st.file_uploader("Or upload mapping.json", type=["json"], key="mapping_file")

st.divider()
go = st.button("🚀 Generate Final Masterfile", type="primary")

log_area = st.container()
download_area = st.container()

# =========================
# Main Action
# =========================
if go:
    with log_area:
        st.markdown("### 📝 Log")
        log = st.empty()
        def slog(msg): log.markdown(msg)

        # Validate inputs
        if not masterfile_file or not onboarding_file:
            st.error("Please upload both **Masterfile Template** and **Onboarding Sheet**.")
            st.stop()

        # Parse mapping JSON
        mapping_raw = None
        if mapping_json_text.strip():
            try:
                mapping_raw = json.loads(mapping_json_text)
            except Exception as e:
                st.error(f"Mapping JSON could not be parsed. Error: {e}")
                st.stop()
        elif mapping_json_file is not None:
            try:
                mapping_raw = json.load(mapping_json_file)
            except Exception as e:
                st.error(f"Mapping JSON file could not be parsed. Error: {e}")
                st.stop()
        else:
            st.error("Please provide mapping JSON (paste or upload).")
            st.stop()

        # Normalize mapping keys
        MAPPING = {}
        for k, v in mapping_raw.items():
            MAPPING[norm(k)] = v[:] if isinstance(v, list) else [v]

        slog("⏳ Reading master (preserving styles)…")
        try:
            master_wb = load_workbook(masterfile_file, keep_links=False)
            master_ws = master_wb.active
        except Exception as e:
            st.error(f"Could not read **Masterfile**: {e}")
            st.stop()

        # --- Read onboarding raw (no header) and auto-detect header row ---
        slog("⏳ Reading onboarding (auto-detecting header row)…")
        try:
            raw_df = pd.read_excel(onboarding_file, header=None, dtype=str)
            raw_df = raw_df.fillna("")
        except Exception as e:
            st.error(f"Could not read **Onboarding**: {e}")
            st.stop()

        # Master headers (Row 1 display)
        used_cols = worksheet_used_cols(master_ws, header_rows=(1, 2))
        master_displays = [master_ws.cell(row=1, column=c).value or "" for c in range(1, used_cols + 1)]

        def build_mapping_for_df(df):
            """Given an onboarding df with proper headers, build mapping, return stats+details."""
            on_headers = list(df.columns)
            series_by_alias = {norm(h): df[h] for h in on_headers}

            master_to_source = {}
            chosen_alias = {}
            unmatched = []
            report = []

            resolved_count = 0
            for c, m_disp in enumerate(master_displays, start=1):
                disp_norm = norm(m_disp)
                aliases = []
                aliases += MAPPING.get(disp_norm, [])
                if m_disp:
                    aliases.append(m_disp)

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
                    resolved_count += 1
                    report.append(f"- ✅ **{m_disp}** ← `{resolved_alias}`")
                else:
                    if disp_norm == norm("Listing Action (List or Unlist)"):
                        master_to_source[c] = SENTINEL_LISTING_ACTION
                        report.append(f"- 🟨 **{m_disp}** ← (will fill `'List'`)")
                    else:
                        unmatched.append(m_disp)
                        suggestions = top_matches(m_disp, on_headers, 3)
                        sug_txt = ", ".join(f"`{name}` ({round(sc*100,1)}%)" for sc, name in suggestions) if suggestions else "*none*"
                        report.append(f"- ❌ **{m_disp}** ← *no match*. Suggestions: {sug_txt}")
            return resolved_count, master_to_source, chosen_alias, unmatched, report

        # Try header rows 0..9 (first 10 rows)
        best = None
        best_header_row = None
        max_try = min(10, len(raw_df)-1)
        for h in range(0, max_try+1):
            headers = list(raw_df.iloc[h].astype(str))
            headers = ["" if x.lower() == "nan" else x for x in headers]
            headers = uniquify_headers(headers)
            candidate_df = raw_df.iloc[h+1:].copy()
            candidate_df.columns = headers
            candidate_df = candidate_df.fillna("")

            resolved_count, m2s, chosen, unmatch, rep = build_mapping_for_df(candidate_df)

            # score by resolved columns; tiebreaker = number of non-empty headers
            nonempty_headers = sum(1 for hh in headers if str(hh).strip())
            score = (resolved_count, nonempty_headers)

            if (best is None) or (score > best[0]):
                best = ((resolved_count, nonempty_headers), m2s, chosen, unmatch, rep, candidate_df)
                best_header_row = h

        if best is None:
            st.error("Could not detect a valid header row in onboarding.")
            st.stop()

        (resolved_count, _), master_to_source, chosen_alias, unmatched, report_lines, on_df = best

        st.info(f"✅ Detected onboarding header row: **Row {best_header_row+1}** (1-based). "
                f"Resolved **{resolved_count}** master columns.")

        st.markdown("#### 🔎 Mapping Summary (Master → Onboarding)")
        st.markdown("\n".join(report_lines))

        # Write values to master starting row 3
        slog("🛠️ Writing data…")
        out_row = 3
        num_rows = len(on_df)

        for i in range(num_rows):
            for c in range(1, used_cols + 1):
                src = master_to_source.get(c, None)
                if src is None:
                    continue
                if src is SENTINEL_LISTING_ACTION:
                    master_ws.cell(row=out_row + i, column=c, value="List")
                elif isinstance(src, pd.Series):
                    if i < len(src):
                        master_ws.cell(row=out_row + i, column=c, value=src.iloc[i])

        # Save to buffer
        slog("💾 Saving…")
        bio = io.BytesIO()
        master_wb.save(bio)
        bio.seek(0)

        with download_area:
            st.success("✅ Final masterfile is ready!")
            st.download_button(
                "⬇️ Download Final Masterfile",
                data=bio.getvalue(),
                file_name="final_masterfile_real.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            if unmatched:
                st.info(
                    "Some master columns had no match and were left blank:\n\n- " +
                    "\n- ".join(unmatched)
                )
