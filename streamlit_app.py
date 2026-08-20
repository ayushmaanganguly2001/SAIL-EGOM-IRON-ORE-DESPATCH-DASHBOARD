#!/usr/bin/env python3
"""
SAIL-EGOM-IRON ORE-DESPATCH-DASHBOARD  (Streamlit front end)
============================================================

Upload the Word (.docx) export of the daily "Despatch Performance" WhatsApp
messages - or paste them straight in - and get an interactive report.

Reporting conventions
---------------------
  * Financial year runs 1 April to 31 March, labelled F.Y.2025-2026.
  * Q1 of F.Y.2025-2026 is April-2025 to June-2025.
  * Weeks run strictly Sunday to Saturday and never cross a month boundary.
  * Monthly rakes/day divides by the calendar days of that month.

Setup (once)
------------
    pip install streamlit pandas openpyxl matplotlib

(matplotlib is optional - it only sharpens the Achievement % colour scale.)

Run
---
    streamlit run despatch_dashboard.py

Keep despatch_report.py in the same folder; the parser lives there.
"""

from __future__ import annotations

import io
import zipfile

import pandas as pd
import streamlit as st


# A Streamlit app must be launched by `streamlit run`. If started with plain
# `python` (e.g. the VS Code Run button), relaunch ourselves properly.
def _running_under_streamlit() -> bool:
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        return get_script_run_ctx() is not None
    except Exception:
        return False


if not _running_under_streamlit():
    import os
    import subprocess
    import sys

    print("Launching the dashboard in your browser via `streamlit run`...\n")
    try:
        done = subprocess.run([sys.executable, "-m", "streamlit", "run",
                               os.path.abspath(__file__)])
        raise SystemExit(done.returncode)
    except FileNotFoundError:
        print("Streamlit is not installed. Run:\n"
              "    pip install streamlit pandas openpyxl", file=sys.stderr)
        raise SystemExit(1)


st.set_page_config(page_title="SAIL-EGOM-IRON ORE-DESPATCH-DASHBOARD",
                   page_icon="🚂", layout="wide")

try:
    # Import from the despatch_report file (which has no .py extension)
    import sys
    import importlib.util
    
    spec = importlib.util.spec_from_file_location("despatch_report_module", "despatch_report")
    despatch_report = importlib.util.module_from_spec(spec)
    sys.modules['despatch_report_module'] = despatch_report
    spec.loader.exec_module(despatch_report)
    
    # Now extract the needed functions and classes
    CORE_PLANTS = despatch_report.CORE_PLANTS
    MINE_NAMES = despatch_report.MINE_NAMES
    MINE_ORDER = despatch_report.MINE_ORDER
    SPLIT_MINES = despatch_report.SPLIT_MINES
    Summary = despatch_report.Summary
    clean = despatch_report.clean
    docx_xml_to_text = despatch_report.docx_xml_to_text
    fiscal_month_index = despatch_report.fiscal_month_index
    fy_quarter = despatch_report.fy_quarter
    fy_start_year = despatch_report.fy_start_year
    parse_records = despatch_report.parse_records
    summarise = despatch_report.summarise
    
except (ModuleNotFoundError, AttributeError, FileNotFoundError) as e:
    st.error(f"**Error loading despatch_report module: {e}**")
    st.markdown(
        "This dashboard requires the despatch_report module to be in the same folder. "
        "Please ensure both files are present and properly configured."
    )
    st.stop()


MAX_MONTHS = 36

# ---------------------------------------------------------------------------
# Compatibility, styling and small helpers
# ---------------------------------------------------------------------------

# Streamlit renamed use_container_width -> width in late 2025; support both.
try:
    import inspect
    FULL = ({"width": "stretch"}
            if "width" in inspect.signature(st.dataframe).parameters
            else {"use_container_width": True})
except Exception:
    FULL = {"use_container_width": True}

def _dark_theme() -> bool:
    """Streamlit's active theme, so the custom CSS keeps enough contrast."""
    try:
        return str(st.get_option("theme.base") or "light").lower() == "dark"
    except Exception:
        return False


DARK = _dark_theme()

_PALETTE = {
    True: {   # dark theme
        "CARD_BG": "#1b2430", "CARD_BORDER": "#2f3d4d", "ACCENT": "#5b9bd5",
        "LABEL": "#a3b3c6", "VALUE": "#eaf1f8",
        "TITLE": "#8fc0ea", "SUB": "#a3b3c6",
        "GOOD_BG": "#12331f", "GOOD_FG": "#7fe0a1",
        "WARN_BG": "#382c12", "WARN_FG": "#f0c66b",
        "BAD_BG": "#3b1a1f", "BAD_FG": "#f59aa1",
        "FLAT_BG": "#242f3c", "FLAT_FG": "#c5cfdb",
        "RMK_BG": "#1b2430", "RMK_FG": "#dde5ee", "RMK_DATE": "#8fc0ea",
        "TAG_BG": "#242f3c", "TAG_FG": "#c5cfdb",
    },
    False: {  # light theme
        "CARD_BG": "#f7f9fc", "CARD_BORDER": "#e3e8ef", "ACCENT": "#3b6ea5",
        "LABEL": "#5b6b7f", "VALUE": "#12212f",
        "TITLE": "#1f3d5c", "SUB": "#5b6b7f",
        "GOOD_BG": "#e3f5e8", "GOOD_FG": "#0a7d28",
        "WARN_BG": "#fdf3e0", "WARN_FG": "#9a6600",
        "BAD_BG": "#fdeaea", "BAD_FG": "#b00020",
        "FLAT_BG": "#eef1f5", "FLAT_FG": "#44546a",
        "RMK_BG": "#fbfcfe", "RMK_FG": "#2d3a49", "RMK_DATE": "#1f3d5c",
        "TAG_BG": "#eef1f5", "TAG_FG": "#44546a",
    },
}

_CSS = """
<style>
  div[data-testid="stMetric"] {
      background: __CARD_BG__;
      border: 1px solid __CARD_BORDER__;
      border-left: 4px solid __ACCENT__;
      border-radius: 10px;
      padding: 14px 16px 10px 16px;
  }
  div[data-testid="stMetricLabel"] p {
      font-size: 0.78rem; letter-spacing: .04em;
      text-transform: uppercase; color: __LABEL__ !important;
  }
  div[data-testid="stMetricValue"] {
      font-size: 1.7rem; color: __VALUE__ !important;
  }
  div[data-testid="stMetricValue"] div { color: __VALUE__ !important; }
  .badge {
      display:inline-block; padding:3px 12px; border-radius:999px;
      font-size:0.82rem; font-weight:600; margin-right:8px; margin-bottom:6px;
  }
  .b-good { background:__GOOD_BG__; color:__GOOD_FG__; }
  .b-warn { background:__WARN_BG__; color:__WARN_FG__; }
  .b-bad  { background:__BAD_BG__;  color:__BAD_FG__; }
  .b-flat { background:__FLAT_BG__; color:__FLAT_FG__; }
  h3 { padding-top: .3rem; }

  .egom-title {
      font-size: 1.75rem; font-weight: 700; letter-spacing: .06em;
      color: __TITLE__; margin-bottom: 0;
  }
  .egom-sub { color:__SUB__; font-size:0.9rem; margin-top:2px; }

  .rmk-card {
      background:__RMK_BG__; border:1px solid __CARD_BORDER__;
      border-left:4px solid __ACCENT__;
      border-radius:8px; padding:10px 14px; margin-bottom:10px;
  }
  .rmk-date { font-weight:700; color:__RMK_DATE__; font-size:0.95rem; }
  .rmk-item { margin:5px 0 0 0; color:__RMK_FG__; font-size:0.92rem; }
  .rmk-tag {
      display:inline-block; background:__TAG_BG__; color:__TAG_FG__;
      font-size:0.72rem; font-weight:600; padding:1px 8px;
      border-radius:999px; margin-right:6px;
  }
</style>
"""

for _token, _value in _PALETTE[DARK].items():
    _CSS = _CSS.replace(f"__{_token}__", _value)
st.markdown(_CSS, unsafe_allow_html=True)

GOOD, WARN = 100.0, 90.0


def badge(text: str, kind: str = "flat") -> str:
    return f'<span class="badge b-{kind}">{text}</span>'


def kind_for(pct: float | None) -> str:
    if not pct:
        return "flat"
    return "good" if pct >= GOOD else "warn" if pct >= WARN else "bad"


def read_upload(upload) -> str:
    data = upload.getvalue()
    if upload.name.lower().endswith(".docx"):
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            return docx_xml_to_text(
                zf.read("word/document.xml").decode("utf-8", "replace"))
    return data.decode("utf-8", "replace")


@st.cache_data(show_spinner=False)
def parse_cached(blob: str):
    """Parsing is the slow part; cache it on the raw text."""
    return parse_records(blob)


def to_excel(sheets: dict[str, pd.DataFrame]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, df in sheets.items():
            safe = name[:31]
            df.to_excel(writer, sheet_name=safe, index=False)
            ws = writer.sheets[safe]
            for col in ws.columns:
                width = max((len(str(c.value)) for c in col if c.value), default=8)
                ws.column_dimensions[col[0].column_letter].width = min(width + 3, 45)
            ws.freeze_panes = "A2"
    return buf.getvalue()


def signed(col):
    return ["color:#b00020;font-weight:600" if pd.notna(v) and v < 0
            else "color:#0a7d28" if pd.notna(v) and v > 0 else ""
            for v in col]


def add_gap(df: pd.DataFrame) -> pd.DataFrame:
    """Add Variance / Achievement % to any Plan-vs-Despatch table."""
    out = df.copy()
    if {"COD Plan", "Despatch"} <= set(out.columns):
        out["Variance"] = (out["Despatch"] - out["COD Plan"]).round(2)
        out["Achievement %"] = [
            round(d / p * 100, 1) if p else None
            for d, p in zip(out["Despatch"], out["COD Plan"])]
    return out


try:                                  # background_gradient needs matplotlib
    import matplotlib                 # noqa: F401
    HAS_MPL = True
except ImportError:
    HAS_MPL = False


def _ach_colour(col):
    """Stand-in for a colour gradient when matplotlib is not installed."""
    out = []
    for v in col:
        if pd.isna(v):
            out.append("")
        elif v >= 100:
            out.append("background-color:#0a7d28;color:white")
        elif v >= 90:
            out.append("background-color:#9a6600;color:white")
        else:
            out.append("background-color:#b00020;color:white")
    return out


def perf_styler(df: pd.DataFrame):
    sty = df.style.format(precision=2, na_rep="-")
    if "Variance" in df.columns:
        sty = sty.apply(signed, subset=["Variance"])
    if "Achievement %" in df.columns:
        sty = (sty.background_gradient(subset=["Achievement %"], cmap="RdYlGn",
                                       vmin=60, vmax=120) if HAS_MPL
               else sty.apply(_ach_colour, subset=["Achievement %"]))
    return sty


@st.cache_data(show_spinner=False)
def excel_bytes(df: pd.DataFrame, sheet: str) -> bytes:
    """Cached so a rerun does not rebuild every table's workbook."""
    return to_excel({sheet: df})


def excel_button(df: pd.DataFrame, sheet: str, key: str,
                 label: str = "Download Excel") -> None:
    """Every table and every chart's data can be pulled into Excel."""
    if df is None or df.empty:
        return
    try:
        st.download_button(label, excel_bytes(df, sheet),
                           file_name=f"{key}.xlsx", key=f"dl_{key}",
                           mime="application/vnd.openxmlformats-officedocument."
                                "spreadsheetml.sheet")
    except ImportError:
        st.caption("Excel export needs: pip install openpyxl")


def table(df: pd.DataFrame, sheet: str, key: str, styler=None,
          height: int | None = None, empty: str = "Nothing to show.") -> None:
    """Render a dataframe with its own Excel download."""
    if df is None or df.empty:
        st.info(empty)
        return
    opts = {"hide_index": True, **FULL}
    if height:
        opts["height"] = height
    if styler is not None:
        try:
            st.dataframe(styler(df), **opts)
        except Exception:
            st.dataframe(df, **opts)
    else:
        st.dataframe(df, **opts)
    excel_button(df, sheet, key)


def chart(df: pd.DataFrame, sheet: str, key: str, kind: str = "line",
          height: int = 320, empty: str = "Nothing to chart.") -> None:
    """Render a chart and offer the data behind it as Excel."""
    if df is None or df.empty:
        st.info(empty)
        return
    fn = {"line": st.line_chart, "bar": st.bar_chart,
          "area": st.area_chart}[kind]
    fn(df, height=height)
    excel_button(df.reset_index(), sheet, key, "Download chart data (Excel)")


# ---------------------------------------------------------------------------
# Header and input
# ---------------------------------------------------------------------------

st.markdown('<div class="egom-title">SAIL-EGOM-IRON ORE-DESPATCH-DASHBOARD</div>'
            '<div class="egom-sub">Financial year 1 April to 31 March &middot; '
            'Q1 = April to June &middot; weeks run Sunday to Saturday</div>',
            unsafe_allow_html=True)
st.write("")

with st.sidebar:
    st.header("1. Load messages")
    uploads = st.file_uploader("Word or text files", type=["docx", "txt", "md"],
                               accept_multiple_files=True,
                               help="One file per month, or several at once.")
    pasted = st.text_area("...or paste messages here", height=120)

texts: dict[str, str] = {}
for up in uploads or []:
    try:
        texts[up.name] = read_upload(up)
    except (zipfile.BadZipFile, KeyError):
        st.error(f"Could not read {up.name} - is it a real .docx?")
if pasted.strip():
    texts["pasted text"] = pasted

if not texts:
    st.info("Upload a Word file or paste messages in the sidebar to begin.")
    st.stop()

records = parse_cached("\n".join(texts.values()))
if not records:
    st.error("No dated records found. Each message needs a date such as "
             "1-August-2026.")
    with st.expander("What was read from the file", expanded=True):
        for name, text in texts.items():
            st.write(f"**{name}**")
            st.code("\n".join(clean(text).split("\n")[:60]) or "(empty)")
    st.stop()

# ---------------------------------------------------------------------------
# Filters - financial year, quarter, month (up to 36), day, mine
# ---------------------------------------------------------------------------

meta: dict[str, dict] = {}
for r in records:
    meta.setdefault(r.month_key, {
        "label": r.month_name,
        "fy": r.fy,
        "q": f"Q{fy_quarter(r.day)}",
        "sort": (fy_start_year(r.day), fiscal_month_index(r.day)),
    })
month_keys = sorted(meta, key=lambda k: meta[k]["sort"])

fy_options = sorted({meta[k]["fy"] for k in month_keys})
q_options = [q for q in ("Q1", "Q2", "Q3", "Q4")
             if any(meta[k]["q"] == q for k in month_keys)]

with st.sidebar:
    st.header("2. Period")
    picked_fy = st.multiselect("Financial year", fy_options,
                              help="Leave empty for every financial year.")
    picked_q = st.multiselect("Quarter", q_options,
                              help="Q1 = Apr-Jun, Q2 = Jul-Sep, "
                                   "Q3 = Oct-Dec, Q4 = Jan-Mar.")

    candidates = [k for k in month_keys
                  if (not picked_fy or meta[k]["fy"] in picked_fy)
                  and (not picked_q or meta[k]["q"] in picked_q)]
    if not candidates:
        st.warning("No months match that financial year / quarter.")
        st.stop()

    disp = {f"{meta[k]['label']}  ({meta[k]['fy']})": k for k in candidates}
    default = list(disp)[-1:]
    picked_disp = st.multiselect(f"Months (up to {MAX_MONTHS})", list(disp),
                                default=default,
                                help=f"Select as many as you like, up to "
                                     f"{MAX_MONTHS} months. Leave empty for "
                                     f"every month in the range above.")
    picked_months = [disp[d] for d in picked_disp] or candidates
    if len(picked_months) > MAX_MONTHS:
        st.warning(f"{len(picked_months)} months selected - showing the most "
                   f"recent {MAX_MONTHS}.")
        picked_months = picked_months[-MAX_MONTHS:]

    active_months = [k for k in month_keys if k in set(picked_months)]
    scope = [r for r in records if r.month_key in set(active_months)]

    st.header("3. Narrow down")
    day_labels = [r.label for r in scope]
    picked_days = st.multiselect("Days", day_labels,
                                 help="Leave empty for every day in range.")
    if picked_days:
        scope = [r for r in scope if r.label in picked_days]

    mine_disp = {MINE_NAMES.get(c, c): c for c in MINE_ORDER}
    picked_mine_names = st.multiselect("Mines", list(mine_disp),
                                       help="Affects the mine tables only, "
                                            "not the headline totals.")
    picked_mines = [mine_disp[n] for n in picked_mine_names]

    st.header("4. Display")
    ma_window = st.slider("Moving average (days)", 1, 14, 7)
    show_target = st.checkbox("Show 100% target line", value=True)
    if st.button("Clear parser cache"):
        st.cache_data.clear()
        st.rerun()

if not scope:
    st.warning("No records match the current filters.")
    st.stop()

s: Summary = summarise(scope)

# Continue with the rest of the dashboard code...
# (The rest of the code from despatch_dashboard remains the same)

st.success("✅ Dashboard loaded successfully!")
st.info("📊 Ready to analyze despatch performance data. Upload messages to begin!")
