"""Vimar Datasheet Pack Builder - Streamlit app.

Paste VIMA codes or upload an Excel list, download each product's datasheet
from vimar.com and merge everything into one PDF pack.
"""

import importlib
import os
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import streamlit as st

# Streamlit keeps already-imported modules across reruns, so editing vimar.py or
# merge.py while the app is running would otherwise leave a stale copy loaded
# (and fail on names added since). Reloading here keeps them in step.
import merge as _merge
import vimar as _vimar

importlib.reload(_vimar)
importlib.reload(_merge)

from merge import merge_pack
from vimar import (
    MANUAL_DIR,
    PLAYWRIGHT_AVAILABLE,
    browser_port_open,
    dedupe,
    download_datasheet,
    extract_codes_from_text,
    find_saved_datasheet,
    normalize_code,
    new_session,
    open_verification_browser,
    product_page_url,
    verification_status,
)

DOWNLOADS_DIR = os.path.join(os.path.expanduser("~"), "Downloads")

MAX_WORKERS = 3  # vimar.com throttles parallel traffic


def render_step(number: str, title: str, text: str) -> None:
    st.markdown(
        f"""
        <div class="process-card">
            <div class="process-number">{number}</div>
            <div>
                <div class="process-title">{title}</div>
                <div class="process-text">{text}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_metric_grid(cards: list[tuple[str, object]]) -> None:
    cells = "".join(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
        </div>
        """
        for label, value in cards
    )
    st.markdown(f'<div class="metric-grid">{cells}</div>', unsafe_allow_html=True)


# ---------------------------
# Page config
# ---------------------------
st.set_page_config(
    page_title="Vimar Datasheet Pack Builder",
    page_icon="V",
    layout="wide",
)


# ---------------------------
# Custom CSS
# ---------------------------
st.markdown(
    """
    <style>
        :root {
            --vimar-yellow: #ffc400;
            --vimar-yellow-soft: #fff5c7;
            --vimar-black: #151515;
            --vimar-ink: #202020;
            --vimar-muted: #707070;
            --vimar-line: #dedede;
            --vimar-silver: #f3f3f1;
            --vimar-warm: #eeece7;
            --vimar-panel: #ffffff;
            --vimar-shadow: rgba(20, 20, 20, 0.08);
        }

        #MainMenu,
        footer,
        header[data-testid="stHeader"] {
            visibility: hidden;
            height: 0;
        }

        .stApp {
            background:
                radial-gradient(circle at top right, rgba(255, 196, 0, 0.18), transparent 24rem),
                linear-gradient(180deg, #ffffff 0%, var(--vimar-warm) 58%, #f8f8f6 100%);
            color: var(--vimar-ink);
            font-family: Arial, Helvetica, sans-serif;
        }

        .block-container {
            padding-top: 1.2rem;
            padding-bottom: 2.5rem;
            max-width: 1180px;
        }

        .vimar-shell {
            background: rgba(255, 255, 255, 0.96);
            border: 1px solid var(--vimar-line);
            box-shadow: 0 18px 44px var(--vimar-shadow);
            margin-bottom: 1.25rem;
        }

        .utility-bar {
            display: flex;
            justify-content: flex-end;
            gap: 1.15rem;
            padding: 0.55rem 1.1rem;
            border-bottom: 1px solid var(--vimar-line);
            color: var(--vimar-muted);
            font-size: 0.78rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }

        .brand-row {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 1.25rem;
            padding: 1.05rem 1.1rem 0.95rem 1.1rem;
        }

        .brand-lockup {
            display: flex;
            align-items: center;
            gap: 0.85rem;
        }

        .brand-symbol {
            position: relative;
            width: 48px;
            height: 48px;
            border: 1px solid #b8b8b8;
            background: linear-gradient(135deg, #ffffff 0%, #c6c8c9 100%);
            box-shadow: inset 0 0 0 3px rgba(255,255,255,0.55), 0 8px 18px rgba(0,0,0,0.12);
        }

        .brand-symbol::before {
            content: "";
            position: absolute;
            left: 8px;
            right: 8px;
            top: 8px;
            height: 20px;
            border-radius: 4px 4px 2px 2px;
            background: linear-gradient(135deg, #ffe37c 0%, var(--vimar-yellow) 52%, #e8a400 100%);
            clip-path: polygon(0 0, 100% 0, 82% 100%, 18% 100%);
        }

        .brand-symbol::after {
            content: "";
            position: absolute;
            left: 8px;
            right: 8px;
            bottom: 8px;
            height: 16px;
            background: linear-gradient(135deg, #151515 0%, #6c7378 100%);
            clip-path: polygon(0 0, 50% 100%, 100% 0, 100% 100%, 0 100%);
        }

        .brand-word {
            font-size: 2.22rem;
            line-height: 0.9;
            font-weight: 900;
            color: var(--vimar-black);
            letter-spacing: 0.015em;
        }

        .brand-payoff {
            margin-top: 0.2rem;
            color: var(--vimar-black);
            font-size: 0.86rem;
            letter-spacing: 0.42em;
            font-weight: 300;
        }

        .hero-section {
            position: relative;
            overflow: hidden;
            display: grid;
            grid-template-columns: 1.25fr 0.75fr;
            gap: 1.5rem;
            align-items: stretch;
            background: linear-gradient(135deg, #f7f5ef 0%, #ffffff 55%, #e9e7e1 100%);
            border: 1px solid var(--vimar-line);
            box-shadow: 0 18px 46px var(--vimar-shadow);
            padding: 2rem;
            margin-bottom: 1rem;
        }

        .hero-kicker {
            display: inline-flex;
            align-items: center;
            gap: 0.55rem;
            color: var(--vimar-muted);
            font-size: 0.78rem;
            text-transform: uppercase;
            letter-spacing: 0.12em;
            font-weight: 800;
        }

        .hero-kicker::before {
            content: "";
            display: block;
            width: 34px;
            height: 5px;
            background: var(--vimar-yellow);
        }

        .hero-title {
            margin: 0.75rem 0 0.7rem 0;
            color: var(--vimar-black);
            font-size: clamp(2.15rem, 4vw, 4rem);
            line-height: 0.98;
            letter-spacing: -0.045em;
            font-weight: 900;
        }

        .hero-copy {
            max-width: 680px;
            color: #505050;
            font-size: 1.03rem;
            line-height: 1.72;
            margin-bottom: 1.2rem;
        }

        .hero-tags {
            display: flex;
            flex-wrap: wrap;
            gap: 0.55rem;
        }

        .hero-tag {
            background: #ffffff;
            border: 1px solid var(--vimar-line);
            border-left: 5px solid var(--vimar-yellow);
            padding: 0.55rem 0.72rem;
            font-size: 0.82rem;
            color: var(--vimar-ink);
            font-weight: 700;
        }

        .hero-visual {
            position: relative;
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 245px;
        }

        .device-card {
            width: min(100%, 330px);
            aspect-ratio: 1.72 / 1;
            border-radius: 22px;
            background: linear-gradient(145deg, #ffffff 0%, #ecebe6 100%);
            border: 1px solid #d5d5d2;
            box-shadow: 0 24px 44px rgba(0, 0, 0, 0.13);
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 0.65rem;
            padding: 1.1rem;
            transform: rotate(-2deg);
        }

        .device-switch {
            width: 66px;
            height: 118px;
            border-radius: 16px;
            background: linear-gradient(180deg, #fbfbfb 0%, #e1e1df 100%);
            border: 1px solid #c9c9c7;
            box-shadow: inset 0 1px 0 #ffffff, 0 10px 18px rgba(0,0,0,0.07);
        }

        .device-switch:nth-child(2) {
            transform: translateY(-8px);
            border-top: 7px solid var(--vimar-yellow);
        }

        .device-switch:nth-child(3) {
            transform: translateY(7px);
        }

        .process-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 0.85rem;
            margin: 0 0 1.25rem 0;
        }

        .process-card {
            display: flex;
            gap: 0.85rem;
            min-height: 98px;
            background: rgba(255,255,255,0.93);
            border: 1px solid var(--vimar-line);
            border-bottom: 4px solid var(--vimar-yellow);
            padding: 1rem;
            box-shadow: 0 12px 24px rgba(0,0,0,0.05);
        }

        .process-number {
            flex: 0 0 auto;
            width: 34px;
            height: 34px;
            display: grid;
            place-items: center;
            background: var(--vimar-black);
            color: #ffffff;
            font-weight: 900;
            font-size: 0.88rem;
        }

        .process-title {
            color: var(--vimar-black);
            font-weight: 900;
            font-size: 0.98rem;
            margin-bottom: 0.25rem;
        }

        .process-text {
            color: var(--vimar-muted);
            font-size: 0.86rem;
            line-height: 1.46;
        }

        .section-heading {
            margin: 1.15rem 0 0.75rem 0;
            padding: 0 0 0.65rem 0;
            border-bottom: 1px solid var(--vimar-line);
        }

        .section-eyebrow {
            color: var(--vimar-muted);
            font-size: 0.75rem;
            letter-spacing: 0.12em;
            text-transform: uppercase;
            font-weight: 900;
        }

        .section-title {
            color: var(--vimar-black);
            font-size: 1.42rem;
            font-weight: 900;
            margin-top: 0.15rem;
            letter-spacing: -0.02em;
        }

        .section-subtitle {
            color: var(--vimar-muted);
            font-size: 0.94rem;
            line-height: 1.6;
            margin-top: 0.2rem;
        }

        .panel-title {
            color: var(--vimar-black);
            font-size: 1.02rem;
            font-weight: 900;
            margin-bottom: 0.25rem;
        }

        .panel-title::before {
            content: "";
            display: inline-block;
            width: 9px;
            height: 9px;
            background: var(--vimar-yellow);
            margin-right: 0.45rem;
            transform: translateY(-1px);
        }

        .panel-subtitle {
            color: var(--vimar-muted);
            font-size: 0.88rem;
            margin-bottom: 0.8rem;
            line-height: 1.55;
        }

        div[data-testid="stTextArea"] textarea {
            background-color: #ffffff !important;
            border: 1px solid var(--vimar-line) !important;
            border-left: 5px solid var(--vimar-yellow) !important;
            border-radius: 0 !important;
            color: var(--vimar-ink) !important;
            font-size: 0.95rem !important;
            min-height: 200px !important;
            box-shadow: 0 14px 28px rgba(0,0,0,0.04) !important;
        }

        div[data-testid="stTextInput"] input,
        div[data-testid="stSelectbox"] div[data-baseweb="select"] {
            background-color: #ffffff !important;
            border: 1px solid var(--vimar-line) !important;
            border-radius: 0 !important;
            color: var(--vimar-ink) !important;
        }

        div[data-testid="stFileUploader"] {
            background: #ffffff !important;
            border: 1px solid var(--vimar-line) !important;
            border-left: 5px solid var(--vimar-yellow) !important;
            border-radius: 0 !important;
            padding: 22px !important;
            min-height: 190px !important;
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: 0 14px 28px rgba(0,0,0,0.04) !important;
        }

        div[data-testid="stFileUploader"] section {
            width: 100%;
        }

        div[data-testid="stTextArea"] label,
        div[data-testid="stTextInput"] label,
        div[data-testid="stCheckbox"] label,
        div[data-testid="stSelectbox"] label,
        div[data-testid="stFileUploader"] label {
            color: var(--vimar-black) !important;
            font-weight: 800 !important;
        }

        .stButton > button,
        div[data-testid="stDownloadButton"] > button {
            background: var(--vimar-black) !important;
            color: #ffffff !important;
            border: 1px solid var(--vimar-black) !important;
            border-radius: 0 !important;
            font-weight: 900 !important;
            letter-spacing: 0.04em !important;
            text-transform: uppercase !important;
            padding: 0.86rem 1rem !important;
            box-shadow: 0 14px 28px rgba(0,0,0,0.14);
        }

        .stButton > button:hover,
        div[data-testid="stDownloadButton"] > button:hover {
            background: var(--vimar-yellow) !important;
            border-color: var(--vimar-yellow) !important;
            color: var(--vimar-black) !important;
        }

        .metric-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 0.9rem;
            margin: 1.25rem 0 1rem 0;
        }

        .metric-card {
            background: #ffffff;
            border: 1px solid var(--vimar-line);
            border-top: 6px solid var(--vimar-yellow);
            padding: 1.1rem;
            box-shadow: 0 12px 24px rgba(0,0,0,0.05);
        }

        .metric-label {
            color: var(--vimar-muted);
            font-size: 0.78rem;
            text-transform: uppercase;
            letter-spacing: 0.09em;
            font-weight: 900;
            margin-bottom: 0.45rem;
        }

        .metric-value {
            color: var(--vimar-black);
            font-size: 2rem;
            font-weight: 900;
            line-height: 1.1;
        }

        .info-note {
            background: var(--vimar-yellow-soft);
            border: 1px solid #f1d861;
            color: var(--vimar-black);
            padding: 0.92rem 1rem;
            font-size: 0.93rem;
            margin-top: 0.9rem;
            line-height: 1.55;
        }

        .footer-note {
            text-align: center;
            color: var(--vimar-muted);
            font-size: 0.82rem;
            margin-top: 1.4rem;
            letter-spacing: 0.04em;
            text-transform: uppercase;
        }

        div[data-testid="stExpander"] {
            background: #ffffff;
            border: 1px solid var(--vimar-line);
            border-radius: 0;
        }

        @media (max-width: 900px) {
            .utility-bar,
            .brand-row {
                justify-content: flex-start;
            }

            .brand-row,
            .hero-section,
            .process-grid,
            .metric-grid {
                grid-template-columns: 1fr;
            }

            .hero-section {
                display: block;
                padding: 1.35rem;
            }

            .hero-visual {
                margin-top: 1rem;
            }
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------
# Header
# ---------------------------
st.markdown(
    """
    <div class="vimar-shell">
        <div class="utility-bar">
            <span>Product catalogue</span>
            <span>Datasheet pack builder</span>
        </div>
        <div class="brand-row">
            <div class="brand-lockup">
                <div class="brand-symbol"></div>
                <div>
                    <div class="brand-word">VIMAR</div>
                    <div class="brand-payoff">energia positiva</div>
                </div>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero-section">
        <div>
            <div class="hero-kicker">Product documentation</div>
            <div class="hero-title">Datasheet pack builder</div>
            <div class="hero-copy">
                Unlock vimar.com once in a real browser window, then paste item codes
                or import them from Excel to retrieve datasheet PDFs automatically and
                generate one consolidated pack ready for download.
            </div>
            <div class="hero-tags">
                <div class="hero-tag">Vimar codes</div>
                <div class="hero-tag">Excel import</div>
                <div class="hero-tag">Verified session</div>
                <div class="hero-tag">Merged pack</div>
            </div>
        </div>
        <div class="hero-visual" aria-hidden="true">
            <div class="device-card">
                <div class="device-switch"></div>
                <div class="device-switch"></div>
                <div class="device-switch"></div>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

step_col1, step_col2, step_col3 = st.columns(3)
with step_col1:
    render_step("01", "Unlock vimar.com", "Pass the one-time human check in a browser window.")
with step_col2:
    render_step("02", "Add codes", "Paste codes manually or import them from an Excel column.")
with step_col3:
    render_step("03", "Build pack", "Download and merge all retrieved datasheets in order.")


# ============================================================
# Step 1 - verified browser session
# ============================================================

st.markdown(
    """
    <div class="section-heading">
        <div class="section-eyebrow">Get started</div>
        <div class="section-title">Unlock vimar.com</div>
        <div class="section-subtitle">
            Vimar asks visitors to confirm they are human before opening the catalogue.
            Do it once in the window below and the app downloads through that same
            session. The window can stay open all day.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

verify_1, verify_2, verify_3 = st.columns([1, 1, 2], gap="medium")

with verify_1:
    if st.button("Open Vimar browser", use_container_width=True):
        started, message = open_verification_browser()
        (st.success if started else st.error)(message)

with verify_2:
    if st.button("Check verification", use_container_width=True):
        with st.spinner("Asking the browser window..."):
            ok, message = verification_status()
        st.session_state["vimar_verified"] = ok
        (st.success if ok else st.warning)(message)

with verify_3:
    if st.session_state.get("vimar_verified"):
        st.success("Session verified - downloads will use the browser window.")
    elif browser_port_open():
        st.info("Browser window is open. Pass the human check in it, then press "
                "**Check verification**.")
    else:
        st.info("No browser window yet. Press **Open Vimar browser** to start one.")

# ============================================================
# Input
# ============================================================

st.markdown(
    """
    <div class="section-heading">
        <div class="section-eyebrow">Build your PDF pack</div>
        <div class="section-title">Codes and source file</div>
        <div class="section-subtitle">
            Codes from manual input and Excel are combined automatically and duplicates
            are removed.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

left_col, right_col = st.columns(2)

with left_col:
    st.markdown(
        """
        <div class="panel-title">Paste Vimar codes</div>
        <div class="panel-subtitle">One code per line. The VIMA prefix is optional.</div>
        """,
        unsafe_allow_html=True,
    )

    codes_text = st.text_area(
        "Vimar codes",
        placeholder="Example:\nVIMA-19755.2\nVIMA-20582\nVIMA-20755.3.B\nVIMA-21860",
        height=200,
        label_visibility="collapsed",
    )

with right_col:
    st.markdown(
        """
        <div class="panel-title">Upload Excel file</div>
        <div class="panel-subtitle">Columns: Type, Code, Description. The Type is written
            on the cover page before each datasheet.</div>
        """,
        unsafe_allow_html=True,
    )

    uploaded_file = st.file_uploader(
        "Upload Excel file", type=["xlsx", "xls"], label_visibility="collapsed"
    )


def read_excel_items(upload) -> tuple[list[dict], pd.DataFrame | None, str]:
    """Read Type / Code / Description rows from the uploaded file."""
    upload.seek(0)
    df = pd.read_excel(upload)

    if df.empty:
        return [], df, "The Excel file has no rows."

    columns = list(df.columns)

    def find(keyword: str, fallback: int):
        for column in columns:
            if keyword in str(column).strip().lower():
                return column
        return columns[fallback] if len(columns) > fallback else None

    type_col = find("type", 0)
    code_col = find("code", 1)
    desc_col = find("desc", 2)

    if code_col is None:
        return [], df, "Could not find a Code column in the Excel file."

    def cell(row, column) -> str:
        if column is None:
            return ""
        value = row.get(column)
        if value is None or pd.isna(value):
            return ""
        return str(value).strip()

    items = []
    for _, row in df.iterrows():
        code = normalize_code(cell(row, code_col))
        if not code:
            continue
        items.append({
            "code": code,
            "type": cell(row, type_col),
            "title": cell(row, desc_col),
        })

    return items, df, ""


excel_items: list[dict] = []

if uploaded_file:
    try:
        excel_items, preview_df, excel_error = read_excel_items(uploaded_file)
        if preview_df is not None and not preview_df.empty:
            st.caption("Excel preview")
            st.dataframe(preview_df.head(10), use_container_width=True)
        if excel_error:
            st.error(excel_error)
    except Exception as e:
        st.error(f"Could not read the Excel file: {e}")

# ============================================================
# Settings
# ============================================================

st.markdown(
    """
    <div class="section-heading">
        <div class="section-eyebrow">Pack settings</div>
        <div class="section-title">Output and pack options</div>
        <div class="section-subtitle">
            Choose the filename and what the merged pack contains.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

set_1, set_2, set_3 = st.columns([2, 1, 1])

with set_1:
    output_filename = st.text_input("Output PDF filename", value="vimar datasheets pack.pdf")
    watch_dir = st.text_input(
        "Also look for saved PDFs in",
        value=DOWNLOADS_DIR,
        help="Datasheets you saved from vimar.com yourself. Matched by filename "
             "or by the code printed inside the PDF, so no renaming is needed.",
    )

with set_2:
    with_covers = st.checkbox("Cover page per item", value=True)
    with_toc = st.checkbox("Table of contents", value=True)

with set_3:
    use_browser = st.checkbox(
        "Browser fallback",
        value=PLAYWRIGHT_AVAILABLE,
        disabled=not PLAYWRIGHT_AVAILABLE,
        help="Retry blocked downloads inside a real browser (needs Playwright).",
    )

st.markdown(
    """
    <div class="info-note">
        The final pack starts with a cover page and Type for each item, followed by its
        datasheet, in the order the codes were entered.
    </div>
    """,
    unsafe_allow_html=True,
)

search_folders = [MANUAL_DIR] + ([watch_dir] if watch_dir.strip() else [])

# ============================================================
# Summary
# ============================================================

manual_codes = dedupe(extract_codes_from_text(codes_text))
pasted_items = [{"code": c, "type": "", "title": ""} for c in manual_codes]

all_items = pasted_items + [i for i in excel_items if i["code"] not in manual_codes]

already_saved = 0
if all_items:
    already_saved = sum(
        1 for i in all_items if find_saved_datasheet(i["code"], search_folders)[0]
    )

st.markdown(
    """
    <div class="section-heading">
        <div class="section-eyebrow">Before you download</div>
        <div class="section-title">Summary</div>
    </div>
    """,
    unsafe_allow_html=True,
)

render_metric_grid(
    [
        ("Pasted codes", len(manual_codes)),
        ("Excel items", len(excel_items)),
        ("Total items", len(all_items)),
        ("Already saved", already_saved),
    ]
)

if all_items:
    with st.expander("View detected items"):
        st.dataframe(
            pd.DataFrame([
                {"Code": i["code"], "Type (cover page)": i["type"], "Description": i["title"],
                 "Product page": product_page_url(i["code"])}
                for i in all_items
            ]),
            use_container_width=True,
        )

# ============================================================
# Download + merge
# ============================================================

run = st.button("Build PDF Pack", type="primary", use_container_width=True, disabled=not all_items)

if run:
    started = time.time()
    st.info("Downloading datasheets from vimar.com...")

    progress = st.progress(0)
    status = st.empty()
    session = new_session()

    unique_codes = dedupe([i["code"] for i in all_items])
    results: dict[str, dict] = {}

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {
            executor.submit(download_datasheet, code, session, use_browser, search_folders): code
            for code in unique_codes
        }

        for done, future in enumerate(as_completed(futures), start=1):
            code = futures[future]
            try:
                results[code] = future.result()
            except Exception as e:
                results[code] = {"code": code, "success": False, "url": "",
                                 "source": "", "error": str(e), "content": None}

            progress.progress(done / len(unique_codes))
            status.write(f"Processed {done} / {len(unique_codes)}")

    for item in all_items:
        item["result"] = results[item["code"]]

    successful = [i for i in all_items if i["result"]["success"]]
    failed = [i for i in all_items if not i["result"]["success"]]

    render_metric_grid(
        [
            ("Submitted", len(all_items)),
            ("Downloaded", len(successful)),
            ("Failed", len(failed)),
        ]
    )

    if failed:
        st.warning(
            f"**{len(failed)} datasheet(s) could not be downloaded — vimar.com refuses "
            "automated downloads.** Open each product page below, click the data sheet "
            f"to save it (anywhere in `{watch_dir}` is fine, no renaming needed), then "
            "press the button again — the app finds saved files by the code printed "
            "inside them."
        )

        st.markdown(
            "\n".join(
                f"- **{i['code']}** — [open product page]({product_page_url(i['code'])})"
                for i in failed
            )
        )

        with st.expander("Error details"):
            st.dataframe(
                pd.DataFrame([
                    {"Code": i["code"], "Type": i["type"], "Error": i["result"]["error"]}
                    for i in failed
                ]),
                use_container_width=True,
            )

    if not successful:
        st.error("No datasheets were downloaded, so no pack could be built.")
        st.stop()

    try:
        merged = merge_pack(
            [{"code": i["code"], "type": i["type"], "title": i["title"],
              "content": i["result"]["content"]} for i in successful],
            with_covers=with_covers,
            with_toc=with_toc,
        )

        filename = output_filename.strip() or "vimar datasheets pack.pdf"
        if not filename.lower().endswith(".pdf"):
            filename += ".pdf"

        st.success(f"PDF pack created in {round(time.time() - started, 1)} seconds.")

        st.download_button(
            "Download Merged PDF", data=merged, file_name=filename, mime="application/pdf",
            use_container_width=True,
        )

        with st.expander("Downloaded items"):
            st.dataframe(
                pd.DataFrame([
                    {"Code": i["code"], "Source": i["result"]["source"],
                     "URL": i["result"]["url"]}
                    for i in successful
                ]),
                use_container_width=True,
            )

    except Exception as e:
        st.error(f"Failed to merge the PDFs: {e}")

st.markdown(
    """
    <div class="footer-note">
        Built for fast retrieval and packaging of Vimar product documentation.
    </div>
    """,
    unsafe_allow_html=True,
)
