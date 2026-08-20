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

st.set_page_config(page_title="Vimar Datasheet Pack Builder", page_icon="🔌", layout="wide")

st.markdown(
    """
<style>
    :root {
        --vimar-red: #E2001A;
        --vimar-dark: #1A1A1A;
        --app-bg: #F6F8FB;
        --text-main: #102033;
        --text-muted: #64748B;
        --border-soft: #DDE6F0;
    }

    .stApp {
        background:
            radial-gradient(circle at top left, rgba(226, 0, 26, 0.10), transparent 30%),
            linear-gradient(180deg, #ffffff 0%, var(--app-bg) 60%);
        color: var(--text-main);
    }

    .block-container { padding-top: 70px; padding-bottom: 60px; max-width: 1180px; }

    .brand-topbar {
        display: flex;
        justify-content: center;
        align-items: center;
        margin-bottom: 22px;
    }

    .vimar-logo {
        background: var(--vimar-red);
        color: #ffffff;
        border: 1px solid var(--vimar-red);
        border-radius: 999px;
        padding: 10px 26px;
        font-weight: 800;
        font-size: 20px;
        letter-spacing: 3px;
        line-height: 1;
        box-shadow: 0 8px 22px rgba(226, 0, 26, 0.22);
    }

    .hero {
        position: relative;
        overflow: hidden;
        padding: 32px;
        border-radius: 28px;
        background: linear-gradient(135deg, var(--vimar-dark) 0%, #3A0A10 55%, var(--vimar-red) 100%);
        color: white;
        margin-bottom: 26px;
        box-shadow: 0 22px 46px rgba(226, 0, 26, 0.20);
    }

    .hero-kicker {
        text-transform: uppercase;
        letter-spacing: 1.9px;
        font-size: 12px;
        font-weight: 800;
        opacity: 0.85;
        margin-bottom: 10px;
    }

    .hero h1 { margin: 0 0 12px 0; font-size: 40px; line-height: 1.08; font-weight: 850; }
    .hero p { margin: 0; font-size: 16px; line-height: 1.6; opacity: 0.94; }

    .tool-card {
        background: rgba(255, 255, 255, 0.92);
        border: 1px solid var(--border-soft);
        border-radius: 22px;
        padding: 20px;
        box-shadow: 0 14px 36px rgba(15, 23, 42, 0.07);
        margin-bottom: 16px;
    }

    .section-title { font-size: 18px; font-weight: 850; color: var(--vimar-dark); margin-bottom: 6px; }
    .section-subtitle { color: var(--text-muted); font-size: 14px; margin-bottom: 12px; }

    div[data-testid="stMetric"] {
        background: white;
        border: 1px solid var(--border-soft);
        border-radius: 20px;
        padding: 16px;
        box-shadow: 0 12px 28px rgba(15, 23, 42, 0.06);
    }

    .stTextArea textarea, .stTextInput input { border-radius: 14px !important; }

    .stFileUploader section,
    .stFileUploader [data-testid="stFileUploaderDropzone"] {
        min-height: 190px; display: flex; align-items: center;
    }

    .stButton > button {
        background: linear-gradient(135deg, var(--vimar-red) 0%, #8E0010 100%);
        color: white; border: 0; border-radius: 999px;
        padding: 12px 26px; font-weight: 850;
        box-shadow: 0 14px 28px rgba(226, 0, 26, 0.26);
    }
    .stButton > button:hover { transform: translateY(-1px); color: white; }

    .stDownloadButton > button {
        background: #ffffff; color: var(--vimar-dark);
        border: 1px solid var(--vimar-red); border-radius: 999px;
        padding: 12px 26px; font-weight: 850;
    }
</style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
<div class="brand-topbar"><div class="vimar-logo">VIMAR</div></div>

<div class="hero">
    <div class="hero-kicker">Product documentation tool</div>
    <h1>Vimar Datasheet Pack Builder</h1>
    <p>Paste Vimar codes or upload an Excel list, download the datasheets from
       vimar.com and merge everything into one PDF.</p>
</div>
""",
    unsafe_allow_html=True,
)

# ============================================================
# Step 1 - verified browser session
# ============================================================

st.markdown(
    """
<div class="tool-card">
    <div class="section-title">Step 1 &mdash; unlock vimar.com</div>
    <div class="section-subtitle">
        Vimar asks visitors to confirm they are human before opening the catalogue.
        Do it once in the window below and the app downloads through that same
        session. The window can stay open all day.
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

st.markdown("</div>", unsafe_allow_html=True)

# ============================================================
# Input
# ============================================================

left_col, right_col = st.columns([1, 1], gap="large")

with left_col:
    st.markdown(
        """
<div class="tool-card">
    <div class="section-title">Paste Vimar codes</div>
    <div class="section-subtitle">One code per line. The VIMA prefix is optional.</div>
""",
        unsafe_allow_html=True,
    )

    codes_text = st.text_area(
        "Vimar codes",
        placeholder="Example:\nVIMA-19755.2\nVIMA-20582\nVIMA-20755.3.B\nVIMA-21860",
        height=200,
        label_visibility="collapsed",
    )

    st.markdown("</div>", unsafe_allow_html=True)

with right_col:
    st.markdown(
        """
<div class="tool-card">
    <div class="section-title">Upload Excel file</div>
    <div class="section-subtitle">Columns: Type, Code, Description. The Type is written
        on the cover page before each datasheet.</div>
""",
        unsafe_allow_html=True,
    )

    uploaded_file = st.file_uploader(
        "Upload Excel file", type=["xlsx", "xls"], label_visibility="collapsed"
    )

    st.markdown("</div>", unsafe_allow_html=True)


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
<div class="tool-card">
    <div class="section-title">Export settings</div>
    <div class="section-subtitle">Choose the filename and what the merged pack contains.</div>
""",
    unsafe_allow_html=True,
)

set_1, set_2, set_3 = st.columns([2, 1, 1], gap="large")

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

st.markdown("</div>", unsafe_allow_html=True)

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

st.markdown("### Summary before download")

m1, m2, m3, m4 = st.columns(4)
with m1:
    st.metric("Pasted codes", len(manual_codes))
with m2:
    st.metric("Excel items", len(excel_items))
with m3:
    st.metric("Total items", len(all_items))
with m4:
    st.metric("Already saved", already_saved)

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

run = st.button("Download and merge datasheets", type="primary", disabled=not all_items)

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

    st.divider()

    r1, r2, r3 = st.columns(3)
    with r1:
        st.metric("Submitted", len(all_items))
    with r2:
        st.metric("Downloaded", len(successful))
    with r3:
        st.metric("Failed", len(failed))

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
            "Download merged PDF", data=merged, file_name=filename, mime="application/pdf"
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
