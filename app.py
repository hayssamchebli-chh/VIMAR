import re
from io import BytesIO
from pathlib import Path
from typing import Iterable, List, Tuple

import pandas as pd
import requests
import streamlit as st
from pypdf import PdfReader, PdfWriter
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE_URLS = [
    "https://www.vimar.com/en/int/catalog/product/download-pdf/code/{code}?type=.pdf",
    "https://www.vimar.com/en/int/catalog/obsolete/download-pdf/code/{code}?type=.pdf",
    "https://www.vimar.com/en/int/catalog/document/download-pdf/code/{code}?type=.pdf",
]

DEFAULT_TIMEOUT = (20, 40)
PDF_DOWNLOAD_RETRIES = 2

APP_DIR = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
DEFAULT_COVER_PATHS = [
    APP_DIR / "cover.pdf",
    APP_DIR / "cover",
    APP_DIR / "cover" / "cover.pdf",
]


# ---------------------------
# Helpers
# ---------------------------
def clean_code(value: str) -> str:
    value = str(value).strip()
    if not value:
        return ""

    if "-" in value:
        value = value.split("-", 1)[1].strip()

    return value


def normalize_codes(raw_codes: Iterable[str]) -> List[str]:
    codes = []

    for item in raw_codes:
        if item is None:
            continue

        item_str = str(item).strip()
        if not item_str:
            continue

        parts = re.split(r"[\s,;]+", item_str)
        for part in parts:
            part = clean_code(part)
            if part:
                codes.append(part)

    seen = set()
    unique_codes = []
    for code in codes:
        if code not in seen:
            seen.add(code)
            unique_codes.append(code)

    return unique_codes


def looks_like_pdf(response: requests.Response) -> bool:
    content_type = (response.headers.get("Content-Type") or "").lower()
    content_disp = (response.headers.get("Content-Disposition") or "").lower()

    if "pdf" in content_type or ".pdf" in content_disp:
        return True

    return response.content[:5] == b"%PDF-"


def is_valid_pdf_bytes(pdf_bytes: bytes) -> bool:
    try:
        reader = PdfReader(BytesIO(pdf_bytes), strict=False)
        _ = len(reader.pages)
        return True
    except Exception:
        return False


def read_default_cover_pdf_bytes() -> bytes | None:
    for cover_path in DEFAULT_COVER_PATHS:
        if cover_path.is_file():
            cover_pdf_bytes = cover_path.read_bytes()
            if is_valid_pdf_bytes(cover_pdf_bytes):
                return cover_pdf_bytes

    return None


def get_cover_pdf_bytes(uploaded_cover) -> Tuple[bytes | None, str | None]:
    if uploaded_cover is not None:
        cover_pdf_bytes = uploaded_cover.getvalue()
        if is_valid_pdf_bytes(cover_pdf_bytes):
            return cover_pdf_bytes, None

        return None, "The uploaded cover file is not a valid PDF."

    cover_pdf_bytes = read_default_cover_pdf_bytes()
    if cover_pdf_bytes is not None:
        return cover_pdf_bytes, None

    return (
        None,
        "Default cover PDF was not found or is invalid. "
        "Add cover.pdf to the repository root, or upload a custom cover PDF.",
    )


def get_session() -> requests.Session:
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": "Mozilla/5.0 (compatible; VimarPdfAutomation/1.0)",
            "Accept": "application/pdf,application/octet-stream;q=0.9,*/*;q=0.8",
        }
    )
    return session


def download_pdf_bytes_for_code(
    session: requests.Session, code: str
) -> Tuple[bool, bytes | None, str | None]:
    for url_template in BASE_URLS:
        url = url_template.format(code=code)

        for _ in range(PDF_DOWNLOAD_RETRIES):
            try:
                response = session.get(url, timeout=DEFAULT_TIMEOUT, allow_redirects=True)
                if response.status_code == 200 and looks_like_pdf(response):
                    pdf_bytes = response.content
                    if is_valid_pdf_bytes(pdf_bytes):
                        return True, pdf_bytes, url
            except requests.RequestException:
                pass

    return False, None, None


def merge_pdf_bytes(pdf_byte_list: List[bytes], cover_pdf_bytes: bytes | None = None) -> bytes:
    writer = PdfWriter()

    if cover_pdf_bytes:
        cover_reader = PdfReader(BytesIO(cover_pdf_bytes), strict=False)
        for page in cover_reader.pages:
            writer.add_page(page)

    for pdf_bytes in pdf_byte_list:
        reader = PdfReader(BytesIO(pdf_bytes), strict=False)
        for page in reader.pages:
            writer.add_page(page)

    output = BytesIO()
    writer.write(output)
    output.seek(0)
    return output.getvalue()


def ensure_pdf_filename(filename: str) -> str:
    filename = filename.strip() or "vimar_datasheet_pack.pdf"
    if not filename.lower().endswith(".pdf"):
        filename += ".pdf"
    return filename


def read_excel_file(uploaded_file) -> pd.DataFrame:
    return pd.read_excel(uploaded_file)


def extract_codes_from_selected_column(df: pd.DataFrame, selected_column: str) -> List[str]:
    if selected_column not in df.columns:
        return []

    values = df[selected_column].dropna().astype(str).tolist()
    return normalize_codes(values)


def process_code(index: int, code: str) -> dict:
    session = get_session()
    ok, pdf_bytes, used_url = download_pdf_bytes_for_code(session, code)

    return {
        "index": index,
        "code": code,
        "ok": ok,
        "pdf_bytes": pdf_bytes,
        "used_url": used_url,
    }


def download_pdfs_parallel(codes: List[str], max_workers: int = 8):
    downloaded_pdfs = []
    success_rows = []
    failed_codes = []

    results = []

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_code = {
            executor.submit(process_code, index, code): code
            for index, code in enumerate(codes)
        }

        completed = 0
        progress_bar = st.progress(0)
        status_text = st.empty()

        for future in as_completed(future_to_code):
            result = future.result()
            completed += 1

            code = result["code"]
            status_text.info(f"Processed {completed} of {len(codes)} - {code}")
            progress_bar.progress(completed / len(codes))

            results.append(result)

        status_text.empty()

    # Sort back to the original input order
    results.sort(key=lambda x: x["index"])

    for result in results:
        if result["ok"] and result["pdf_bytes"]:
            downloaded_pdfs.append(result["pdf_bytes"])
            success_rows.append(
                {
                    "Code": result["code"],
                    "Status": "Downloaded",
                    "Source URL": result["used_url"],
                }
            )
        else:
            failed_codes.append(result["code"])

    return downloaded_pdfs, success_rows, failed_codes, results


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

        .search-pill {
            display: inline-flex;
            align-items: center;
            gap: 0.45rem;
            border: 1px solid var(--vimar-line);
            background: var(--vimar-silver);
            color: var(--vimar-muted);
            border-radius: 999px;
            padding: 0.62rem 0.95rem;
            font-size: 0.88rem;
            min-width: 210px;
            justify-content: space-between;
        }

        .search-dot {
            width: 10px;
            height: 10px;
            border-radius: 50%;
            background: var(--vimar-yellow);
            box-shadow: 0 0 0 5px rgba(255, 196, 0, 0.18);
        }

        .nav-row {
            display: flex;
            flex-wrap: wrap;
            gap: 0;
            border-top: 1px solid var(--vimar-line);
        }

        .nav-item {
            padding: 0.82rem 1.05rem;
            border-right: 1px solid var(--vimar-line);
            color: var(--vimar-ink);
            font-size: 0.85rem;
            text-transform: uppercase;
            letter-spacing: 0.045em;
            font-weight: 700;
        }

        .nav-item:first-child {
            background: var(--vimar-yellow);
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
            min-height: 232px !important;
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
            grid-template-columns: repeat(3, 1fr);
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
            .brand-row,
            .nav-row {
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

            .search-pill {
                display: none;
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
            <span>Work with us</span>
            <span>MyVIMAR</span>
        </div>
        <div class="brand-row">
            <div class="brand-lockup">
                <div class="brand-symbol" aria-hidden="true"></div>
                <div>
                    <div class="brand-word">VIMAR</div>
                    <div class="brand-payoff">energia positiva</div>
                </div>
            </div>
            <div class="search-pill">
                <span>Search on the site</span>
                <span class="search-dot"></span>
            </div>
        </div>
        <div class="nav-row">
            <div class="nav-item">Products</div>
            <div class="nav-item">Solutions</div>
            <div class="nav-item">Services for professionals</div>
            <div class="nav-item">News &amp; documentation</div>
            <div class="nav-item">Contacts</div>
            <div class="nav-item">Company</div>
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
                Enter Vimar item codes, retrieve datasheet PDFs automatically,
                add the cover page, and generate one consolidated PDF pack ready for download.
            </div>
            <div class="hero-tags">
                <div class="hero-tag">Vimar codes</div>
                <div class="hero-tag">Excel import</div>
                <div class="hero-tag">Cover PDF</div>
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
    render_step("01", "Add codes", "Paste codes manually or import them from an Excel column.")
with step_col2:
    render_step("02", "Choose cover", "Use the repository cover or upload a custom PDF cover.")
with step_col3:
    render_step("03", "Build pack", "Download and merge all retrieved datasheets in order.")


# ---------------------------
# Input section
# ---------------------------
st.markdown(
    """
    <div class="section-heading">
        <div class="section-eyebrow">Build your PDF pack</div>
        <div class="section-title">Codes and source file</div>
        <div class="section-subtitle">
            Codes from manual input and Excel are combined automatically and duplicates are removed.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

manual_codes = []
excel_codes = []
excel_df = None
uploaded_excel = None

input_col1, input_col2 = st.columns(2)

with input_col1:
    st.markdown(
        """
        <div class="panel-title">Paste item codes</div>
        <div class="panel-subtitle">
            Enter one code per line, or separate them with commas, spaces, or semicolons.
        </div>
        """,
        unsafe_allow_html=True,
    )

    codes_text = st.text_area(
        "Paste item codes",
        height=232,
        placeholder="Example:\n00200\nK40930\nVIM-09208.C",
        label_visibility="collapsed",
    )

    manual_codes = normalize_codes(codes_text.splitlines())

with input_col2:
    st.markdown(
        """
        <div class="panel-title">Upload Excel file</div>
        <div class="panel-subtitle">
            Drag and drop your Excel file here, then choose the column containing item codes.
        </div>
        """,
        unsafe_allow_html=True,
    )

    uploaded_excel = st.file_uploader(
        "Upload Excel file",
        type=["xlsx", "xls"],
        label_visibility="collapsed",
    )

    if uploaded_excel is not None:
        try:
            excel_df = read_excel_file(uploaded_excel)

            if excel_df.empty:
                st.warning("The uploaded Excel file is empty.")
            else:
                column_options = excel_df.columns.tolist()

                default_index = 0
                if "Item No.1" in column_options:
                    default_index = column_options.index("Item No.1")

                selected_column = st.selectbox(
                    "Select the column containing item codes",
                    options=column_options,
                    index=default_index,
                )

                excel_codes = extract_codes_from_selected_column(excel_df, selected_column)
                st.caption(f"{len(excel_codes)} code(s) detected from Excel.")

        except Exception as e:
            st.error(f"Could not read Excel file: {e}")

st.markdown(
    """
    <div class="section-heading">
        <div class="section-eyebrow">Pack settings</div>
        <div class="section-title">Cover and output</div>
        <div class="section-subtitle">
            Leave the cover field empty to use the default cover PDF included in the repository.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

settings_col1, settings_col2 = st.columns(2)
with settings_col1:
    keep_going = st.checkbox("Skip failed codes and continue", value=True)
    output_name = st.text_input("Output file name", value="vimar_datasheet_pack.pdf")

with settings_col2:
    uploaded_cover = st.file_uploader(
        "Use another cover page (optional)",
        type=["pdf"],
        help="Leave empty to use the default cover PDF included in the repository.",
    )

st.markdown(
    """
    <div class="info-note">
        The final PDF uses the selected cover first, followed by the downloaded Vimar datasheets.
    </div>
    """,
    unsafe_allow_html=True,
)

run_clicked = st.button("Build PDF Pack", type="primary", use_container_width=True)


# ---------------------------
# Action / Processing
# ---------------------------
if run_clicked:
    codes = normalize_codes(manual_codes + excel_codes)

    if not codes:
        st.error("Please enter item codes manually or upload an Excel file.")
    else:
        cover_pdf_bytes, cover_error = get_cover_pdf_bytes(uploaded_cover)
        if cover_error:
            st.error(cover_error)
            st.stop()

        max_workers = min(8, max(1, len(codes)))

        downloaded_pdfs, success_rows, failed_codes, _ = download_pdfs_parallel(
            codes=codes,
            max_workers=max_workers,
        )
        submitted_count = len(codes)
        downloaded_count = len(downloaded_pdfs)
        failed_count = len(failed_codes)

        st.markdown(
            f"""
            <div class="metric-grid">
                <div class="metric-card">
                    <div class="metric-label">Submitted</div>
                    <div class="metric-value">{submitted_count}</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Downloaded</div>
                    <div class="metric-value">{downloaded_count}</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Failed</div>
                    <div class="metric-value">{failed_count}</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if downloaded_count == 0:
            st.error("No PDFs were downloaded, so no merged file could be created.")
        else:
            try:
                merged_pdf = merge_pdf_bytes(downloaded_pdfs, cover_pdf_bytes=cover_pdf_bytes)

                st.success("Your consolidated PDF pack is ready.")

                st.download_button(
                    label="Download Merged PDF",
                    data=merged_pdf,
                    file_name=ensure_pdf_filename(output_name),
                    mime="application/pdf",
                    use_container_width=True,
                )

                with st.expander("Downloaded items", expanded=False):
                    st.dataframe(success_rows, use_container_width=True)

                if failed_codes:
                    with st.expander("Failed codes", expanded=True):
                        st.warning(", ".join(failed_codes))

            except Exception as e:
                st.error(f"Failed to merge PDFs: {e}")

st.markdown(
    """
    <div class="footer-note">
        Built for fast retrieval and packaging of Vimar product documentation.
    </div>
    """,
    unsafe_allow_html=True,
)

