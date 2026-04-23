import re
from io import BytesIO
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


def merge_pdf_bytes(pdf_byte_list: List[bytes]) -> bytes:
    writer = PdfWriter()

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


def process_code(code: str) -> dict:
    session = get_session()
    ok, pdf_bytes, used_url = download_pdf_bytes_for_code(session, code)

    return {
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
        future_to_code = {executor.submit(process_code, code): code for code in codes}

        completed = 0
        progress_bar = st.progress(0)
        status_text = st.empty()

        for future in as_completed(future_to_code):
            result = future.result()
            completed += 1

            code = result["code"]
            status_text.info(f"Processed {completed} of {len(codes)} — {code}")
            progress_bar.progress(completed / len(codes))

            if result["ok"] and result["pdf_bytes"]:
                downloaded_pdfs.append(result["pdf_bytes"])
                success_rows.append(
                    {
                        "Code": code,
                        "Status": "Downloaded",
                        "Source URL": result["used_url"],
                    }
                )
            else:
                failed_codes.append(code)

            results.append(result)

        status_text.empty()

    return downloaded_pdfs, success_rows, failed_codes, results


# ---------------------------
# Page config
# ---------------------------
st.set_page_config(
    page_title="Vimar Datasheet Pack Builder",
    page_icon="📘",
    layout="centered",
)


# ---------------------------
# Custom CSS
# ---------------------------
st.markdown(
    """
    <style>
        :root {
            --navy: #1f2a44;
            --navy-2: #2f3d5c;
            --navy-soft: #e9eef7;
            --grey-bg: #f3f4f6;
            --grey-card: #ffffff;
            --grey-border: #d1d5db;
            --grey-muted: #6b7280;
        }

        .stApp {
            background-color: var(--grey-bg);
        }

        .block-container {
            padding-top: 2rem;
            padding-bottom: 2rem;
            max-width: 1000px;
        }

        .hero-card {
            background: linear-gradient(135deg, var(--navy) 0%, var(--navy-2) 100%);
            color: white;
            border-radius: 18px;
            padding: 1.6rem 1.6rem 1.4rem 1.6rem;
            box-shadow: 0 10px 25px rgba(31, 42, 68, 0.18);
            margin-bottom: 1.2rem;
        }

        .hero-badge {
            display: inline-block;
            font-size: 0.78rem;
            font-weight: 600;
            letter-spacing: 0.4px;
            background: rgba(255,255,255,0.14);
            border: 1px solid rgba(255,255,255,0.16);
            padding: 0.35rem 0.7rem;
            border-radius: 999px;
            margin-bottom: 0.8rem;
        }

        .hero-title {
            font-size: 2rem;
            font-weight: 700;
            line-height: 1.2;
            margin-bottom: 0.45rem;
        }

        .hero-subtitle {
            font-size: 1rem;
            color: #e5e7eb;
            line-height: 1.55;
            margin-bottom: 0;
        }

        .section-card {
            background: var(--grey-card);
            border: 1px solid var(--grey-border);
            border-radius: 16px;
            padding: 1.2rem 1.2rem 1rem 1.2rem;
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.04);
            margin-bottom: 1rem;
        }

        .section-title {
            font-size: 1.05rem;
            font-weight: 700;
            color: var(--navy);
            margin-bottom: 0.2rem;
        }

        .section-subtitle {
            font-size: 0.93rem;
            color: var(--grey-muted);
            margin-bottom: 0.8rem;
        }

        .panel-title {
            color: var(--navy);
            font-size: 1rem;
            font-weight: 700;
            margin-bottom: 0.25rem;
        }

        .panel-subtitle {
            color: var(--grey-muted);
            font-size: 0.88rem;
            margin-bottom: 0.8rem;
        }

        div[data-testid="stTextArea"] textarea {
            background-color: #fbfbfc !important;
            border: 1px solid var(--grey-border) !important;
            border-radius: 12px !important;
            color: #111827 !important;
            font-size: 0.95rem !important;
            min-height: 220px !important;
        }

        div[data-testid="stTextInput"] input,
        div[data-testid="stSelectbox"] div[data-baseweb="select"] {
            background-color: #fbfbfc !important;
            border-radius: 12px !important;
        }

        div[data-testid="stFileUploader"] {
            background: #fbfbfc !important;
            border: 2px dashed #b9c2d0 !important;
            border-radius: 12px !important;
            padding: 24px !important;
            min-height: 220px !important;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        div[data-testid="stFileUploader"] section {
            width: 100%;
        }

        div[data-testid="stTextArea"] label,
        div[data-testid="stTextInput"] label,
        div[data-testid="stCheckbox"] label,
        div[data-testid="stSelectbox"] label,
        div[data-testid="stFileUploader"] label {
            color: var(--navy) !important;
            font-weight: 600 !important;
        }

        .stButton > button,
        div[data-testid="stDownloadButton"] > button {
            background: var(--navy) !important;
            color: white !important;
            border: none !important;
            border-radius: 12px !important;
            font-weight: 600 !important;
            padding: 0.7rem 1rem !important;
            box-shadow: 0 8px 18px rgba(31, 42, 68, 0.14);
        }

        .stButton > button:hover,
        div[data-testid="stDownloadButton"] > button:hover {
            background: #172033 !important;
        }

        .metric-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 0.9rem;
            margin: 1rem 0 1rem 0;
        }

        .metric-card {
            background: white;
            border: 1px solid var(--grey-border);
            border-radius: 14px;
            padding: 1rem;
            box-shadow: 0 6px 16px rgba(15, 23, 42, 0.04);
        }

        .metric-label {
            color: var(--grey-muted);
            font-size: 0.85rem;
            margin-bottom: 0.35rem;
        }

        .metric-value {
            color: var(--navy);
            font-size: 1.7rem;
            font-weight: 700;
            line-height: 1.1;
        }

        .info-note {
            background: var(--navy-soft);
            border: 1px solid #d7deea;
            color: var(--navy);
            border-radius: 12px;
            padding: 0.85rem 1rem;
            font-size: 0.93rem;
            margin-top: 0.5rem;
        }

        .footer-note {
            text-align: center;
            color: var(--grey-muted);
            font-size: 0.85rem;
            margin-top: 1rem;
        }

        div[data-testid="stExpander"] {
            background: white;
            border: 1px solid var(--grey-border);
            border-radius: 12px;
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
    <div class="hero-card">
        <div class="hero-badge">PDF AUTOMATION TOOL</div>
        <div class="hero-title">Vimar Datasheet Pack Builder</div>
        <div class="hero-subtitle">
            Enter Vimar item codes, retrieve their datasheet PDFs automatically,
            and generate one consolidated PDF pack ready for download.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ---------------------------
# Input section
# ---------------------------
st.markdown(
    """
    <div class="section-card">
        <div class="section-title">Build your PDF pack</div>
        <div class="section-subtitle">
            Add codes manually or upload an Excel file and select the column containing the item codes.
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
        height=220,
        placeholder="Example:\n01423\nK40930\nABC-20211.B",
        label_visibility="collapsed",
    )

    manual_codes = normalize_codes(codes_text.splitlines())

with input_col2:
    st.markdown(
        """
        <div class="panel-title">Upload Excel file</div>
        <div class="panel-subtitle">
            Drag and drop your Excel file here, or browse to upload it.
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

col1, col2 = st.columns([1, 1])
with col1:
    keep_going = st.checkbox("Skip failed codes and continue", value=True)
# with col2:
    output_name = st.text_input("Output file name", value="vimar_datasheet_pack.pdf")

st.markdown(
    """
    <div class="info-note">
        Codes from manual input and Excel are combined automatically and duplicates are removed.
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
                merged_pdf = merge_pdf_bytes(downloaded_pdfs)

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
