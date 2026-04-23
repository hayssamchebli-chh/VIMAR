import re
from io import BytesIO
from typing import Iterable, List, Tuple

import requests
import streamlit as st
from pypdf import PdfReader, PdfWriter

BASE_URLS = [
    "https://www.vimar.com/en/int/catalog/product/download-pdf/code/{code}?type=.pdf",
    "https://www.vimar.com/en/int/catalog/obsolete/download-pdf/code/{code}?type=.pdf",
    "https://www.vimar.com/en/int/catalog/document/download-pdf/code/{code}?type=.pdf",
]

DEFAULT_TIMEOUT = 120


# ---------------------------
# Helpers
# ---------------------------
def normalize_codes(raw_codes: Iterable[str]) -> List[str]:
    codes = []
    for item in raw_codes:
        if not item:
            continue
        parts = re.split(r"[\s,;]+", item.strip())
        for part in parts:
            part = part.strip()
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
        try:
            response = session.get(url, timeout=DEFAULT_TIMEOUT, allow_redirects=True)
            if response.status_code == 200 and looks_like_pdf(response):
                return True, response.content, url
        except requests.RequestException:
            pass

    return False, None, None


def merge_pdf_bytes(pdf_byte_list: List[bytes]) -> bytes:
    writer = PdfWriter()

    for pdf_bytes in pdf_byte_list:
        reader = PdfReader(BytesIO(pdf_bytes))
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


# ---------------------------
# Page config
# ---------------------------
st.set_page_config(
    page_title="Vimar Datasheet Pack Builder",
    page_icon="📘",
    layout="centered",
)


# ---------------------------
# Custom CSS (Grey + Navy Theme)
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
            --grey-text: #4b5563;
            --grey-muted: #6b7280;
            --success-bg: #eef6f1;
            --warning-bg: #fff7ed;
        }

        .stApp {
            background-color: var(--grey-bg);
        }

        .block-container {
            padding-top: 2rem;
            padding-bottom: 2rem;
            max-width: 900px;
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

        div[data-testid="stTextArea"] textarea {
            background-color: #fbfbfc !important;
            border: 1px solid var(--grey-border) !important;
            border-radius: 12px !important;
            color: #111827 !important;
            font-size: 0.95rem !important;
        }

        div[data-testid="stTextInput"] input {
            background-color: #fbfbfc !important;
            border: 1px solid var(--grey-border) !important;
            border-radius: 12px !important;
            color: #111827 !important;
        }

        div[data-testid="stTextArea"] label,
        div[data-testid="stTextInput"] label,
        div[data-testid="stCheckbox"] label {
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
            Paste one code per line, or separate codes with commas, spaces, or semicolons.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

codes_text = st.text_area(
    "Item codes",
    height=220,
    placeholder="Example:\n01423\nK40930\n20211.B",
    label_visibility="collapsed",
)

col1, col2 = st.columns([1, 1])
with col1:
    keep_going = st.checkbox("Skip failed codes and continue", value=True)
with col2:
    output_name = st.text_input("Output file name", value="vimar_datasheet_pack.pdf")

st.markdown(
    """
    <div class="info-note">
        The application will try multiple Vimar PDF endpoints for each code and merge all successful results into one file.
    </div>
    """,
    unsafe_allow_html=True,
)

run_clicked = st.button("Build PDF Pack", type="primary", use_container_width=True)


# ---------------------------
# Action / Processing
# ---------------------------
if run_clicked:
    codes = normalize_codes(codes_text.splitlines())

    if not codes:
        st.error("Please enter at least one item code.")
    else:
        session = get_session()
        downloaded_pdfs = []
        success_rows = []
        failed_codes = []

        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, code in enumerate(codes, start=1):
            status_text.info(f"Processing code {idx} of {len(codes)} — {code}")
            ok, pdf_bytes, used_url = download_pdf_bytes_for_code(session, code)

            if ok and pdf_bytes:
                downloaded_pdfs.append(pdf_bytes)
                success_rows.append(
                    {
                        "Code": code,
                        "Status": "Downloaded",
                        "Source URL": used_url,
                    }
                )
            else:
                failed_codes.append(code)
                if not keep_going:
                    progress_bar.progress(idx / len(codes))
                    break

            progress_bar.progress(idx / len(codes))

        status_text.empty()

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
