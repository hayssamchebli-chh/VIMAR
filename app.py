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


st.set_page_config(
    page_title="Vimar Datasheet Pack Builder",
    page_icon="📘",
    layout="centered",
)

st.markdown(
    """
    <style>
    .main-title {
        font-size: 2.2rem;
        font-weight: 700;
        margin-bottom: 0.3rem;
    }
    .sub-text {
        color: #6b7280;
        font-size: 1rem;
        margin-bottom: 1.5rem;
    }
    .summary-box {
        padding: 1rem;
        border-radius: 12px;
        background-color: #f8f9fb;
        border: 1px solid #e6e8ec;
        margin-top: 1rem;
        margin-bottom: 1rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="main-title">Vimar Datasheet Pack Builder</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-text">Enter one or more Vimar item codes, download their PDFs, and combine them into a single document.</div>',
    unsafe_allow_html=True,
)

with st.container():
    codes_text = st.text_area(
        "Item codes",
        height=220,
        placeholder="Example:\n01423\nK40930\n20211.B",
        help="Paste one code per line, or separate them with commas, spaces, or semicolons.",
    )

    col1, col2 = st.columns([1, 1])
    with col1:
        keep_going = st.checkbox("Skip failed codes", value=True)
    with col2:
        output_name = st.text_input("Output file name", value="vimar_datasheet_pack.pdf")

    run_clicked = st.button("Build PDF Pack", type="primary", use_container_width=True)

if run_clicked:
    codes = normalize_codes(codes_text.splitlines())

    if not codes:
        st.error("Please enter at least one item code.")
    else:
        session = get_session()
        downloaded_pdfs = []
        success_rows = []
        failed_codes = []

        st.markdown(
            f"""
            <div class="summary-box">
                <strong>Total codes submitted:</strong> {len(codes)}
            </div>
            """,
            unsafe_allow_html=True,
        )

        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, code in enumerate(codes, start=1):
            status_text.info(f"Processing code {idx}/{len(codes)}: {code}")
            ok, pdf_bytes, used_url = download_pdf_bytes_for_code(session, code)

            if ok and pdf_bytes:
                downloaded_pdfs.append(pdf_bytes)
                success_rows.append({"Code": code, "Status": "Downloaded", "Source URL": used_url})
            else:
                failed_codes.append(code)
                if not keep_going:
                    progress_bar.progress(idx / len(codes))
                    break

            progress_bar.progress(idx / len(codes))

        status_text.empty()

        success_count = len(downloaded_pdfs)
        failed_count = len(failed_codes)

        col1, col2, col3 = st.columns(3)
        col1.metric("Submitted", len(codes))
        col2.metric("Downloaded", success_count)
        col3.metric("Failed", failed_count)

        if success_count == 0:
            st.error("No PDFs were downloaded, so no merged file could be created.")
        else:
            try:
                merged_pdf = merge_pdf_bytes(downloaded_pdfs)

                st.success("Your PDF pack is ready.")

                st.download_button(
                    label="Download merged PDF",
                    data=merged_pdf,
                    file_name=output_name.strip() or "vimar_datasheet_pack.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )

                with st.expander("View downloaded items", expanded=False):
                    if success_rows:
                        st.dataframe(success_rows, use_container_width=True)

                if failed_codes:
                    with st.expander("View failed codes", expanded=True):
                        st.warning(", ".join(failed_codes))

            except Exception as e:
                st.error(f"Failed to merge PDFs: {e}")
