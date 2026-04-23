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


st.set_page_config(page_title="Vimar PDF Merger", page_icon="📄", layout="centered")

st.title("Vimar PDF Merger")
st.write("Paste Vimar item codes, download their PDFs, and merge them into one file.")

codes_text = st.text_area(
    "Item codes",
    height=180,
    placeholder="Example:\n01423\nK40930\n20211.B",
)

keep_going = st.checkbox("Skip failed codes and continue", value=True)

if st.button("Download and merge PDFs", type="primary"):
    codes = normalize_codes(codes_text.splitlines())

    if not codes:
        st.error("Please provide at least one item code.")
    else:
        session = get_session()
        downloaded_pdfs = []
        results = []
        failed_codes = []

        progress = st.progress(0)
        status_box = st.empty()

        for idx, code in enumerate(codes, start=1):
            status_box.write(f"Checking code: {code}")
            ok, pdf_bytes, used_url = download_pdf_bytes_for_code(session, code)

            if ok and pdf_bytes:
                downloaded_pdfs.append(pdf_bytes)
                results.append((code, "Downloaded", used_url))
            else:
                failed_codes.append(code)
                results.append((code, "Failed", None))
                if not keep_going:
                    break

            progress.progress(idx / len(codes))

        status_box.empty()

        st.subheader("Results")
        for code, status, url in results:
            if status == "Downloaded":
                st.success(f"{code} — downloaded")
                if url:
                    st.caption(url)
            else:
                st.error(f"{code} — failed")

        if not downloaded_pdfs:
            st.error("No PDFs were downloaded, so nothing could be merged.")
        else:
            try:
                merged_pdf = merge_pdf_bytes(downloaded_pdfs)

                st.success(f"Merged {len(downloaded_pdfs)} PDF(s) successfully.")

                if failed_codes:
                    st.warning("Failed codes: " + ", ".join(failed_codes))

                st.download_button(
                    label="Download merged PDF",
                    data=merged_pdf,
                    file_name="vimar_combined.pdf",
                    mime="application/pdf",
                )
            except Exception as e:
                st.error(f"Failed to merge PDFs: {e}")
