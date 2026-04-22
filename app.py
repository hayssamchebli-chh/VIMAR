import streamlit as st
import requests
import zipfile
import io

def download_vimar_pdfs(codes):
    # Create an in-memory zip file
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for code in codes:
            code = code.strip()
            if not code:
                continue
                
            url = f"https://www.vimar.com/en/int/catalog/product/download-pdf/code/{code}?type=.pdf"
            
            try:
                response = requests.get(url, timeout=100)
                # Check if the request was successful and it's actually a PDF
                if response.status_code == 200 and 'application/pdf' in response.headers.get('Content-Type', ''):
                    zip_file.writestr(f"{code}.pdf", response.content)
                else:
                    st.warning(f"Could not find PDF for code: {code}")
            except Exception as e:
                st.error(f"Error downloading {code}: {e}")
                
    return zip_buffer.getvalue()

# --- Streamlit UI ---
st.title("Vimar PDF Bulk Downloader")
st.write("Enter product codes (one per line) to download their technical sheets.")

# Input area
input_codes = st.text_area("Product Codes", placeholder="e.g.\n19001\n16005\n14008")

if st.button("Prepare Download"):
    code_list = input_codes.split('\n')
    code_list = [c.strip() for c in code_list if c.strip()]
    
    if not code_list:
        st.error("Please enter at least one code.")
    else:
        with st.spinner(f"Downloading {len(code_list)} files..."):
            zip_data = download_vimar_pdfs(code_list)
            
            if zip_data:
                st.success("Files ready!")
                st.download_button(
                    label="Download ZIP Archive",
                    data=zip_data,
                    file_name="vimar_products.zip",
                    mime="application/zip"
                )