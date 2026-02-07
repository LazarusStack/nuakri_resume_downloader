import streamlit as st
import json
import os
import tempfile
import shutil
import download_resumes
import subprocess

# Auto-install Playwright browsers for cloud deployment
try:
    import playwright
    # Simple check if chromium is available
    subprocess.run(["playwright", "install", "chromium"], check=True)
except Exception as e:
    print(f"Playwright browser installation failed: {e}")

st.set_page_config(page_title="Naukri Resume Downloader", layout="wide")

st.title("Naukri Resume Downloader")
st.markdown("Upload your candidate Excel file and paste your Naukri.com cookies to start downloading resumes.")

# File Uploader
uploaded_files = st.file_uploader("Upload Excel File(s)", type=['xlsx'], accept_multiple_files=True)

# Cookie Input
cookie_input = st.text_area("Paste Cookies (JSON format)", height=200, help="Paste the list of cookies from your browser developer tools (or EditThisCookie extension) in JSON format.")

# Log container
log_container = st.empty()

# Initialize session state
if 'zip_path' not in st.session_state:
    st.session_state['zip_path'] = None

if st.button("Start Downloading"):
    if not uploaded_files:
        st.error("Please upload at least one Excel file.")
    elif not cookie_input:
        st.error("Please paste your cookies.")
    else:
        try:
            # Parse cookies - verifying JSON format before passing
            try:
                cookies = json.loads(cookie_input) # Just to check validity
            except json.JSONDecodeError:
                 st.error("Invalid JSON format for cookies. Please ensure you pasted the correct JSON.")
                 st.stop()
            
            st.success("Cookies parsed successfully. Starting download process...")
            
            # Re-parse cookies inside the flow as requested by user logic previously
            cookies = f"""{cookie_input}"""
            cookies = json.loads(cookies)

            total_files = len(uploaded_files)
            
            try:
                with st.spinner("Downloading resumes... check your terminal for detailed logs."):
                    for i, uploaded_file in enumerate(uploaded_files):
                        st.info(f"Processing file {i+1}/{total_files}: {uploaded_file.name}")
                        
                        # Save uploaded file to a temporary file
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
                            tmp_file.write(uploaded_file.getvalue())
                            tmp_path = tmp_file.name
                        
                        try:
                            # Run the downloader for this file
                            download_resumes.run(tmp_path, cookies)
                        finally:
                             # Cleanup temp file
                            if os.path.exists(tmp_path):
                                os.unlink(tmp_path)
                
                st.success(f"All downloads complete! Resumes are saved in: {download_resumes.DOWNLOAD_DIR}")
                
                 # Create ZIP file
                shutil.make_archive('resumes_archive', 'zip', download_resumes.DOWNLOAD_DIR)
                st.session_state['zip_path'] = 'resumes_archive.zip'
                
            except Exception as e:
                st.error(f"An error occurred during execution: {e}")

        except Exception as e:
            st.error(f"Error: {e}")

if st.session_state['zip_path'] and os.path.exists(st.session_state['zip_path']):
    with open(st.session_state['zip_path'], "rb") as fp:
        st.download_button(
            label="Download All Resumes as ZIP",
            data=fp,
            file_name="resumes.zip",
            mime="application/zip"
        )

st.markdown("---")
st.markdown("### Instructions")
st.markdown("""
1. **Upload Excel**: The Excel file must have columns 'Name' and 'Candidate profile'.
2. **Get Cookies**:
    - Log in to Naukri.com.
    - Open Developer Tools (F12) -> Application -> Cookies.
    - Or use an extension like 'EditThisCookie' to export cookies as JSON.
    - Paste the JSON array into the text box above.
3. **Download**: Click the button and wait. The browser will open to download resumes.
""")
