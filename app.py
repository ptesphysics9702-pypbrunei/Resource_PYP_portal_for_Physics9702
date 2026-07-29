import io
import os
import re
import fitz  # PyMuPDF
import streamlit as st
from docx import Document
from docx.shared import Inches

# Google API Libraries
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from googleapiclient.errors import HttpError

# ==========================================
# 1. CONFIGURATION & DRIVE FOLDER MAPPING (PHYSICS 9702)
# ==========================================
SYLLABUS_CODE = "9702"

# Google Drive Folder IDs mapped to your live Google Drive folders for Physics 9702
FOLDER_IDS = {
    "theory": "YOUR_9702_THEORY_DRIVE_FOLDER_ID",      # Papers 1, 2, 4, 5
    "practical": "YOUR_9702_PRACTICAL_DRIVE_FOLDER_ID", # Paper 3 (33, 34, 35, 36)
    "zips": "YOUR_9702_ZIPS_DRIVE_FOLDER_ID"           # Practical Data / Source ZIPs
}

# Local directories for mirroring files on the server
LOCAL_FOLDERS = {
    "theory": f"{SYLLABUS_CODE}_theory",
    "practical": f"{SYLLABUS_CODE}_practical",
    "zips": f"{SYLLABUS_CODE}_zips"
}

# Ensure local storage directories exist on server startup
for folder_path in LOCAL_FOLDERS.values():
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)


# ==========================================
# 2. AUTOMATIC ROUTING & GOOGLE DRIVE API
# ==========================================
def determine_target_folder(filename: str) -> tuple[str, str]:
    """
    Analyzes the Physics 9702 filename using Regular Expressions to determine target folder.
    Returns a tuple of (folder_key, folder_display_name).
    """
    filename_lower = filename.lower()
    
    # 1. Zip / Source Files
    if filename_lower.endswith(".zip") or "_sf_" in filename_lower:
        return "zips", f"{SYLLABUS_CODE}_zips (Source / Data Files)"
        
    # 2. Practical Papers (Paper 3 variants: 33, 34, 35, 36)
    if re.search(r'_(qp|ms)_3[3456]\b', filename_lower):
        return "practical", f"{SYLLABUS_CODE}_practical (Paper 3)"
        
    # 3. Theory Papers (Papers 1, 2, 4, 5 variants: 12, 13, 22, 23, 42, 43, 52, 53)
    if re.search(r'_(qp|ms)_(1[23]|2[23]|4[23]|5[23])\b', filename_lower):
        return "theory", f"{SYLLABUS_CODE}_theory (Papers 1, 2, 4, 5)"
        
    return None, None


def build_drive_service():
    """
    Authenticates with Google Drive API using OAuth credentials stored in Streamlit Secrets.
    """
    try:
        creds = Credentials(
            token=None,
            refresh_token=st.secrets["refresh_token"],
            token_uri="https://oauth2.googleapis.com/token",
            client_id=st.secrets["client_id"],
            client_secret=st.secrets["client_secret"]
        )
        return build('drive', 'v3', credentials=creds)
    except KeyError as ke:
        st.error(f"❌ Missing Secret Key: {ke}. Please check your Streamlit Secrets configuration.")
        return None
    except Exception as e:
        st.error(f"❌ Authentication Error: {e}")
        return None


def upload_file_to_drive(file_bytes, filename, folder_id, mime_type):
    """
    Uploads file binary stream to Google Drive with HTTP error handling for quotas.
    """
    service = build_drive_service()
    if not service:
        return None

    try:
        file_stream = io.BytesIO(file_bytes)
        file_metadata = {
            'name': filename,
            'parents': [folder_id]
        }
        media = MediaIoBaseUpload(file_stream, mimetype=mime_type, resumable=True)

        uploaded_file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id, webViewLink'
        ).execute()

        return uploaded_file

    except HttpError as http_err:
        if http_err.resp.status in [403, 429]:
            st.error("❌ Drive API Quota Limit Exceeded. Please try again later.")
        else:
            st.error(f"❌ Google Drive API HTTP Error ({http_err.resp.status}): {http_err}")
        return None
    except Exception as error:
        st.error(f"❌ Drive API Upload Failed: {error}")
        return None


def sync_drive_folder_to_local(folder_key: str) -> tuple[int, str]:
    """
    Queries Google Drive for a folder and downloads missing files to local mirror.
    """
    service = build_drive_service()
    if not service:
        return 0, "Failed to authenticate with Google Drive."

    drive_folder_id = FOLDER_IDS[folder_key]
    local_path = LOCAL_FOLDERS[folder_key]

    try:
        query = f"'{drive_folder_id}' in parents and trashed = false"
        results = service.files().list(q=query, fields="files(id, name, mimeType)").execute()
        drive_files = results.get('files', [])

        downloaded_count = 0

        for file_info in drive_files:
            file_name = file_info['name']
            file_id = file_info['id']
            local_file_path = os.path.join(local_path, file_name)

            if not os.path.exists(local_file_path):
                request = service.files().get_media(fileId=file_id)
                with open(local_file_path, "wb") as f:
                    downloader = MediaIoBaseDownload(f, request)
                    done = False
                    while not done:
                        status, done = downloader.next_chunk()
                downloaded_count += 1

        return downloaded_count, f"Synced {downloaded_count} new file(s) for folder `{folder_key}`."

    except HttpError as http_err:
        return 0, f"Sync HTTP Error ({http_err.resp.status}): API limits may be reached."
    except Exception as e:
        return 0, f"Sync error on folder `{folder_key}`: {e}"


# ==========================================
# 3. SEARCH ENGINE & HELPER FUNCTIONS
# ==========================================
def render_pdf_page_preview(filepath: str, page_num: int):
    """
    Safely renders a single PDF page into PNG bytes for Streamlit previewing.
    """
    try:
        doc = fitz.open(filepath)
        page = doc.load_page(page_num)
        pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
        img_bytes = pix.tobytes("png")
        doc.close()
        return img_bytes
    except Exception as e:
        st.error(f"Unable to render page preview: {e}")
        return None


def search_pdfs(keyword_list, folder_path, allowed_variants):
    """
    Scans local PDF files for keywords using flexible, case-insensitive matching.
    """
    results = []
    if not os.path.exists(folder_path):
        return results

    cleaned_keywords = [k.strip().lower() for k in keyword_list if k.strip()]
    if not cleaned_keywords:
        return results

    for file in os.listdir(folder_path):
        if file.endswith(".pdf"):
            base_name = os.path.splitext(file)[0]
            
            is_valid_variant = any(base_name.endswith(f"_{variant}") for variant in allowed_variants)
            if not is_valid_variant:
                continue

            filepath = os.path.join(folder_path, file)
            try:
                doc = fitz.open(filepath)
                for page_num in range(len(doc)):
                    page_text = doc[page_num].get_text()
                    
                    matches_all = True
                    for kw in cleaned_keywords:
                        escaped_kw = re.escape(kw)
                        pattern = r'\b' + escaped_kw + r'(s|es)?\b'
                        
                        if not re.search(pattern, page_text, re.IGNORECASE) and kw not in page_text.lower():
                            matches_all = False
                            break
                    
                    if matches_all:
                        results.append({
                            "file": file,
                            "page": page_num,
                            "path": filepath,
                            "type": "QP" if "_qp_" in file else "MS"
                        })
                doc.close()
            except Exception:
                continue
                
    return results


# ==========================================
# 4. APP STATE INITIALIZATION
# ==========================================
if 'handout_basket' not in st.session_state:
    st.session_state.handout_basket = []
if 'theory_results' not in st.session_state:
    st.session_state.theory_results = []
if 'practical_results' not in st.session_state:
    st.session_state.practical_results = []


# ==========================================
# 5. STREAMLIT UI LAYOUT & STYLING
# ==========================================
st.set_page_config(page_title="9702 Physics Resource Platform", layout="wide")

# --- CUSTOM COLOR PALETTE STYLING ---
MAIN_BG_COLOR = "#FEE7F9"     # Main screen background (Soft Pink)
SIDEBAR_BG_COLOR = "#FCBBEF"  # Sidebar background (Matching Accent Pink)
INPUT_BAR_COLOR = "#E7FCBB"   # Input field background (Soft Lime Yellow for visibility)

st.markdown(
    f"""
    <style>
    /* Main application background */
    .stAppViewContainer {{
        background-color: {MAIN_BG_COLOR};
    }}
    
    /* Top sticky header background */
    .stHeader {{
        background-color: {MAIN_BG_COLOR};
    }}

    /* Sidebar background color */
    [data-testid="stSidebar"] {{
        background-color: {SIDEBAR_BG_COLOR};
    }}

    /* Text Inputs, Text Areas, Number Inputs */
    div[data-baseweb="input"] > div {{
        background-color: {INPUT_BAR_COLOR} !important;
        border-radius: 8px;
    }}

    /* Dropdown / Selectbox Containers */
    div[data-baseweb="select"] > div {{
        background-color: {INPUT_BAR_COLOR} !important;
        border-radius: 8px;
    }}

    /* Ensure text inside inputs remains clear and readable */
    div[data-baseweb="input"] input,
    div[data-baseweb="select"] span {{
        color: #000000 !important;
    }}
    </style>
    """,
    unsafe_allow_html=True
)

#####################################################################
st.title("BRUNEI FORM SIXTH CENTRE")
st.subheader("⚡ 9702 Physics PYP Resource Platform")

# Sidebar - Handout Basket Summary
with st.sidebar:
    st.header("Handout Basket Summary")
    st.metric(label="Saved Pages in Basket", value=len(st.session_state.handout_basket))
    if st.button("🗑️ Clear Entire Basket"):
        st.session_state.handout_basket = []
        st.rerun()

# Navigation Tabs
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🔍 Theory Search (P1, P2, P4, P5)", 
    "⚙️ Practical Search (P3)", 
    "🛒 Handout Cart", 
    "📦 Source/Data Files (ZIP)", 
    "🔒 Admin & Sync Panel"
])


# --- TAB 1: THEORY SEARCH ---
with tab1:
    st.header("Search Physics Theory Papers")
    st.caption("Variants: P1 (12, 13) | P2 (22, 23) | P4 (42, 43) | P5 (52, 53)")
    keyword_t1 = st.text_input("Enter Theory Keywords (e.g., 'Velocity', 'Quantum', 'Gravitational')", key="t1_kw")

    if st.button("Search Theory Papers", type="primary"):
        if keyword_t1:
            with st.spinner("Scanning Theory PDFs..."):
                keywords = [k.strip() for k in keyword_t1.split(",") if k.strip()]
                theory_variants = ["12", "13", "22", "23", "42", "43", "52", "53"]
                st.session_state.theory_results = search_pdfs(keywords, LOCAL_FOLDERS["theory"], theory_variants)
        else:
            st.warning("Please enter a keyword.")

    if st.session_state.theory_results:
        st.write(f"Found **{len(st.session_state.theory_results)}** matching pages:")
        for idx, item in enumerate(st.session_state.theory_results):
            doc_kind = "📝 Question Paper" if item["type"] == "QP" else "🔑 Marking Scheme"
            
            with st.expander(f"📄 {item['file']} | {doc_kind} | Page {item['page'] + 1}"):
                c1, c2 = st.columns([3, 1])
                with c1:
                    preview_img = render_pdf_page_preview(item["path"], item["page"])
                    if preview_img:
                        st.image(preview_img, caption=f"Preview Page {item['page'] + 1}", use_column_width=True)
                with c2:
                    if st.button("➕ Add to Basket", key=f"add_t1_{idx}"):
                        st.session_state.handout_basket.append(item)
                        st.toast("Added to basket!")


# --- TAB 2: PRACTICAL SEARCH ---
with tab2:
    st.header("Search Physics Practical Papers")
    st.caption("Variants: Paper 3 (33, 34, 35, 36)")
    keyword_t2 = st.text_input("Enter Practical Keywords (e.g., 'Oscillation', 'Resistance', 'Uncertainty')", key="t2_kw")

    if st.button("Search Practical Papers", type="primary"):
        if keyword_t2:
            with st.spinner("Scanning Practical PDFs..."):
                keywords = [k.strip() for k in keyword_t2.split(",") if k.strip()]
                practical_variants = ["33", "34", "35", "36"]
                st.session_state.practical_results = search_pdfs(keywords, LOCAL_FOLDERS["practical"], practical_variants)
        else:
            st.warning("Please enter a keyword.")

    if st.session_state.practical_results:
        st.write(f"Found **{len(st.session_state.practical_results)}** matching pages:")
        for idx, item in enumerate(st.session_state.practical_results):
            doc_kind = "📝 Question Paper" if item["type"] == "QP" else "🔑 Marking Scheme"
            
            with st.expander(f"📄 {item['file']} | {doc_kind} | Page {item['page'] + 1}"):
                c1, c2 = st.columns([3, 1])
                with c1:
                    preview_img = render_pdf_page_preview(item["path"], item["page"])
                    if preview_img:
                        st.image(preview_img, caption=f"Preview Page {item['page'] + 1}", use_column_width=True)
                with c2:
                    if st.button("➕ Add to Basket", key=f"add_t2_{idx}"):
                        st.session_state.handout_basket.append(item)
                        st.toast("Added to basket!")


# --- TAB 3: HANDOUT CART ---
with tab3:
    st.header("Worksheet / Handout Builder")
    if st.session_state.handout_basket:
        st.subheader(f"Selected Pages: {len(st.session_state.handout_basket)}")

        for idx, item in enumerate(st.session_state.handout_basket):
            col_info, col_remove = st.columns([4, 1])
            col_info.write(f"{idx + 1}. **{item['file']}** (Page {item['page'] + 1})")
            if col_remove.button("❌ Remove", key=f"remove_basket_{idx}"):
                st.session_state.handout_basket.pop(idx)
                st.rerun()

        st.markdown("---")
        if st.button("🪄 Export Handout to Word Document", type="primary"):
            try:
                doc = Document()
                doc.add_heading(f'PTES {SYLLABUS_CODE} Physics Handout', 0)

                for item in st.session_state.handout_basket:
                    doc.add_heading(f"Source: {item['file']} (Page {item['page'] + 1})", level=2)
                    pdf_doc = fitz.open(item['path'])
                    page = pdf_doc.load_page(item['page'])
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                    img_data = io.BytesIO(pix.tobytes("png"))
                    doc.add_picture(img_data, width=Inches(6.5))
                    doc.add_page_break()
                    pdf_doc.close()

                target_filename = f"{SYLLABUS_CODE}_Physics_Handout.docx"
                doc.save(target_filename)

                with open(target_filename, "rb") as f:
                    st.download_button("📥 Click to Download Word Handout", f, file_name=target_filename)
            except Exception as e:
                st.error(f"❌ Handout Export Failed: {e}")
    else:
        st.info("Your basket is empty. Add pages from Tab 1 or Tab 2.")


# --- TAB 4: SOURCE / DATA FILES (ZIP) ---
with tab4:
    st.header("Download Practical Data / Source Files (ZIP)")
    c1, c2, c3 = st.columns(3)
    with c1:
        z_year = st.selectbox("Select Year", [str(y) for y in range(2026, 2018, -1)])
    with c2:
        z_session = st.selectbox("Select Session", ["March (m)", "June (s)", "Nov (w)"])
        session_code = z_session.split("(")[1].replace(")", "")
    with c3:
        z_paper = st.selectbox("Select Paper Component", ["33", "34", "35", "36"])

    short_year = z_year[-2:]
    expected_zip_name = f"{SYLLABUS_CODE}_{session_code}{short_year}_sf_{z_paper}.zip"
    zip_path = os.path.join(LOCAL_FOLDERS["zips"], expected_zip_name)

    st.markdown("---")
    if os.path.exists(zip_path):
        st.success(f"Found Source File: `{expected_zip_name}`")
        with open(zip_path, "rb") as zf:
            st.download_button(
                label=f"📦 Download {expected_zip_name}",
                data=zf,
                file_name=expected_zip_name,
                mime="application/zip"
            )
    else:
        st.warning(f"File `{expected_zip_name}` is not available locally. Use Admin Sync to check Google Drive.")


# --- TAB 5: ADMIN & SYNC PANEL ---
with tab5:
    st.header("Admin & Google Drive Sync Panel")

    admin_password = st.secrets.get("ADMIN_PASSWORD")

    if not admin_password:
        st.error("🚨 `ADMIN_PASSWORD` is not configured in Streamlit Secrets.")
    else:
        pwd = st.text_input("Enter Admin Password", type="password")

        if pwd == admin_password:
            st.success("Admin Access Granted")
            
            st.subheader("🔄 Bulk Sync with Google Drive")
            if st.button("🔄 Sync All Files from Google Drive", type="primary"):
                with st.spinner("Scanning Google Drive..."):
                    total_synced = 0
                    for f_key in ["theory", "practical", "zips"]:
                        count, msg = sync_drive_folder_to_local(f_key)
                        total_synced += count
                        st.info(msg)
                    st.success(f"🎉 Sync Complete! **{total_synced}** new file(s) downloaded.")

            st.markdown("---")
            st.subheader("📤 Single File Direct Upload")
            uploaded_file = st.file_uploader("Upload Past Paper (PDF) or Source File (ZIP)", type=["pdf", "zip"])

            if uploaded_file is not None:
                folder_key, folder_name = determine_target_folder(uploaded_file.name)

                if folder_key is None:
                    st.warning(f"⚠️ Filename `{uploaded_file.name}` does not match Cambridge 9702 naming conventions.")
                else:
                    st.info(f"🎯 Target Destination: **{folder_name}**")

                    if st.button("🚀 Upload File"):
                        with st.spinner("Processing file..."):
                            file_bytes = uploaded_file.read()

                            local_dest_dir = LOCAL_FOLDERS[folder_key]
                            local_save_path = os.path.join(local_dest_dir, uploaded_file.name)
                            with open(local_save_path, "wb") as f:
                                f.write(file_bytes)

                            drive_result = upload_file_to_drive(
                                file_bytes, 
                                uploaded_file.name, 
                                FOLDER_IDS[folder_key], 
                                uploaded_file.type
                            )

                            if drive_result:
                                st.success(f"✅ Uploaded `{uploaded_file.name}` to Google Drive & local storage!")


# ==========================================
# 6. FOOTER
# ==========================================
st.markdown("---")
st.markdown(
    """
    <div style="text-align: center; width: 100%;">
        <p style="font-size: 20px; font-weight: bold; margin-bottom: 5px;">✨ Digital 9702 Physics Resource Portal ✨</p>
        <p style="color: gray; font-size: 14px;">Creator: HNHaziqah Computer Science PTES</p>
    </div>
    """,
    unsafe_allow_html=True
)
