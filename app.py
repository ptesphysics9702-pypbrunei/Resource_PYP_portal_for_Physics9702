import io
import os
import re
import fitz  # PyMuPDF
import streamlit as st

# Word Document Libraries
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Google API Libraries
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from googleapiclient.errors import HttpError

# ==========================================
# 1. CONFIGURATION & DRIVE FOLDER MAPPING (PHYSICS 9702)
# ==========================================
SYLLABUS_CODE = "9702"

# Google Drive Folder IDs mapped to your live Google Drive folders
FOLDER_IDS = {
    "theory": "180eJGPSt53c0u3Fx-39G8ltcopB2eT-b",    # Papers 1, 2, 4, 5
    "practical": "10yymDTMshSmyjBB5VGhRp6TjGbS3QsK7", # Paper 3 (33, 34, 35, 36)
    "zips": "17I6Cm9H_BMfP_TMoqurTjay-29BkD1km"       # Confidential Instructions (_ci_) / Data ZIPs (_sf_)
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
# 2. HELPER FUNCTIONS: WORD DOCUMENT XML
# ==========================================
def add_page_number_to_run(run):
    """
    Inserts a dynamic Word PAGE field into a document text run using OpenXML.
    """
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "PAGE"
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    
    r = run._r
    r.append(fldChar1)
    r.append(instrText)
    r.append(fldChar2)
    r.append(fldChar3)

# ==========================================
# 3. AUTOMATIC ROUTING & GOOGLE DRIVE API
# ==========================================
def determine_target_folder(filename: str) -> tuple[str, str]:
    """
    Analyzes the Physics 9702 filename using Regular Expressions to determine target folder.
    Returns a tuple of (folder_key, folder_display_name).
    """
    filename_lower = filename.lower()
    
    # 1. Practical Instructions / Source Files / Zip Archives (_ci_, _sf_, or .zip)
    if filename_lower.endswith(".zip") or "_sf_" in filename_lower or "_ci_" in filename_lower:
        return "zips", f"{SYLLABUS_CODE}_zips (Confidential Instructions Files)"
        
    # 2. Practical Question Papers & Mark Schemes (Paper 3 variants: 33, 34, 35, 36)
    if re.search(r'_(qp|ms)_3[3456]\b', filename_lower):
        return "practical", f"{SYLLABUS_CODE}_practical (Paper 3)"
        
    # 3. Theory Question Papers & Mark Schemes (Papers 1, 2, 4, 5 variants: 12, 13, 22, 23, 42, 43, 52, 53)
    if re.search(r'_(qp|ms)_(1[23]|2[23]|4[23]|5[23])\b', filename_lower):
        return "theory", f"{SYLLABUS_CODE}_theory (Papers 1, 2, 4, 5)"
        
    return None, None

def build_drive_service():
    """
    Authenticates with Google Drive API using OAuth credentials stored in Streamlit Secrets.
    Safely checks for required secrets keys to prevent runtime crashes.
    """
    required_keys = ["refresh_token", "client_id", "client_secret"]
    missing_keys = [k for k in required_keys if k not in st.secrets]

    if missing_keys:
        st.error(
            f"❌ Missing Secret Key(s): {', '.join(missing_keys)}. "
            "Please ensure all API keys are placed at the top level of your secrets configuration."
        )
        return None

    try:
        creds = Credentials(
            token=None,
            refresh_token=st.secrets["refresh_token"],
            token_uri="https://oauth2.googleapis.com/token",
            client_id=st.secrets["client_id"],
            client_secret=st.secrets["client_secret"]
        )
        return build('drive', 'v3', credentials=creds)
    except Exception as e:
        st.error(f"❌ Google Drive Authentication Error: {e}")
        return None

def upload_file_to_drive(file_bytes, filename, folder_id, mime_type):
    """
    Uploads a file binary stream to a specific Google Drive folder.
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
            st.error("❌ Google Drive API Quota Limit Exceeded.")
        else:
            st.error(f"❌ Google Drive API HTTP Error ({http_err.resp.status}): {http_err}")
        return None
    except Exception as error:
        st.error(f"❌ Drive API Upload Failed: {error}")
        return None

def sync_drive_folder_to_local(folder_key: str) -> tuple[int, str]:
    """
    Queries Google Drive for a folder and downloads missing files to local server mirror.
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
# 4. SEARCH ENGINE & HELPER FUNCTIONS
# ==========================================
def render_pdf_page_preview(filepath: str, page_num: int):
    """
    Renders a single PDF page into PNG bytes for Streamlit previewing.
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
            
            # Exclude confidential instructions from general QP/MS searches if stored in same folder
            if "_ci_" in file:
                continue

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
# 5. APP STATE INITIALIZATION
# ==========================================
if 'handout_basket' not in st.session_state:
    st.session_state.handout_basket = []
if 'theory_results' not in st.session_state:
    st.session_state.theory_results = []
if 'practical_results' not in st.session_state:
    st.session_state.practical_results = []

# ==========================================
# 6. STREAMLIT UI LAYOUT & STYLING
# ==========================================
st.set_page_config(page_title="9702 Physics PYP Archives", layout="wide")

MAIN_BG_COLOR = "#FEE7F9"     # Soft Pink
SIDEBAR_BG_COLOR = "#FCBBEF"  # Accent Pink
INPUT_BAR_COLOR = "#E7FCBB"   # Soft Green

st.markdown(
    f"""
    <style>
    .stAppViewContainer {{ background-color: {MAIN_BG_COLOR}; }}
    .stHeader {{ background-color: {MAIN_BG_COLOR}; }}
    [data-testid="stSidebar"] {{ background-color: {SIDEBAR_BG_COLOR}; }}
    div[data-testid="stTextInput"] input, div[data-baseweb="input"] {{
        background-color: {INPUT_BAR_COLOR} !important; border-radius: 8px !important;
    }}
    div[data-testid="stSelectbox"] div {{
        background-color: {INPUT_BAR_COLOR} !important; border-radius: 8px !important;
    }}
    [data-testid="stFileUploaderDropzone"] {{
        background-color: {INPUT_BAR_COLOR} !important; border-radius: 8px !important;
    }}
    </style>
    """,
    unsafe_allow_html=True
)

st.title("GCE A/AS LEVEL PHYSICS")
st.subheader("⚡ 9702 Physics PYP Resource Library")

# Sidebar - Handout Basket Summary
with st.sidebar:
    st.header("Handout Basket Summary")
    st.metric(label="Saved Pages in Basket", value=len(st.session_state.handout_basket))
    if st.button("🗑️ Clear Entire Basket"):
        st.session_state.handout_basket = []
        st.rerun()

# Navigation Tabs
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🔍 Theory Search (P1,2,4,5)", 
    "⚙️ Practical Search (P3)", 
    "🛒 Handout Cart", 
    "📦 Practical Instructions", 
    "🔒 Upload PYP & Sync"
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
                    
                    if os.path.exists(item["path"]):
                        with open(item["path"], "rb") as pdf_file:
                            st.download_button(
                                label="📥 Download Full PDF",
                                data=pdf_file,
                                file_name=item["file"],
                                mime="application/pdf",
                                key=f"dl_t1_{idx}"
                            )


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
                    
                    if os.path.exists(item["path"]):
                        with open(item["path"], "rb") as pdf_file:
                            st.download_button(
                                label="📥 Download Full PDF",
                                data=pdf_file,
                                file_name=item["file"],
                                mime="application/pdf",
                                key=f"dl_t2_{idx}"
                            )


# --- TAB 3: HANDOUT CART & WORD BUILDER ---
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
                section = doc.sections[0]

                # 1. Document Page Setup (Portrait, 8.5" x 11.5")
                section.orientation = WD_ORIENT.PORTRAIT
                section.page_width = Inches(8.5)
                section.page_height = Inches(11.5)

                # 2. Document Margins (Top: 0.6", Left/Right/Bottom: 0.4")
                section.top_margin = Inches(0.6)
                section.bottom_margin = Inches(0.4)
                section.left_margin = Inches(0.4)
                section.right_margin = Inches(0.4)

                # 3. Top Center Dynamic Page Numbering in Header
                header = section.header
                header_paragraph = header.paragraphs[0]
                header_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                header_run = header_paragraph.add_run("Page ")
                header_run.font.name = "Calibri"
                header_run.font.size = Pt(10)
                add_page_number_to_run(header_run)

                # 4. Populate Document Content
                doc.add_heading(f'PTES {SYLLABUS_CODE} Physics Handout', level=1)

                for i, item in enumerate(st.session_state.handout_basket):
                    doc.add_heading(f"Source: {item['file']} (Page {item['page'] + 1})", level=2)
                    
                    # Render PDF page at high DPI for Word insertion
                    pdf_doc = fitz.open(item['path'])
                    page = pdf_doc.load_page(item['page'])
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                    img_data = io.BytesIO(pix.tobytes("png"))
                    
                    # Maximum width constrained to 7.7 inches (8.5 - 0.4 - 0.4)
                    doc.add_picture(img_data, width=Inches(7.7))
                    
                    # Page break after each image except the last
                    if i < len(st.session_state.handout_basket) - 1:
                        doc.add_page_break()
                        
                    pdf_doc.close()

                target_filename = f"{SYLLABUS_CODE}_Physics_Handout.docx"
                doc.save(target_filename)

                with open(target_filename, "rb") as f:
                    st.download_button(
                        label="📥 Click to Download Word Handout", 
                        data=f, 
                        file_name=target_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            except Exception as e:
                st.error(f"❌ Handout Export Failed: {e}")
    else:
        st.info("Your basket is empty. Add pages from Tab 1 or Tab 2.")


# --- TAB 4: CONFIDENTIAL INSTRUCTIONS (_ci_) & SOURCE FILES ---
with tab4:
    st.header("Download Practical Confidential Instructions")
    c1, c2, c3 = st.columns(3)
    with c1:
        z_year = st.selectbox("Select Year", [str(y) for y in range(2026, 2018, -1)])
    with c2:
        z_session = st.selectbox("Select Session", ["May/June (s)", "October/November (w)"])
        session_code = z_session.split("(")[1].replace(")", "")
    with c3:
        z_paper = st.selectbox("Select Paper Component", ["Paper 33", "Paper 34", "Paper 35", "Paper 36"])

    short_year = z_year[-2:]
    
    possible_filenames = [
        f"{SYLLABUS_CODE}_{session_code}{short_year}_ci_{z_paper}.pdf",
        f"{SYLLABUS_CODE}_{session_code}{short_year}_ci_{z_paper}.zip",
        f"{SYLLABUS_CODE}_{session_code}{short_year}_sf_{z_paper}.zip",
        f"{SYLLABUS_CODE}_{session_code}{short_year}_sf_{z_paper}.pdf",
    ]

    found_file_path = None
    matched_filename = None

    for fname in possible_filenames:
        check_path = os.path.join(LOCAL_FOLDERS["zips"], fname)
        if os.path.exists(check_path):
            found_file_path = check_path
            matched_filename = fname
            break

    st.markdown("---")
    if found_file_path and matched_filename:
        st.success(f"Found File: `{matched_filename}`")
        
        mime_type = "application/pdf" if matched_filename.endswith(".pdf") else "application/zip"
        
        with open(found_file_path, "rb") as f:
            st.download_button(
                label=f"📥 Download {matched_filename}",
                data=f,
                file_name=matched_filename,
                mime=mime_type
            )
            
        if matched_filename.endswith(".pdf"):
            preview = render_pdf_page_preview(found_file_path, 0)
            if preview:
                st.image(preview, caption=f"Preview of {matched_filename} (Page 1)", width=600)
    else:
        st.warning(
            f"No Instructions Source File found for `{SYLLABUS_CODE}_{session_code}{short_year}` Paper `{z_paper}`. "
            "Use Tab 5 (Admin Panel) to sync or upload."
        )


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
            st.caption("Downloads all newly added papers from your 3 Google Drive folders to the app server.")
            if st.button("🔄 Sync All Files from Google Drive", type="primary"):
                with st.spinner("Scanning Google Drive folders..."):
                    total_synced = 0
                    for f_key in ["theory", "practical", "zips"]:
                        count, msg = sync_drive_folder_to_local(f_key)
                        total_synced += count
                        st.info(msg)
                    st.success(f"🎉 Sync Complete! **{total_synced}** new file(s) downloaded.")

            st.markdown("---")
            st.subheader("📤 Single File Direct Upload")
            st.caption("Upload a file directly to local server storage and sync it to Google Drive.")
            uploaded_file = st.file_uploader(
                "Upload Past Paper (PDF) or Source File (ZIP)", 
                type=["pdf", "zip"],
                key="admin_file_uploader"
            )

            if uploaded_file is not None:
                folder_key, folder_name = determine_target_folder(uploaded_file.name)

                if folder_key is None:
                    st.warning(f"⚠️ Filename `{uploaded_file.name}` does not match Cambridge 9702 naming conventions.")
                else:
                    st.info(f"🎯 Target Destination: **{folder_name}**")

                    if st.button("🚀 Upload File"):
                        with st.spinner("Uploading file to Google Drive and local storage..."):
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
                                st.toast(f"Successfully added {uploaded_file.name}!", icon="🎉")


# ==========================================
# 7. FOOTER
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
