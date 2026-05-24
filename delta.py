# ==========================================
# DELTA | PREMIUM CONTRACT REVIEWER PLATFORM
# ==========================================
# Dependencies required: pip install streamlit pymupdf python-docx

import streamlit as st
import difflib
import io
import fitz  # PyMuPDF
from docx import Document

# ==========================================
# BLOCK 1: SESSION STATE INITIALIZATION
# ==========================================
# [WHAT]: Initialize fundamental session keys to prevent Streamlit from wiping data on reactive reruns.
# [HOW]: Checks if the 'initialized' flag exists. If not, it sets up empty data structures.
if 'initialized' not in st.session_state:
    st.session_state.initialized = True
    # Dictionary mapping filename to the extracted list of paragraphs
    st.session_state.files_data = {}  
    # List maintaining the exact sequence/order of uploaded files
    st.session_state.file_order = []  
    # Dictionary mapping filename to its assigned role string (e.g., "Standard Template")
    st.session_state.file_roles = {}  
    # Boolean toggle to switch between Phase 1 (Staging) and Phase 2 (Dashboard)
    st.session_state.dashboard_active = False 
    # Boolean toggle to show/hide the Phase 3 Standard Template Sidebar
    st.session_state.peek_active = False 

# ==========================================
# BLOCK 2: LUXURY BRAND CSS INJECTION
# ==========================================
# [WHAT]: Enforces the custom luxury color palette and typography across the UI.
# [HOW]: Injects an HTML <style> block targeting specific Streamlit classes and custom divs.
def inject_luxury_css():
    st.markdown("""
        <style>
            /* Base Theme overrides: Deep Navy Blue Background, Sans-Serif Typography */
            .stApp {
                background-color: #0A192F !important;
                color: #FFFFFF !important;
                font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif !important;
            }
            
            /* Sidebar styling for the Peek Window: Dark Charcoal Gray */
            [data-testid="stSidebar"] {
                background-color: #1E1E1E !important;
                border-right: 1px solid #333333;
            }
            
            /* Accent Headers: Premium Metallic Gold */
            h1, h2, h3, h4, h5, h6 {
                color: #D4AF37 !important;
            }
            
            /* Primary Action Buttons: Metallic Gold with Dark Navy Text */
            .stButton > button {
                background-color: #D4AF37 !important;
                color: #0A192F !important;
                font-weight: 600 !important;
                border: none !important;
                border-radius: 4px !important;
                transition: all 0.3s ease;
            }
            .stButton > button:hover {
                background-color: #F3C94F !important; /* Lighter gold on hover */
                transform: translateY(-2px);
            }
            
            /* Crisp White Document Containers for Flawless Readability */
            .doc-card {
                background-color: #FFFFFF !important;
                color: #111111 !important; /* Dark text for contrast */
                padding: 1.5rem !important;
                margin-bottom: 1.5rem !important;
                border-radius: 6px !important;
                position: relative; /* Essential for absolute positioning of the badge */
                box-shadow: 0 4px 6px rgba(0,0,0,0.1);
                line-height: 1.6;
            }
            
            /* The Edit Count Badge (Top-Right of Modified Blocks) */
            .edit-badge {
                position: absolute;
                top: -12px;
                right: 15px;
                background-color: #D4AF37;
                color: #0A192F;
                font-size: 11px;
                font-weight: bold;
                padding: 4px 10px;
                border-radius: 12px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.2);
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }
            
            /* Micro Word-Level Highlights inside the white container */
            .highlight-add {
                background-color: #dcfce7; /* Soft pastel green */
                color: #166534;
                padding: 2px 4px;
                border-radius: 3px;
            }
            .highlight-del {
                background-color: #fee2e2; /* Soft pastel red */
                color: #991b1b;
                text-decoration: line-through;
                padding: 2px 4px;
                border-radius: 3px;
            }
            
            /* Delta Summary Sidebar Cards */
            .summary-card {
                background-color: #1E1E1E;
                border-left: 4px solid #D4AF37;
                padding: 1rem;
                margin-bottom: 0.75rem;
                border-radius: 0 4px 4px 0;
                font-size: 14px;
            }
            
            /* Hide Streamlit default styling elements where necessary */
            .stFileUploader > div > div { background-color: #1E1E1E !important; }
        </style>
    """, unsafe_allow_html=True)

# ==========================================
# BLOCK 3: DETERMINISTIC DOCUMENT PARSING
# ==========================================
# [WHAT]: Extracts flat paragraph lists from PDF and DOCX files.
# [HOW]: PyMuPDF iterates page blocks; python-docx iterates paragraphs.

def extract_pdf(file_bytes):
    # Initializes the PyMuPDF document from binary stream
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    paragraphs = []
    # Loops through pages and extracts structurally grouped blocks of text
    for page in doc:
        for b in page.get_text("blocks"):
            text = b[4].strip()
            if text:
                # Replaces internal newlines to keep paragraph logic intact
                paragraphs.append(" ".join(text.split()))
    return paragraphs

def extract_docx(file_bytes):
    # Initializes the DOCX Document object from binary stream
    doc = Document(io.BytesIO(file_bytes))
    paragraphs = []
    # Loops through all document paragraphs sequentially
    for p in doc.paragraphs:
        if p.text.strip():
            paragraphs.append(" ".join(p.text.split()))
    return paragraphs

def handle_uploads(uploaded_files):
    # Parses un-ingested files and maps them into session state
    for file in uploaded_files:
        if file.name not in st.session_state.files_data:
            b_data = file.read()
            if file.name.endswith('.pdf'):
                parsed_list = extract_pdf(b_data)
            elif file.name.endswith('.docx'):
                parsed_list = extract_docx(b_data)
            else:
                continue
            
            # Save parsed text to the dictionary
            st.session_state.files_data[file.name] = parsed_list
            # Append to file sequence tracker
            if file.name not in st.session_state.file_order:
                st.session_state.file_order.append(file.name)
                # Assign a default role
                st.session_state.file_roles[file.name] = "v1: Baseline"

# ==========================================
# BLOCK 4: STATE MUTATION UTILITIES (Move Up/Down)
# ==========================================
# [WHAT]: Callbacks for shifting the order of files in Phase 1.
# [HOW]: Swaps indices in the st.session_state.file_order list.

def shift_up(idx):
    if idx > 0:
        st.session_state.file_order[idx], st.session_state.file_order[idx-1] = \
            st.session_state.file_order[idx-1], st.session_state.file_order[idx]

def shift_down(idx):
    if idx < len(st.session_state.file_order) - 1:
        st.session_state.file_order[idx], st.session_state.file_order[idx+1] = \
            st.session_state.file_order[idx+1], st.session_state.file_order[idx]

# ==========================================
# BLOCK 5: ALGORITHMIC TEXT DIFFING ENGINE
# ==========================================
# [WHAT]: Compares Base and Counter paragraphs, tracks modifications, and generates HTML.
# [HOW]: SequenceMatcher compares macro-structure. If a replace happens, a micro word-level diff is computed.

def compute_inline_diff(txt1, txt2):
    # Word-level deterministic diffing
    matcher = difflib.SequenceMatcher(None, txt1.split(), txt2.split())
    out1, out2 = [], []
    deltas = 0
    
    # Process the opcodes to wrap modified words in pastel spans
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        w1 = " ".join(txt1.split()[i1:i2])
        w2 = " ".join(txt2.split()[j1:j2])
        
        if tag == 'equal':
            out1.append(w1)
            out2.append(w2)
        elif tag == 'replace':
            deltas += 1
            out1.append(f'<span class="highlight-del">{w1}</span>')
            out2.append(f'<span class="highlight-add">{w2}</span>')
        elif tag == 'delete':
            deltas += 1
            out1.append(f'<span class="highlight-del">{w1}</span>')
        elif tag == 'insert':
            deltas += 1
            out2.append(f'<span class="highlight-add">{w2}</span>')
            
    return deltas, " ".join(out1), " ".join(out2)

def generate_dashboard_payload(base_list, cntr_list):
    # Compares full paragraph arrays
    matcher = difflib.SequenceMatcher(None, base_list, cntr_list)
    html_base = []
    html_cntr = []
    summary_deck = []
    
    # Build structural markup based on comparison opcodes
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            for i in range(i1, i2):
                # Standard white card without any gold badges
                card = f'<div class="doc-card">{base_list[i]}</div>'
                html_base.append(card)
                html_cntr.append(card)
                
        elif tag == 'replace':
            for i, j in zip(range(i1, i2), range(j1, j2)):
                d_count, out1, out2 = compute_inline_diff(base_list[i], cntr_list[j])
                # Wrap in white card and inject the gold Edit Count Badge
                html_base.append(f'<div class="doc-card"><div class="edit-badge">[ {d_count} Deltas ]</div>{out1}</div>')
                html_cntr.append(f'<div class="doc-card"><div class="edit-badge">[ {d_count} Deltas ]</div>{out2}</div>')
                summary_deck.append(f"Modification Tracked: {d_count} structural inline edits executed.")
                
        elif tag == 'delete':
            for i in range(i1, i2):
                # Only exists in baseline; mark as removed
                html_base.append(f'<div class="doc-card"><div class="edit-badge">[ Block Removed ]</div><span class="highlight-del">{base_list[i]}</span></div>')
                html_cntr.append(f'<div class="doc-card" style="background-color: transparent !important; color:#777 !important; border: 1px dashed #555;"><i>[Paragraph purged in counter-offer]</i></div>')
                summary_deck.append("Omission: Standard block was structurally purged.")
                
        elif tag == 'insert':
            for j in range(j1, j2):
                # Only exists in counter; mark as inserted
                html_base.append(f'<div class="doc-card" style="background-color: transparent !important; color:#777 !important; border: 1px dashed #555;"><i>[Paragraph missing from baseline]</i></div>')
                html_cntr.append(f'<div class="doc-card"><div class="edit-badge">[ Block Inserted ]</div><span class="highlight-add">{cntr_list[j]}</span></div>')
                summary_deck.append("Insertion: New clause parameter appended by counter-party.")
                
    return html_base, html_cntr, summary_deck

# ==========================================
# BLOCK 6: PHASE 1 - INTERIM STAGING VIEW
# ==========================================
def render_phase_one():
    # Renders the executive brand header
    st.markdown("<h1>▲ DELTA | Contract Revision Tree</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color: #8892B0;'>Upload and stage your document version matrix.</p>", unsafe_allow_html=True)
    
    # Renders the multi-file uploader component
    upload_batch = st.file_uploader("Ingest Contract Files", type=['pdf', 'docx'], accept_multiple_files=True)
    
    if upload_batch:
        handle_uploads(upload_batch)
        st.markdown("### Document Staging Sequence")
        
        # Iterates through the current sequence array to render the interactive list
        for idx, fname in enumerate(st.session_state.file_order):
            col_name, col_role, col_up, col_dn = st.columns([4, 4, 1, 1])
            with col_name:
                st.write(f"📄 **{fname}**")
            with col_role:
                roles_opts = ["Standard Template", "v1: Baseline", "v2: Counter", "v3: Counter", "v4: Counter"]
                st.session_state.file_roles[fname] = st.selectbox(
                    "Classification", roles_opts, 
                    index=roles_opts.index(st.session_state.file_roles.get(fname, "v1: Baseline")), 
                    key=f"role_{fname}", label_visibility="collapsed"
                )
            with col_up:
                # Hooks to the shift_up state mutation callback
                st.button("🔼", key=f"up_{fname}", on_click=shift_up, args=(idx,))
            with col_dn:
                # Hooks to the shift_down state mutation callback
                st.button("🔽", key=f"dn_{fname}", on_click=shift_down, args=(idx,))
        
        st.markdown("---")
        # Action button that mutates state to transition to Phase 2
        if st.button("🚀 Process & Compare Contracts"):
            # Ensure at least two files exist to compare
            if len(st.session_state.file_order) >= 2:
                st.session_state.dashboard_active = True
                st.rerun()
            else:
                st.warning("Insufficient files. Ingest at least 2 files to initialize comparison.")

# ==========================================
# BLOCK 7: PHASE 2 & 3 - DASHBOARD & PEEK WINDOW
# ==========================================
def render_phase_two():
    st.markdown("<h1>▲ DELTA | Review Dashboard</h1>", unsafe_allow_html=True)
    
    seq = st.session_state.file_order
    # Map friendly names displaying role and filename for the slider UI
    slider_options = [f"{st.session_state.file_roles[f]} ({f})" for f in seq]
    
    # [Top Timeline Controls] Uses select_slider for ranged mapping (Base and Counter selection)
    st.markdown("### Version Navigation Track")
    selected_range = st.select_slider(
        "Select Baseline and Counter Versions",
        options=slider_options,
        value=(slider_options[0], slider_options[min(1, len(slider_options)-1)])
    )
    
    # Isolate the precise filenames based on slider selection
    idx_base = slider_options.index(selected_range[0])
    idx_cntr = slider_options.index(selected_range[1])
    file_base = seq[idx_base]
    file_cntr = seq[idx_cntr]
    
    st.markdown("---")
    
    # Phase 3 Toggle inside main workspace
    st.session_state.peek_active = st.checkbox("👁️ Peek Standard Template", value=st.session_state.peek_active)
    
    # [PHASE 3] Extends the Sidebar to render the raw Standard Template
    if st.session_state.peek_active:
        with st.sidebar:
            st.markdown("### 👁️ Standard Template Vault")
            st.markdown("<hr/>", unsafe_allow_html=True)
            # Find the file explicitly mapped to 'Standard Template'
            template_fname = next((f for f in seq if "Standard Template" in st.session_state.file_roles[f]), None)
            
            if template_fname:
                st.markdown(f"**Target:** {template_fname}")
                for para in st.session_state.files_data[template_fname]:
                    st.markdown(f'<div class="doc-card" style="font-size: 13px; padding: 1rem !important;">{para}</div>', unsafe_allow_html=True)
            else:
                st.warning("No Matrix configured as 'Standard Template' detected in Phase 1 mapping.")
                
    # Layout definition: 3-Column Grid
    c1, c2, c3 = st.columns([4, 4, 3])
    
    # Fetch paragraph lists and execute the deterministic difference algorithms
    base_paras = st.session_state.files_data[file_base]
    cntr_paras = st.session_state.files_data[file_cntr]
    base_ui, cntr_ui, summaries = generate_dashboard_payload(base_paras, cntr_paras)
    
    # [Column 1] Render Baseline Framework
    with c1:
        st.markdown(f"### 📄 BASELINE\n**{selected_range[0]}**")
        for markup in base_ui:
            st.markdown(markup, unsafe_allow_html=True)
            
    # [Column 2] Render Counter-Party Framework
    with c2:
        st.markdown(f"### 📝 COUNTER-OFFER\n**{selected_range[1]}**")
        for markup in cntr_ui:
            st.markdown(markup, unsafe_allow_html=True)
            
    # [Column 3] Render Smart Delta Output
    with c3:
        st.markdown("### 🔍 SMART DELTA")
        if not summaries:
            st.success("100% Convergence. No structural variables modified.")
        else:
            for idx, sum_text in enumerate(summaries):
                st.markdown(f'<div class="summary-card"><strong>Delta 0{idx+1}:</strong><br/>{sum_text}</div>', unsafe_allow_html=True)
                
        st.markdown("<br/><br/>", unsafe_allow_html=True)
        # Teardown button to return to Phase 1 staging
        if st.button("⬅️ Back to Staging"):
            st.session_state.dashboard_active = False
            st.rerun()

# ==========================================
# BLOCK 8: RUNTIME ORCHESTRATOR
# ==========================================
# [WHAT]: Sets page config, injects CSS, and acts as the router between logical views.
# [HOW]: Evaluates the dashboard_active state parameter to render the appropriate block.

def main():
    st.set_page_config(page_title="DELTA | Contract Reviewer", layout="wide", initial_sidebar_state="collapsed")
    inject_luxury_css()
    
    if not st.session_state.dashboard_active:
        render_phase_one()
    else:
        render_phase_two()

if __name__ == "__main__":
    main()
