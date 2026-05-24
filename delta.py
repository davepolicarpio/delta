# ==========================================
# DELTA | PREMIUM CONTRACT REVIEWER PLATFORM
# ==========================================
# Dependencies: pip install streamlit pymupdf python-docx

import streamlit as st
import difflib
import io
import fitz  # PyMuPDF
from docx import Document

# ==========================================
# BLOCK 1: SESSION STATE INITIALIZATION
# ==========================================
# [DIRECTIVE: Explicitly declare all session state keys before any UI renders.]
# [DIRECTIVE: This prevents key errors during Streamlit's reactive reruns.]
if 'initialized' not in st.session_state:
    st.session_state.initialized = True
    st.session_state.uploaded_files_data = {}  # Stores raw text mapped by filename
    st.session_state.file_order = []           # Tracks the manually sorted sequence of files
    st.session_state.file_roles = {}           # Maps filename -> Dropdown Role (Template, v1, etc.)
    st.session_state.processing_complete = False # Toggles the Phase 2 Dashboard view
    st.session_state.current_baseline = 0      # Timeline index for Col 1
    st.session_state.current_counter = 1       # Timeline index for Col 2
    st.session_state.peek_template = False     # Toggles Phase 3 Sidebar

# ==========================================
# BLOCK 2: LUXURY BRAND CSS INJECTION
# ==========================================
# [DIRECTIVE: Inject custom HTML/CSS to override Streamlit's default theme.]
# [DIRECTIVE: Enforce Navy Blue (#0a192f), Charcoal (#1e293b), Crisp White, and Gold (#d4af37).]
def inject_custom_css():
    st.markdown("""
        <style>
            /* Global Backgrounds: Deep Navy Blue */
            .stApp {
                background-color: #0a192f;
                color: #e2e8f0;
                font-family: 'Helvetica Neue', sans-serif;
            }
            /* Sidebar Background: Dark Charcoal Gray */
            [data-testid="stSidebar"] {
                background-color: #1e293b;
            }
            /* Headers and Titles */
            h1, h2, h3 {
                color: #d4af37 !important; /* Metallic Gold Accent */
            }
            /* Primary Action Buttons */
            .stButton > button {
                background-color: #d4af37 !important;
                color: #0a192f !important;
                font-weight: bold !important;
                border: none !important;
                border-radius: 4px;
                transition: all 0.3s ease;
            }
            .stButton > button:hover {
                background-color: #fcd34d !important;
                transform: scale(1.02);
            }
            /* Document Text Containers: Crisp White for Readability */
            .doc-container {
                background-color: #ffffff;
                color: #0f172a;
                padding: 1.5rem;
                border-radius: 8px;
                box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
                margin-bottom: 1rem;
                position: relative;
                font-size: 14px;
                line-height: 1.6;
            }
            /* The Edit Count Badge: Gold Accent */
            .delta-badge {
                position: absolute;
                top: -10px;
                right: -10px;
                background-color: #d4af37;
                color: #0a192f;
                padding: 4px 10px;
                border-radius: 12px;
                font-size: 0.75rem;
                font-weight: bold;
                box-shadow: 0 2px 4px rgba(0,0,0,0.2);
            }
            /* Highlights inside the white container */
            .highlight-add {
                background-color: #dcfce7; /* Pastel Green */
                color: #166534;
                padding: 2px 4px;
                border-radius: 3px;
            }
            .highlight-del {
                background-color: #fee2e2; /* Pastel Red */
                color: #991b1b;
                text-decoration: line-through;
                padding: 2px 4px;
                border-radius: 3px;
            }
            /* Delta Summary Cards in Sidebar */
            .summary-card {
                background-color: #1e293b;
                border-left: 4px solid #d4af37;
                padding: 1rem;
                margin-bottom: 0.5rem;
                border-radius: 0 4px 4px 0;
            }
        </style>
    """, unsafe_allow_html=True)

# ==========================================
# BLOCK 3: DOCUMENT PARSING ENGINE
# ==========================================
# [DIRECTIVE: Pure deterministic text extraction. No ML/LLM models.]
def parse_pdf(file_bytes):
    # [DIRECTIVE: Load PDF bytes via PyMuPDF]
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    paragraphs = []
    # [DIRECTIVE: Iterate through pages and extract blocks to maintain paragraph structure]
    for page in doc:
        blocks = page.get_text("blocks")
        for b in blocks:
            text = b[4].strip()
            if text:
                paragraphs.append(text)
    return paragraphs

def parse_docx(file_bytes):
    # [DIRECTIVE: Load DOCX bytes via python-docx]
    doc = Document(io.BytesIO(file_bytes))
    paragraphs = []
    # [DIRECTIVE: Append non-empty paragraph strings to list]
    for p in doc.paragraphs:
        if p.text.strip():
            paragraphs.append(p.text.strip())
    return paragraphs

def process_uploaded_files(uploaded_files):
    # [DIRECTIVE: Detect file type and route to appropriate parser]
    for file in uploaded_files:
        if file.name not in st.session_state.uploaded_files_data:
            bytes_data = file.read()
            if file.name.endswith('.pdf'):
                text_list = parse_pdf(bytes_data)
            elif file.name.endswith('.docx'):
                text_list = parse_docx(bytes_data)
            else:
                continue
            
            # [DIRECTIVE: Store parsed lists in session state mapped by filename]
            st.session_state.uploaded_files_data[file.name] = text_list
            # [DIRECTIVE: Auto-append new files to the staging order tracker]
            if file.name not in st.session_state.file_order:
                st.session_state.file_order.append(file.name)
                # Default role assignment
                st.session_state.file_roles[file.name] = "v1: Baseline"

# ==========================================
# BLOCK 4: STATE MUTATION CALLBACKS
# ==========================================
# [DIRECTIVE: Handle array shifting for Phase 1 Move Up/Down buttons]
def move_up(idx):
    if idx > 0:
        # [DIRECTIVE: Swap current item with the one above it in the order array]
        st.session_state.file_order[idx], st.session_state.file_order[idx-1] = \
            st.session_state.file_order[idx-1], st.session_state.file_order[idx]

def move_down(idx):
    if idx < len(st.session_state.file_order) - 1:
        # [DIRECTIVE: Swap current item with the one below it in the order array]
        st.session_state.file_order[idx], st.session_state.file_order[idx+1] = \
            st.session_state.file_order[idx+1], st.session_state.file_order[idx]

# ==========================================
# BLOCK 5: ALGORITHMIC TEXT DIFFING
# ==========================================
def compute_inline_diff(text1, text2):
    # [DIRECTIVE: Compare two strings word-by-word to find micro-edits]
    matcher = difflib.SequenceMatcher(None, text1.split(), text2.split())
    out1, out2 = [], []
    delta_count = 0
    
    # [DIRECTIVE: Map opcodes to HTML highlight classes]
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        w1 = " ".join(text1.split()[i1:i2])
        w2 = " ".join(text2.split()[j1:j2])
        if tag == 'equal':
            out1.append(w1)
            out2.append(w2)
        elif tag == 'replace':
            delta_count += 1
            out1.append(f'<span class="highlight-del">{w1}</span>')
            out2.append(f'<span class="highlight-add">{w2}</span>')
        elif tag == 'delete':
            delta_count += 1
            out1.append(f'<span class="highlight-del">{w1}</span>')
        elif tag == 'insert':
            delta_count += 1
            out2.append(f'<span class="highlight-add">{w2}</span>')
            
    return delta_count, " ".join(out1), " ".join(out2)

def generate_diff_view(base_paras, counter_paras):
    # [DIRECTIVE: Align macro-paragraphs before computing word-level diffs]
    matcher = difflib.SequenceMatcher(None, base_paras, counter_paras)
    base_html_blocks = []
    counter_html_blocks = []
    summary_cards = []
    
    # [DIRECTIVE: Process structural changes based on block opcodes]
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            for i in range(i1, i2):
                # No edits: render plain text inside standard white container
                html = f'<div class="doc-container">{base_paras[i]}</div>'
                base_html_blocks.append(html)
                counter_html_blocks.append(html)
                
        elif tag == 'replace':
            # Blocks changed: compute inline word diffs
            for i, j in zip(range(i1, i2), range(j1, j2)):
                deltas, out1, out2 = compute_inline_diff(base_paras[i], counter_paras[j])
                base_html_blocks.append(f'<div class="doc-container"><div class="delta-badge">[ {deltas} Deltas ]</div>{out1}</div>')
                counter_html_blocks.append(f'<div class="doc-container"><div class="delta-badge">[ {deltas} Deltas ]</div>{out2}</div>')
                summary_cards.append(f"Block Modification: {deltas} changes detected.")
                
        elif tag == 'delete':
            # Block removed entirely in counter-offer
            for i in range(i1, i2):
                base_html_blocks.append(f'<div class="doc-container"><div class="delta-badge">[ Block Removed ]</div><span class="highlight-del">{base_paras[i]}</span></div>')
                counter_html_blocks.append(f'<div class="doc-container" style="opacity: 0.5;"><i>[Paragraph structurally deleted in this version]</i></div>')
                summary_cards.append("Structural Deletion: A full paragraph was removed.")
                
        elif tag == 'insert':
            # Block added entirely in counter-offer
            for j in range(j1, j2):
                base_html_blocks.append(f'<div class="doc-container" style="opacity: 0.5;"><i>[Paragraph missing in baseline]</i></div>')
                counter_html_blocks.append(f'<div class="doc-container"><div class="delta-badge">[ Block Inserted ]</div><span class="highlight-add">{counter_paras[j]}</span></div>')
                summary_cards.append("Structural Addition: A new paragraph was inserted.")
                
    return base_html_blocks, counter_html_blocks, summary_cards

# ==========================================
# BLOCK 6: UI PHASE 1 - STAGING VIEW
# ==========================================
def render_staging_view():
    st.markdown("<h1>▲ DELTA | Contract Revision Tree</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color: #e2e8f0;'>Upload, sequence, and classify document versions to build the revision tree.</p>", unsafe_allow_html=True)
    
    # [DIRECTIVE: Native file uploader accepting specified formats]
    uploaded_files = st.file_uploader("Ingest Contracts", type=['pdf', 'docx'], accept_multiple_files=True)
    
    if uploaded_files:
        process_uploaded_files(uploaded_files)
        st.markdown("### Document Staging Sequence")
        
        # [DIRECTIVE: Iterate through ordered files and render the interactive sorting UI]
        for idx, filename in enumerate(st.session_state.file_order):
            col1, col2, col3, col4 = st.columns([4, 3, 1, 1])
            with col1:
                st.write(f"📄 **{filename}**")
            with col2:
                roles = ["Standard Template", "v1: Baseline", "v2: Counter", "v3: Counter", "v4: Counter"]
                # [DIRECTIVE: Selectbox assigns semantic role to the specific document]
                st.session_state.file_roles[filename] = st.selectbox(
                    "Role", roles, index=roles.index(st.session_state.file_roles.get(filename, "v1: Baseline")), key=f"role_{filename}"
                )
            with col3:
                # [DIRECTIVE: Move Up button with state mutation callback]
                st.button("🔼", key=f"up_{filename}", on_click=move_up, args=(idx,))
            with col4:
                # [DIRECTIVE: Move Down button with state mutation callback]
                st.button("🔽", key=f"down_{filename}", on_click=move_down, args=(idx,))
                
        st.markdown("---")
        # [DIRECTIVE: Primary action button transitions application to Phase 2]
        if st.button("🚀 Process & Compare Contracts"):
            st.session_state.processing_complete = True
            st.rerun()

# ==========================================
# BLOCK 7: UI PHASE 2 & 3 - DASHBOARD & PEEK
# ==========================================
def render_dashboard_view():
    st.markdown("<h1>▲ DELTA | Workspace</h1>", unsafe_allow_html=True)
    
    # [DIRECTIVE: Top Timeline Controls for Version Navigation]
    ordered_files = st.session_state.file_order
    roles = [st.session_state.file_roles[f] for f in ordered_files]
    
    st.markdown("### Version Navigation Timeline")
    t_col1, t_col2 = st.columns(2)
    with t_col1:
        st.session_state.current_baseline = st.selectbox("Select Baseline", range(len(ordered_files)), format_func=lambda x: f"{roles[x]} ({ordered_files[x]})")
    with t_col2:
        st.session_state.current_counter = st.selectbox("Select Counter-Offer", range(len(ordered_files)), format_func=lambda x: f"{roles[x]} ({ordered_files[x]})", index=min(1, len(ordered_files)-1))
        
    # [DIRECTIVE: Toggle switch for Phase 3 Split-Pane Peek Window]
    st.session_state.peek_template = st.checkbox("👁️ Peek Standard Template", value=st.session_state.peek_template)
    
    # [DIRECTIVE: Phase 3 Sidebar Logic - Render Standard Template if toggled]
    if st.session_state.peek_template:
        with st.sidebar:
            st.markdown("### 👁️ Standard Template Vault")
            template_file = next((f for f in ordered_files if "Standard Template" in st.session_state.file_roles[f]), None)
            if template_file:
                st.markdown(f"**Viewing:** {template_file}")
                # Render raw template text inside white containers
                for p in st.session_state.uploaded_files_data[template_file]:
                    st.markdown(f'<div class="doc-container" style="font-size: 12px;">{p}</div>', unsafe_allow_html=True)
            else:
                st.warning("No document tagged as 'Standard Template' in staging.")

    # [DIRECTIVE: Phase 2 Main Reviewer Dashboard (3-Column Grid)]
    st.markdown("---")
    col1, col2, col3 = st.columns([4, 4, 3])
    
    # Fetch text arrays for selected versions
    base_file = ordered_files[st.session_state.current_baseline]
    counter_file = ordered_files[st.session_state.current_counter]
    base_text = st.session_state.uploaded_files_data[base_file]
    counter_text = st.session_state.uploaded_files_data[counter_file]
    
    # Run algorithms
    base_blocks, counter_blocks, summaries = generate_diff_view(base_text, counter_text)
    
    # [DIRECTIVE: Render Column 1 (Baseline)]
    with col1:
        st.markdown(f"#### Baseline: {roles[st.session_state.current_baseline]}")
        for html_block in base_blocks:
            st.markdown(html_block, unsafe_allow_html=True)
            
    # [DIRECTIVE: Render Column 2 (Counter-Offer)]
    with col2:
        st.markdown(f"#### Counter: {roles[st.session_state.current_counter]}")
        for html_block in counter_blocks:
            st.markdown(html_block, unsafe_allow_html=True)
            
    # [DIRECTIVE: Render Column 3 (Smart Delta Sidebar)]
    with col3:
        st.markdown("#### Smart Delta Output")
        st.markdown("<p style='color:#cbd5e1; font-size:14px;'>Aggregated block-level modifications.</p>", unsafe_allow_html=True)
        if not summaries:
            st.success("Documents are identical.")
        else:
            for idx, summary in enumerate(summaries):
                st.markdown(f'<div class="summary-card"><strong>Delta {idx+1}:</strong><br/>{summary}</div>', unsafe_allow_html=True)
        
        st.markdown("<br/>", unsafe_allow_html=True)
        if st.button("⬅️ Return to Staging"):
            st.session_state.processing_complete = False
            st.rerun()

# ==========================================
# BLOCK 8: MAIN EXECUTION THREAD
# ==========================================
# [DIRECTIVE: Top-level orchestration. Injects CSS, then routes to appropriate view.]
def main():
    st.set_page_config(page_title="DELTA | Contract Reviewer", layout="wide", initial_sidebar_state="collapsed")
    inject_custom_css()
    
    if not st.session_state.processing_complete:
        render_staging_view()
    else:
        render_dashboard_view()

if __name__ == "__main__":
    main()
