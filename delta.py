# ==========================================
# DELTA | PREMIUM CONSULTANCY REVIEWER SUITE
# ==========================================

import streamlit as st
import difflib
import io
import re
import fitz  # PyMuPDF
from docx import Document

# ==========================================
# BLOCK 1: SESSION STATE INTENT STATE
# ==========================================
if 'initialized' not in st.session_state:
    st.session_state.initialized = True
    st.session_state.uploaded_files_data = {}  
    st.session_state.file_order = []           
    st.session_state.file_roles = {}           
    st.session_state.processing_complete = False 
    st.session_state.current_baseline = 0      
    st.session_state.current_counter = 1       
    st.session_state.vault_expanded = False     
    st.session_state.consultancy_notes = {}    # Tracks action tags and advisor annotations

# ==========================================
# BLOCK 2: SYSTEM DESIGN SYSTEM (CSS)
# ==========================================
def inject_luxury_system_css():
    st.markdown("""
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Cinzel:wght@400;500&family=Inter:wght@300;400;500;600&display=swap');
            
            /* Enforce Strict Grayscale Foundation */
            .stApp {
                background-color: #0c0c0c !important;
                color: #e5e5e5 !important;
                font-family: 'Inter', sans-serif !important;
            }
            
            /* Structural Vault (Sidebar Overrides) */
            [data-testid="stSidebar"] {
                background-color: #111111 !important;
                border-right: 1px solid #1a1a1a;
            }
            
            /* Luxury Brand Elements */
            .brand-title {
                font-family: 'Cinzel', serif !important;
                color: #d4af37 !important;
                letter-spacing: 3px;
                font-weight: 400;
                font-size: 26px;
                margin-bottom: 0.2rem;
            }
            .brand-subtitle {
                font-family: 'Inter', sans-serif;
                color: #525252;
                font-size: 11px;
                letter-spacing: 2px;
                text-transform: uppercase;
                margin-bottom: 2rem;
            }
            
            /* Minimal Sharp Action Controls */
            .stButton > button {
                background-color: transparent !important;
                color: #d4af37 !important;
                border: 1px solid #d4af37 !important;
                border-radius: 0px !important;
                font-family: 'Inter', sans-serif;
                font-size: 11px !important;
                letter-spacing: 1.5px;
                text-transform: uppercase;
                padding: 0.6rem 2rem !important;
                transition: all 0.3s ease;
                width: auto;
            }
            .stButton > button:hover {
                background-color: #d4af37 !important;
                color: #0c0c0c !important;
                box-shadow: 0 0 15px rgba(212, 175, 55, 0.15);
            }
            
            /* Continuous Unified Parchment Canvas */
            .document-canvas {
                background-color: #121212;
                border: 1px solid #1c1c1c;
                padding: 2.5rem;
                color: #e5e5e5;
                font-size: 14px;
                line-height: 1.8;
            }
            
            /* Flush Paragraph Stream (One Whole Page Scroll) */
            .stream-paragraph {
                padding: 1.2rem 0;
                border-bottom: 1px solid #1a1a1a;
                position: relative;
            }
            .stream-paragraph:last-child {
                border-bottom: none;
            }
            
            /* High-End Clean Diff Overlays */
            .add-token {
                background-color: #0f2d19 !important;
                color: #34d399 !important;
                padding: 2px 4px;
            }
            .del-token {
                background-color: #3b1414 !important;
                color: #f87171 !important;
                text-decoration: line-through;
                padding: 2px 4px;
            }
            
            /* Parallel Structural Indicators */
            .trace-flag {
                font-family: 'Cinzel', serif;
                color: #d4af37;
                font-size: 10px;
                letter-spacing: 1px;
                text-transform: uppercase;
                margin-bottom: 0.4rem;
                display: block;
            }
            
            /* Premium Parallel Delta Advisory Blocks */
            .advisory-panel {
                background-color: #121212;
                border: 1px solid #1c1c1c;
                border-left: 2px solid #d4af37;
                padding: 1.2rem;
                margin-bottom: 1.5rem;
                min-height: 180px; /* Forces height distribution grid alignment */
            }
            .advisory-header {
                font-family: 'Cinzel', serif;
                color: #ffffff;
                font-size: 11px;
                letter-spacing: 1px;
                text-transform: uppercase;
                margin-bottom: 0.5rem;
            }
            
            /* Premium Vault Switch Framework */
            .vault-switch-box {
                border: 1px solid #262626;
                padding: 10px 20px;
                background-color: #121212;
                text-align: center;
                cursor: pointer;
            }
            
            /* Clean form text overrides */
            div[data-baseweb="select"] {
                background-color: #121212 !important;
                border-radius: 0px !important;
            }
        </style>
    """, unsafe_allow_html=True)

# ==========================================
# BLOCK 3: CORE DETERMINISTIC EXTRACTION
# ==========================================
def parse_pdf(file_bytes):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    paragraphs = []
    for page in doc:
        blocks = page.get_text("blocks")
        for b in blocks:
            text = b[4].strip()
            if text:
                paragraphs.append(" ".join(text.split()))
    return paragraphs

def parse_docx(file_bytes):
    doc = Document(io.BytesIO(file_bytes))
    paragraphs = []
    for p in doc.paragraphs:
        if p.text.strip():
            paragraphs.append(" ".join(p.text.split()))
    return paragraphs

def load_staged_matrices(uploaded_files):
    for file in uploaded_files:
        if file.name not in st.session_state.uploaded_files_data:
            bytes_data = file.read()
            if file.name.endswith('.pdf'):
                parsed_text = parse_pdf(bytes_data)
            elif file.name.endswith('.docx'):
                parsed_text = parse_docx(bytes_data)
            else:
                continue
            st.session_state.uploaded_files_data[file.name] = parsed_text
            if file.name not in st.session_state.file_order:
                st.session_state.file_order.append(file.name)
                st.session_state.file_roles[file.name] = "v1: Baseline"

# ==========================================
# BLOCK 4: CONTEXTUAL TITLE ISOLATION ENGINE
# ==========================================
def extract_clause_signature(text):
    match = re.match(r'^([A-Za-z0-9\.\s]+(?:\b[A-Z]{2,}\b|\bRent\b|\bTerm\b|\bDeposit\b|\bUse\b))', text)
    if match:
        return match.group(1).strip()
    words = text.split()
    return " ".join(words[:2]) if len(words) >= 2 else "Provision Block"

def generate_advisory_trace_text(text1, text2, signature):
    matcher = difflib.SequenceMatcher(None, text1.split(), text2.split())
    traces = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'replace':
            o = " ".join(text1.split()[i1:i2])
            n = " ".join(text2.split()[j1:j2])
            if len(o) < 30 and len(n) < 30:
                traces.append(f"Altered '{o}' to '{n}'")
        elif tag == 'delete':
            d = " ".join(text1.split()[i1:i2])
            if len(d) < 30:
                traces.append(f"Omitted sequence '{d}'")
        elif tag == 'insert':
            ins = " ".join(text2.split()[j1:j2])
            if len(ins) < 30:
                traces.append(f"Appended sequence '{ins}'")
    if traces:
        return f"{signature} // " + "; ".join(traces)
    return f"{signature} // Structural parameter recalibration detected."

def compute_token_diff_html(text1, text2):
    matcher = difflib.SequenceMatcher(None, text1.split(), text2.split())
    out1, out2 = [], []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        w1 = " ".join(text1.split()[i1:i2])
        w2 = " ".join(text2.split()[j1:j2])
        if tag == 'equal':
            out1.append(w1)
            out2.append(w2)
        elif tag == 'replace':
            out1.append(f'<span class="del-token">{w1}</span>')
            out2.append(f'<span class="add-token">{w2}</span>')
        elif tag == 'delete':
            out1.append(f'<span class="del-token">{w1}</span>')
        elif tag == 'insert':
            out2.append(f'<span class="add-token">{w2}</span>')
    return " ".join(out1), " ".join(out2)

# ==========================================
# BLOCK 5: PHASE 1 - MINIMALIST STAGING CANVAS
# ==========================================
def render_premium_landing_view():
    st.markdown('<div class="brand-title">DELTA</div>', unsafe_allow_html=True)
    st.markdown('<div class="brand-subtitle">Automated Legal Infrastructure & Revision Tree Mapping</div>', unsafe_allow_html=True)
    
    # Drag & drop baseline input
    uploaded_files = st.file_uploader("Ingest Core Binaries", type=['pdf', 'docx'], accept_multiple_files=True, label_visibility="collapsed")
    
    if uploaded_files:
        load_staged_matrices(uploaded_files)
        st.markdown("<br/>", unsafe_allow_html=True)
        
        # Clean array management without cluttering up/down controls
        for idx, filename in enumerate(st.session_state.file_order):
            col1, col2 = st.columns([7, 3])
            with col1:
                st.markdown(f'<p style="font-size:13px; color:#ffffff; letter-spacing:0.5px; padding-top:8px;">◼ {filename}</p>', unsafe_allow_html=True)
            with col2:
                roles = ["Standard Template", "v1: Baseline", "v2: Counter", "v3: Counter", "v4: Counter"]
                st.session_state.file_roles[filename] = st.selectbox(
                    f"Classification_{filename}", roles, 
                    index=roles.index(st.session_state.file_roles.get(filename, "v1: Baseline")), 
                    label_visibility="collapsed", key=f"stage_role_{filename}"
                )
                
        st.markdown("<br/><br/>", unsafe_allow_html=True)
        if st.button("Compile Revision Tree"):
            st.session_state.processing_complete = True
            st.rerun()

# ==========================================
# BLOCK 6: PHASE 2 & 3 - DELTA CONTRACTS & CONSULTING MATRIX
# ==========================================
def render_delta_contracts_view():
    st.markdown('<div class="brand-title">DELTA CONTRACTS</div>', unsafe_allow_html=True)
    st.markdown("<br/>", unsafe_allow_html=True)
    
    ordered_files = st.session_state.file_order
    roles = [st.session_state.file_roles[f] for f in ordered_files]
    
    # Premium Menu Control Row
    t_col1, t_col2, t_col3 = st.columns([4, 4, 3])
    with t_col1:
        st.session_state.current_baseline = st.selectbox("Baseline File Model", range(len(ordered_files)), format_func=lambda x: f"BASE // {roles[x]}")
    with t_col2:
        st.session_state.current_counter = st.selectbox("Counter File Model", range(len(ordered_files)), format_func=lambda x: f"CNTR // {roles[x]}", index=min(1, len(ordered_files)-1))
    with t_col3:
        # High-end Luxury Typography Toggle representation replacing checkboxes
        st.markdown("<p style='font-size:11px; color:#525252; margin-bottom:5px; text-transform:uppercase; letter-spacing:1px;'>Vault Reference Tracker</p>", unsafe_allow_html=True)
        toggle_label = "ACTIVE // CLOSED" if not st.session_state.vault_expanded else "ACTIVE // OPENED"
        if st.button(toggle_label, key="vault_lux_toggle"):
            st.session_state.vault_expanded = not st.session_state.vault_expanded
            st.rerun()

    # Phase 3: Vault Sidebar Integration
    if st.session_state.vault_expanded:
        with st.sidebar:
            st.markdown('<div class="brand-title" style="font-size:16px;">Vault Reference</div>', unsafe_allow_html=True)
            template_file = next((f for f in ordered_files if "Standard Template" in st.session_state.file_roles[f]), None)
            if template_file:
                st.markdown(f'<p style="color:#525252; font-size:11px; margin-bottom:1.5rem;">MODEL: {template_file}</p>', unsafe_allow_html=True)
                for p in st.session_state.uploaded_files_data[template_file]:
                    st.markdown(f'<div style="font-size:12px; color:#a3a3a3; margin-bottom:1.2rem; line-height:1.6;">{p}</div>', unsafe_allow_html=True)
            else:
                st.markdown('<p style="color:#525252; font-size:12px;">Standard Matrix tracking signature missing from execution array.</p>', unsafe_allow_html=True)

    st.markdown("<br/>", unsafe_allow_html=True)
    
    # Pull data matrices
    base_file = ordered_files[st.session_state.current_baseline]
    counter_file = ordered_files[st.session_state.current_counter]
    base_paras = st.session_state.uploaded_files_data[base_file]
    counter_paras = st.session_state.uploaded_files_data[counter_file]
    
    # Process deterministic sequence matcher matrix
    matcher = difflib.SequenceMatcher(None, base_paras, counter_paras)
    
    # BUILD SYNCHRONIZED SPATIAL ROWS TO KEEP SUMMARY KAPANTAAY WITH CLAUSES
    # Every aligned structural block is printed on a matching row array across columns [4, 4, 3]
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            for i in range(i1, i2):
                col1, col2, col3 = st.columns([4, 4, 3])
                block_html = f'<div class="stream-paragraph">{base_paras[i]}</div>'
                with col1:
                    st.markdown(block_html, unsafe_allow_html=True)
                with col2:
                    st.markdown(block_html, unsafe_allow_html=True)
                with col3:
                    # Keep empty space to maintain alignment context
                    st.write("")
                    
        elif tag == 'replace':
            for i, j in zip(range(i1, i2), range(j1, j2)):
                col1, col2, col3 = st.columns([4, 4, 3])
                signature = extract_clause_signature(base_paras[i])
                trace_detail = generate_advisory_trace_text(base_paras[i], counter_paras[j], signature)
                h1, h2 = compute_token_diff_html(base_paras[i], counter_paras[j])
                
                with col1:
                    st.markdown(f'<div class="stream-paragraph"><span class="trace-flag">▲ Modified</span>{h1}</div>', unsafe_allow_html=True)
                with col2:
                    st.markdown(f'<div class="stream-paragraph"><span class="trace-flag">▲ Modified</span>{h2}</div>', unsafe_allow_html=True)
                with col3:
                    # Parallel advisory block containing what's next metrics
                    unique_id = f"mod_{i}_{j}"
                    st.markdown(f"""
                        <div class="advisory-panel">
                            <div class="advisory-header">Trace // {signature}</div>
                            <p style='color:#a3a3a3; font-size:12px;'>{trace_detail}</p>
                        </div>
                    """, unsafe_allow_html=True)
                    
                    # Consultancy Next Steps Option Blocks
                    opts = ["Select Protocol Action", "Issue Counter-Omission Challenge", "Accept Alteration via Carve-Out", "Escalate to Senior Partner Review"]
                    st.selectbox("Strategic Action Path", opts, key=f"act_{unique_id}", label_visibility="collapsed")
                    st.text_input("Internal Risk Annotation Notes", key=f"note_{unique_id}", placeholder="Log advisory stance notes...", label_visibility="collapsed")
                    st.markdown("<br/>", unsafe_allow_html=True)
                    
        elif tag == 'delete':
            for i in range(i1, i2):
                col1, col2, col3 = st.columns([4, 4, 3])
                signature = extract_clause_signature(base_paras[i])
                
                with col1:
                    st.markdown(f'<div class="stream-paragraph"><span class="trace-flag" style="color:#f87171;">◼ Omitted</span><span class="del-token">{base_paras[i]}</span></div>', unsafe_allow_html=True)
                with col2:
                    st.markdown('<div class="stream-paragraph" style="color:#404040; font-style:italic;">Provision completely absent from this iteration.</div>', unsafe_allow_html=True)
                with col3:
                    unique_id = f"del_{i}"
                    st.markdown(f"""
                        <div class="advisory-panel">
                            <div class="advisory-header">Trace // {signature}</div>
                            <p style='color:#f87171; font-size:12px;'>Structural Deletion: This entire provision was completely purged by counter-party.</p>
                        </div>
                    """, unsafe_allow_html=True)
                    opts = ["Select Protocol Action", "Demand Immediate Clause Restoration", "Approve Deletion Conditionally", "Substitute Alternative Indemnity Provision"]
                    st.selectbox("Strategic Action Path", opts, key=f"act_{unique_id}", label_visibility="collapsed")
                    st.text_input("Internal Risk Annotation Notes", key=f"note_{unique_id}", placeholder="Log advisory stance notes...", label_visibility="collapsed")
                    st.markdown("<br/>", unsafe_allow_html=True)
                    
        elif tag == 'insert':
            for j in range(j1, j2):
                col1, col2, col3 = st.columns([4, 4, 3])
                signature = extract_clause_signature(counter_paras[j])
                
                with col1:
                    st.markdown('<div class="stream-paragraph" style="color:#404040; font-style:italic;">Provision absent in baseline framework template.</div>', unsafe_allow_html=True)
                with col2:
                    st.markdown(f'<div class="stream-paragraph"><span class="trace-flag" style="color:#34d399;">◆ Insertion</span><span class="add-token">{counter_paras[j]}</span></div>', unsafe_allow_html=True)
                with col3:
                    unique_id = f"ins_{j}"
                    st.markdown(f"""
                        <div class="advisory-panel">
                            <div class="advisory-header">Trace // {signature}</div>
                            <p style='color:#34d399; font-size:12px;'>Structural Insertion: Counter-party injected a complete new provision set.</p>
                        </div>
                    """, unsafe_allow_html=True)
                    opts = ["Select Protocol Action", "Issue Total Strike-Out Directive", "Accept New Provision Terms", "Redraft Clause Borders"]
                    st.selectbox("Strategic Action Path", opts, key=f"act_{unique_id}", label_visibility="collapsed")
                    st.text_input("Internal Risk Annotation Notes", key=f"note_{unique_id}", placeholder="Log advisory stance notes...", label_visibility="collapsed")
                    st.markdown("<br/>", unsafe_allow_html=True)

    # Global structural exit button footer
    st.markdown("<br/><br/><hr style='border-color:#1a1a1a;'/>", unsafe_allow_html=True)
    if st.button("Teardown Workspace Session"):
        st.session_state.processing_complete = False
        st.rerun()

# ==========================================
# BLOCK 7: RUNTIME SYSTEM ENTRY
# ==========================================
def main():
    st.set_page_config(page_title="DELTA CONTRACTS", layout="wide", initial_sidebar_state="collapsed")
    inject_luxury_system_css()
    
    if not st.session_state.processing_complete:
        render_premium_landing_view()
    else:
        render_delta_contracts_view()

if __name__ == "__main__":
    main()
