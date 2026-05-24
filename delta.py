# ==========================================
# DELTA | PREMIUM CONSULTANCY REVIEWER SUITE
# ==========================================

import streamlit as st
import difflib
import io
import re
from datetime import datetime
import fitz  # PyMuPDF
from docx import Document

# ==========================================
# BLOCK 1: SESSION STATE MATRIX INITIALIZATION
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
    st.session_state.comparison_mode = "Baseline vs Counter" # Alternative: "Standard vs Baseline", "Standard vs Counter"

# ==========================================
# BLOCK 2: PREMIUM HIGH-END DARK CSS
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
                width: 100%;
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
                padding: 2rem;
                color: #e5e5e5;
                font-size: 13px;
                line-height: 1.8;
            }
            
            /* Flush Paragraph Stream (One Whole Page Scroll) */
            .stream-paragraph {
                padding: 1rem 0;
                border-bottom: 1px solid #1a1a1a;
                position: relative;
                min-height: 190px; /* Keeps horizontal row tracking locked sync */
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
                padding: 1rem;
                margin-bottom: 0.5rem;
            }
            .advisory-header {
                font-family: 'Cinzel', serif;
                color: #ffffff;
                font-size: 11px;
                letter-spacing: 1px;
                text-transform: uppercase;
                margin-bottom: 0.5rem;
            }
            
            /* Clean input box overrides */
            div[data-baseweb="select"] {
                background-color: #121212 !important;
                border-radius: 0px !important;
            }
            input {
                border-radius: 0px !important;
                background-color: #161616 !important;
                color: #ffffff !important;
                border: 1px solid #262626 !important;
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
# BLOCK 4: CONTEXTUAL IDENTIFIER LOGIC
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
            if len(o) < 40 and len(n) < 40:
                traces.append(f"Changed '{o}' to '{n}'")
        elif tag == 'delete':
            d = " ".join(text1.split()[i1:i2])
            if len(d) < 40:
                traces.append(f"Omitted '{d}'")
        elif tag == 'insert':
            ins = " ".join(text2.split()[j1:j2])
            if len(ins) < 40:
                traces.append(f"Injected '{ins}'")
    if traces:
        return f"{signature} | " + "; ".join(traces)
    return f"{signature} | Variant modification detected."

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
# BLOCK 5: REPORT GENERATION COMPILER
# ==========================================
def compile_markdown_report(base_name, counter_name):
    stream = []
    stream.append("# DELTA CONTRACT ADVISORY REPORT")
    stream.append(f"**Execution Date:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    stream.append(f"**Primary Baseline Target:** {base_name}")
    stream.append(f"**Counter-Party Iteration:** {counter_name}")
    stream.append("---")
    stream.append("## STRATEGIC CLAUSE INTERVENTIONS\n")
    
    has_content = False
    for key, val in st.session_state.items():
        if key.startswith("act_") and val != "Select Protocol Action":
            has_content = True
            id_suffix = key.replace("act_", "")
            note_val = st.session_state.get(f"note_{id_suffix}", "No risk notes logged.")
            
            stream.append(f"### Target ID: {id_suffix.upper()}")
            stream.append(f"* **Strategic Action Directive:** {val}")
            stream.append(f"* **Consultant Risk Annotation:** {note_val}\n")
            
    if not has_content:
        stream.append("No manual risk actions or notes were recorded for this legal brief configuration.")
        
    return "\n".join(stream)

# ==========================================
# BLOCK 6: PHASE 1 - STAGING ENVIRONMENT
# ==========================================
def render_premium_landing_view():
    st.markdown('<div class="brand-title">DELTA</div>', unsafe_allow_html=True)
    st.markdown('<div class="brand-subtitle">Automated Legal Infrastructure & Revision Tree Mapping</div>', unsafe_allow_html=True)
    
    uploaded_files = st.file_uploader("Ingest Core Binaries", type=['pdf', 'docx'], accept_multiple_files=True, label_visibility="collapsed")
    
    if uploaded_files:
        load_staged_matrices(uploaded_files)
        st.markdown("<br/>", unsafe_allow_html=True)
        
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
# BLOCK 7: PHASE 2 & 3 - SYSTEM REVIEW WORKSPACE
# ==========================================
def render_delta_contracts_view():
    st.markdown('<div class="brand-title">DELTA CONTRACTS</div>', unsafe_allow_html=True)
    st.markdown("<br/>", unsafe_allow_html=True)
    
    ordered_files = st.session_state.file_order
    roles = [st.session_state.file_roles[f] for f in ordered_files]
    
    base_file = ordered_files[st.session_state.current_baseline]
    counter_file = ordered_files[st.session_state.current_counter]
    template_file = next((f for f in ordered_files if "Standard Template" in st.session_state.file_roles[f]), None)

    # Toolbar Config Row
    t_col1, t_col2, t_col3 = st.columns([4, 4, 3])
    with t_col1:
        st.session_state.current_baseline = st.selectbox("Baseline Selection", range(len(ordered_files)), format_func=lambda x: f"BASE // {roles[x]}")
    with t_col2:
        st.session_state.current_counter = st.selectbox("Counter Selection", range(len(ordered_files)), format_func=lambda x: f"CNTR // {roles[x]}", index=min(1, len(ordered_files)-1))
    with t_col3:
        st.markdown("<p style='font-size:11px; color:#525252; margin-bottom:5px; text-transform:uppercase; letter-spacing:1px;'>Vault Core Track</p>", unsafe_allow_html=True)
        toggle_label = "VAULT VIEW // CLOSED" if not st.session_state.vault_expanded else "VAULT VIEW // ACTIVE"
        if st.button(toggle_label, key="vault_lux_toggle"):
            st.session_state.vault_expanded = not st.session_state.vault_expanded
            if not st.session_state.vault_expanded:
                st.session_state.comparison_mode = "Baseline vs Counter"
            st.rerun()

    # Dynamic Strategy Matrix Selection Row (Only shows when Vault is active)
    if st.session_state.vault_expanded:
        st.markdown("<br/>", unsafe_allow_html=True)
        m_col1, m_col2, m_col3 = st.columns([4, 4, 3])
        with m_col1:
            if st.button("Compare: Standard vs Baseline"):
                st.session_state.comparison_mode = "Standard vs Baseline"
        with m_col2:
            if st.button("Compare: Standard vs Counter"):
                st.session_state.comparison_mode = "Standard vs Counter"
        with m_col3:
            if st.button("Reset Default Comparison"):
                st.session_state.comparison_mode = "Baseline vs Counter"

    st.markdown("<br/>", unsafe_allow_html=True)

    # Route programmatic text targets based on the chosen luxury operational strategy mode
    if st.session_state.vault_expanded and st.session_state.comparison_mode == "Standard vs Baseline":
        left_paras = st.session_state.uploaded_files_data[template_file] if template_file else base_paras
        right_paras = st.session_state.uploaded_files_data[base_file]
        col1_title, col2_title = "STANDARD MODEL", roles[st.session_state.current_baseline]
    elif st.session_state.vault_expanded and st.session_state.comparison_mode == "Standard vs Counter":
        left_paras = st.session_state.uploaded_files_data[template_file] if template_file else base_paras
        right_paras = st.session_state.uploaded_files_data[counter_file]
        col1_title, col2_title = "STANDARD MODEL", roles[st.session_state.current_counter]
    else:
        left_paras = st.session_state.uploaded_files_data[base_file]
        right_paras = st.session_state.uploaded_files_data[counter_file]
        col1_title, col2_title = roles[st.session_state.current_baseline], roles[st.session_state.current_counter]

    # Compute Structural Differences Matrix
    matcher = difflib.SequenceMatcher(None, left_paras, right_paras)
    template_paras = st.session_state.uploaded_files_data[template_file] if template_file else []
    v_idx = 0

    # Layout Column Assignment Block
    if st.session_state.vault_expanded:
        # Dynamic 4-Column Layout Grid
        c_vault, c_left, c_right, c_delta = st.columns([3, 3, 3, 3])
    else:
        # Standard 3-Column Layout Grid
        c_left, c_right, c_delta = st.columns([4, 4, 3])

    # Header Rendering Row
    if st.session_state.vault_expanded:
        with c_vault: st.markdown("<p style='font-size:11px; text-transform:uppercase; color:#525252;'>STATIC STANDARD VAULT</p>", unsafe_allow_html=True)
    with c_left: st.markdown(f"<p style='font-size:11px; text-transform:uppercase; color:#525252;'>{col1_title}</p>", unsafe_allow_html=True)
    with c_right: st.markdown(f"<p style='font-size:11px; text-transform:uppercase; color:#525252;'>{col2_title}</p>", unsafe_allow_html=True)
    with c_delta: st.markdown("<p style='font-size:11px; text-transform:uppercase; color:#525252;'>SMART DELTA OUTPUT</p>", unsafe_allow_html=True)

    # Main One-Whole-Scroll Synchronized Paragraph Loop
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            for i in range(i1, i2):
                block_html = f'<div class="stream-paragraph">{left_paras[i]}</div>'
                
                if st.session_state.vault_expanded:
                    with c_vault:
                        v_text = template_paras[v_idx] if v_idx < len(template_paras) else ""
                        st.markdown(f'<div class="stream-paragraph" style="color:#666666;">{v_text}</div>', unsafe_allow_html=True)
                        v_idx += 1
                        
                with c_left: st.markdown(block_html, unsafe_allow_html=True)
                with c_right: st.markdown(block_html, unsafe_allow_html=True)
                with c_delta:
                    # Spatial layout spacer
                    st.markdown('<div class="stream-paragraph" style="border-bottom:none;"></div>', unsafe_allow_html=True)
                    
        elif tag == 'replace':
            for i, j in zip(range(i1, i2), range(j1, j2)):
                signature = extract_clause_signature(left_paras[i])
                trace_detail = generate_advisory_trace_text(left_paras[i], right_paras[j], signature)
                h1, h2 = compute_token_diff_html(left_paras[i], right_paras[j])
                unique_id = f"mod_{i}_{j}"

                if st.session_state.vault_expanded:
                    with c_vault:
                        v_text = template_paras[v_idx] if v_idx < len(template_paras) else ""
                        st.markdown(f'<div class="stream-paragraph" style="color:#666666;">{v_text}</div>', unsafe_allow_html=True)
                        v_idx += 1

                with c_left: st.markdown(f'<div class="stream-paragraph"><span class="trace-flag">▲ Modified</span>{h1}</div>', unsafe_allow_html=True)
                with c_right: st.markdown(f'<div class="stream-paragraph"><span class="trace-flag">▲ Modified</span>{h2}</div>', unsafe_allow_html=True)
                with c_delta:
                    st.markdown(f"""
                        <div class="stream-paragraph" style="border-bottom:none; padding-bottom:0;">
                            <div class="advisory-panel">
                                <div class="advisory-header">{signature}</div>
                                <p style='color:#a3a3a3; font-size:12px; margin:0;'>{trace_detail}</p>
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
                    opts = ["Select Protocol Action", "Issue Counter-Omission Challenge", "Accept Alteration via Carve-Out", "Escalate to Senior Partner Review"]
                    st.selectbox("Action Directive", opts, key=f"act_{unique_id}", label_visibility="collapsed")
                    st.text_input("Risk Notes", key=f"note_{unique_id}", placeholder="Log advisory stance...", label_visibility="collapsed")

        elif tag == 'delete':
            for i in range(i1, i2):
                signature = extract_clause_signature(left_paras[i])
                unique_id = f"del_{i}"

                if st.session_state.vault_expanded:
                    with c_vault:
                        v_text = template_paras[v_idx] if v_idx < len(template_paras) else ""
                        st.markdown(f'<div class="stream-paragraph" style="color:#666666;">{v_text}</div>', unsafe_allow_html=True)
                        v_idx += 1

                with c_left: st.markdown(f'<div class="stream-paragraph"><span class="trace-flag" style="color:#f87171;">◼ Omitted</span><span class="del-token">{left_paras[i]}</span></div>', unsafe_allow_html=True)
                with c_right: st.markdown('<div class="stream-paragraph" style="color:#404040; font-style:italic; padding-top:2rem;">Provision structurally deleted in this version.</div>', unsafe_allow_html=True)
                with c_delta:
                    st.markdown(f"""
                        <div class="stream-paragraph" style="border-bottom:none; padding-bottom:0;">
                            <div class="advisory-panel" style="border-left-color:#f87171;">
                                <div class="advisory-header">{signature}</div>
                                <p style='color:#f87171; font-size:12px; margin:0;'>Omission: This clause was stripped from the contract body.</p>
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
                    opts = ["Select Protocol Action", "Demand Immediate Clause Restoration", "Approve Deletion Conditionally"]
                    st.selectbox("Action Directive", opts, key=f"act_{unique_id}", label_visibility="collapsed")
                    st.text_input("Risk Notes", key=f"note_{unique_id}", placeholder="Log advisory stance...", label_visibility="collapsed")

        elif tag == 'insert':
            for j in range(j1, j2):
                signature = extract_clause_signature(right_paras[j])
                unique_id = f"ins_{j}"

                if st.session_state.vault_expanded:
                    with c_vault:
                        v_text = template_paras[v_idx] if v_idx < len(template_paras) else ""
                        st.markdown(f'<div class="stream-paragraph" style="color:#666666;">{v_text}</div>', unsafe_allow_html=True)
                        v_idx += 1

                with c_left: st.markdown('<div class="stream-paragraph" style="color:#404040; font-style:italic; padding-top:2rem;">Provision absent in baseline framework.</div>', unsafe_allow_html=True)
                with c_right: st.markdown(f'<div class="stream-paragraph"><span class="trace-flag" style="color:#34d399;">◆ Insertion</span><span class="add-token">{right_paras[j]}</span></div>', unsafe_allow_html=True)
                with c_delta:
                    st.markdown(f"""
                        <div class="stream-paragraph" style="border-bottom:none; padding-bottom:0;">
                            <div class="advisory-panel" style="border-left-color:#34d399;">
                                <div class="advisory-header">{signature}</div>
                                <p style='color:#34d399; font-size:12px; margin:0;'>Insertion: Counterparty injected a non-standard provision block.</p>
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
                    opts = ["Select Protocol Action", "Issue Total Strike-Out Directive", "Accept New Provision Terms"]
                    st.selectbox("Action Directive", opts, key=f"act_{unique_id}", label_visibility="collapsed")
                    st.text_input("Risk Notes", key=f"note_{unique_id}", placeholder="Log advisory stance...", label_visibility="collapsed")

    # Persistent Sticky Control Deliverable Row
    st.markdown("<br/><br/><hr style='border-color:#1a1a1a;'/>", unsafe_allow_html=True)
    b_col1, b_col2 = st.columns(2)
    
    with b_col1:
        # Compiles state text buffer on demand for manual file system output
        report_data = compile_markdown_report(base_file, counter_file)
        st.download_button(
            label="📥 Download Advisory Brief",
            data=report_data,
            file_name=f"DELTA_Advisory_Brief_{datetime.now().strftime('%Y%m%d')}.md",
            mime="text/markdown"
        )
    with b_col2:
        if st.button("Teardown Workspace Session"):
            st.session_state.processing_complete = False
            st.rerun()

# ==========================================
# BLOCK 8: RUNTIME MATRIX CONTROL
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
