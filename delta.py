# ==========================================
# DELTA | LUXURY LEGAL-TECH SYSTEM ARCHITECTURE
# ==========================================

import streamlit as st
import difflib
import io
import re
import fitz  # PyMuPDF
from docx import Document

# ==========================================
# BLOCK 1: SESSION STATE INITIALIZATION
# ==========================================
# [DIRECTIVE: Strict initialization of processing array arrays and states.]
if 'initialized' not in st.session_state:
    st.session_state.initialized = True
    st.session_state.uploaded_files_data = {}  
    st.session_state.file_order = []           
    st.session_state.file_roles = {}           
    st.session_state.processing_complete = False 
    st.session_state.current_baseline = 0      
    st.session_state.current_counter = 1       
    st.session_state.peek_template = False     

# ==========================================
# BLOCK 2: PREMIUM HIGH-END DARK CSS INJECTION
# ==========================================
# [DIRECTIVE: Force complete dark mode grayscale aesthetic with gold accents and elegant typography rules.]
def inject_premium_luxury_css():
    st.markdown("""
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Cinzel:wght@400;600&family=Inter:wght@300;400;500&display=swap');
            
            /* Enforce Global Dark Grayscale Canvas */
            .stApp {
                background-color: #121212 !important;
                color: #e5e5e5 !important;
                font-family: 'Inter', sans-serif !important;
            }
            
            /* Sidebar Restyling */
            [data-testid="stSidebar"] {
                background-color: #161616 !important;
                border-right: 1px solid #262626;
            }
            
            /* Luxury Headers */
            h1, h2, .luxury-header {
                font-family: 'Cinzel', serif !important;
                color: #d4af37 !important; /* Premium Metallic Gold */
                letter-spacing: 2px;
                font-weight: 400;
                margin-bottom: 0.5rem;
            }
            h3, h4 {
                font-family: 'Inter', sans-serif !important;
                color: #ffffff !important;
                font-weight: 500;
                letter-spacing: 0.5px;
            }
            
            /* Minimal Buttons styling */
            .stButton > button {
                background-color: transparent !important;
                color: #d4af37 !important;
                border: 1px solid #d4af37 !important;
                border-radius: 0px !important; /* Sharp technical edges */
                font-family: 'Inter', sans-serif;
                font-size: 13px !important;
                letter-spacing: 1px;
                text-transform: uppercase;
                padding: 0.5rem 1.5rem !important;
                transition: all 0.4s ease;
            }
            .stButton > button:hover {
                background-color: #d4af37 !important;
                color: #121212 !important;
                box-shadow: 0 0 10px rgba(212, 175, 55, 0.2);
            }
            
            /* Document Page Simulation (Continuous Text Sheet) */
            .document-canvas {
                background-color: #1c1c1c;
                color: #ffffff;
                padding: 2.5rem;
                border: 1px solid #262626;
                min-height: 700px;
                font-size: 14px;
                line-height: 1.8;
                letter-spacing: 0.3px;
                white-space: pre-wrap;
            }
            
            /* Inline Paragraph Container Block mapping to continuous view */
            .para-block {
                position: relative;
                margin-bottom: 1.5rem;
                padding-right: 2rem;
            }
            
            /* Micro Structural Badges */
            .structural-indicator {
                color: #d4af37;
                font-size: 11px;
                font-family: 'Cinzel', serif;
                text-transform: uppercase;
                margin-bottom: 0.25rem;
                display: block;
                letter-spacing: 1px;
            }
            
            /* High Contrast Diff Highlighting inside dark sheets */
            .add-segment {
                background-color: #14321a !important; /* Deep Forest Green Accent */
                color: #a7f3d0 !important;
                padding: 1px 3px;
                border-bottom: 1px solid #34d399;
            }
            .del-segment {
                background-color: #3b1818 !important; /* Deep Crimson Maroon */
                color: #fca5a5 !important;
                text-decoration: line-through;
                padding: 1px 3px;
                border-bottom: 1px solid #f87171;
            }
            
            /* Minimalist Delta Side Cards */
            .delta-meta-card {
                background-color: #161616;
                border-left: 1px solid #d4af37;
                padding: 1rem;
                margin-bottom: 1rem;
                font-size: 13px;
                color: #a3a3a3;
                line-height: 1.5;
            }
            .delta-meta-title {
                font-family: 'Cinzel', serif;
                color: #ffffff;
                font-size: 12px;
                margin-bottom: 0.25rem;
                letter-spacing: 1px;
            }
            
            /* Suppress system visuals */
            .stCheckbox { color: #ffffff; }
        </style>
    """, unsafe_allow_html=True)

# ==========================================
# BLOCK 3: PURE DETERMINISTIC DOCUMENT PARSING
# ==========================================
def parse_pdf(file_bytes):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    paragraphs = []
    for page in doc:
        blocks = page.get_text("blocks")
        for b in blocks:
            text = b[4].strip()
            if text:
                # Normalize line breaks within a structural block
                text_clean = " ".join(text.split())
                paragraphs.append(text_clean)
    return paragraphs

def parse_docx(file_bytes):
    doc = Document(io.BytesIO(file_bytes))
    paragraphs = []
    for p in doc.paragraphs:
        if p.text.strip():
            text_clean = " ".join(p.text.split())
            paragraphs.append(text_clean)
    return paragraphs

def process_uploaded_files(uploaded_files):
    for file in uploaded_files:
        if file.name not in st.session_state.uploaded_files_data:
            bytes_data = file.read()
            if file.name.endswith('.pdf'):
                text_list = parse_pdf(bytes_data)
            elif file.name.endswith('.docx'):
                text_list = parse_docx(bytes_data)
            else:
                continue
            st.session_state.uploaded_files_data[file.name] = text_list
            if file.name not in st.session_state.file_order:
                st.session_state.file_order.append(file.name)
                st.session_state.file_roles[file.name] = "v1: Baseline"

# ==========================================
# BLOCK 4: REVISION STRUCTURAL MUTATIONS
# ==========================================
def shift_up(index):
    if index > 0:
        st.session_state.file_order[index], st.session_state.file_order[index-1] = \
            st.session_state.file_order[index-1], st.session_state.file_order[index]

def shift_down(index):
    if index < len(st.session_state.file_order) - 1:
        st.session_state.file_order[index], st.session_state.file_order[index+1] = \
            st.session_state.file_order[index+1], st.session_state.file_order[index]

# ==========================================
# BLOCK 5: DESCRIPTIVE ANALYSIS ENGINE
# ==========================================
def isolate_clause_title(text):
    # [FIXED]: Correctly balanced regex capture groups to avoid re.PatternError
    match = re.match(r'^([A-Za-z0-9\.\s]+(?:\b[A-Z]{2,}\b|\bRent\b|\bTerm\b|\bDeposit\b|\bUse\b))', text)
    if match:
        return match.group(1).strip()
    words = text.split()
    return " ".join(words[:3]) if len(words) >= 3 else "Clause Block"

def generate_descriptive_diff(text1, text2, clause_name):
    # Generates exact contextual text trace transformations
    matcher = difflib.SequenceMatcher(None, text1.split(), text2.split())
    changes = []
    
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'replace':
            old_words = " ".join(text1.split()[i1:i2])
            new_words = " ".join(text2.split()[j1:j2])
            # Keep trace window brief and functional
            if len(old_words) < 40 and len(new_words) < 40:
                changes.append(f"Changed '{old_words}' to '{new_words}'")
            else:
                changes.append(f"Terms modified within segment execution")
        elif tag == 'delete':
            deleted_words = " ".join(text1.split()[i1:i2])
            if len(deleted_words) < 40:
                changes.append(f"Removed '{deleted_words}'")
        elif tag == 'insert':
            inserted_words = " ".join(text2.split()[j1:j2])
            if len(inserted_words) < 40:
                changes.append(f"Inserted '{inserted_words}'")
                
    if changes:
        return f"Modification in {clause_name}: " + "; ".join(changes)
    return f"Modification in {clause_name}"

def compute_word_diff_html(text1, text2):
    matcher = difflib.SequenceMatcher(None, text1.split(), text2.split())
    out1, out2 = [], []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        w1 = " ".join(text1.split()[i1:i2])
        w2 = " ".join(text2.split()[j1:j2])
        if tag == 'equal':
            out1.append(w1)
            out2.append(w2)
        elif tag == 'replace':
            out1.append(f'<span class="del-segment">{w1}</span>')
            out2.append(f'<span class="add-segment">{w2}</span>')
        elif tag == 'delete':
            out1.append(f'<span class="del-segment">{w1}</span>')
        elif tag == 'insert':
            out2.append(f'<span class="add-segment">{w2}</span>')
    return " ".join(out1), " ".join(out2)

def generate_continuous_document_flow(base_paras, counter_paras):
    matcher = difflib.SequenceMatcher(None, base_paras, counter_paras)
    base_html_collector = []
    counter_html_collector = []
    descriptive_summaries = []
    
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            for i in range(i1, i2):
                block = f'<div class="para-block">{base_paras[i]}</div>'
                base_html_collector.append(block)
                counter_html_collector.append(block)
                
        elif tag == 'replace':
            for i, j in zip(range(i1, i2), range(j1, j2)):
                clause_title = isolate_clause_title(base_paras[i])
                desc = generate_descriptive_diff(base_paras[i], counter_paras[j], clause_title)
                descriptive_summaries.append(desc)
                
                h1, h2 = compute_word_diff_html(base_paras[i], counter_paras[j])
                base_html_collector.append(f'<div class="para-block"><span class="structural-indicator">▲ Modified</span>{h1}</div>')
                counter_html_collector.append(f'<div class="para-block"><span class="structural-indicator">▲ Modified</span>{h2}</div>')
                
        elif tag == 'delete':
            for i in range(i1, i2):
                clause_title = isolate_clause_title(base_paras[i])
                descriptive_summaries.append(f"Structural Deletion: {clause_title} was removed completely.")
                base_html_collector.append(f'<div class="para-block"><span class="structural-indicator" style="color:#f87171;">◼ Omitted Clause</span><span class="del-segment">{base_paras[i]}</span></div>')
                
        elif tag == 'insert':
            for j in range(j1, j2):
                clause_title = isolate_clause_title(counter_paras[j])
                descriptive_summaries.append(f"Structural Insertion: Introduced new provision under {clause_title}.")
                counter_html_collector.append(f'<div class="para-block"><span class="structural-indicator" style="color:#34d399;">◆ New Provision</span><span class="add-segment">{counter_paras[j]}</span></div>')
                
    return "".join(base_html_collector), "".join(counter_html_collector), descriptive_summaries

# ==========================================
# BLOCK 6: PHASE 1 - STAGING WORKSPACE
# ==========================================
def render_staging_view():
    st.markdown('<div class="luxury-header">DELTA</div>', unsafe_allow_html=True)
    st.markdown('<p style="color:#737373; font-size:12px; letter-spacing:1px; text-transform:uppercase;">Revision Tree Staging</p>', unsafe_allow_html=True)
    st.markdown("<br/>", unsafe_allow_html=True)
    
    uploaded_files = st.file_uploader("Ingest Core Binaries", type=['pdf', 'docx'], accept_multiple_files=True, label_visibility="collapsed")
    
    if uploaded_files:
        process_uploaded_files(uploaded_files)
        st.markdown("<br/>", unsafe_allow_html=True)
        
        for idx, filename in enumerate(st.session_state.file_order):
            col1, col2, col3, col4 = st.columns([5, 3, 1, 1])
            with col1:
                st.markdown(f'<p style="font-size:14px; padding-top:10px; color:#ffffff;">◼ {filename}</p>', unsafe_allow_html=True)
            with col2:
                roles = ["Standard Template", "v1: Baseline", "v2: Counter", "v3: Counter", "v4: Counter"]
                st.session_state.file_roles[filename] = st.selectbox(
                    "Classification", roles, index=roles.index(st.session_state.file_roles.get(filename, "v1: Baseline")), key=f"role_{filename}", label_visibility="collapsed"
                )
            with col3:
                st.button("▲", key=f"up_{filename}", on_click=shift_up, args=(idx,))
            with col4:
                st.button("▼", key=f"down_{filename}", on_click=shift_down, args=(idx,))
                
        st.markdown("<br/><br/>", unsafe_allow_html=True)
        if st.button("Execute Core Comparison"):
            st.session_state.processing_complete = True
            st.rerun()

# ==========================================
# BLOCK 7: PHASE 2 & 3 - SYSTEM REVIEW WORKSPACE
# ==========================================
def render_dashboard_view():
    st.markdown('<div class="luxury-header">DELTA Workspace</div>', unsafe_allow_html=True)
    st.markdown("<br/>", unsafe_allow_html=True)
    
    ordered_files = st.session_state.file_order
    roles = [st.session_state.file_roles[f] for f in ordered_files]
    
    # Minimalist Control Toolbar
    t_col1, t_col2, t_col3 = st.columns([4, 4, 3])
    with t_col1:
        st.session_state.current_baseline = st.selectbox("Baseline File", range(len(ordered_files)), format_func=lambda x: f"BASE // {roles[x]}")
    with t_col2:
        st.session_state.current_counter = st.selectbox("Counter File", range(len(ordered_files)), format_func=lambda x: f"CNTR // {roles[x]}", index=min(1, len(ordered_files)-1))
    with t_col3:
        st.markdown("<p style='font-size:11px; color:#737373; margin-bottom:5px; text-transform:uppercase;'>Vault View</p>", unsafe_allow_html=True)
        st.session_state.peek_template = st.checkbox("Peek Template", value=st.session_state.peek_template)
        
    # Phase 3 Sidebar Integration
    if st.session_state.peek_template:
        with st.sidebar:
            st.markdown('<div class="luxury-header" style="font-size:16px;">Vault View</div>', unsafe_allow_html=True)
            template_file = next((f for f in ordered_files if "Standard Template" in st.session_state.file_roles[f]), None)
            if template_file:
                st.markdown(f'<p style="color:#737373; font-size:11px;">Target: {template_file}</p>', unsafe_allow_html=True)
                template_paras = st.session_state.uploaded_files_data[template_file]
                t_flow = "".join([f'<div class="para-block" style="font-size:12px; color:#a3a3a3;">{p}</div>' for p in template_paras])
                st.markdown(f'<div class="document-canvas" style="padding:1.5rem; min-height:auto;">{t_flow}</div>', unsafe_allow_html=True)
            else:
                st.markdown('<p style="color:#737373; font-size:12px;">Standard Template reference missing in staging execution tree.</p>', unsafe_allow_html=True)

    st.markdown("<br/>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([4, 4, 3])
    
    base_file = ordered_files[st.session_state.current_baseline]
    counter_file = ordered_files[st.session_state.current_counter]
    
    # Process modifications
    base_html_flow, counter_html_flow, summaries = generate_continuous_document_flow(
        st.session_state.uploaded_files_data[base_file],
        st.session_state.uploaded_files_data[counter_file]
    )
    
    # Render Column 1: Baseline Document
    with col1:
        st.markdown(f"<p style='font-size:11px; text-transform:uppercase; color:#737373; letter-spacing:1px;'>{roles[st.session_state.current_baseline]}</p>", unsafe_allow_html=True)
        st.markdown(f'<div class="document-canvas">{base_html_flow}</div>', unsafe_allow_html=True)
        
    # Render Column 2: Counter Document
    with col2:
        st.markdown(f"<p style='font-size:11px; text-transform:uppercase; color:#737373; letter-spacing:1px;'>{roles[st.session_state.current_counter]}</p>", unsafe_allow_html=True)
        st.markdown(f'<div class="document-canvas">{counter_html_flow}</div>', unsafe_allow_html=True)
        
    # Render Column 3: Smart Delta Output Sidebar
    with col3:
        st.markdown("<p style='font-size:11px; text-transform:uppercase; color:#737373; letter-spacing:1px;'>Smart Delta Output</p>", unsafe_allow_html=True)
        if not summaries:
            st.markdown('<div class="delta-meta-card"><div class="delta-meta-title">Status Verified</div>Documents align identically across structural checkpoints.</div>', unsafe_allow_html=True)
        else:
            for idx, summary in enumerate(summaries):
                st.markdown(f"""
                    <div class="delta-meta-card">
                        <div class="delta-meta-title">Trace Modification 0{idx+1}</div>
                        {summary}
                    </div>
                """, unsafe_allow_html=True)
                
        st.markdown("<br/>", unsafe_allow_html=True)
        if st.button("Reset Staging Environment"):
            st.session_state.processing_complete = False
            st.rerun()

# ==========================================
# BLOCK 8: MAIN RUNTIME ORCHESTRATOR
# ==========================================
def main():
    st.set_page_config(page_title="DELTA", layout="wide", initial_sidebar_state="collapsed")
    inject_premium_luxury_css()
    
    if not st.session_state.processing_complete:
        render_staging_view()
    else:
        render_dashboard_view()

if __name__ == "__main__":
    main()
