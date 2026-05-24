# ==========================================
# DELTA | PREMIUM LEGAL ARCHITECTURE SUITE
# ==========================================

import streamlit as st
import difflib
import io
import re
import hashlib
from datetime import datetime
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ==========================================
# BLOCK 1: STATE HYDRATION
# ==========================================
if 'initialized' not in st.session_state:
    st.session_state.initialized = True
    st.session_state.uploaded_files_data = {}   
    st.session_state.uploaded_files_hashes = {} 
    st.session_state.file_order = []           
    st.session_state.file_roles = {}           
    st.session_state.processing_complete = False 
    st.session_state.current_baseline = 0      
    st.session_state.current_counter = 1       
    st.session_state.comparison_mode = "Baseline vs Counter"

# ==========================================
# BLOCK 2: SYSTEM DESIGN SYSTEM (CSS)
# ==========================================
def inject_luxury_system_css():
    st.markdown("""
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Cinzel:wght@400;500&family=Inter:wght@300;400;500;600&display=swap');
            
            /* Enforce Global Dark Grayscale Foundation */
            .stApp {
                background-color: #0c0c0c !important;
                color: #e5e5e5 !important;
                font-family: 'Inter', sans-serif !important;
            }
            
            /* Sidebar Custom Dark Aesthetic */
            [data-testid="stSidebar"] {
                background-color: #111111 !important;
                border-right: 1px solid #1a1a1a;
            }
            
            /* High-End Clean Corporate Typography */
            .brand-title {
                font-family: 'Cinzel', serif !important;
                color: #d4af37 !important;
                letter-spacing: 4px;
                font-weight: 400;
                font-size: 28px;
                margin-bottom: 0.2rem;
            }
            .brand-subtitle {
                font-family: 'Inter', sans-serif;
                color: #525252;
                font-size: 10px;
                letter-spacing: 2px;
                text-transform: uppercase;
                margin-bottom: 2.5rem;
            }
            
            /* Strict Centering System for Landing View */
            .landing-wrapper {
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
                text-align: center;
                width: 100%;
                margin: 0 auto;
            }
            
            /* Skinned Premium Drag & Drop Uploader */
            .stFileUploader {
                border: 1px dashed #d4af37 !important;
                background-color: #111111 !important;
                border-radius: 0px !important;
                padding: 1.5rem !important;
                width: 100% !important;
            }
            
            /* Minimalist Sharp Buttons */
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
            }
            
            /* Row alignment padding */
            .aligned-matrix-row {
                padding: 0.8rem 0;
                margin: 0;
                background: transparent;
            }
            
            /* Text Flow Formatting */
            .stream-paragraph {
                color: #e5e5e5;
                font-size: 13px;
                line-height: 1.8;
                word-wrap: break-word;
            }
            
            /* Clean Text Tokens with Light Pastel Overlays */
            .add-token { background-color: #1c3d25 !important; color: #6ee7b7 !important; padding: 2px 4px; }
            .del-token { background-color: #4c1d1d !important; color: #fca5a5 !important; text-decoration: line-through; padding: 2px 4px; }
            .trace-flag { font-family: 'Cinzel', serif; color: #d4af37; font-size: 10px; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 0.4rem; display: block; }
            
            /* Standard Advisory Cards */
            .advisory-panel {
                background-color: #161616;
                border-left: 2px solid #d4af37;
                padding: 0.8rem;
                margin-bottom: 0.5rem;
            }
            .advisory-header { font-family: 'Cinzel', serif; color: #ffffff; font-size: 11px; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 0.3rem; }
            
            /* FEATURE 1: Compliance Gatekeeper Crimson Alerts */
            .alert-crimson {
                background-color: #1f0b0b !important;
                border-left: 3px solid #ef4444 !important;
            }
            .alert-gold-text {
                color: #d4af37;
                font-weight: 600;
                font-size: 11px;
                letter-spacing: 1px;
                text-transform: uppercase;
            }

            /* FEATURE 2: Surgical Drift Tracker Badges */
            .drift-badge {
                background-color: #d4af37;
                color: #0c0c0c;
                padding: 2px 6px;
                font-weight: 600;
                font-size: 9px;
                margin-left: 8px;
                border-radius: 2px;
                vertical-align: middle;
            }
            .high-drift-panel {
                border-left: 3px solid #d4af37 !important;
                background-color: #1a1710 !important;
            }
            
            /* Crypto Manifest Status Bar */
            .crypto-banner {
                font-family: monospace;
                font-size: 10px;
                background-color: #111111;
                border: 1px solid #1a1a1a;
                padding: 6px 12px;
                color: #737373;
                margin-bottom: 1rem;
            }
            
            /* Form Style Overrides */
            div[data-baseweb="select"] { background-color: #121212 !important; border-radius: 0px !important; }
            input {
                border-radius: 0px !important;
                background-color: #161616 !important;
                color: #ffffff !important;
                border: 1px solid #262626 !important;
                font-size: 12px !important;
                padding: 6px 10px !important;
            }

            /* Sidebar Tabs Styling override */
            button[data-baseweb="tab"] { background-color: transparent !important; font-family: 'Inter', sans-serif !important; }
            div[data-baseweb="tab-list"] { background-color: #111111 !important; border-bottom: 1px solid #1a1a1a !important; }
        </style>
    """, unsafe_allow_html=True)

# ==========================================
# BLOCK 3: ENGINE PARSING & CRYPTO HASHING
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
            file_hash = hashlib.sha256(bytes_data).hexdigest()
            st.session_state.uploaded_files_hashes[file.name] = file_hash
            
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
# BLOCK 4: FUZZY ALIGNMENT ENGINE
# ==========================================
def compute_fuzzy_alignment_matrix(left_paras, right_paras, threshold=0.45):
    matched_right_indices = set()
    alignment_opcodes = []
    
    for i, lp in enumerate(left_paras):
        best_ratio = 0.0
        best_j = None
        lp_tokens = sorted(lp.lower().split())
        
        for j, rp in enumerate(right_paras):
            if j in matched_right_indices:
                continue
            rp_tokens = sorted(rp.lower().split())
            ratio = difflib.SequenceMatcher(None, lp_tokens, rp_tokens).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_j = j
                
        if best_ratio >= threshold and best_j is not None:
            matched_right_indices.add(best_j)
            if best_ratio > 0.98 and lp == right_paras[best_j]:
                alignment_opcodes.append(('equal', i, i, best_j, best_j))
            else:
                alignment_opcodes.append(('replace', i, i, best_j, best_j))
        else:
            alignment_opcodes.append(('delete', i, i, None, None))
            
    for j in range(len(right_paras)):
        if j not in matched_right_indices:
            alignment_opcodes.append(('insert', None, None, j, j))
            
    alignment_opcodes.sort(key=lambda x: (x[1] if x[1] is not None else float('inf'), x[3] if x[3] is not None else 0))
    return alignment_opcodes

# ==========================================
# BLOCK 5: STRING TRANSFORMATION ANALYTICS
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
            if len(o) < 30 and len(n) < 30: traces.append(f"Changed '{o}' to '{n}'")
        elif tag == 'delete':
            d = " ".join(text1.split()[i1:i2])
            if len(d) < 30: traces.append(f"Omitted '{d}'")
        elif tag == 'insert':
            ins = " ".join(text2.split()[j1:j2])
            if len(ins) < 30: traces.append(f"Injected '{ins}'")
    return f"{signature} | " + ("; ".join(traces) if traces else "Terms modified.")

def compute_token_diff_html(text1, text2):
    matcher = difflib.SequenceMatcher(None, text1.split(), text2.split())
    out1, out2 = [], []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        w1 = " ".join(text1.split()[i1:i2])
        w2 = " ".join(text2.split()[j1:j2])
        if tag == 'equal':
            out1.append(w1); out2.append(w2)
        elif tag == 'replace':
            out1.append(f'<span class="del-token">{w1}</span>')
            out2.append(f'<span class="add-token">{w2}</span>')
        elif tag == 'delete':
            out1.append(f'<span class="del-token">{w1}</span>')
        elif tag == 'insert':
            out2.append(f'<span class="add-token">{w2}</span>')
    return " ".join(out1), " ".join(out2)

# ==========================================
# BLOCK 6: HIGH-FIDELITY COMPLIANCE EXPORTERS
# ==========================================
def set_run_background(run, color_hex):
    rPr = run._r.get_or_add_rPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    rPr.append(shd)

def export_landscape_docx(left_paras, right_paras, title_left, title_right, alignment_opcodes, hash_l, hash_r):
    doc = Document()
    section = doc.sections[-1]
    section.orientation = WD_ORIENT.LANDSCAPE
    new_width, new_height = section.page_height, section.page_width
    section.page_width = new_width
    section.page_height = new_height
    
    title = doc.add_paragraph()
    title.add_run("DELTA CONTRACT ADVISORY MATRIX").bold = True
    
    meta = doc.add_paragraph()
    meta.add_run(f"INTEGRITY MANIFEST LOG\nBASE SHA-256: {hash_l}\nCNTR SHA-256: {hash_r}\n").font.size = Pt(8)
    
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = f"BASELINE ({title_left})"
    hdr_cells[1].text = f"COUNTERPART ({title_right})"
    hdr_cells[2].text = "SMART DELTA EVALUATION"
    
    for tag, i1, _, j1, _ in alignment_opcodes:
        row = table.add_row()
        if tag == 'equal':
            row.cells[0].text = left_paras[i1]
            row.cells[1].text = right_paras[j1]
            row.cells[2].text = "No variance detected."
        elif tag == 'replace':
            p1 = row.cells[0].paragraphs[0]
            p2 = row.cells[1].paragraphs[0]
            m_words = difflib.SequenceMatcher(None, left_paras[i1].split(), right_paras[j1].split())
            
            for w_tag, w_i1, w_i2, w_j1, w_j2 in m_words.get_opcodes():
                w1_str = " ".join(left_paras[i1].split()[w_i1:w_i2]) + " "
                w2_str = " ".join(right_paras[j1].split()[w_j1:w_j2]) + " "
                if w_tag == 'equal':
                    p1.add_run(w1_str); p2.add_run(w2_str)
                else:
                    r1 = p1.add_run(w1_str); set_run_background(r1, "fee2e2")
                    r2 = p2.add_run(w2_str); set_run_background(r2, "dcfce7")
            
            act_val = st.session_state.get(f"act_mod_{i1}_{j1}", "Unassigned")
            note_val = st.session_state.get(f"note_mod_{i1}_{j1}", "")
            sig = extract_clause_signature(left_paras[i1])
            trace = generate_advisory_trace_text(left_paras[i1], right_paras[j1], sig)
            
            # Export Drift Data
            ratio = difflib.SequenceMatcher(None, left_paras[i1], right_paras[j1]).ratio()
            drift_score = (1.0 - ratio) * 100
            drift_str = f" [DRIFT: {drift_score:.1f}%]" if drift_score > 40 else ""
            
            row.cells[2].text = f"Evaluation: {trace}{drift_str}\nAction: {act_val}\nComments: {note_val}"
            
        elif tag == 'delete':
            r1 = row.cells[0].paragraphs[0].add_run(left_paras[i1])
            set_run_background(r1, "fee2e2")
            row.cells[1].text = "[Clause Omitted]"
            act_val = st.session_state.get(f"act_del_{i1}", "Unassigned")
            note_val = st.session_state.get(f"note_del_{i1}", "")
            row.cells[2].text = f"Evaluation: MISSING MANDATORY CLAUSE\nAction: {act_val}\nComments: {note_val}"
            
        elif tag == 'insert':
            row.cells[0].text = "[Absent from Baseline Template]"
            r2 = row.cells[1].paragraphs[0].add_run(right_paras[j1])
            set_run_background(r2, "dcfce7")
            act_val = st.session_state.get(f"act_ins_{j1}", "Unassigned")
            note_val = st.session_state.get(f"note_ins_{j1}", "")
            row.cells[2].text = f"Evaluation: Structural Insertion\nAction: {act_val}\nComments: {note_val}"

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def export_landscape_pdf(left_paras, right_paras, title_left, title_right, alignment_opcodes, hash_l, hash_r):
    doc = fitz.open()
    page_w, page_h = 792, 612
    page = doc.new_page(width=page_w, height=page_h)
    
    page.insert_text(fitz.Point(36, 30), "DELTA CONTRACT REVIEW MATRIX", fontsize=12, color=(0.83, 0.68, 0.21))
    page.insert_text(fitz.Point(36, 45), f"BASE SHA-256: {hash_l} | CNTR SHA-256: {hash_r}", fontsize=7, color=(0.4, 0.4, 0.4))
    
    y = 75
    col_w = 230
    c1_x, c2_x, c3_x = 36, 280, 524
    
    for tag, i1, _, j1, _ in alignment_opcodes:
        if y > (page_h - 90):
            page = doc.new_page(width=page_w, height=page_h)
            y = 50
            
        if tag == 'equal':
            page.insert_textbox(fitz.Rect(c1_x, y, c1_x + col_w, y + 65), left_paras[i1], fontsize=8, color=(0.1, 0.1, 0.1))
            page.insert_textbox(fitz.Rect(c2_x, y, c2_x + col_w, y + 65), right_paras[j1], fontsize=8, color=(0.1, 0.1, 0.1))
            page.insert_textbox(fitz.Rect(c3_x, y, c3_x + col_w, y + 65), "No variance detected.", fontsize=8, color=(0.5, 0.5, 0.5))
            y += 70
        elif tag == 'replace':
            page.draw_rect(fitz.Rect(c1_x - 4, y, c1_x + col_w + 4, y + 70), color=None, fill=(0.99, 0.94, 0.94))
            page.draw_rect(fitz.Rect(c2_x - 4, y, c2_x + col_w + 4, y + 70), color=None, fill=(0.92, 0.99, 0.95))
            
            page.insert_textbox(fitz.Rect(c1_x, y + 5, c1_x + col_w, y + 65), left_paras[i1], fontsize=8, color=(0.4, 0.1, 0.1))
            page.insert_textbox(fitz.Rect(c2_x, y + 5, c2_x + col_w, y + 65), right_paras[j1], fontsize=8, color=(0.1, 0.3, 0.1))
            
            act_val = st.session_state.get(f"act_mod_{i1}_{j1}", "Unassigned")
            note_val = st.session_state.get(f"note_mod_{i1}_{j1}", "")
            sig = extract_clause_signature(left_paras[i1])
            trace = generate_advisory_trace_text(left_paras[i1], right_paras[j1], sig)
            
            ratio = difflib.SequenceMatcher(None, left_paras[i1], right_paras[j1]).ratio()
            drift_score = (1.0 - ratio) * 100
            drift_str = f" [DRIFT: {drift_score:.1f}%]" if drift_score > 40 else ""
            
            page.insert_textbox(fitz.Rect(c3_x, y + 5, c3_x + col_w, y + 65), f"Eval: {trace}{drift_str}\nAction: {act_val}\nComments: {note_val}", fontsize=8, color=(0.1, 0.1, 0.1))
            y += 75
        elif tag == 'delete':
            page.draw_rect(fitz.Rect(c1_x - 4, y, c1_x + col_w + 4, y + 55), color=None, fill=(0.99, 0.94, 0.94))
            page.insert_textbox(fitz.Rect(c1_x, y + 5, c1_x + col_w, y + 50), left_paras[i1], fontsize=8, color=(0.4, 0.1, 0.1))
            page.insert_text(fitz.Point(c2_x, y + 15), "[Clause Omitted]", fontsize=8, color=(0.5, 0.5, 0.5))
            
            act_val = st.session_state.get(f"act_del_{i1}", "Unassigned")
            note_val = st.session_state.get(f"note_del_{i1}", "")
            page.insert_textbox(fitz.Rect(c3_x, y + 5, c3_x + col_w, y + 50), f"Eval: MISSING MANDATORY CLAUSE\nAction: {act_val}\nComments: {note_val}", fontsize=8, color=(0.1, 0.1, 0.1))
            y += 60
        elif tag == 'insert':
            page.draw_rect(fitz.Rect(c2_x - 4, y, c2_x + col_w + 4, y + 55), color=None, fill=(0.92, 0.99, 0.95))
            page.insert_text(fitz.Point(c1_x, y + 15), "[Absent from Baseline Template]", fontsize=8, color=(0.5, 0.5, 0.5))
            page.insert_textbox(fitz.Rect(c2_x, y + 5, c2_x + col_w, y + 50), right_paras[j1], fontsize=8, color=(0.1, 0.3, 0.1))
            
            act_val = st.session_state.get(f"act_ins_{j1}", "Unassigned")
            note_val = st.session_state.get(f"note_ins_{j1}", "")
            page.insert_textbox(fitz.Rect(c3_x, y + 5, c3_x + col_w, y + 50), f"Eval: Structural Insertion\nAction: {act_val}\nComments: {note_val}", fontsize=8, color=(0.1, 0.1, 0.1))
            y += 60
                
    return doc.write()

# ==========================================
# BLOCK 7: UI PHASE 1 - CENTERED PORTAL
# ==========================================
def render_premium_landing_view():
    st.markdown('<div class="landing-wrapper">', unsafe_allow_html=True)
    st.markdown('<div class="brand-title">DELTA</div>', unsafe_allow_html=True)
    st.markdown('<div class="brand-subtitle">Legal Infrastructure Pipeline</div>', unsafe_allow_html=True)
    
    uploaded_files = st.file_uploader("Ingest Core Binaries", type=['pdf', 'docx'], accept_multiple_files=True, label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)
    
    if uploaded_files:
        load_staged_matrices(uploaded_files)
        st.markdown("<div style='max-width:500px; margin: 0 auto;'>", unsafe_allow_html=True)
        
        for idx, filename in enumerate(st.session_state.file_order):
            col1, col2 = st.columns([6, 4])
            with col1:
                st.markdown(f'<p style="font-size:13px; color:#ffffff; padding-top:8px;">◼ {filename}</p>', unsafe_allow_html=True)
            with col2:
                roles = ["Standard Template", "v1: Baseline", "v2: Counter", "v3: Counter", "v4: Counter"]
                st.session_state.file_roles[filename] = st.selectbox(
                    f"Classification_{filename}", roles, 
                    index=roles.index(st.session_state.file_roles.get(filename, "v1: Baseline")), 
                    label_visibility="collapsed", key=f"stage_role_{filename}"
                )
                
        st.markdown("<br/>", unsafe_allow_html=True)
        if st.button("Compile Revision Tree"):
            st.session_state.processing_complete = True
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# BLOCK 8: UI PHASE 2 - REVIEW MATRIX
# ==========================================
def render_delta_contracts_view():
    st.markdown('<div class="brand-title">DELTA CONTRACTS</div>', unsafe_allow_html=True)
    st.markdown("<br/>", unsafe_allow_html=True)
    
    ordered_files = st.session_state.file_order
    roles = [st.session_state.file_roles[f] for f in ordered_files]
    
    base_file = ordered_files[st.session_state.current_baseline]
    counter_file = ordered_files[st.session_state.current_counter]
    template_file = next((f for f in ordered_files if "Standard Template" in st.session_state.file_roles[f]), None)

    # Core System Model Selection Layout
    t_col1, t_col2 = st.columns(2)
    with t_col1:
        st.session_state.current_baseline = st.selectbox("Baseline Selection Frame", range(len(ordered_files)), format_func=lambda x: f"BASE // {roles[x]}")
    with t_col2:
        st.session_state.current_counter = st.selectbox("Counterparty Selection Frame", range(len(ordered_files)), format_func=lambda x: f"CNTR // {roles[x]}", index=min(1, len(ordered_files)-1))

    # FEATURE 3: Split-Pane Executive Hub inside the Vault
    with st.sidebar:
        st.markdown('<div class="brand-title" style="font-size:16px; margin-top:1rem;">Vault Core</div>', unsafe_allow_html=True)
        if template_file:
            modes = ["Baseline vs Counter", "Standard vs Baseline", "Standard vs Counter"]
            st.session_state.comparison_mode = st.radio("Target Pipeline Routing", modes, index=modes.index(st.session_state.comparison_mode))
            st.markdown("<hr style='border-color:#1a1a1a;'/>", unsafe_allow_html=True)
            
            # Implementing the Tabbed Split-Pane Logic
            tab_std, tab_playbook = st.tabs(["STANDARD", "PLAYBOOK"])
            
            with tab_std:
                for p in st.session_state.uploaded_files_data[template_file]:
                    st.markdown(f'<div style="font-size:12px; color:#737373; margin-bottom:1rem; line-height:1.5;">{p}</div>', unsafe_allow_html=True)
                    
            with tab_playbook:
                st.markdown('<p style="font-size:11px; color:#d4af37; font-weight:bold; letter-spacing:1px;">ENTERPRISE PLAYBOOK RULES</p>', unsafe_allow_html=True)
                for p in st.session_state.uploaded_files_data[template_file]:
                    sig = extract_clause_signature(p)
                    # Automatically generate a mapped rule for each signature found in the template
                    st.markdown(f'''
                        <div style="background-color:#161616; border-left:2px solid #525252; padding:8px; margin-bottom:8px;">
                            <span style="font-size:10px; color:#a3a3a3; text-transform:uppercase;">Rule Mapping</span><br/>
                            <b style="font-size:11px; color:#e5e5e5;">{sig}</b><br/>
                            <span style="font-size:10px; color:#737373;">Standard inclusion required. Deviations over 40% drift require VP override.</span>
                        </div>
                    ''', unsafe_allow_html=True)
        else:
            st.markdown('<p style="color:#525252; font-size:12px;">Standard Template absent in staging array. Playbook tracking offline.</p>', unsafe_allow_html=True)

    st.markdown("<br/>", unsafe_allow_html=True)

    if st.session_state.comparison_mode == "Standard vs Baseline":
        left_paras = st.session_state.uploaded_files_data[template_file] if template_file else st.session_state.uploaded_files_data[base_file]
        right_paras = st.session_state.uploaded_files_data[base_file]
        hash_l = st.session_state.uploaded_files_hashes.get(template_file, "N/A")
        hash_r = st.session_state.uploaded_files_hashes.get(base_file, "N/A")
        col1_title, col2_title = "STANDARD MATRIX", roles[st.session_state.current_baseline]
    elif st.session_state.comparison_mode == "Standard vs Counter":
        left_paras = st.session_state.uploaded_files_data[template_file] if template_file else st.session_state.uploaded_files_data[base_file]
        right_paras = st.session_state.uploaded_files_data[counter_file]
        hash_l = st.session_state.uploaded_files_hashes.get(template_file, "N/A")
        hash_r = st.session_state.uploaded_files_hashes.get(counter_file, "N/A")
        col1_title, col2_title = "STANDARD MATRIX", roles[st.session_state.current_counter]
    else:
        left_paras = st.session_state.uploaded_files_data[base_file]
        right_paras = st.session_state.uploaded_files_data[counter_file]
        hash_l = st.session_state.uploaded_files_hashes.get(base_file, "N/A")
        hash_r = st.session_state.uploaded_files_hashes.get(counter_file, "N/A")
        col1_title, col2_title = roles[st.session_state.current_baseline], roles[st.session_state.current_counter]

    alignment_opcodes = compute_fuzzy_alignment_matrix(left_paras, right_paras)

    st.markdown(f"""
        <div class="crypto-banner">
            VALIDATION MANIFEST LOG // LOCKED STATUS<br/>
            [BASE_SIG]: {hash_l}<br/>[CNTR_SIG]: {hash_r}
        </div>
    """, unsafe_allow_html=True)

    # Column Headings
    h_col1, h_col2, h_col3 = st.columns([4, 4, 3])
    with h_col1: st.markdown(f"<p style='font-size:11px; text-transform:uppercase; color:#525252; font-weight:600;'>{col1_title}</p>", unsafe_allow_html=True)
    with h_col2: st.markdown(f"<p style='font-size:11px; text-transform:uppercase; color:#525252; font-weight:600;'>{col2_title}</p>", unsafe_allow_html=True)
    with h_col3: st.markdown("<p style='font-size:11px; text-transform:uppercase; color:#525252; font-weight:600;'>SMART DELTA EVALUATION</p>", unsafe_allow_html=True)

    # Row Aligned Processing Stream
    for tag, i1, _, j1, _ in alignment_opcodes:
        if tag == 'equal':
            st.markdown('<div class="aligned-matrix-row">', unsafe_allow_html=True)
            col1, col2, col3 = st.columns([4, 4, 3])
            with col1: st.markdown(f'<div class="stream-paragraph">{left_paras[i1]}</div>', unsafe_allow_html=True)
            with col2: st.markdown(f'<div class="stream-paragraph">{right_paras[j1]}</div>', unsafe_allow_html=True)
            with col3: st.write("")
            st.markdown('</div>', unsafe_allow_html=True)
                
        elif tag == 'replace':
            unique_id = f"mod_{i1}_{j1}"
            signature = extract_clause_signature(left_paras[i1])
            trace_detail = generate_advisory_trace_text(left_paras[i1], right_paras[j1], signature)
            h1, h2 = compute_token_diff_html(left_paras[i1], right_paras[j1])
            
            # FEATURE 2: Surgical Drift Tracker Logic
            raw_ratio = difflib.SequenceMatcher(None, left_paras[i1], right_paras[j1]).ratio()
            drift_score = (1.0 - raw_ratio) * 100
            
            if drift_score > 40:
                badge_html = f'<span class="drift-badge">{drift_score:.1f}% DRIFT</span>'
                panel_class = "advisory-panel high-drift-panel"
            else:
                badge_html = ""
                panel_class = "advisory-panel"

            st.markdown('<div class="aligned-matrix-row">', unsafe_allow_html=True)
            col1, col2, col3 = st.columns([4, 4, 3])
            with col1: st.markdown(f'<div class="stream-paragraph"><span class="trace-flag">▲ Modified</span>{h1}</div>', unsafe_allow_html=True)
            with col2: st.markdown(f'<div class="stream-paragraph"><span class="trace-flag">▲ Modified</span>{h2}</div>', unsafe_allow_html=True)
            with col3:
                st.markdown(f'<div class="{panel_class}"><div class="advisory-header">{signature}{badge_html}</div><p style="color:#a3a3a3; font-size:11px; margin:0;">{trace_detail}</p></div>', unsafe_allow_html=True)
                st.radio("Action Protocol", ["✓", "⚠", "⇄", "✕"], key=f"act_{unique_id}", horizontal=True, label_visibility="collapsed", index=None, help="✓ Accept | ⚠ Escalate | ⇄ Counter | ✕ Reject")
                st.text_input("Comments", key=f"note_{unique_id}", placeholder="Comments", label_visibility="collapsed")
            st.markdown('</div>', unsafe_allow_html=True)
                
        elif tag == 'delete':
            unique_id = f"del_{i1}"
            signature = extract_clause_signature(left_paras[i1])
            
            st.markdown('<div class="aligned-matrix-row">', unsafe_allow_html=True)
            col1, col2, col3 = st.columns([4, 4, 3])
            with col1: st.markdown(f'<div class="stream-paragraph"><span class="trace-flag" style="color:#f87171;">◼ Omitted</span><span class="del-token">{left_paras[i1]}</span></div>', unsafe_allow_html=True)
            with col2: st.markdown('<div class="stream-paragraph" style="color:#404040; font-style:italic;">Provision completely absent from this iteration track.</div>', unsafe_allow_html=True)
            with col3:
                # FEATURE 1: Compliance Gatekeeper logic replacing the standard delete panel
                st.markdown(f'''
                    <div class="advisory-panel alert-crimson">
                        <div class="alert-gold-text">⚠ MISSING CLAUSE ALERT</div>
                        <p style="color:#fca5a5; font-size:11px; margin:5px 0 0 0; line-height:1.4;">
                            Warning: Mandatory <b>{signature}</b> block has been entirely removed by the counter-party track.
                        </p>
                    </div>
                ''', unsafe_allow_html=True)
                st.radio("Action Protocol", ["✓", "⚠", "⇄", "✕"], key=f"act_{unique_id}", horizontal=True, label_visibility="collapsed", index=None, help="✓ Accept | ⚠ Escalate | ⇄ Counter | ✕ Reject")
                st.text_input("Comments", key=f"note_{unique_id}", placeholder="Comments", label_visibility="collapsed")
            st.markdown('</div>', unsafe_allow_html=True)
                
        elif tag == 'insert':
            unique_id = f"ins_{j1}"
            signature = extract_clause_signature(right_paras[j1])
            
            st.markdown('<div class="aligned-matrix-row">', unsafe_allow_html=True)
            col1, col2, col3 = st.columns([4, 4, 3])
            with col1: st.markdown('<div class="stream-paragraph" style="color:#404040; font-style:italic;">Provision absent in baseline architectural blueprint.</div>', unsafe_allow_html=True)
            with col2: st.markdown(f'<div class="stream-paragraph"><span class="trace-flag" style="color:#34d399;">◆ Insertion</span><span class="add-token">{right_paras[j1]}</span></div>', unsafe_allow_html=True)
            with col3:
                st.markdown(f'<div class="advisory-panel" style="border-left-color:#34d399;"><div class="advisory-header">{signature}</div><p style="color:#34d399; font-size:11px; margin:0;">Injected Clause Block</p></div>', unsafe_allow_html=True)
                st.radio("Action Protocol", ["✓", "⚠", "⇄", "✕"], key=f"act_{unique_id}", horizontal=True, label_visibility="collapsed", index=None, help="✓ Accept | ⚠ Escalate | ⇄ Counter | ✕ Reject")
                st.text_input("Comments", key=f"note_{unique_id}", placeholder="Comments", label_visibility="collapsed")
            st.markdown('</div>', unsafe_allow_html=True)

    # Branded Consolidated Single-Action Export Footer Row
    st.markdown("<br/><br/><hr style='border-color:#1a1a1a;'/>", unsafe_allow_html=True)
    b_col1, b_col2 = st.columns(2)
    
    with b_col1:
        target_format = st.selectbox("Export Options", ["Microsoft Word (.docx)", "Adobe Portable Document (.pdf)"], label_visibility="collapsed")
        if "Word" in target_format:
            export_bytes = export_landscape_docx(left_paras, right_paras, col1_title, col2_title, alignment_opcodes, hash_l, hash_r)
            mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            extension = "docx"
        else:
            export_bytes = export_landscape_pdf(left_paras, right_paras, col1_title, col2_title, alignment_opcodes, hash_l, hash_r)
            mime_type = "application/pdf"
            extension = "pdf"
            
        st.download_button(
            label="📥 Export Review Matrix",
            data=export_bytes,
            file_name=f"DELTA_Matrix_{datetime.now().strftime('%Y%m%d')}.{extension}",
            mime=mime_type
        )
        
    with b_col2:
        st.write("<div style='height:1px;'></div>", unsafe_allow_html=True)
        if st.button("Terminate Session"):
            st.session_state.processing_complete = False
            st.rerun()

# ==========================================
# BLOCK 9: ENTRYPOINT ORCHESTRATION
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
