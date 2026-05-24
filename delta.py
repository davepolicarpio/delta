# ==========================================
# DELTA | PREMIUM LEGAL ARCHITECTURE SUITE
# ==========================================

import streamlit as st
import difflib
import io
import re
from datetime import datetime
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_COLOR_INDEX

# ==========================================
# BLOCK 1: STATE HYDRATION
# ==========================================
# [DIRECTIVE: Explicitly preserve text indices, file mappings, and workflow annotations across session lifecycle.]
if 'initialized' not in st.session_state:
    st.session_state.initialized = True
    st.session_state.uploaded_files_data = {}  
    st.session_state.file_order = []           
    st.session_state.file_roles = {}           
    st.session_state.processing_complete = False 
    st.session_state.current_baseline = 0      
    st.session_state.current_counter = 1       
    st.session_state.comparison_mode = "Baseline vs Counter"

# ==========================================
# BLOCK 2: SYSTEM DESIGN SYSTEM (CSS)
# ==========================================
# [DIRECTIVE: Inject custom CSS to create an elite, minimal dark interface.]
# [DIRECTIVE: Continuous document-canvas sheets hide row boundaries, resolving visual dizziness.]
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
            
            /* Ultra-Minimal Centered Landing Box */
            .landing-center-box {
                max-width: 600px;
                margin: 0 auto;
                text-align: center;
                padding: 3rem 0;
            }
            
            /* Skinned Premium Drag & Drop Uploader */
            .stFileUploader {
                border: 1px dashed #d4af37 !important;
                background-color: #111111 !important;
                border-radius: 0px !important;
                padding: 1.5rem !important;
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
            
            /* Continuous Canvas View: No boxes or dark borders to remove visual fatigue */
            .continuous-doc-sheet {
                background-color: #121212;
                border: 1px solid #1c1c1c;
                padding: 1rem 1.5rem;
            }
            
            /* Row alignment padding without clunky visual boxes */
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
            
            /* Clean Text Tokens with Minimal Dark Overlays */
            .add-token { background-color: #0f2d19 !important; color: #34d399 !important; padding: 2px 4px; }
            .del-token { background-color: #3b1414 !important; color: #f87171 !important; text-decoration: line-through; padding: 2px 4px; }
            .trace-flag { font-family: 'Cinzel', serif; color: #d4af37; font-size: 10px; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 0.4rem; display: block; }
            
            /* Advisory Cards */
            .advisory-panel {
                background-color: #161616;
                border-left: 2px solid #d4af37;
                padding: 0.8rem;
                margin-bottom: 0.5rem;
            }
            .advisory-header { font-family: 'Cinzel', serif; color: #ffffff; font-size: 11px; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 0.3rem; }
            
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
        </style>
    """, unsafe_allow_html=True)

# ==========================================
# BLOCK 3: ENGINE PARSING DATA
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
# BLOCK 4: ALIGNMENT & EXTRACTION TOOLS
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
# BLOCK 5: MULTI-FORMAT LANDSCAPE EXPORTERS
# ==========================================
# [DIRECTIVE: Build landscape Word export engine featuring run-level text highlighting indices.]
def export_landscape_docx(left_paras, right_paras, title_left, title_right, matcher):
    doc = Document()
    section = doc.sections[-1]
    section.orientation = WD_ORIENT.LANDSCAPE
    new_width, new_height = section.page_height, section.page_width
    section.page_width = new_width
    section.page_height = new_height
    
    title = doc.add_paragraph()
    title.add_run("DELTA CONTRACT ADVISORY MATRIX").bold = True
    
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = f"BASELINE ({title_left})"
    hdr_cells[1].text = f"COUNTERPART ({title_right})"
    hdr_cells[2].text = "STRATEGIC EXECUTABLE COMMENTS"
    
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            for i in range(i1, i2):
                row = table.add_row()
                row.cells[0].text = left_paras[i]
                row.cells[1].text = left_paras[i]
        elif tag == 'replace':
            for i, j in zip(range(i1, i2), range(j1, j2)):
                row = table.add_row()
                # Run word level matcher to assign native word highlights
                p1 = row.cells[0].paragraphs[0]
                p2 = row.cells[1].paragraphs[0]
                m_words = difflib.SequenceMatcher(None, left_paras[i].split(), right_paras[j].split())
                
                for w_tag, w_i1, w_i2, w_j1, w_j2 in m_words.get_opcodes():
                    w1_str = " ".join(left_paras[i].split()[w_i1:w_i2]) + " "
                    w2_str = " ".join(right_paras[j].split()[w_j1:w_j2]) + " "
                    if w_tag == 'equal':
                        p1.add_run(w1_str); p2.add_run(w2_str)
                    else:
                        r1 = p1.add_run(w1_str); r1.font.highlight_color = WD_COLOR_INDEX.RED
                        r2 = p2.add_run(w2_str); r2.font.highlight_color = WD_COLOR_INDEX.GREEN
                
                act_val = st.session_state.get(f"act_mod_{i}_{j}", "Unassigned")
                note_val = st.session_state.get(f"note_mod_{i}_{j}", "")
                row.cells[2].text = f"Action: {act_val}\nComments: {note_val}"
        elif tag == 'delete':
            for i in range(i1, i2):
                row = table.add_row()
                r1 = row.cells[0].paragraphs[0].add_run(left_paras[i])
                r1.font.highlight_color = WD_COLOR_INDEX.RED
                row.cells[1].text = "[Clause Omitted]"
                act_val = st.session_state.get(f"act_del_{i}", "Unassigned")
                note_val = st.session_state.get(f"note_del_{i}", "")
                row.cells[2].text = f"Action: {act_val}\nComments: {note_val}"
        elif tag == 'insert':
            for j in range(j1, j2):
                row = table.add_row()
                row.cells[0].text = "[Absent from Baseline Template]"
                r2 = row.cells[1].paragraphs[0].add_run(right_paras[j])
                r2.font.highlight_color = WD_COLOR_INDEX.GREEN
                act_val = st.session_state.get(f"act_ins_{j}", "Unassigned")
                note_val = st.session_state.get(f"note_ins_{j}", "")
                row.cells[2].text = f"Action: {act_val}\nComments: {note_val}"

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# [DIRECTIVE: Implement highlighted landscape PDF export using standard PyMuPDF canvas streams.]
def export_landscape_pdf(left_paras, right_paras, title_left, title_right, matcher):
    doc = fitz.open()
    # Standard landscape bounds (11 x 8.5 inches)
    page_w, page_h = 792, 612
    page = doc.new_page(width=page_w, height=page_h)
    
    # Title Header Elements
    page.insert_text(fitz.Point(36, 40), "DELTA CONTRACT REVIEW MATRIX", fontsize=14, color=(0.83, 0.68, 0.21))
    page.insert_text(fitz.Point(36, 60), f"Executed: {datetime.now().strftime('%Y-%m-%d')} | Left: {title_left} vs Right: {title_right}", fontsize=9, color=(0.5, 0.5, 0.5))
    
    y = 90
    col_w = 230
    c1_x, c2_x, c3_x = 36, 280, 524
    
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if y > (page_h - 60): # Pagination break check
            page = doc.new_page(width=page_w, height=page_h)
            y = 50
            
        if tag == 'equal':
            for i in range(i1, i2):
                page.insert_textbox(fitz.Rect(c1_x, y, c1_x + col_w, y + 80), left_paras[i], fontsize=8, color=(0.1, 0.1, 0.1))
                page.insert_textbox(fitz.Rect(c2_x, y, c2_x + col_w, y + 80), left_paras[i], fontsize=8, color=(0.1, 0.1, 0.1))
                y += 75
        elif tag == 'replace':
            for i, j in zip(range(i1, i2), range(j1, j2)):
                # Draw minimal high-contrast status blocks behind the tracking strings
                page.draw_rect(fitz.Rect(c1_x - 4, y, c1_x + col_w + 4, y + 70), color=(0.95, 0.9, 0.9), fill=(0.99, 0.92, 0.92))
                page.draw_rect(fitz.Rect(c2_x - 4, y, c2_x + col_w + 4, y + 70), color=(0.9, 0.95, 0.9), fill=(0.92, 0.99, 0.94))
                
                page.insert_textbox(fitz.Rect(c1_x, y + 5, c1_x + col_w, y + 65), f"[MODIFIED] {left_paras[i]}", fontsize=8, color=(0.5, 0.1, 0.1))
                page.insert_textbox(fitz.Rect(c2_x, y + 5, c2_x + col_w, y + 65), f"[MODIFIED] {right_paras[j]}", fontsize=8, color=(0.1, 0.4, 0.1))
                
                act_val = st.session_state.get(f"act_mod_{i}_{j}", "Unassigned")
                note_val = st.session_state.get(f"note_mod_{i}_{j}", "")
                page.insert_textbox(fitz.Rect(c3_x, y + 5, c3_x + col_w, y + 65), f"Action: {act_val}\nComments: {note_val}", fontsize=8, color=(0.2, 0.2, 0.2))
                y += 80
        elif tag == 'delete':
            for i in range(i1, i2):
                page.draw_rect(fitz.Rect(c1_x - 4, y, c1_x + col_w + 4, y + 50), color=(0.95, 0.9, 0.9), fill=(0.99, 0.92, 0.92))
                page.insert_textbox(fitz.Rect(c1_x, y + 5, c1_x + col_w, y + 45), left_paras[i], fontsize=8, color=(0.5, 0.1, 0.1))
                page.insert_text(fitz.Point(c2_x, y + 15), "[Provision Omitted]", fontsize=8, color=(0.6, 0.6, 0.6))
                
                act_val = st.session_state.get(f"act_del_{i}", "Unassigned")
                note_val = st.session_state.get(f"note_del_{i}", "")
                page.insert_textbox(fitz.Rect(c3_x, y + 5, c3_x + col_w, y + 45), f"Action: {act_val}\nComments: {note_val}", fontsize=8, color=(0.2, 0.2, 0.2))
                y += 60
        elif tag == 'insert':
            for j in range(j1, j2):
                page.draw_rect(fitz.Rect(c2_x - 4, y, c2_x + col_w + 4, y + 50), color=(0.9, 0.95, 0.9), fill=(0.92, 0.99, 0.94))
                page.insert_text(fitz.Point(c1_x, y + 15), "[Absent from Baseline Template]", fontsize=8, color=(0.6, 0.6, 0.6))
                page.insert_textbox(fitz.Rect(c2_x, y + 5, c2_x + col_w, y + 45), right_paras[j], fontsize=8, color=(0.1, 0.4, 0.1))
                
                act_val = st.session_state.get(f"act_ins_{j}", "Unassigned")
                note_val = st.session_state.get(f"note_ins_{j}", "")
                page.insert_textbox(fitz.Rect(c3_x, y + 5, c3_x + col_w, y + 45), f"Action: {act_val}\nComments: {note_val}", fontsize=8, color=(0.2, 0.2, 0.2))
                y += 60
                
    return doc.write()

# ==========================================
# BLOCK 6: PHASE 1 - ULTRA-MINIMAL LANDING
# ==========================================
def render_premium_landing_view():
    # Strict alignment wrapper
    st.markdown('<div class="landing-center-box">', unsafe_allow_html=True)
    st.markdown('<div class="brand-title">DELTA</div>', unsafe_allow_html=True)
    st.markdown('<p style="color:#525252; font-size:10px; letter-spacing:2px; text-transform:uppercase; margin-bottom:2.5rem;">Legal Infrastructure Pipeline</p>', unsafe_allow_html=True)
    
    uploaded_files = st.file_uploader("Ingest Core Binaries", type=['pdf', 'docx'], accept_multiple_files=True, label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)
    
    if uploaded_files:
        load_staged_matrices(uploaded_files)
        st.markdown("<div style='max-width:600px; margin:0 auto;'>", unsafe_allow_html=True)
        
        for idx, filename in enumerate(st.session_state.file_order):
            col1, col2 = st.columns([7, 3])
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

    # File target selector row
    t_col1, t_col2 = st.columns(2)
    with t_col1:
        st.session_state.current_baseline = st.selectbox("Baseline Document", range(len(ordered_files)), format_func=lambda x: f"BASE // {roles[x]}")
    with t_col2:
        st.session_state.current_counter = st.selectbox("Counter-Party Document", range(len(ordered_files)), format_func=lambda x: f"CNTR // {roles[x]}", index=min(1, len(ordered_files)-1))

    # Phase 3: Left Native Sidebar Reference Vault
    with st.sidebar:
        st.markdown('<div class="brand-title" style="font-size:16px; margin-top:1rem;">Vault</div>', unsafe_allow_html=True)
        if template_file:
            st.markdown(f'<p style="color:#525252; font-size:11px;">Model Matrix: {template_file}</p>', unsafe_allow_html=True)
            modes = ["Baseline vs Counter", "Standard vs Baseline", "Standard vs Counter"]
            st.session_state.comparison_mode = st.radio("Target Strategy Pivot", modes, index=modes.index(st.session_state.comparison_mode))
            
            st.markdown("<hr style='border-color:#1a1a1a;'/>", unsafe_allow_html=True)
            for p in st.session_state.uploaded_files_data[template_file]:
                st.markdown(f'<div style="font-size:12px; color:#737373; margin-bottom:1rem; line-height:1.5;">{p}</div>', unsafe_allow_html=True)
        else:
            st.markdown('<p style="color:#525252; font-size:12px;">Standard Reference Matrix absent in Staging Tree.</p>', unsafe_allow_html=True)

    st.markdown("<br/>", unsafe_allow_html=True)

    # Route execution models based on sidebar active configuration
    if st.session_state.comparison_mode == "Standard vs Baseline":
        left_paras = st.session_state.uploaded_files_data[template_file] if template_file else st.session_state.uploaded_files_data[base_file]
        right_paras = st.session_state.uploaded_files_data[base_file]
        col1_title, col2_title = "STANDARD MATRIX", roles[st.session_state.current_baseline]
    elif st.session_state.comparison_mode == "Standard vs Counter":
        left_paras = st.session_state.uploaded_files_data[template_file] if template_file else st.session_state.uploaded_files_data[base_file]
        right_paras = st.session_state.uploaded_files_data[counter_file]
        col1_title, col2_title = "STANDARD MATRIX", roles[st.session_state.current_counter]
    else:
        left_paras = st.session_state.uploaded_files_data[base_file]
        right_paras = st.session_state.uploaded_files_data[counter_file]
        col1_title, col2_title = roles[st.session_state.current_baseline], roles[st.session_state.current_counter]

    matcher = difflib.SequenceMatcher(None, left_paras, right_paras)

    # Structural Document Header Track
    h_col1, h_col2, h_col3 = st.columns([4, 4, 3])
    with h_col1: st.markdown(f"<p style='font-size:11px; text-transform:uppercase; color:#525252; font-weight:600;'>{col1_title}</p>", unsafe_allow_html=True)
    with h_col2: st.markdown(f"<p style='font-size:11px; text-transform:uppercase; color:#525252; font-weight:600;'>{col2_title}</p>", unsafe_allow_html=True)
    with h_col3: st.markdown("<p style='font-size:11px; text-transform:uppercase; color:#525252; font-weight:600;'>SMART DELTA EVALUATION</p>", unsafe_allow_html=True)

    # Seamless Aligned Grid Mapping Stream
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            for i in range(i1, i2):
                # Row-isolated wrapper with 0 margin gaps to retain document parchment look
                st.markdown('<div class="aligned-matrix-row">', unsafe_allow_html=True)
                col1, col2, col3 = st.columns([4, 4, 3])
                with col1: st.markdown(f'<div class="stream-paragraph">{left_paras[i]}</div>', unsafe_allow_html=True)
                with col2: st.markdown(f'<div class="stream-paragraph">{left_paras[i]}</div>', unsafe_allow_html=True)
                with col3: st.write("")
                st.markdown('</div>', unsafe_allow_html=True)
                    
        elif tag == 'replace':
            for i, j in zip(range(i1, i2), range(j1, j2)):
                unique_id = f"mod_{i}_{j}"
                signature = extract_clause_signature(left_paras[i])
                trace_detail = generate_advisory_trace_text(left_paras[i], right_paras[j], signature)
                h1, h2 = compute_token_diff_html(left_paras[i], right_paras[j])
                
                st.markdown('<div class="aligned-matrix-row">', unsafe_allow_html=True)
                col1, col2, col3 = st.columns([4, 4, 3])
                with col1: st.markdown(f'<div class="stream-paragraph"><span class="trace-flag">▲ Modified</span>{h1}</div>', unsafe_allow_html=True)
                with col2: st.markdown(f'<div class="stream-paragraph"><span class="trace-flag">▲ Modified</span>{h2}</div>', unsafe_allow_html=True)
                with col3:
                    st.markdown(f'<div class="advisory-panel"><div class="advisory-header">{signature}</div><p style="color:#a3a3a3; font-size:11px; margin:0;">{trace_detail}</p></div>', unsafe_allow_html=True)
                    # Non-selected radio array mapping crisp custom tooltips natively via the help property
                    st.radio("Action Protocol", ["✓", "⚠", "⇄", "✕"], key=f"act_{unique_id}", horizontal=True, label_visibility="collapsed", index=None, help="✓ Accept | ⚠ Escalate | ⇄ Counter | ✕ Reject")
                    st.text_input("Comments", key=f"note_{unique_id}", placeholder="Comments", label_visibility="collapsed")
                st.markdown('</div>', unsafe_allow_html=True)
                    
        elif tag == 'delete':
            for i in range(i1, i2):
                unique_id = f"del_{i}"
                signature = extract_clause_signature(left_paras[i])
                
                st.markdown('<div class="aligned-matrix-row">', unsafe_allow_html=True)
                col1, col2, col3 = st.columns([4, 4, 3])
                with col1: st.markdown(f'<div class="stream-paragraph"><span class="trace-flag" style="color:#f87171;">◼ Omitted</span><span class="del-token">{left_paras[i]}</span></div>', unsafe_allow_html=True)
                with col2: st.markdown('<div class="stream-paragraph" style="color:#404040; font-style:italic;">Provision completely absent from this iteration track.</div>', unsafe_allow_html=True)
                with col3:
                    st.markdown(f'<div class="advisory-panel" style="border-left-color:#f87171;"><div class="advisory-header">{signature}</div><p style="color:#f87171; font-size:11px; margin:0;">Omission Directive</p></div>', unsafe_allow_html=True)
                    st.radio("Action Protocol", ["✓", "⚠", "⇄", "✕"], key=f"act_{unique_id}", horizontal=True, label_visibility="collapsed", index=None, help="✓ Accept | ⚠ Escalate | ⇄ Counter | ✕ Reject")
                    st.text_input("Comments", key=f"note_{unique_id}", placeholder="Comments", label_visibility="collapsed")
                st.markdown('</div>', unsafe_allow_html=True)
                    
        elif tag == 'insert':
            for j in range(j1, j2):
                unique_id = f"ins_{j}"
                signature = extract_clause_signature(right_paras[j])
                
                st.markdown('<div class="aligned-matrix-row">', unsafe_allow_html=True)
                col1, col2, col3 = st.columns([4, 4, 3])
                with col1: st.markdown('<div class="stream-paragraph" style="color:#404040; font-style:italic;">Provision absent in baseline architectural blueprint.</div>', unsafe_allow_html=True)
                with col2: st.markdown(f'<div class="stream-paragraph"><span class="trace-flag" style="color:#34d399;">◆ Insertion</span><span class="add-token">{right_paras[j]}</span></div>', unsafe_allow_html=True)
                with col3:
                    st.markdown(f'<div class="advisory-panel" style="border-left-color:#34d399;"><div class="advisory-header">{signature}</div><p style="color:#34d399; font-size:11px; margin:0;">Injected Clause Block</p></div>', unsafe_allow_html=True)
                    st.radio("Action Protocol", ["✓", "⚠", "⇄", "✕"], key=f"act_{unique_id}", horizontal=True, label_visibility="collapsed", index=None, help="✓ Accept | ⚠ Escalate | ⇄ Counter | ✕ Reject")
                    st.text_input("Comments", key=f"note_{unique_id}", placeholder="Comments", label_visibility="collapsed")
                st.markdown('</div>', unsafe_allow_html=True)

    # Dual Format Landscape Export Action Grid
    st.markdown("<br/><br/><hr style='border-color:#1a1a1a;'/>", unsafe_allow_html=True)
    b_col1, b_col2, b_col3 = st.columns(3)
    
    with b_col1:
        docx_bytes = export_landscape_docx(left_paras, right_paras, col1_title, col2_title, matcher)
        st.download_button(label="📥 Export Landscape Word Matrix", data=docx_bytes, file_name=f"DELTA_Matrix_{datetime.now().strftime('%Y%m%d')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    with b_col2:
        pdf_bytes = export_landscape_pdf(left_paras, right_paras, col1_title, col2_title, matcher)
        st.download_button(label="📥 Export Landscape PDF Report", data=pdf_bytes, file_name=f"DELTA_Report_{datetime.now().strftime('%Y%m%d')}.pdf", mime="application/pdf")
    with b_col3:
        if st.button("Teardown Workspace Session"):
            st.session_state.processing_complete = False
            st.rerun()

# ==========================================
# BLOCK 8: SYSTEM RUNTIME ENTRY
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
