# ==============================================================================
# DELTA | PREMIUM LEGAL ARCHITECTURE SUITE
# ==============================================================================
# VERSION: 2.0 (Integrated Synchronized Dual-Viewport UI)
# DESCRIPTION: High-fidelity contract comparison framework utilizing fuzzy token
#              alignment, SHA-256 cryptographic hashing, and a luxury dark-mode
#              UI to review structural and semantic modifications in legal texts.
# ==============================================================================

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

# ==============================================================================
# BLOCK 1: STATE HYDRATION & LIFECYCLE MANAGEMENT
# ==============================================================================
# This block ensures that session states persist across Streamlit re-renders.
# We store raw parsed paragraphs, hashes, file hierarchies, and tracking counters.
# ==============================================================================
if 'initialized' not in st.session_state:
    st.session_state.initialized = True
    
    # Stores arrays of string paragraphs for each uploaded file
    st.session_state.uploaded_files_data = {}   
    
    # Stores SHA-256 cryptographic hashes for integrity validation
    st.session_state.uploaded_files_hashes = {} 
    
    # Tracks the sequential order of files as they are uploaded
    st.session_state.file_order = []           
    
    # Maps specific workflow roles (e.g., "v1: Baseline") to specific files
    st.session_state.file_roles = {}           
    
    # Flags whether the initial pipeline ingestion is finished
    st.session_state.processing_complete = False 
    
    # Index reference for the currently selected Baseline document
    st.session_state.current_baseline = 0      
    
    # Index reference for the currently selected Counterparty document
    st.session_state.current_counter = 1       
    
    # Defines the active routing matrix mode for document comparison
    st.session_state.comparison_mode = "Baseline vs Counter"


# ==============================================================================
# BLOCK 2: SYSTEM DESIGN SYSTEM (CSS)
# ==============================================================================
# Injects the custom corporate luxury CSS to override Streamlit's default theme.
# Utilizes 'Inter' for clean body text and 'Cinzel' for premium branding.
# ==============================================================================
def inject_luxury_system_css():
    """
    Injects custom HTML/CSS into the Streamlit DOM to enforce a dark grayscale
    foundation, precise typography, and minimalistic component styling.
    """
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
                padding-top: 5vh;
            }
            
            /* Skinned Premium Drag & Drop Uploader */
            .stFileUploader {
                border: 1px dashed #d4af37 !important;
                background-color: #111111 !important;
                border-radius: 0px !important;
                padding: 1.5rem !important;
                width: 100% !important;
                transition: border 0.3s ease;
            }
            
            .stFileUploader:hover {
                border: 1px dashed #ffffff !important;
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
                margin-top: 1rem;
            }
            
            .stButton > button:hover {
                background-color: #d4af37 !important;
                color: #0c0c0c !important;
            }
            
            /* Row alignment padding without clunky visual boxes */
            .aligned-matrix-row {
                padding: 0.8rem 0;
                margin: 0;
                background: transparent;
                display: block;
            }
            
            /* Text Flow Formatting */
            .stream-paragraph {
                color: #e5e5e5;
                font-size: 13px;
                line-height: 1.8;
                word-wrap: break-word;
            }
            
            /* Clean Text Tokens with Light Pastel Overlays */
            .add-token { 
                background-color: #1c3d25 !important; 
                color: #6ee7b7 !important; 
                padding: 2px 4px; 
                border-radius: 2px;
            }
            
            .del-token { 
                background-color: #4c1d1d !important; 
                color: #fca5a5 !important; 
                text-decoration: line-through; 
                padding: 2px 4px; 
                border-radius: 2px;
            }
            
            .trace-flag { 
                font-family: 'Cinzel', serif; 
                color: #d4af37; 
                font-size: 10px; 
                letter-spacing: 1px; 
                text-transform: uppercase; 
                margin-bottom: 0.6rem; 
                display: block; 
                font-weight: 600;
            }
            
            /* Advisory Cards for Review Insights */
            .advisory-panel {
                background-color: #161616;
                border-left: 2px solid #d4af37;
                padding: 0.8rem 1rem;
                margin-bottom: 0.5rem;
                border-radius: 0 4px 4px 0;
            }
            
            .advisory-header { 
                font-family: 'Cinzel', serif; 
                color: #ffffff; 
                font-size: 11px; 
                letter-spacing: 1px; 
                text-transform: uppercase; 
                margin-bottom: 0.4rem; 
            }
            
            /* Crypto Manifest Status Bar */
            .crypto-banner {
                font-family: monospace;
                font-size: 10px;
                background-color: #111111;
                border: 1px solid #1a1a1a;
                padding: 8px 14px;
                color: #737373;
                margin-bottom: 1.5rem;
                border-radius: 2px;
                line-height: 1.6;
            }
            
            /* Form Style Overrides */
            div[data-baseweb="select"] { 
                background-color: #121212 !important; 
                border-radius: 0px !important; 
            }
            
            input {
                border-radius: 0px !important;
                background-color: #161616 !important;
                color: #ffffff !important;
                border: 1px solid #262626 !important;
                font-size: 12px !important;
                padding: 8px 12px !important;
                width: 100%;
            }
            
            /* Hide empty label spacing for inputs */
            label {
                display: none !important;
            }
        </style>
    """, unsafe_allow_html=True)


# ==============================================================================
# BLOCK 3: ENGINE PARSING & CRYPTO HASHING
# ==============================================================================
# Handlers for extracting raw text from PDFs and Word Documents, alongside
# real-time SHA-256 generation for document integrity verification.
# ==============================================================================

def parse_pdf(file_bytes):
    """
    Parses a PDF document byte stream using PyMuPDF (fitz).
    Iterates through pages and blocks to extract raw text, normalizing whitespace.
    
    Args:
        file_bytes (bytes): The raw byte stream of the uploaded PDF file.
        
    Returns:
        list: A list of string paragraphs extracted from the document.
    """
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    paragraphs = []
    
    for page_num, page in enumerate(doc):
        # Extract dictionary blocks (text, images, etc.)
        blocks = page.get_text("blocks")
        for b in blocks:
            # Index 4 in the block tuple contains the actual string data
            text = b[4].strip()
            if text:
                # Normalize arbitrary line breaks into single spaces for clean diffing
                paragraphs.append(" ".join(text.split()))
                
    return paragraphs

def parse_docx(file_bytes):
    """
    Parses a Microsoft Word (.docx) document using python-docx.
    Iterates through paragraph objects to extract text.
    
    Args:
        file_bytes (bytes): The raw byte stream of the uploaded DOCX file.
        
    Returns:
        list: A list of string paragraphs extracted from the document.
    """
    doc = Document(io.BytesIO(file_bytes))
    paragraphs = []
    
    for p in doc.paragraphs:
        if p.text.strip():
            # Normalize whitespace
            paragraphs.append(" ".join(p.text.split()))
            
    return paragraphs

def load_staged_matrices(uploaded_files):
    """
    Processes all staged files from the uploader. Computes crypto hashes,
    parses text arrays, and appends them to session state memory.
    
    Args:
        uploaded_files (list): List of Streamlit UploadedFile objects.
    """
    for file in uploaded_files:
        if file.name not in st.session_state.uploaded_files_data:
            bytes_data = file.read()
            
            # [ADVANCED FEATURE 3: Generate SHA-256 Cryptographic Hash]
            # Used to establish strict version control and prevent tampering
            file_hash = hashlib.sha256(bytes_data).hexdigest()
            st.session_state.uploaded_files_hashes[file.name] = file_hash
            
            # Route processing based on file extension
            if file.name.lower().endswith('.pdf'):
                parsed_text = parse_pdf(bytes_data)
            elif file.name.lower().endswith('.docx'):
                parsed_text = parse_docx(bytes_data)
            else:
                continue
                
            # Store data and initialize default hierarchy roles
            st.session_state.uploaded_files_data[file.name] = parsed_text
            if file.name not in st.session_state.file_order:
                st.session_state.file_order.append(file.name)
                # Default role assignment upon ingestion
                st.session_state.file_roles[file.name] = "v1: Baseline"


# ==============================================================================
# BLOCK 4: FEATURE 2 - FUZZY ALIGNMENT ENGINE
# ==============================================================================
def compute_fuzzy_alignment_matrix(left_paras, right_paras, threshold=0.45):
    """
    Position-agnostic token-ratio optimization algorithm.
    Compares two lists of paragraphs and aligns them based on token similarity,
    allowing for structural rearrangement without breaking the comparison mapping.
    
    Args:
        left_paras (list): Array of baseline paragraphs.
        right_paras (list): Array of counterparty paragraphs.
        threshold (float): Minimum similarity ratio to consider a match.
        
    Returns:
        list of tuples: Opcodes detailing the layout instructions.
                        Format: (tag, i1, i2, j1, j2)
    """
    matched_right_indices = set()
    alignment_opcodes = []
    
    for i, lp in enumerate(left_paras):
        best_ratio = 0.0
        best_j = None
        
        # Token sort optimization neutralizes minor structural rearrangements
        lp_tokens = sorted(lp.lower().split())
        
        for j, rp in enumerate(right_paras):
            # Skip if target paragraph has already been aligned
            if j in matched_right_indices:
                continue
                
            rp_tokens = sorted(rp.lower().split())
            ratio = difflib.SequenceMatcher(None, lp_tokens, rp_tokens).ratio()
            
            # Keep track of the highest scoring match
            if ratio > best_ratio:
                best_ratio = ratio
                best_j = j
                
        # If the best match exceeds our confidence threshold, lock it in
        if best_ratio >= threshold and best_j is not None:
            matched_right_indices.add(best_j)
            
            # Check for exact structural parity
            if best_ratio > 0.98 and lp == right_paras[best_j]:
                alignment_opcodes.append(('equal', i, i, best_j, best_j))
            else:
                alignment_opcodes.append(('replace', i, i, best_j, best_j))
        else:
            # If no suitable match is found, the baseline clause was omitted
            alignment_opcodes.append(('delete', i, i, None, None))
            
    # Sweep remaining unaligned paragraphs from the counterparty document
    # These represent newly injected clauses (insertions)
    for j in range(len(right_paras)):
        if j not in matched_right_indices:
            alignment_opcodes.append(('insert', None, None, j, j))
            
    # Normalize sorting to preserve the structural execution flow of the baseline file
    alignment_opcodes.sort(key=lambda x: (x[1] if x[1] is not None else float('inf'), x[3] if x[3] is not None else 0))
    return alignment_opcodes


# ==============================================================================
# BLOCK 5: STRING TRANSFORMATION ANALYTICS & DELTA HELPERS
# ==============================================================================

def extract_clause_signature(text):
    """
    Extracts the semantic "name" or header of a clause for advisory labeling.
    Looks for capital letter patterns or specific legal keywords.
    """
    match = re.match(r'^([A-Za-z0-9\.\s]+(?:\b[A-Z]{2,}\b|\bRent\b|\bTerm\b|\bDeposit\b|\bUse\b))', text)
    if match:
        return match.group(1).strip()
    words = text.split()
    return " ".join(words[:2]) if len(words) >= 2 else "Provision Block"

def generate_advisory_trace_text(text1, text2, signature):
    """
    Generates human-readable summary logs of the specific modifications.
    Used within the evaluation action panel.
    """
    matcher = difflib.SequenceMatcher(None, text1.split(), text2.split())
    traces = []
    
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'replace':
            o = " ".join(text1.split()[i1:i2])
            n = " ".join(text2.split()[j1:j2])
            if len(o) < 30 and len(n) < 30: 
                traces.append(f"Changed '{o}' to '{n}'")
        elif tag == 'delete':
            d = " ".join(text1.split()[i1:i2])
            if len(d) < 30: 
                traces.append(f"Omitted '{d}'")
        elif tag == 'insert':
            ins = " ".join(text2.split()[j1:j2])
            if len(ins) < 30: 
                traces.append(f"Injected '{ins}'")
                
    return f"{signature} | " + ("; ".join(traces) if traces else "Terms modified.")

def compute_token_diff_html(text1, text2):
    """
    Executes deep token diffing and wraps differences in styled HTML spans.
    Returns dual string outputs corresponding to the left and right viewports.
    """
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

def count_block_deltas(text1, text2):
    """
    NEW HELPER: Executes a micro-diff on paragraph tokens to count exact modifications.
    Used by the luxury UI injection to render the [ X Deltas ] badge.
    
    Args:
        text1 (str): Baseline paragraph string.
        text2 (str): Counterparty paragraph string.
        
    Returns:
        int: The integer count of discrete token blocks changed, added, or removed.
    """
    # Regex fallback to separate trailing punctuation could be implemented here
    # For now, we split by whitespace as per the original design constraints
    matcher = difflib.SequenceMatcher(None, text1.split(), text2.split())
    
    # Generator sum of all non-equal token blocks
    return sum(1 for tag, _, _, _, _ in matcher.get_opcodes() if tag != 'equal')


# ==============================================================================
# BLOCK 6: HIGH-FIDELITY COMPLIANCE EXPORTERS (DOCX & PDF)
# ==============================================================================

def set_run_background(run, color_hex):
    """
    Helper function to inject background shading XML into a docx Run object.
    Used to highlight modifications in the offline export.
    """
    rPr = run._r.get_or_add_rPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    rPr.append(shd)

def export_landscape_docx(left_paras, right_paras, title_left, title_right, alignment_opcodes, hash_l, hash_r):
    """
    Compiles the synchronized review matrix into a landscape Microsoft Word doc.
    Includes crypto manifests and captures user input logic.
    """
    doc = Document()
    
    # Force Landscape Orientation for optimal matrix viewing
    section = doc.sections[-1]
    section.orientation = WD_ORIENT.LANDSCAPE
    new_width, new_height = section.page_height, section.page_width
    section.page_width = new_width
    section.page_height = new_height
    
    # Header Branding
    title = doc.add_paragraph()
    title.add_run("DELTA CONTRACT ADVISORY MATRIX").bold = True
    
    # Embed Crypto hashes into document layout header
    meta = doc.add_paragraph()
    meta.add_run(f"INTEGRITY MANIFEST LOG\nBASE SHA-256: {hash_l}\nCNTR SHA-256: {hash_r}\n").font.size = Pt(8)
    
    # Construct Evaluation Table
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
                    r1 = p1.add_run(w1_str); set_run_background(r1, "fee2e2") # Pastel Red
                    r2 = p2.add_run(w2_str); set_run_background(r2, "dcfce7") # Pastel Green
            
            # Fetch user advisory states applied during session
            act_val = st.session_state.get(f"act_mod_{i1}_{j1}", "Unassigned")
            note_val = st.session_state.get(f"note_mod_{i1}_{j1}", "")
            sig = extract_clause_signature(left_paras[i1])
            trace = generate_advisory_trace_text(left_paras[i1], right_paras[j1], sig)
            
            row.cells[2].text = f"Evaluation: {trace}\nAction: {act_val}\nComments: {note_val}"
            
        elif tag == 'delete':
            r1 = row.cells[0].paragraphs[0].add_run(left_paras[i1])
            set_run_background(r1, "fee2e2")
            row.cells[1].text = "[Clause Omitted]"
            
            act_val = st.session_state.get(f"act_del_{i1}", "Unassigned")
            note_val = st.session_state.get(f"note_del_{i1}", "")
            row.cells[2].text = f"Evaluation: Structural Omission\nAction: {act_val}\nComments: {note_val}"
            
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
    """
    Compiles the synchronized review matrix into a clean PDF stream.
    Utilizes PyMuPDF low-level canvas drawing primitives.
    """
    doc = fitz.open()
    page_w, page_h = 792, 612
    page = doc.new_page(width=page_w, height=page_h)
    
    # Title Branding
    page.insert_text(fitz.Point(36, 30), "DELTA CONTRACT REVIEW MATRIX", fontsize=12, color=(0.83, 0.68, 0.21))
    page.insert_text(fitz.Point(36, 45), f"BASE SHA-256: {hash_l} | CNTR SHA-256: {hash_r}", fontsize=7, color=(0.4, 0.4, 0.4))
    
    y = 75
    col_w = 230
    c1_x, c2_x, c3_x = 36, 280, 524
    
    for tag, i1, _, j1, _ in alignment_opcodes:
        # Paging execution
        if y > (page_h - 90):
            page = doc.new_page(width=page_w, height=page_h)
            y = 50
            
        if tag == 'equal':
            page.insert_textbox(fitz.Rect(c1_x, y, c1_x + col_w, y + 65), left_paras[i1], fontsize=8, color=(0.1, 0.1, 0.1))
            page.insert_textbox(fitz.Rect(c2_x, y, c2_x + col_w, y + 65), right_paras[j1], fontsize=8, color=(0.1, 0.1, 0.1))
            page.insert_textbox(fitz.Rect(c3_x, y, c3_x + col_w, y + 65), "No variance detected.", fontsize=8, color=(0.5, 0.5, 0.5))
            y += 70
            
        elif tag == 'replace':
            # Background fills for evaluation states
            page.draw_rect(fitz.Rect(c1_x - 4, y, c1_x + col_w + 4, y + 70), color=None, fill=(0.99, 0.94, 0.94))
            page.draw_rect(fitz.Rect(c2_x - 4, y, c2_x + col_w + 4, y + 70), color=None, fill=(0.92, 0.99, 0.95))
            
            page.insert_textbox(fitz.Rect(c1_x, y + 5, c1_x + col_w, y + 65), left_paras[i1], fontsize=8, color=(0.4, 0.1, 0.1))
            page.insert_textbox(fitz.Rect(c2_x, y + 5, c2_x + col_w, y + 65), right_paras[j1], fontsize=8, color=(0.1, 0.3, 0.1))
            
            act_val = st.session_state.get(f"act_mod_{i1}_{j1}", "Unassigned")
            note_val = st.session_state.get(f"note_mod_{i1}_{j1}", "")
            sig = extract_clause_signature(left_paras[i1])
            trace = generate_advisory_trace_text(left_paras[i1], right_paras[j1], sig)
            
            page.insert_textbox(fitz.Rect(c3_x, y + 5, c3_x + col_w, y + 65), f"Eval: {trace}\nAction: {act_val}\nComments: {note_val}", fontsize=8, color=(0.1, 0.1, 0.1))
            y += 75
            
        elif tag == 'delete':
            page.draw_rect(fitz.Rect(c1_x - 4, y, c1_x + col_w + 4, y + 55), color=None, fill=(0.99, 0.94, 0.94))
            page.insert_textbox(fitz.Rect(c1_x, y + 5, c1_x + col_w, y + 50), left_paras[i1], fontsize=8, color=(0.4, 0.1, 0.1))
            page.insert_text(fitz.Point(c2_x, y + 15), "[Clause Omitted]", fontsize=8, color=(0.5, 0.5, 0.5))
            
            act_val = st.session_state.get(f"act_del_{i1}", "Unassigned")
            note_val = st.session_state.get(f"note_del_{i1}", "")
            page.insert_textbox(fitz.Rect(c3_x, y + 5, c3_x + col_w, y + 50), f"Eval: Structural Omission\nAction: {act_val}\nComments: {note_val}", fontsize=8, color=(0.1, 0.1, 0.1))
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


# ==============================================================================
# BLOCK 7: UI PHASE 1 - CENTERED PORTAL (DOCUMENT INGESTION)
# ==============================================================================

def render_premium_landing_view():
    """
    Renders the initial application state. Features a centered, branded portal
    for drag-and-drop document ingestion and initial role assignment.
    """
    st.markdown('<div class="landing-wrapper">', unsafe_allow_html=True)
    st.markdown('<div class="brand-title">DELTA</div>', unsafe_allow_html=True)
    st.markdown('<div class="brand-subtitle">Legal Infrastructure Pipeline</div>', unsafe_allow_html=True)
    
    # Core ingestion component
    uploaded_files = st.file_uploader("Ingest Core Binaries", type=['pdf', 'docx'], accept_multiple_files=True, label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)
    
    if uploaded_files:
        load_staged_matrices(uploaded_files)
        st.markdown("<div style='max-width:600px; margin: 2rem auto 0 auto;'>", unsafe_allow_html=True)
        
        # Iterates through uploaded files to assign hierarchical tagging
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
        
        # Trigger pipeline state change
        if st.button("Compile Revision Tree"):
            st.session_state.processing_complete = True
            st.rerun()
            
        st.markdown("</div>", unsafe_allow_html=True)


# ==============================================================================
# BLOCK 8: UI PHASE 2 - SYNCHRONIZED REVIEW MATRIX
# ==============================================================================

def render_delta_contracts_view():
    """
    Renders the primary application dashboard where document arrays are evaluated
    side-by-side using the synchronized layout constraints and delta badges.
    """
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
        st.session_state.current_baseline = st.selectbox(
            "Baseline Selection Frame", range(len(ordered_files)), 
            format_func=lambda x: f"BASE // {roles[x]}"
        )
    with t_col2:
        st.session_state.current_counter = st.selectbox(
            "Counterparty Selection Frame", range(len(ordered_files)), 
            format_func=lambda x: f"CNTR // {roles[x]}", 
            index=min(1, len(ordered_files)-1)
        )

    # Left Native Sidebar Track Configuration (The Vault)
    with st.sidebar:
        st.markdown('<div class="brand-title" style="font-size:16px; margin-top:1rem;">Vault</div>', unsafe_allow_html=True)
        
        if template_file:
            st.markdown(f'<p style="color:#525252; font-size:11px;">Matrix Reference: {template_file}</p>', unsafe_allow_html=True)
            modes = ["Baseline vs Counter", "Standard vs Baseline", "Standard vs Counter"]
            st.session_state.comparison_mode = st.radio("Target Pipeline Strategy Routing", modes, index=modes.index(st.session_state.comparison_mode))
            
            st.markdown("<hr style='border-color:#1a1a1a;'/>", unsafe_allow_html=True)
            for p in st.session_state.uploaded_files_data[template_file]:
                st.markdown(f'<div style="font-size:12px; color:#737373; margin-bottom:1rem; line-height:1.5;">{p}</div>', unsafe_allow_html=True)
        else:
            st.markdown('<p style="color:#525252; font-size:12px;">Standard Matrix signature absent in workspace staging array.</p>', unsafe_allow_html=True)

    st.markdown("<br/>", unsafe_allow_html=True)

    # Route execution models based on active state parameters
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


    # Execute Core Matrix Mapping Algorithm
    alignment_opcodes = compute_fuzzy_alignment_matrix(left_paras, right_paras)

    # Render Cryptographic Validation Header
    st.markdown(f"""
        <div class="crypto-banner">
            VALIDATION MANIFEST LOG // LOCKED STATUS<br/>
            [BASE_SIG]: {hash_l}<br/>[CNTR_SIG]: {hash_r}
        </div>
    """, unsafe_allow_html=True)

    # Structure Grid Headings
    h_col1, h_col2, h_col3 = st.columns([4, 4, 3])
    with h_col1: st.markdown(f"<p style='font-size:11px; text-transform:uppercase; color:#525252; font-weight:600;'>{col1_title}</p>", unsafe_allow_html=True)
    with h_col2: st.markdown(f"<p style='font-size:11px; text-transform:uppercase; color:#525252; font-weight:600;'>{col2_title}</p>", unsafe_allow_html=True)
    with h_col3: st.markdown("<p style='font-size:11px; text-transform:uppercase; color:#525252; font-weight:600;'>SMART DELTA EVALUATION</p>", unsafe_allow_html=True)

    # Base CSS Templates for the Luxury Containers (Inline for isolated mapping without polluting global space)
    BASE_BOX = "position: relative; padding: 1.2rem; border-radius: 4px; border: 1px solid;"
    BADGE_CSS = "position: absolute; top: -10px; right: 15px; background: #111111; color: #d4af37; font-family: 'Cinzel', serif; font-size: 10px; padding: 2px 8px; border: 1px solid #d4af37; letter-spacing: 1px; text-transform: uppercase; z-index: 10;"

    # Row Aligned Processing Stream (Synchronized Dual-Viewport Injection)
    for tag, i1, _, j1, _ in alignment_opcodes:
        
        # ==========================================
        # STATE 1: NO VARIANCE DETECTED
        # ==========================================
        if tag == 'equal':
            st.markdown('<div class="aligned-matrix-row">', unsafe_allow_html=True)
            col1, col2, col3 = st.columns([4, 4, 3])
            
            # Standard neutral box for baseline
            with col1: 
                st.markdown(f'<div style="{BASE_BOX} border-color: #262626; background-color: transparent;"><div class="stream-paragraph">{left_paras[i1]}</div></div>', unsafe_allow_html=True)
            
            # Standard neutral box for counter
            with col2: 
                st.markdown(f'<div style="{BASE_BOX} border-color: #262626; background-color: transparent;"><div class="stream-paragraph">{right_paras[j1]}</div></div>', unsafe_allow_html=True)
            
            # Empty evaluation column
            with col3: 
                st.write("")
                
            st.markdown('</div>', unsafe_allow_html=True)
                
        # ==========================================
        # STATE 2: MODIFIED BLOCK DETECTED
        # ==========================================
        elif tag == 'replace':
            unique_id = f"mod_{i1}_{j1}"
            signature = extract_clause_signature(left_paras[i1])
            trace_detail = generate_advisory_trace_text(left_paras[i1], right_paras[j1], signature)
            h1, h2 = compute_token_diff_html(left_paras[i1], right_paras[j1])
            
            # Execute micro-diff to calculate the absolute integer for the luxury badge
            delta_count = count_block_deltas(left_paras[i1], right_paras[j1])
            
            st.markdown('<div class="aligned-matrix-row">', unsafe_allow_html=True)
            col1, col2, col3 = st.columns([4, 4, 3])
            
            # Left Box: Soft Red (Pastel) overlay for baseline modifications
            with col1: 
                st.markdown(f'''
                    <div style="{BASE_BOX} border-color: #f87171; background-color: rgba(248, 113, 113, 0.05);">
                        <div style="{BADGE_CSS}">[ {delta_count} Deltas ]</div>
                        <div class="stream-paragraph">{h1}</div>
                    </div>
                ''', unsafe_allow_html=True)
                
            # Right Box: Soft Green (Pastel) overlay for counter modifications
            with col2: 
                st.markdown(f'''
                    <div style="{BASE_BOX} border-color: #34d399; background-color: rgba(52, 211, 153, 0.05);">
                        <div style="{BADGE_CSS}">[ {delta_count} Deltas ]</div>
                        <div class="stream-paragraph">{h2}</div>
                    </div>
                ''', unsafe_allow_html=True)
                
            # Evaluation Advisory Panel & Action Controls
            with col3:
                st.markdown(f'<div class="advisory-panel"><div class="advisory-header">{signature}</div><p style="color:#a3a3a3; font-size:11px; margin:0;">{trace_detail}</p></div>', unsafe_allow_html=True)
                st.radio("Action Protocol", ["✓", "⚠", "⇄", "✕"], key=f"act_{unique_id}", horizontal=True, label_visibility="collapsed", index=None)
                st.text_input("Comments", key=f"note_{unique_id}", placeholder="Comments", label_visibility="collapsed")
                
            st.markdown('</div>', unsafe_allow_html=True)
                
        # ==========================================
        # STATE 3: DELETED / DROPPED BLOCK
        # ==========================================
        elif tag == 'delete':
            unique_id = f"del_{i1}"
            signature = extract_clause_signature(left_paras[i1])
            
            st.markdown('<div class="aligned-matrix-row">', unsafe_allow_html=True)
            col1, col2, col3 = st.columns([4, 4, 3])
            
            # Left Box: Explicit Red styling for completely removed provisions
            with col1: 
                st.markdown(f'''
                    <div style="{BASE_BOX} border-color: #ef4444; background-color: rgba(239, 68, 68, 0.08);">
                        <div style="{BADGE_CSS} border-color: #ef4444; color: #ef4444;">[ Block Removed ]</div>
                        <div class="stream-paragraph"><span class="del-token">{left_paras[i1]}</span></div>
                    </div>
                ''', unsafe_allow_html=True)
                
            # Right Box: Explicit empty spacer constraint to enforce strict horizontal parallelism
            with col2: 
                st.markdown(f'''
                    <div style="{BASE_BOX} border-style: dashed; border-color: #262626; background-color: transparent; display: flex; align-items: center; justify-content: center; min-height: 100px;">
                        <span style="color:#404040; font-style:italic; font-size: 11px;">Provision dropped from counter.</span>
                    </div>
                ''', unsafe_allow_html=True)
                
            # Action Panel routing for escalated structural removals
            with col3:
                st.markdown(f'<div class="advisory-panel" style="border-left-color:#f87171;"><div class="advisory-header">{signature}</div><p style="color:#f87171; font-size:11px; margin:0;">Omission Directive</p></div>', unsafe_allow_html=True)
                st.radio("Action Protocol", ["✓", "⚠", "⇄", "✕"], key=f"act_{unique_id}", horizontal=True, label_visibility="collapsed", index=None)
                st.text_input("Comments", key=f"note_{unique_id}", placeholder="Comments", label_visibility="collapsed")
                
            st.markdown('</div>', unsafe_allow_html=True)
                
        # ==========================================
        # STATE 4: INSERTED / NEW BLOCK
        # ==========================================
        elif tag == 'insert':
            unique_id = f"ins_{j1}"
            signature = extract_clause_signature(right_paras[j1])
            
            st.markdown('<div class="aligned-matrix-row">', unsafe_allow_html=True)
            col1, col2, col3 = st.columns([4, 4, 3])
            
            # Left Box: Explicit empty spacer constraint to enforce strict horizontal parallelism
            with col1: 
                st.markdown(f'''
                    <div style="{BASE_BOX} border-style: dashed; border-color: #262626; background-color: transparent; display: flex; align-items: center; justify-content: center; min-height: 100px;">
                        <span style="color:#404040; font-style:italic; font-size: 11px;">No baseline equivalent.</span>
                    </div>
                ''', unsafe_allow_html=True)
                
            # Right Box: Explicit Green styling for brand new provisions identified on the counter file
            with col2: 
                st.markdown(f'''
                    <div style="{BASE_BOX} border-color: #10b981; background-color: rgba(16, 185, 129, 0.08);">
                        <div style="{BADGE_CSS} border-color: #10b981; color: #10b981;">[ New Block ]</div>
                        <div class="stream-paragraph"><span class="add-token">{right_paras[j1]}</span></div>
                    </div>
                ''', unsafe_allow_html=True)
                
            # Action panel routing for structurally injected terms
            with col3:
                st.markdown(f'<div class="advisory-panel" style="border-left-color:#34d399;"><div class="advisory-header">{signature}</div><p style="color:#34d399; font-size:11px; margin:0;">Injected Clause Block</p></div>', unsafe_allow_html=True)
                st.radio("Action Protocol", ["✓", "⚠", "⇄", "✕"], key=f"act_{unique_id}", horizontal=True, label_visibility="collapsed", index=None)
                st.text_input("Comments", key=f"note_{unique_id}", placeholder="Comments", label_visibility="collapsed")
                
            st.markdown('</div>', unsafe_allow_html=True)

    # ==========================================
    # BRANDED CONSOLIDATED EXPORT FOOTER
    # ==========================================
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
        # Flush session state trigger returning the app back to pipeline ingestion view
        if st.button("Terminate Session"):
            st.session_state.processing_complete = False
            st.rerun()


# ==============================================================================
# BLOCK 9: ENTRYPOINT ORCHESTRATION
# ==============================================================================

def main():
    """
    Main orchestration loop for the Streamlit execution tree.
    Sets wide-layout configuration and routes users based on data ingestion status.
    """
    st.set_page_config(page_title="DELTA CONTRACTS", layout="wide", initial_sidebar_state="collapsed")
    inject_luxury_system_css()
    
    if not st.session_state.processing_complete:
        render_premium_landing_view()
    else:
        render_delta_contracts_view()

if __name__ == "__main__":
    main()
