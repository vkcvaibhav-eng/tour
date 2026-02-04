import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import tempfile
import os
import json
import pandas as pd
from datetime import datetime

# --- CONFIGURE PAGE ---
st.set_page_config(page_title="NAU Tour Diary Generator", layout="wide")

st.title("📝 Automated Tour Diary Generator (Word Output)")
st.markdown("""
**Generates a .docx Tour Diary matching the specific NAU format.**
* **Upload:** Tour Orders, Tickets (Railway/Bus/Flight), and Salary Slip.
* **Process:** Extracts dates, places, and distances automatically using Gemini AI.
* **Output:** MS Word file with the required Headers, B.H. Code, and Signature blocks.
""")

# --- SIDEBAR: API KEY ---
with st.sidebar:
    st.header("🔑 API Configuration")
    GEMINI_API_KEY = st.text_input("Gemini API Key", type="password")
    st.info("Get your key from Google AI Studio.")

# --- HELPER FUNCTIONS ---

def set_cell_border(cell, **kwargs):
    """
    Helper to set cell borders in python-docx (which is tricky by default).
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)
            element = tcPr.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcPr.append(element)
            
            for key in ["val", "sz", "space", "color"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))

def extract_doc_data(uploaded_file, api_key):
    """
    Uses Gemini to extract data from Tour Orders, Tickets, or Salary Slips.
    """
    genai.configure(api_key=api_key)
    
    # Upload file to Gemini
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name

    try:
        sample_file = genai.upload_file(path=tmp_path, display_name="NAU_Doc")
        
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        prompt = """
        Analyze this document. It is either a 'Tour Approval', a 'Ticket/Fare Enquiry', or a 'Salary Slip'.
        
        1. If **Salary Slip**: Extract 'type': 'salary', 'name', 'designation', 'basic_pay'.
        
        2. If **Tour Approval** OR **Ticket**: Extract 'type': 'tour'.
           Create a list of trips found. For each trip extract:
           - departure_date (DD/MM/YYYY)
           - departure_time (HH:MM format, 24hr)
           - departure_place (City/Campus name)
           - arrival_date (DD/MM/YYYY)
           - arrival_time (HH:MM format, 24hr)
           - arrival_place (City/Campus name)
           - mode_of_journey (e.g., Govt Vehicle, Private Vehicle, Train, Bus)
           - distance_km (Numeric only. Look for 'KM', 'Distance', or infer from ticket details. If not found, put 0).
           - purpose (The specific purpose of the journey mentioned).
        
        Return ONLY valid JSON. Structure:
        {
          "type": "salary" or "tour",
          ... fields ...
        }
        """
        
        response = model.generate_content([sample_file, prompt])
        # Clean response
        text = response.text.strip()
        if text.startswith('```json'):
            text = text.replace('```json', '').replace('```', '')
        return json.loads(text)
        
    except Exception as e:
        st.error(f"Error processing {uploaded_file.name}: {str(e)}")
        return None
    finally:
        os.remove(tmp_path)

def generate_word_doc(tour_data, user_details):
    doc = Document()
    
    # --- STYLES ---
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    # --- HEADER ---
    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("TOUR DIARY")
    run_title.bold = True
    run_title.font.size = Pt(14)
    run_title.font.underline = True

    # User Info Block
    # Determine Month Range
    dates = [t['departure_date'] for t in tour_data if t.get('departure_date')]
    month_str = "Month: [Date Range]"
    if dates:
        try:
            # Simple logic to find start and end month
            date_objs = [datetime.strptime(d, "%d/%m/%Y") for d in dates]
            min_date = min(date_objs)
            max_date = max(date_objs)
            month_str = f"Month: {min_date.strftime('%B-%Y')} to {max_date.strftime('%B-%Y')}"
        except:
            pass

    # Header Text Block
    header_text = (
        f"Designation: {user_details.get('designation', 'Associate Professor')}\n"
        f"Name: {user_details.get('name', 'Vaibhav Kumar Kanubhai Chaudhari')}\n"
        f"Basic salary: {user_details.get('basic_pay', 'N/A')}\n"
        f"B.H: 303/2092\n"
        f"Dept. of Entomology, N. M. Collage of Agriculture, NAU, Navsari - 396 450\n"
        f"{month_str}"
    )
    
    p_header = doc.add_paragraph(header_text)
    p_header_fmt = p_header.paragraph_format
    p_header_fmt.space_after = Pt(12)

    # --- TABLE ---
    # 7 Columns: Dep Place, Dep Date, Dep Time, Arr Place, Arr Date, Arr Time, Mode, KM, Purpose
    # Actually, let's match the visual reference:
    # Dep (Date, Time, Place) | Arr (Date, Time, Place) | Mode | KM | Purpose
    
    table = doc.add_table(rows=1, cols=9)
    table.style = 'Table Grid'
    table.autofit = False 
    
    # Set Column Headers
    hdr_cells = table.rows[0].cells
    headers = ["Dep. Place", "Date", "Time", "Arr. Place", "Date", "Time", "Mode", "KM", "Purpose"]
    
    for i, text in enumerate(headers):
        hdr_cells[i].text = text
        run = hdr_cells[i].paragraphs[0].runs[0]
        run.bold = True
        run.font.size = Pt(10)

    # Fill Data
    for trip in tour_data:
        row_cells = table.add_row().cells
        
        # Mapping data to columns
        row_cells[0].text = str(trip.get('departure_place', ''))
        row_cells[1].text = str(trip.get('departure_date', ''))
        row_cells[2].text = str(trip.get('departure_time', ''))
        
        row_cells[3].text = str(trip.get('arrival_place', ''))
        row_cells[4].text = str(trip.get('arrival_date', ''))
        row_cells[5].text = str(trip.get('arrival_time', ''))
        
        row_cells[6].text = str(trip.get('mode_of_journey', ''))
        row_cells[7].text = str(trip.get('distance_km', ''))
        row_cells[8].text = str(trip.get('purpose', ''))
        
        # Set font size for row
        for cell in row_cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(9)

    doc.add_paragraph().paragraph_format.space_after = Pt(24)

    # --- SIGNATURE BLOCK ---
    # We use a table with invisible borders to arrange the signatures
    
    # Row 1: User Signature (Right Aligned usually, or as per letter pdf)
    # The letter pdf shows User sign, then Recommended, then Approved.
    
    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.autofit = True
    
    # Left cell empty, Right cell has User Sign
    cell_user = sig_table.rows[0].cells[1]
    p_user = cell_user.add_paragraph()
    p_user.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_u = p_user.add_run("(V. K. Chaudhari)\nSenior Acarologist\nDepartment of Entomology\nN.M. College of Agriculture\nNAU, Navsari")
    run_u.bold = True
    
    doc.add_paragraph().paragraph_format.space_after = Pt(36)

    # Row 2: Recommended (Left) and Approved (Right)
    approval_table = doc.add_table(rows=1, cols=2)
    approval_table.autofit = True
    
    # Recommended
    cell_rec = approval_table.rows[0].cells[0]
    p_rec = cell_rec.add_paragraph()
    p_rec.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_r = p_rec.add_run("Recommended\n\n\nProfessor and Head\nDept. of Entomology\nN. M. College of Agriculture\nNAU, Navsari")
    run_r.bold = True
    
    # Approved
    cell_app = approval_table.rows[0].cells[1]
    p_app = cell_app.add_paragraph()
    p_app.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_a = p_app.add_run("Approved\n\n\nPrincipal and Dean\nN. M. College of Agriculture\nNAU, Navsari")
    run_a.bold = True

    return doc

# --- MAIN APP LOGIC ---

uploaded_files = st.file_uploader("Upload Documents (PDF)", 
                                  type=['pdf'], 
                                  accept_multiple_files=True,
                                  help="Upload Tour Approvals, Tickets/Fare Enquiries, and Salary Slip.")

if uploaded_files and st.button("Generate Word Diary"):
    if not GEMINI_API_KEY:
        st.error("Please enter your Gemini API key.")
    else:
        with st.spinner("Analyzing documents..."):
            
            all_trips = []
            user_info = {}
            
            for file in uploaded_files:
                data = extract_doc_data(file, GEMINI_API_KEY)
                
                if data:
                    if data.get('type') == 'salary':
                        # Update user info if salary slip found
                        user_info.update(data)
                    elif data.get('type') == 'tour':
                        # It might be a list of trips or a single object
                        # The prompt asks for a list, but let's handle both
                        if 'trips' in data and isinstance(data['trips'], list):
                             all_trips.extend(data['trips'])
                        elif 'departure_place' in data:
                             all_trips.append(data)
                        
                        # Sometimes Gemini returns the list directly
                        if isinstance(data, list):
                            all_trips.extend(data)

            if all_trips:
                # Sort trips by date (simple string sort, preferably convert to date obj)
                try:
                    all_trips.sort(key=lambda x: datetime.strptime(x['departure_date'], "%d/%m/%Y") if x.get('departure_date') else datetime.min)
                except:
                    pass # Keep order if date parsing fails

                # Preview
                st.success(f"Found {len(all_trips)} trip entries.")
                df = pd.DataFrame(all_trips)
                if not df.empty:
                    st.dataframe(df[['departure_date', 'departure_place', 'arrival_place', 'distance_km', 'purpose']])

                # Generate DOCX
                doc = generate_word_doc(all_trips, user_info)
                
                # Save to buffer
                bio = io.BytesIO() if 'io' in locals() else None
                import io
                bio = io.BytesIO()
                doc.save(bio)
                
                st.download_button(
                    label="Download Tour Diary (.docx)",
                    data=bio.getvalue(),
                    file_name="NAU_Tour_Diary.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            else:
                st.warning("No tour data extracted. Please check the uploaded files.")
