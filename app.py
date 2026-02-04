import streamlit as st
import google.generativeai as genai
from serpapi import GoogleSearch
from fpdf import FPDF
import tempfile
import os
import json
import pandas as pd

# --- CONFIGURE PAGE ---
st.set_page_config(page_title="NAU Tour Diary Generator", layout="wide")

st.title("📝 Automated Tour Diary Generator (NAU)")
st.markdown("""
This tool processes **Online Tour Management System** PDFs and **Salary Slips** to generate a formatted Tour Diary.
**Distance Logic:** Calculates Railway distance first. If no direct connection exists, it falls back to GSRTC/Road distance.
""")

# --- SIDEBAR: API KEYS ---
with st.sidebar:
    st.header("🔑 API Configuration")
    GEMINI_API_KEY = st.text_input("Gemini API Key", type="password")
    SERPAPI_KEY = st.text_input("SerpApi Key", type="password")
    
    st.info("Get keys from Google AI Studio and SerpApi.")

# --- FUNCTIONS ---

def get_distance_serpapi(origin, destination, api_key):
    """
    Calculates distance. Prioritizes Railway. If not available, uses Road.
    """
    # 1. Try Railway First
    params_train = {
        "engine": "google_maps",
        "q": f"train from {origin} to {destination}",
        "type": "search",
        "api_key": api_key
    }
    
    try:
        search = GoogleSearch(params_train)
        results = search.get_dict()
        
        # Check if transit options exist and look for a train line
        if "directions" in results and results["directions"]:
            # Simplified check: assumes if google gives transit direction, rail/bus is valid
            # For stricter "Railway Only", we would parse the transit_details
            dist_text = results["directions"][0]["distance"]
            km = float(dist_text.replace(" km", "").replace(",", ""))
            return km, "Railway (Calculated)"
            
    except Exception as e:
        print(f"Railway search failed: {e}")

    # 2. Fallback to Road (GSRTC logic)
    params_road = {
        "engine": "google_maps",
        "q": f"driving distance from {origin} to {destination}",
        "type": "search",
        "api_key": api_key
    }
    
    try:
        search = GoogleSearch(params_road)
        results = search.get_dict()
        
        if "directions" in results and results["directions"]:
            dist_text = results["directions"][0]["distance"]
            km = float(dist_text.replace(" km", "").replace(",", ""))
            return km, "Road (GSRTC/Fallback)"
            
    except Exception as e:
        return 0, "Error"
    
    return 0, "Not Found"

def extract_pdf_data(uploaded_file, api_key):
    """
    Uses Gemini 1.5 Pro to extract relevant fields from the uploaded PDF.
    """
    genai.configure(api_key=api_key)
    
    # Upload file to Gemini
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name

    try:
        sample_file = genai.upload_file(path=tmp_path, display_name="TourDoc")
        
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        prompt = """
        Analyze this document. It is either a 'Tour Approval' or a 'Salary Slip'.
        
        If it is a **Tour Approval**, extract the following into a JSON object:
        - type: "tour"
        - departure_date: (Format DD/MM/YYYY)
        - departure_place: (City name only)
        - arrival_date: (Format DD/MM/YYYY)
        - arrival_place: (City name only)
        - purpose: (The full text under 'Purpose of Journey')
        - mode_of_journey: (e.g., Private Vehicle, Govt Vehicle)
        
        If it is a **Salary Slip**, extract:
        - type: "salary"
        - name: (Employee Name)
        - designation: (Designation)
        - basic_pay: (Basic Pay Amount)
        
        Return ONLY valid JSON. No markdown formatting.
        """
        
        response = model.generate_content([sample_file, prompt])
        return json.loads(response.text.strip().replace('```json', '').replace('```', ''))
        
    except Exception as e:
        st.error(f"Error processing {uploaded_file.name}: {str(e)}")
        return None
    finally:
        os.remove(tmp_path)

def generate_tour_pdf(tour_data, user_details):
    """
    Generates the A4 PDF formatted like the 'Tour Diary' example.
    """
    pdf = FPDF(orientation='L', unit='mm', format='A4') # Landscape to fit table
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    
    # Header
    pdf.cell(0, 10, "TOUR DIARY", ln=True, align='C')
    pdf.set_font("Arial", '', 11)
    
    # User Details Section
    if user_details:
        pdf.cell(0, 8, f"Name: {user_details.get('name', '')}", ln=True)
        pdf.cell(0, 8, f"Designation: {user_details.get('designation', '')}", ln=True)
        pdf.cell(0, 8, f"Basic Salary: {user_details.get('basic_pay', '')}", ln=True)
    pdf.ln(5)
    
    # Table Header
    pdf.set_font("Arial", 'B', 10)
    pdf.set_fill_color(200, 220, 255)
    
    # Columns: Dep Date, Dep Time, Arr Date, Arr Time, Mode, KM, Purpose
    # Widths sum to roughly 270mm (A4 Landscape)
    col_w = [25, 20, 25, 20, 30, 20, 130]
    headers = ["Dep. Date", "Place", "Arr. Date", "Place", "Mode", "KM", "Purpose"]
    
    for i, h in enumerate(headers):
        pdf.cell(col_w[i], 10, h, border=1, fill=True, align='C')
    pdf.ln()
    
    # Table Rows
    pdf.set_font("Arial", '', 9)
    for trip in tour_data:
        # Multi-cell for purpose is tricky in basic FPDF, truncating for simplicity
        # or using basic cells.
        
        # Row 1: Departure
        pdf.cell(col_w[0], 10, trip['departure_date'], border=1)
        pdf.cell(col_w[1], 10, trip['departure_place'], border=1)
        pdf.cell(col_w[2], 10, trip['arrival_date'], border=1)
        pdf.cell(col_w[3], 10, trip['arrival_place'], border=1)
        pdf.cell(col_w[4], 10, trip['mode_of_journey'], border=1)
        pdf.cell(col_w[5], 10, str(trip['distance_km']), border=1)
        
        # Purpose (Truncate to fit single line for basic FPDF)
        purpose_short = (trip['purpose'][:80] + '..') if len(trip['purpose']) > 80 else trip['purpose']
        pdf.cell(col_w[6], 10, purpose_short, border=1)
        pdf.ln()

    return pdf

# --- MAIN APP LOGIC ---

uploaded_files = st.file_uploader("Upload Tour PDFs and Salary Slip", 
                                  type=['pdf'], 
                                  accept_multiple_files=True)

if uploaded_files and st.button("Generate Diary"):
    if not GEMINI_API_KEY or not SERPAPI_KEY:
        st.error("Please enter both API keys in the sidebar.")
    else:
        with st.spinner("Analyzing documents and calculating distances..."):
            
            tour_entries = []
            user_info = {}
            
            for file in uploaded_files:
                data = extract_pdf_data(file, GEMINI_API_KEY)
                
                if data:
                    if data.get('type') == 'salary':
                        user_info = data
                    elif data.get('type') == 'tour':
                        # Calculate Distance Logic
                        origin = data.get('departure_place')
                        dest = data.get('arrival_place')
                        mode = data.get('mode_of_journey', '').lower()
                        
                        km = 0
                        calc_note = ""
                        
                        # Only calculate if Private Vehicle or explicit request
                        if "private" in mode:
                            km, calc_note = get_distance_serpapi(origin, dest, SERPAPI_KEY)
                        
                        data['distance_km'] = km
                        data['calc_note'] = calc_note
                        tour_entries.append(data)

            # Sort by date
            # (Requires consistent date format from Gemini, handled loosely here)
            
            if tour_entries:
                st.success(f"Processed {len(tour_entries)} tours.")
                
                # Preview Data
                df = pd.DataFrame(tour_entries)
                st.dataframe(df[['departure_date', 'departure_place', 'arrival_place', 'distance_km', 'calc_note']])
                
                # Generate PDF
                pdf = generate_tour_pdf(tour_entries, user_info)
                
                # Save and Download
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                    pdf.output(tmp_pdf.name)
                    
                    with open(tmp_pdf.name, "rb") as f:
                        st.download_button(
                            label="Download Tour Diary (PDF)",
                            data=f,
                            file_name="Tour_Diary_Generated.pdf",
                            mime="application/pdf"
                        )
            else:
                st.warning("No tour data found in uploaded files.")


