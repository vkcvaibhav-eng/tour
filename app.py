import streamlit as st
import google.generativeai as genai
from serpapi import GoogleSearch
from fpdf import FPDF
import PyPDF2
import json
import io
import datetime

# --- CONFIGURATION & SETUP ---
st.set_page_config(page_title="Auto-Tour Diary Generator", layout="wide")

st.title("📄 Automated Tour Diary Generator")
st.markdown("""
This tool extracts tour details using **Gemini 1.5 Pro**, calculates distances based on **Govt Rules (Rail vs. Road/GSRTC)** via SerpApi, 
and generates a final **Tour Diary PDF** for submission.
""")

# --- SIDEBAR: API KEYS ---
with st.sidebar:
    st.header("🔑 API Configuration")
    GEMINI_API_KEY = st.text_input("Enter Gemini API Key", type="password")
    SERPAPI_KEY = st.text_input("Enter SerpApi Key", type="password")
    
    st.info("Files needed: Upload your 'Online Tour Management' PDFs and your 'Salary Slip'.")

# --- HELPER FUNCTIONS ---

def extract_text_from_pdf(uploaded_file):
    """Extracts raw text from a PDF file."""
    try:
        pdf_reader = PyPDF2.PdfReader(uploaded_file)
        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text()
        return text
    except Exception as e:
        st.error(f"Error reading PDF: {e}")
        return None

def get_distance_logic(origin, destination, api_key):
    """
    Calculates distance based on User Rule:
    1. Check Direct Railway Connection.
    2. If Direct Rail exists -> Return Rail Distance.
    3. Else -> Return Road Distance (GSRTC proxy).
    """
    if not api_key:
        return "N/A (No Key)", "Manual"

    # 1. Try Finding a Train Route
    params_train = {
        "engine": "google_maps_directions",
        "start_addr": origin,
        "end_addr": destination,
        "travel_mode": "transit",
        "transit_mode": "train",
        "api_key": api_key
    }
    
    try:
        search = GoogleSearch(params_train)
        results = search.get_dict()
        
        # Check if a valid train route exists
        if "directions" in results and results["directions"]:
            route = results["directions"][0]
            
            # Check for direct connection (1 step usually means direct or minimal changes)
            # We look at the 'steps' to ensure it's a train
            is_direct_train = False
            train_distance_km = 0
            
            for leg in route.get("legs", []):
                train_distance_km = leg["distance"]["value"] / 1000 # Convert meters to km
                for step in leg.get("steps", []):
                    if step.get("travel_mode") == "TRANSIT" and step.get("transit_details", {}).get("vehicle", {}).get("type") == "TRAIN":
                        is_direct_train = True
            
            if is_direct_train:
                return f"{train_distance_km:.2f}", "Rail (Direct)"

    except Exception as e:
        print(f"Rail search failed: {e}")

    # 2. Fallback to Road (GSRTC/Driving) if no direct train
    params_road = {
        "engine": "google_maps_directions",
        "start_addr": origin,
        "end_addr": destination,
        "travel_mode": "driving", # Standard for road distance calculation
        "api_key": api_key
    }
    
    try:
        search = GoogleSearch(params_road)
        results = search.get_dict()
        
        if "directions" in results and results["directions"]:
            route = results["directions"][0]
            dist_meters = route["legs"][0]["distance"]["value"]
            dist_km = dist_meters / 1000
            return f"{dist_km:.2f}", "Road (GSRTC Logic)"
            
    except Exception as e:
        return "Error", "Manual Check"
        
    return "0", "Unknown"

def analyze_with_gemini(text_content, doc_type):
    """Uses Gemini 1.5 Pro to extract specific JSON data."""
    if not GEMINI_API_KEY:
        return None

    genai.configure(api_key=GEMINI_API_KEY)
    
    # Define schemas based on document type
    if doc_type == "tour":
        prompt = f"""
        Extract the following details from this Tour Report text and return ONLY a valid JSON object.
        Do not include markdown formatting like ```json.
        
        Fields required:
        - tour_date (DD/MM/YYYY)
        - origin_city (City name only, e.g., Navsari)
        - destination_city (City name only, e.g., Vyara)
        - departure_time (HH:MM format, if found, else "00:00")
        - arrival_time (HH:MM format, if found, else "00:00")
        - purpose (Short summary of the purpose)
        
        Text content:
        {text_content[:4000]}
        """
    elif doc_type == "salary":
        prompt = f"""
        Extract the following details from this Salary Slip text and return ONLY a valid JSON object.
        Do not include markdown formatting like ```json.
        
        Fields required:
        - employee_name
        - designation
        - basic_salary
        - headquarters (City name, usually inferred from office address)
        
        Text content:
        {text_content[:4000]}
        """

    model = genai.GenerativeModel('gemini-1.5-pro')
    response = model.generate_content(prompt)
    
    try:
        # Clean response to ensure pure JSON
        cleaned_response = response.text.strip().replace('```json', '').replace('```', '')
        return json.loads(cleaned_response)
    except:
        st.error("Failed to parse Gemini response.")
        return None

class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'TOUR DIARY / T.A. BILL DETAILS', 0, 1, 'C')
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

# --- MAIN APP LOGIC ---

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Upload Salary Slip (For Personal Details)")
    salary_file = st.file_uploader("Upload Salary PDF", type=['pdf'], key="salary")

with col2:
    st.subheader("2. Upload Tour Approvals")
    tour_files = st.file_uploader("Upload Generated Tour PDFs", type=['pdf'], accept_multiple_files=True, key="tours")

if st.button("🚀 Process Files & Calculate Distances"):
    if not GEMINI_API_KEY or not SERPAPI_KEY:
        st.error("Please enter both API Keys in the sidebar.")
    elif not salary_file or not tour_files:
        st.error("Please upload both salary slip and tour documents.")
    else:
        # 1. Process Salary Slip
        st.info("Analyzing Salary Slip...")
        salary_text = extract_text_from_pdf(salary_file)
        salary_data = analyze_with_gemini(salary_text, "salary")
        
        if salary_data:
            st.success(f"Identified: {salary_data.get('employee_name')} ({salary_data.get('designation')})")
        
        # 2. Process Tour Files
        st.info(f"Analyzing {len(tour_files)} Tour Documents & Calculating Distances...")
        
        processed_tours = []
        
        progress_bar = st.progress(0)
        
        for idx, tour_file in enumerate(tour_files):
            # Extract Text
            tour_text = extract_text_from_pdf(tour_file)
            
            # Extract Data via Gemini
            tour_data = analyze_with_gemini(tour_text, "tour")
            
            if tour_data:
                # Calculate Distance via SerpApi
                origin = tour_data.get('origin_city', 'Navsari')
                dest = tour_data.get('destination_city', '')
                
                dist_km, mode = get_distance_logic(origin, dest, SERPAPI_KEY)
                
                tour_data['distance_km'] = dist_km
                tour_data['calc_mode'] = mode
                processed_tours.append(tour_data)
            
            progress_bar.progress((idx + 1) / len(tour_files))

        # Store in session state for PDF generation
        st.session_state['salary_data'] = salary_data
        st.session_state['processed_tours'] = processed_tours
        st.success("Processing Complete!")

# --- DISPLAY RESULTS & GENERATE PDF ---

if 'processed_tours' in st.session_state and st.session_state['processed_tours']:
    st.write("---")
    st.subheader("3. Verified Data")
    
    # Display table of calculated data
    st.table(st.session_state['processed_tours'])
    
    if st.button("📥 Generate Final Tour Diary PDF"):
        salary = st.session_state.get('salary_data', {})
        tours = st.session_state.get('processed_tours', [])
        
        pdf = PDF(orientation='L', unit='mm', format='A4') # Landscape for better table fit
        pdf.add_page()
        pdf.set_font("Arial", size=10)
        
        # Employee Header Info
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(0, 7, f"Name: {salary.get('employee_name', 'Unknown')}", ln=True)
        pdf.cell(0, 7, f"Designation: {salary.get('designation', 'Unknown')}", ln=True)
        pdf.cell(0, 7, f"Basic Salary: {salary.get('basic_salary', 'Unknown')}", ln=True)
        pdf.cell(0, 7, f"Headquarters: {salary.get('headquarters', 'Navsari')}", ln=True)
        pdf.ln(5)
        
        # Table Header
        pdf.set_fill_color(200, 220, 255)
        pdf.cell(30, 10, "Date", 1, 0, 'C', 1)
        pdf.cell(40, 10, "Origin", 1, 0, 'C', 1)
        pdf.cell(40, 10, "Destination", 1, 0, 'C', 1)
        pdf.cell(20, 10, "Mode", 1, 0, 'C', 1)
        pdf.cell(20, 10, "KM", 1, 0, 'C', 1)
        pdf.cell(120, 10, "Purpose", 1, 1, 'C', 1)
        
        # Table Rows
        pdf.set_font("Arial", size=9)
        total_km = 0.0
        
        for tour in tours:
            try:
                km_val = float(tour.get('distance_km', 0))
                total_km += km_val
            except:
                pass
                
            pdf.cell(30, 10, str(tour.get('tour_date')), 1)
            pdf.cell(40, 10, str(tour.get('origin_city')), 1)
            pdf.cell(40, 10, str(tour.get('destination_city')), 1)
            pdf.cell(20, 10, "Pvt/Road" if "Road" in tour.get('calc_mode') else "Rail", 1)
            pdf.cell(20, 10, str(tour.get('distance_km')), 1)
            pdf.cell(120, 10, str(tour.get('purpose')), 1, 1) # 1,1 means line break after
            
        pdf.ln(5)
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(0, 10, f"Total Kilometers: {total_km:.2f} KM", ln=True)
        
        # Output
        pdf_bytes = pdf.output(dest='S').encode('latin-1')
        
        st.download_button(
            label="Download PDF Report",
            data=pdf_bytes,
            file_name="generated_tour_diary.pdf",
            mime="application/pdf"
        )
