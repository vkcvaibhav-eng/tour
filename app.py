import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
from datetime import datetime, timedelta

# ==========================================
# 1. DATA EXTRACTION LOGIC
# ==========================================

def parse_salary_slip(uploaded_file):
    """
    Extracts employee details from NAU Salary Slip PDF.
    Based on format: 'EMP NAME:', 'DESIGNATION', 'Basic', 'Level'
    """
    data = {
        "name": "",
        "designation": "",
        "pay_level": "12", # Defaulting based on user profile
        "basic_pay": 0,
        "headquarters": "N.M.C.A., NAU, Navsari" # Default
    }
    
    with pdfplumber.open(uploaded_file) as pdf:
        text = "\n".join([page.extract_text() for page in pdf.pages])
        
        # Regex extraction based on your file snippets
        name_match = re.search(r"EMP NAME:\s*(.*)", text)
        desig_match = re.search(r"DESIGNATION\s*(.*)", text)
        basic_match = re.search(r"Basic\s*[\"']?([\d,]+\.\d{2})", text)
        
        # Pay Level often appears near Scale. Using a fallback or user input is safer, 
        # but we try to find 'Level' or 'PB SCALE'
        
        if name_match:
            data["name"] = name_match.group(1).strip()
        if desig_match:
            data["designation"] = desig_match.group(1).strip()
        if basic_match:
            # Remove commas and convert to float
            clean_basic = basic_match.group(1).replace(",", "").replace('"', "")
            data["basic_pay"] = float(clean_basic)

    return data

def parse_tour_pdf(uploaded_file):
    """
    Extracts tour details from NAU Online Tour Management System PDF.
    Captures: Date, Location, Purpose, Mode of Journey
    """
    tour_data = {
        "departure_date": None,
        "arrival_date": None,
        "destination": "",
        "purpose": "",
        "distance_km": 0, # Placeholder, will need user input or calculation
        "mode": "Private Vehicle" # Default seen in snippets
    }
    
    with pdfplumber.open(uploaded_file) as pdf:
        text = "\n".join([page.extract_text() for page in pdf.pages])
        
        # Extract Dates
        # Snippet format: "1days, Departure: 2025-08-28, Arrival: 2025-08-28"
        date_match = re.search(r"Departure:\s*(\d{4}-\d{2}-\d{2}).*Arrival:\s*(\d{4}-\d{2}-\d{2})", text)
        if date_match:
            tour_data["departure_date"] = date_match.group(1)
            tour_data["arrival_date"] = date_match.group(2)
            
        # Extract Purpose
        purpose_match = re.search(r"Purpose of Journey:\s*(.*?)Justification", text, re.DOTALL)
        if purpose_match:
            # Clean up newlines and extra spaces
            raw_purpose = purpose_match.group(1).replace("\n", " ").strip()
            # Truncate if too long for Excel
            tour_data["purpose"] = (raw_purpose[:75] + '..') if len(raw_purpose) > 75 else raw_purpose
            
        # Extract Destination (Heuristic based on snippets)
        if "Vyara" in text:
            tour_data["destination"] = "Polytechnic, Vyara"
            tour_data["distance_km"] = 65 # Hardcoded estimate for Navsari-Vyara based on context
        elif "Udaipur" in text:
            tour_data["destination"] = "Udaipur"
        
        # Extract Mode
        if "Private Vehical" in text or "Private Vehicle" in text:
            tour_data["mode"] = "Private Vehicle"
        elif "Public Bus" in text:
            tour_data["mode"] = "Public Bus"

    return tour_data

# ==========================================
# 2. CALCULATION LOGIC (Gujarat Govt Rules)
# ==========================================

def calculate_da(pay_level, hours, city_class="Z"):
    """
    Calculates Daily Allowance based on Pay Level and Duration.
    Simplified logic based on 7th Pay Comm rules.
    """
    # Rate Table (Simplified for Level 9-13 based on 2024 GR)
    # Z Class (Ordinary)
    full_rate = 500 
    if city_class == "X": full_rate = 900
    elif city_class == "Y": full_rate = 700 # Example rates
    
    # Duration Logic
    if hours < 6:
        return 0
    elif 6 <= hours < 12:
        return 0.5 * full_rate
    else:
        return full_rate

def process_tours(tour_files, mileage_rate):
    """
    Process list of tour files and generate line items.
    """
    rows = []
    total_claim = 0
    
    for f in tour_files:
        info = parse_tour_pdf(f)
        
        if not info["departure_date"]:
            continue
            
        # Logic: Create two rows (Outward and Return) for 1-day tours
        # Assuming 8:00 AM start and 7:00 PM return for 1-day tours
        dep_dt = datetime.strptime(info["departure_date"], "%Y-%m-%d")
        
        # Row 1: Outward
        rows.append({
            "Date": info["departure_date"],
            "From": "NAU, Navsari",
            "To": info["destination"],
            "Mode": info["mode"],
            "Distance": info["distance_km"],
            "Amount": info["distance_km"] * mileage_rate,
            "DA": 0, # Usually DA is calculated at the end of the day or return
            "Purpose": info["purpose"]
        })
        
        # Row 2: Return
        # Calculating DA for the day (Assuming >6 hours <12 hours for single day trip)
        da_amount = 250 # 50% of 500
        
        rows.append({
            "Date": info["arrival_date"],
            "From": info["destination"],
            "To": "NAU, Navsari",
            "Mode": info["mode"],
            "Distance": info["distance_km"],
            "Amount": info["distance_km"] * mileage_rate,
            "DA": da_amount,
            "Purpose": "Return Journey"
        })
        
        total_claim += (info["distance_km"] * 2 * mileage_rate) + da_amount

    return pd.DataFrame(rows), total_claim

# ==========================================
# 3. EXCEL GENERATION
# ==========================================

def generate_excel_bill(df, user_data, total_claim):
    output = io.BytesIO()
    
    # Using XlsxWriter for formatting
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Write user details header manually
        workbook = writer.book
        worksheet = workbook.add_worksheet("TA Bill")
        writer.sheets['TA Bill'] = worksheet
        
        # Formats
        header_fmt = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'bg_color': '#D3D3D3'})
        cell_fmt = workbook.add_format({'border': 1, 'text_wrap': True, 'valign': 'top'})
        title_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
        
        # Title
        worksheet.merge_range('A1:H1', "NAVSARI AGRICULTURAL UNIVERSITY", title_fmt)
        worksheet.merge_range('A2:H2', "Traveling Allowance Bill", title_fmt)
        
        # User Info Block
        worksheet.write('A4', f"Name: {user_data['name']}")
        worksheet.write('E4', f"Designation: {user_data['designation']}")
        worksheet.write('A5', f"Basic Pay: {user_data['basic_pay']}")
        worksheet.write('E5', f"Headquarters: {user_data['headquarters']}")
        
        # Write Data Table
        # We write the dataframe starting at row 7
        df.to_excel(writer, sheet_name='TA Bill', startrow=7, startcol=0, index=False)
        
        # Apply formatting to data cells
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(7, col_num, value, header_fmt)
            
        # Adjust column widths
        worksheet.set_column('A:A', 12) # Date
        worksheet.set_column('B:C', 20) # From/To
        worksheet.set_column('H:H', 30) # Purpose
        
        # Total Row
        last_row = 7 + len(df) + 1
        worksheet.write(last_row, 6, "Total Claim:", header_fmt)
        worksheet.write(last_row, 7, f"₹ {total_claim:,.2f}", header_fmt)
        
        # Certifications
        cert_row = last_row + 3
        worksheet.merge_range(f'A{cert_row}:H{cert_row}', 
            "Certified that the journey was performed in the interest of University work.", cell_fmt)
        worksheet.merge_range(f'A{cert_row+4}:C{cert_row+4}', "Signature of Claimant", cell_fmt)
        worksheet.merge_range(f'F{cert_row+4}:H{cert_row+4}', "Controlling Officer", cell_fmt)

    return output.getvalue()

# ==========================================
# 4. STREAMLIT UI
# ==========================================

st.set_page_config(page_title="NAU TA/DA Automation", layout="wide")

st.title("🎓 NAU TA/DA Reimbursement Automation System")
st.markdown("Automate the generation of Gujarat Govt / NAU compliant TA bills from your official PDFs.")

# Sidebar - Settings
with st.sidebar:
    st.header("⚙️ Settings")
    mileage_rate = st.number_input("Mileage Rate (₹/km)", value=11.0, step=0.5)
    default_hq = st.text_input("Headquarters", "N.M.C.A., NAU, Navsari")

# File Upload Section
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Upload Salary Slip")
    salary_file = st.file_uploader("Upload PDF (for Basic Pay & Designation)", type=['pdf'])

with col2:
    st.subheader("2. Upload Tour Approvals")
    tour_files = st.file_uploader("Upload Tour PDFs (OTMS)", type=['pdf'], accept_multiple_files=True)

# Main Processing
if salary_file and tour_files:
    # Parse Data
    user_info = parse_salary_slip(salary_file)
    user_info["headquarters"] = default_hq
    
    st.success(f"✅ Identified Employee: **{user_info['name']}** ({user_info['designation']})")
    
    # Process Tours
    df_tours, total_amt = process_tours(tour_files, mileage_rate)
    
    # Display Preview
    st.subheader("📋 Bill Preview")
    st.dataframe(df_tours)
    
    st.metric(label="Total Claim Amount", value=f"₹ {total_amt:,.2f}")
    
    # Generate Excel
    excel_data = generate_excel_bill(df_tours, user_info, total_amt)
    
    st.download_button(
        label="📥 Download TA Bill (Excel)",
        data=excel_data,
        file_name=f"TA_Bill_{user_info['name'].split()[0]}_{datetime.now().strftime('%b%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

elif not salary_file:
    st.info("👈 Please upload your Salary Slip to begin.")