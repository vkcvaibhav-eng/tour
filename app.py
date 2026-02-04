import streamlit as st
import pandas as pd
import re
from pypdf import PdfReader
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# --- CONFIGURATION & RULES (Gujarat Finance Dept Nov 2024) ---
CITY_CLASS = {
    'Ahmedabad': 'X', 'Surat': 'X', 'Vadodara': 'Y', 'Rajkot': 'Y', 'Bhavnagar': 'Y', 'Jamnagar': 'Y',
    'Gandhinagar': 'Y', 'Junagadh': 'Z', 'Navsari': 'Z', 'Vyara': 'Z', 'Udaipur': 'Z', 'Jaipur': 'Y', 'Ludhiana': 'Y'
}

DA_RATES = {
    "Level 12+": {"X": 1000, "Y": 800, "Z": 500},
    "Level 6-11": {"X": 800, "Y": 500, "Z": 400},
    "Level 1-5": {"X": 500, "Y": 400, "Z": 300}
}

# --- PDF EXTRACTION LOGIC ---
def extract_salary_data(pdf_file):
    reader = PdfReader(pdf_file)
    text = "".join([page.extract_text() for page in reader.pages])
    
    data = {}
    data['name'] = re.search(r"EMP NAME:\s*(.*)", text).group(1).strip() if re.search(r"EMP NAME:\s*(.*)", text) else "N/A"
    data['designation'] = re.search(r"DESIGNATION\s*(.*)", text).group(1).strip() if re.search(r"DESIGNATION\s*(.*)", text) else "N/A"
    data['basic'] = float(re.search(r"Basic\s*([\d,]+\.\d+)", text).group(1).replace(',', '')) if re.search(r"Basic\s*([\d,]+\.\d+)", text) else 0.0
    data['bh'] = re.search(r"BH NO:\s*\((.*?)\)", text).group(1).strip() if re.search(r"BH NO:\s*\((.*?)\)", text) else "303/2092"
    
    # Determine Level based on Basic Pay (Simplified logic for NAU)
    if data['basic'] >= 78800: data['level'] = "Level 12+"
    elif data['basic'] >= 35400: data['level'] = "Level 6-11"
    else: data['level'] = "Level 1-5"
    
    return data

def extract_tour_data(pdf_file):
    reader = PdfReader(pdf_file)
    text = "".join([page.extract_text() for page in reader.pages])
    
    tour = {}
    tour['otms_no'] = re.search(r"(\d{14})", text).group(1) if re.search(r"(\d{14})", text) else "N/A"
    tour['purpose'] = re.search(r"Purpose of Journey:\s*(.*?)\n", text, re.S).group(1).strip() if re.search(r"Purpose of Journey:", text) else "Official Work"
    
    # Extracting Departure/Arrival (Simplified pattern for OTMS)
    tour['from_city'] = "Navsari"
    to_city_match = re.search(r"To\s+(.*?)\s+Private", text)
    tour['to_city'] = to_city_match.group(1).strip() if to_city_match else "Vyara"
    
    date_match = re.search(r"Departure:\s*(\d{4}-\d{2}-\d{2})", text)
    tour['date'] = date_match.group(1) if date_match else datetime.now().strftime("%Y-%m-%d")
    
    return tour

# --- EXCEL GENERATION (NAU FORMAT) ---
def create_nau_excel(emp_data, tours):
    wb = Workbook()
    ws = wb.active
    ws.title = "TA Bill"
    
    # Styling
    bold_font = Font(bold=True, size=11)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Header
    ws.merge_cells('A1:L1')
    ws['A1'] = "GJ;FZL S'lQF lJ`JlJWF,I (NAVSARI AGRICULTURAL UNIVERSITY)"
    ws['A1'].font = Font(bold=True, size=14)
    ws['A1'].alignment = center_align

    # Employee Info
    ws['A3'] = f"Name: {emp_data['name']}"
    ws['F3'] = f"Designation: {emp_data['designation']}"
    ws['A4'] = f"Basic Pay: {emp_data['basic']}"
    ws['F4'] = f"Pay Level: {emp_data['level']}"
    ws['A5'] = f"Budget Head: {emp_data['bh']}"

    # Table Headers
    headers = ["Date", "Departure", "Arrival", "Mode", "Fare", "DA Rate", "DA %", "DA Amt", "Total", "Purpose"]
    for col, text in enumerate(headers, 1):
        cell = ws.cell(row=8, column=col, value=text)
        cell.font = bold_font
        cell.border = border
        cell.alignment = center_align

    # Data Rows
    row_num = 9
    grand_total = 0
    for t in tours:
        city_cls = CITY_CLASS.get(t['to_city'], 'Z')
        da_rate = DA_RATES[emp_data['level']][city_cls]
        
        # Logic: No private vehicle reimbursement (Rule enforced)
        fare = 0.0 # Strict compliance
        
        ws.cell(row=row_num, column=1, value=t['date']).border = border
        ws.cell(row=row_num, column=2, value="Navsari").border = border
        ws.cell(row=row_num, column=3, value=t['to_city']).border = border
        ws.cell(row=row_num, column=4, value="Rail/Bus (Lowest)").border = border
        ws.cell(row=row_num, column=5, value=fare).border = border
        ws.cell(row=row_num, column=6, value=da_rate).border = border
        ws.cell(row=row_num, column=7, value="100%").border = border
        ws.cell(row=row_num, column=8, value=da_rate).border = border
        total = fare + da_rate
        ws.cell(row=row_num, column=9, value=total).border = border
        ws.cell(row=row_num, column=10, value=t['purpose']).border = border
        
        grand_total += total
        row_num += 1

    # Footer
    ws.cell(row=row_num+1, column=8, value="Grand Total:").font = bold_font
    ws.cell(row=row_num+1, column=9, value=grand_total).font = bold_font

    # Certifications
    ws.cell(row=row_num+3, column=1, value="Certified that the shortest route was taken and no private vehicle fare is claimed.")
    
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# --- STREAMLIT UI ---
st.set_page_config(page_title="NAU TA/DA Automation", layout="wide")

st.title("🏛️ NAU TA/DA Reimbursement System")
st.info("Upload your Approved OTMS Tour PDFs and Salary Slip to generate the official NAU TA Bill.")

with st.sidebar:
    st.header("1. Upload Documents")
    salary_pdf = st.file_uploader("Upload Salary Slip (PDF)", type="pdf")
    tour_pdfs = st.file_uploader("Upload Approved OTMS Tours (Multiple PDFs)", type="pdf", accept_multiple_files=True)

if salary_pdf and tour_pdfs:
    try:
        with st.spinner("Processing Documents..."):
            emp_info = extract_salary_data(salary_pdf)
            all_tours = [extract_tour_data(tp) for tp in tour_pdfs]
            
            st.subheader("📋 Verification Preview")
            col1, col2 = st.columns(2)
            with col1:
                st.write("**Employee Details:**")
                st.json(emp_info)
            with col2:
                st.write(f"**Tours Detected:** {len(all_tours)}")
                st.dataframe(pd.DataFrame(all_tours))

            # Generate Excel
            excel_data = create_nau_excel(emp_info, all_tours)
            
            st.success("Reimbursement File Ready!")
            st.download_button(
                label="📥 Download Official NAU TA Bill (Excel)",
                data=excel_data,
                file_name=f"TA_Bill_{emp_info['name'].replace(' ', '_')}_{datetime.now().strftime('%b_%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Error processing files: {e}")
        st.warning("Ensure the PDFs are text-readable and not scanned images.")

else:
    st.warning("Please upload both Salary Slip and at least one Tour PDF to proceed.")
