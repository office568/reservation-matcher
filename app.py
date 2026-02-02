import streamlit as st
import pandas as pd
import pdfplumber
import re
import gspread
from gspread_formatting import *
from urllib.parse import urlparse, parse_qs

# --- CONFIGURATION ---
TARGET_GSHEET_HEADER = "予約番号"
# Headers we look for in your uploaded CSV
CSV_MATCH_HEADERS = ["Confirmation code", "Reference code", "Reference number"]

# --- HELPER FUNCTIONS ---

def get_gspread_client():
    """Authenticates using Streamlit Secrets."""
    return gspread.service_account_from_dict(st.secrets["gcp_service_account"])

def extract_ids_from_file(uploaded_file):
    """Extracts unique reservation codes from CSV or PDF."""
    all_codes = set()
    
    if uploaded_file.name.endswith('.csv'):
        # Try common encodings for Airbnb exports
        try:
            df = pd.read_csv(uploaded_file)
        except:
            df = pd.read_csv(uploaded_file, encoding='cp932') # Common for Japanese systems
            
        # 1. Try specific columns first for accuracy
        found_col = False
        for col in CSV_MATCH_HEADERS:
            if col in df.columns:
                all_codes.update(df[col].dropna().astype(str).str.strip().unique())
                found_col = True
                break
        
        # 2. Fallback: Search the entire CSV textually if headers didn't match
        if not found_col:
            text_blob = df.to_string()
            all_codes.update(re.findall(r'\b[A-Z0-9]{7,15}\b', text_blob))
            
    else:
        # PDF Extraction
        with pdfplumber.open(uploaded_file) as pdf:
            text = " ".join([page.extract_text() for page in pdf.pages if page.extract_text()])
        all_codes.update(re.findall(r'\b[A-Z0-9]{7,15}\b', text))
        
    return all_codes

# --- UI SETUP ---
st.set_page_config(page_title="GSheet Cloud Matcher", layout="wide", page_icon="🟢")
st.title("🟢 Smart Cloud Matcher & Highlighter")
st.markdown("""
This agent finds reservation codes in your file and **paints the rows light green** in your actual Google Spreadsheet.
""")

with st.sidebar:
    st.header("1. Input Data")
    links_input = st.text_area("Google Sheet Links (One per line)", 
                               placeholder="Paste links with gid=...", height=200)
    uploaded_file = st.file_uploader("Upload airbnb_.csv or PDF", type=["csv", "pdf"])
    st.divider()
    st.info(f"Scanning for column: **{TARGET_GSHEET_HEADER}**")

# --- MAIN LOGIC ---
if st.button("🚀 Run Cloud Matching"):
    if not uploaded_file or not links_input:
        st.error("Please provide both a file and at least one Google Sheet link.")
    else:
        try:
            client = get_gspread_client()
            extracted_ids = extract_ids_from_file(uploaded_file)
            
            if not extracted_ids:
                st.warning("No reservation codes (7-15 alphanumeric) found in your file.")
            else:
                st.success(f"Extracted {len(extracted_ids)} unique codes from your file.")
                
                sheet_links = [l.strip() for l in links_input.split('\n') if l.strip()]
                
                for i, link in enumerate(sheet_links):
                    try:
                        # 1. Open Spreadsheet
                        sh = client.open_by_url(link)
                        
                        # 2. Identify the correct Worksheet (Tab) using gid
                        parsed_url = urlparse(link)
                        query_params = parse_qs(parsed_url.query)
                        gid = query_params.get('gid', ['0'])[0]
                        
                        worksheet = next((w for w in sh.worksheets() if str(w.id) == gid), sh.get_worksheet(0))
                        
                        # 3. Get all data to find the column index
                        data = worksheet.get_all_values()
                        if not data:
                            st.warning(f"Sheet {i+1} is empty. Skipping...")
                            continue
                            
                        headers = [str(h).strip() for h in data[0]]
                        col_idx = next((idx for idx, h in enumerate(headers) if TARGET_GSHEET_HEADER in h), None)
                        
                        if col_idx is None:
                            st.warning(f"⚠️ Skipping Sheet {i+1}: Could not find '{TARGET_GSHEET_HEADER}' column.")
                            continue

                        # 4. Prepare Formatting (Light Green)
                        green_fmt = cellFormat(backgroundColor=color(0.82, 0.93, 0.85))
                        
                        # 5. Process Rows & Highlight
                        matches_this_sheet = 0
                        matched_list = []
                        
                        # We use batching for better performance if possible, 
                        # but simple loop is fine for most moderate sheets.
                        for row_num, row_data in enumerate(data[1:], start=2):
                            val = str(row_data[col_idx]).strip()
                            if val in extracted_ids:
                                # format_cell_range uses 1-based indexing for rows. 
                                # A:Z covers the first 26 columns.
                                format_cell_range(worksheet, f"A{row_num}:Z{row_num}", green_fmt)
                                matches_this_sheet += 1
                                matched_list.append(val)

                        # 6. Final Report for this Sheet
                        st.write(f"### Results for: {sh.title} (Tab: {worksheet.title})")
                        
                        c1, c2 = st.columns(2)
                        with c1:
                            st.success(f"✅ Found and Highlighted {matches_this_sheet} rows!")
                        with c2:
                            missing = extracted_ids - set(matched_list)
                            if missing:
                                with st.expander(f"❌ {len(missing)} Codes Missing in this Sheet"):
                                    st.write(list(missing))
                            else:
                                st.balloons()
                                st.success("All codes from file matched in this sheet!")
                        st.divider()

                    except Exception as e:
                        st.error(f"Error with Sheet {i+1}: {e}")
                        continue
                        
        except Exception as auth_e:
            st.error(f"Cloud Connection Failed. Did you share the sheet with the Service Account email? Error: {auth_e}")
