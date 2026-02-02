import streamlit as st
import pandas as pd
import pdfplumber
import re
import gspread
from gspread_formatting import *

# --- CONFIGURATION ---
TARGET_GSHEET_HEADER = "予約番号"
CSV_MATCH_HEADERS = ["Confirmation code", "Reference code", "Reference number"]

def get_gspread_client():
    return gspread.service_account_from_dict(st.secrets["gcp_service_account"])

def extract_ids_from_file(uploaded_file):
    all_codes = set()
    try:
        if uploaded_file.name.endswith('.csv'):
            # Try Airbnb standard encoding first
            try:
                df = pd.read_csv(uploaded_file)
            except:
                df = pd.read_csv(uploaded_file, encoding='cp932')
            
            found_col = False
            for col in CSV_MATCH_HEADERS:
                if col in df.columns:
                    # Clean: Convert to string, uppercase, and strip spaces
                    codes = df[col].dropna().astype(str).str.strip().str.upper().unique()
                    all_codes.update(codes)
                    found_col = True
                    break
            if not found_col:
                all_codes.update(re.findall(r'\b[A-Z0-9]{7,15}\b', df.to_string().upper()))
        else:
            with pdfplumber.open(uploaded_file) as pdf:
                text = " ".join([page.extract_text() for page in pdf.pages if page.extract_text()])
            all_codes.update(re.findall(r'\b[A-Z0-9]{7,15}\b', text.upper()))
    except Exception as e:
        st.error(f"File reading error: {e}")
    return all_codes

st.set_page_config(page_title="Global Spreadsheet Matcher", page_icon="🟢")
st.title("🌍 Global Spreadsheet Matcher & Highlighter")
st.info("System will now scan EVERY tab in EVERY spreadsheet provided.")

with st.sidebar:
    links_input = st.text_area("Spreadsheet URLs (one per line)", height=200)
    uploaded_file = st.file_uploader("Upload airbnb_.csv", type=["csv", "pdf"])

if st.button("🚀 Run Matching Across All Tabs"):
    if not uploaded_file or not links_input:
        st.error("Missing file or links.")
    else:
        try:
            client = get_gspread_client()
            extracted_ids = extract_ids_from_file(uploaded_file)
            sheet_links = [l.strip() for l in links_input.split('\n') if l.strip()]

            if not extracted_ids:
                st.error("No valid reservation codes found in your file.")
            else:
                st.success(f"Extracted {len(extracted_ids)} codes from your file.")
                # Diagnostic: Show first 3 extracted IDs
                st.write(f"Sample IDs from file: `{list(extracted_ids)[:3]}`")

                for link in sheet_links:
                    try:
                        sh = client.open_by_url(link)
                        st.subheader(f"📂 Processing Sheet: {sh.title}")
                        
                        # Loop through ALL tabs (Worksheets)
                        for worksheet in sh.worksheets():
                            data = worksheet.get_all_values()
                            if not data: continue
                            
                            headers = [str(h).strip() for h in data[0]]
                            col_idx = next((idx for idx, h in enumerate(headers) if TARGET_GSHEET_HEADER in h), None)
                            
                            if col_idx is None:
                                st.caption(f"Skipped tab '{worksheet.title}': Column '{TARGET_GSHEET_HEADER}' not found.")
                                continue

                            # Batch formatting setup
                            fmt = cellFormat(backgroundColor=color(0.82, 0.93, 0.85))
                            requests = []
                            match_count = 0
                            
                            for row_num, row_data in enumerate(data[1:], start=2):
                                # Clean GSheet value: String, Uppercase, Strip spaces
                                gsheet_val = str(row_data[col_idx]).strip().upper()
                                
                                if gsheet_val in extracted_ids:
                                    requests.append((f"A{row_num}:Z{row_num}", fmt))
                                    match_count += 1

                            if requests:
                                format_cell_ranges(worksheet, requests)
                                st.success(f"✅ Tab '{worksheet.title}': Found and highlighted {match_count} matches.")
                            else:
                                st.write(f"ℹ️ Tab '{worksheet.title}': No matches found.")

                    except Exception as e:
                        st.error(f"Error processing sheet {link}: {e}")

        except Exception as auth_e:
            st.error(f"Authentication failure: {auth_e}")
