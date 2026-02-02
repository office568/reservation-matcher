import streamlit as st
import pandas as pd
import pdfplumber
import re
import gspread
from gspread_formatting import *

# --- CONFIGURATION ---
TARGET_GSHEET_HEADER = "予約番号"
CSV_MATCH_HEADERS = ["Confirmation code", "Reference code", "Reference number"]

# Your Pre-defined Spreadsheet Links
DEFAULT_LINKS = [
    "https://docs.google.com/spreadsheets/d/1iQZTAYk8mq6j_1H4-u8TO4SZZX61HKULtv2lvLrUf6w/edit",
    "https://docs.google.com/spreadsheets/d/1eD0C4rFpHJDye5lDuS3kIkse2g9KV3qsjQW1NHyn1ug/edit",
    "https://docs.google.com/spreadsheets/d/11ISMzFrZl6rYSdOSYshl_KfzfghNpNUzTx_gCoZRsvI/edit",
    "https://docs.google.com/spreadsheets/d/1nqEuhnHqaPY_-okDlPfwssQwIO2KG8JGCgC65CwKeVc/edit",
    "https://docs.google.com/spreadsheets/d/1PpaSECqXI9YC2Xax7o5G985hk2MuSkF9SWIlrf4WTRI/edit",
    "https://docs.google.com/spreadsheets/d/1rvw82CBs4BTE2iUKwjMkIVMODLT_nyXIsKv1TIFR6zU/edit",
    "https://docs.google.com/spreadsheets/d/1qgJj_7qL68SbOdNVRXqkHQx7gjFPqE5RcbsXJEPhY0Q/edit"
]

def get_gspread_client():
    return gspread.service_account_from_dict(st.secrets["gcp_service_account"])

def extract_ids_from_file(uploaded_file):
    all_codes = set()
    try:
        if uploaded_file.name.endswith('.csv'):
            try:
                df = pd.read_csv(uploaded_file)
            except:
                df = pd.read_csv(uploaded_file, encoding='cp932')
            
            found_col = False
            for col in CSV_MATCH_HEADERS:
                if col in df.columns:
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

# --- UI SETUP ---
st.set_page_config(page_title="Seirai Auto-Matcher", page_icon="🟢", layout="wide")
st.title("🟢 Seirai Group: Auto-Spreadsheet Highlighter")
st.markdown(f"Scanning **{len(DEFAULT_LINKS)}** spreadsheets for reservation codes...")

uploaded_file = st.file_uploader("Upload your airbnb_.csv file", type=["csv", "pdf"])

if st.button("🚀 Start Matching & Highlighting"):
    if not uploaded_file:
        st.error("Please upload a file first.")
    else:
        try:
            client = get_gspread_client()
            extracted_ids = extract_ids_from_file(uploaded_file)

            if not extracted_ids:
                st.error("No valid reservation codes found in your file.")
            else:
                st.success(f"Successfully extracted {len(extracted_ids)} codes from your file.")
                st.write(f"**Sample codes extracted:** `{list(extracted_ids)[:5]}`")
                st.divider()

                progress_bar = st.progress(0)
                
                for index, link in enumerate(DEFAULT_LINKS):
                    try:
                        sh = client.open_by_url(link)
                        st.subheader(f"📂 Spreadsheet: {sh.title}")
                        
                        # Loop through EVERY tab in the current spreadsheet
                        for worksheet in sh.worksheets():
                            data = worksheet.get_all_values()
                            if not data: continue
                            
                            headers = [str(h).strip() for h in data[0]]
                            col_idx = next((idx for idx, h in enumerate(headers) if TARGET_GSHEET_HEADER in h), None)
                            
                            if col_idx is None:
                                continue # Skip tab if column not found

                            fmt = cellFormat(backgroundColor=color(0.82, 0.93, 0.85))
                            requests = []
                            match_count = 0
                            
                            for row_num, row_data in enumerate(data[1:], start=2):
                                # Data Cleaning for Comparison
                                gsheet_val = str(row_data[col_idx]).strip().upper()
                                
                                if gsheet_val in extracted_ids:
                                    requests.append((f"A{row_num}:Z{row_num}", fmt))
                                    match_count += 1

                            if requests:
                                format_cell_ranges(worksheet, requests)
                                st.success(f"✅ Tab '{worksheet.title}': Highlighted {match_count} matches.")
                            else:
                                st.info(f"ℹ️ Tab '{worksheet.title}': No matches.")

                    except Exception as e:
                        st.error(f"Error accessing sheet {index + 1}: {e}")
                    
                    # Update progress bar
                    progress_bar.progress((index + 1) / len(DEFAULT_LINKS))

                st.balloons()
                st.success("Matching Process Complete!")

        except Exception as auth_e:
            st.error(f"System Authentication Failed. Please check the Service Account settings. Error: {auth_e}")
