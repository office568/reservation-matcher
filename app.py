import streamlit as st
import pandas as pd
import pdfplumber
import re
import gspread
from gspread_formatting import *
import time

# --- CONFIGURATION ---
TARGET_GSHEET_HEADER = "予約番号"
TARGET_GSHEET_PRICE = "金額"
CSV_MATCH_HEADERS = ["Confirmation code", "Reference code", "Reference number"]

# CRITICAL FIX: Put "Gross earnings" and "Gross amount" first so the app prioritizes total revenue
CSV_PRICE_HEADERS = ["Gross earnings", "Gross amount", "Amount", "Transaction amount", "Payable amount", "Paid out"]

DEFAULT_LINKS = [
    "https://docs.google.com/spreadsheets/d/1iQZTAYk8mq6j_1H4-u8TO4SZZX61HKULtv2lvLrUf6w/edit",
    "https://docs.google.com/spreadsheets/d/1eD0C4rFpHJDye5lDuS3kIkse2g9KV3qsjQW1NHyn1ug/edit",
    "https://docs.google.com/spreadsheets/d/11ISMzFrZl6rYSdOSYshl_KfzfghNpNUzTx_gCoZRsvI/edit",
    "https://docs.google.com/spreadsheets/d/1nqEuhnHqaPY_-okDlPfwssQwIO2KG8JGCgC65CwKeVc/edit",
    "https://docs.google.com/spreadsheets/d/1PpaSECqXI9YC2Xax7o5G985hk2MuSkF9SWIlrf4WTRI/edit",
    "https://docs.google.com/spreadsheets/d/1rvw82CBs4BTE2iUKwjMkIVMODLT_nyXIsKv1TIFR6zU/edit",
    "https://docs.google.com/spreadsheets/d/1qgJj_7qL68SbOdNVRXqkHQx7gjFPqE5RcbsXJEPhY0Q/edit",
    "PASTE_YOUR_8TH_NEW_PROPERTY_LINK_HERE" # <-- Remember to paste your 8th link here if needed
]

def get_gspread_client():
    """Authenticates using Streamlit Secrets."""
    return gspread.service_account_from_dict(st.secrets["gcp_service_account"])

def clean_numeric_string(val):
    """Converts raw currency or object strings into flat integers for clean matching."""
    if pd.isna(val) or val is None or str(val).strip() == "" or str(val).strip() == "-":
        return None
    cleaned = re.sub(r'[^\d\.\-]', '', str(val))
    try:
        return int(float(cleaned))
    except ValueError:
        return None

def extract_data_from_file(uploaded_file):
    """Extracts a clean dictionary mapping {RESERVATION_CODE: PRICE_INT} from CSV or PDF."""
    file_data = {"type": "csv", "data": {}}
    
    if uploaded_file.name.endswith('.csv'):
        try:
            df = pd.read_csv(uploaded_file)
        except:
            df = pd.read_csv(uploaded_file, encoding='cp932')
        
        # Dynamically find header paths
        type_col = next((col for col in df.columns if "type" in col.lower()), None)
        code_col = next((col for col in CSV_MATCH_HEADERS if col in df.columns), None)
        price_col = next((col for col in CSV_PRICE_HEADERS if col in df.columns), None)
        
        if code_col and price_col:
            for _, row in df.iterrows():
                # Filter strictly for standard bookings to ignore payout distributions/adjustments
                if type_col and str(row[type_col]).strip() != "Reservation":
                    continue
                    
                code = str(row[code_col]).strip().upper()
                if pd.isna(row[code_col]) or code == "NAN" or not code or code == "-":
                    continue
                file_data["data"][code] = clean_numeric_string(row[price_col])
        else:
            file_data["type"] = "text_only"
            file_data["data"] = set(re.findall(r'\b[A-Z0-9]{7,15}\b', df.to_string().upper()))
    else:
        file_data["type"] = "text_only"
        with pdfplumber.open(uploaded_file) as pdf:
            text = " ".join([page.extract_text() for page in pdf.pages if page.extract_text()])
        file_data["data"] = set(re.findall(r'\b[A-Z0-9]{7,15}\b', text.upper()))
        
    return file_data

# --- UI SETUP ---
st.set_page_config(page_title="Seirai Auto-Matcher", page_icon="🟢", layout="wide")
st.title("🟢 Seirai Group: Multi-Property Reconciliation Agent")
st.markdown(f"Automated verification engine synced with **{len(DEFAULT_LINKS)}** properties.")

uploaded_file = st.file_uploader("Upload your transaction CSV record", type=["csv", "pdf"])

if st.button("🚀 Run Matching & Price Reconciliation"):
    if not uploaded_file:
        st.error("Please upload your source file to begin.")
    else:
        try:
            client = get_gspread_client()
            parsed_file_package = extract_data_from_file(uploaded_file)
            is_csv_mode = parsed_file_package["type"] == "csv"
            
            csv_codes = set(parsed_file_package["data"].keys()) if is_csv_mode else parsed_file_package["data"]

            if not csv_codes:
                st.error("No valid transaction fields mapped from layout.")
            else:
                st.success(f"Mapped {len(csv_codes)} active validation targets.")
                st.divider()

                globally_found_codes = set()
                all_price_mismatches = []
                progress_bar = st.progress(0)
                
                green_fmt = cellFormat(backgroundColor=color(0.82, 0.93, 0.85)) # Perfect match color
                red_fmt = cellFormat(backgroundColor=color(0.97, 0.84, 0.85))   # Soft discrepancy red
                
                for index, link in enumerate(DEFAULT_LINKS):
                    if "docs.google.com" not in link: continue
                        
                    try:
                        sh = client.open_by_url(link)
                        st.subheader(f"📂 Spreadsheet: {sh.title}")
                        
                        for worksheet in sh.worksheets():
                            time.sleep(1.2) # Active Anti-Throttling Guard
                            data = worksheet.get_all_values()
                            if not data: continue
                            
                            headers = [str(h).strip() for h in data[0]]
                            col_idx = next((idx for idx, h in enumerate(headers) if TARGET_GSHEET_HEADER in h), None)
                            price_idx = next((idx for idx, h in enumerate(headers) if TARGET_GSHEET_PRICE in h), None)
                            
                            if col_idx is None: continue 

                            formatting_requests = []
                            tab_green_matches = 0
                            tab_red_mismatches = 0
                            
                            for row_num, row_data in enumerate(data[1:], start=2):
                                gsheet_val = str(row_data[col_idx]).strip().upper()
                                
                                if gsheet_val in csv_codes:
                                    globally_found_codes.add(gsheet_val)
                                    
                                    # Perform cross-verification
                                    if is_csv_mode and price_idx is not None:
                                        csv_price = parsed_file_package["data"][gsheet_val]
                                        gsheet_price = clean_numeric_string(row_data[price_idx])
                                        
                                        # Skip flagging blank cells as discrepancies
                                        if gsheet_price is None:
                                            formatting_requests.append((f"A{row_num}:Z{row_num}", green_fmt))
                                            tab_green_matches += 1
                                        elif gsheet_price == csv_price:
                                            formatting_requests.append((f"A{row_num}:Z{row_num}", green_fmt))
                                            tab_green_matches += 1
                                        else:
                                            formatting_requests.append((f"A{row_num}:Z{row_num}", red_fmt))
                                            tab_red_mismatches += 1
                                            all_price_mismatches.append({
                                                "Spreadsheet": sh.title,
                                                "Tab Name": worksheet.title,
                                                "Code": gsheet_val,
                                                "CSV Price": f"¥{csv_price:,}" if csv_price else "¥0",
                                                "GSheet Price": f"¥{gsheet_price:,}"
                                            })
                                    else:
                                        formatting_requests.append((f"A{row_num}:Z{row_num}", green_fmt))
                                        tab_green_matches += 1

                            if formatting_requests:
                                format_cell_ranges(worksheet, formatting_requests)
                                if tab_green_matches > 0:
                                    st.success(f"✅ Tab '{worksheet.title}': Highlighted {tab_green_matches} accurate matches.")
                                if tab_red_mismatches > 0:
                                    st.error(f"❌ Tab '{worksheet.title}': Flagged {tab_red_mismatches} price discrepancies.")

                    except Exception as e:
                        st.error(f"Error accessing sheet {index + 1}: {e}")
                    
                    progress_bar.progress((index + 1) / len(DEFAULT_LINKS))

                # --- SUMMARY REPORT LOGS ---
                st.divider()
                if is_csv_mode and all_price_mismatches:
                    st.error("🚨 PRICE DISCREPANCIES DETECTED (Rows Marked RED in Sheets)")
                    st.markdown("The values in the spreadsheet do not align with your official CSV records:")
                    st.dataframe(pd.DataFrame(all_price_mismatches), use_container_width=True)
                
                missing_codes = csv_codes - globally_found_codes
                if missing_codes:
                    st.warning(f"⚠️ {len(missing_codes)} Unmapped Transactions (Missing From Spreadsheets entirely)")
                    st.markdown("These entries appear on your CSV records but were not located on your property tracking tabs:")
                    missing_df = pd.DataFrame(list(missing_codes), columns=["Missing Reservation Code"])
                    st.dataframe(missing_df, use_container_width=True)
                
                if not all_price_mismatches and not missing_codes:
                    st.balloons()
                    st.success("Reconciliation successful! Complete alignment achieved with zero financial or data leaks.")

        except Exception as auth_e:
            st.error(f"Authentication Failure: {auth_e}")
