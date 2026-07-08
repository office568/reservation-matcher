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
CSV_PRICE_HEADERS = ["Gross earnings", "Gross amount", "Amount", "Transaction amount", "Payable amount", "Paid out"]

IGNORED_TAB_KEYWORDS = ["コピー", "COPY", "ARCHIVE", "古い", "OLD", "バックアップ", "BACKUP", "TEST", "テスト"]

DEFAULT_LINKS = [
    "https://docs.google.com/spreadsheets/d/1iQZTAYk8mq6j_1H4-u8TO4SZZX61HKULtv2lvLrUf6w/edit",
    "https://docs.google.com/spreadsheets/d/1eD0C4rFpHJDye5lDuS3kIkse2g9KV3qsjQW1NHyn1ug/edit",
    "https://docs.google.com/spreadsheets/d/11ISMzFrZl6rYSdOSYshl_KfzfghNpNUzTx_gCoZRsvI/edit",
    "https://docs.google.com/spreadsheets/d/1nqEuhnHqaPY_-okDlPfwssQwIO2KG8JGCgC65CwKeVc/edit",
    "https://docs.google.com/spreadsheets/d/1PpaSECqXI9YC2Xax7o5G985hk2MuSkF9SWIlrf4WTRI/edit",
    "https://docs.google.com/spreadsheets/d/1rvw82CBs4BTE2iUKwjMkIVMODLT_nyXIsKv1TIFR6zU/edit",
    "https://docs.google.com/spreadsheets/d/1qgJj_7qL68SbOdNVRXqkHQx7gjFPqE5RcbsXJEPhY0Q/edit",
    "https://docs.google.com/spreadsheets/d/1b2WPv0ybZc85_CyuyT7KqA9CQPVMvZ-6QB3u2g8hmNY/edit",
    "https://docs.google.com/spreadsheets/d/1p9datgPRonSfbsIyRY0g7tNfKYfnsKWdBaJQSNWV2ng/edit",
    "https://docs.google.com/spreadsheets/d/1Ms-S5qgkaGPv5iUvJZqMYcEUbpUTsn859IYlY8p_W18/edit"
]

def get_gspread_client():
    return gspread.service_account_from_dict(st.secrets["gcp_service_account"])

def clean_numeric_string(val):
    if pd.isna(val) or val is None or str(val).strip() == "" or str(val).strip() == "-":
        return None
    cleaned = re.sub(r'[^\d\.\-]', '', str(val))
    try:
        return int(float(cleaned))
    except ValueError:
        return None

def extract_data_from_file(uploaded_file):
    """Step 1: Gather all reservation numbers and revenues from the uploaded file."""
    file_data = {"type": "csv", "data": {}}
    if uploaded_file.name.endswith('.csv'):
        try:
            df = pd.read_csv(uploaded_file)
        except:
            df = pd.read_csv(uploaded_file, encoding='cp932')
        
        df.columns = [str(c).strip() for c in df.columns]
        
        type_col = next((col for col in df.columns if "type" in col.lower()), None)
        code_col = next((col for col in CSV_MATCH_HEADERS if col in df.columns), None)
        price_col = next((col for col in CSV_PRICE_HEADERS if col in df.columns), None)
        
        if code_col and price_col:
            for _, row in df.iterrows():
                if type_col and str(row[type_col]).strip().lower() not in ["reservation", "booked", "okay"]:
                    continue
                
                code = str(row[code_col]).strip().upper()
                if pd.isna(row[code_col]) or code == "NAN" or not code or code == "-":
                    continue
                
                parsed_price = clean_numeric_string(row[price_col])
                
                # Protect valid prices from being overwritten by blank/zero adjustment rows
                if code in file_data["data"] and parsed_price is None:
                    continue
                    
                file_data["data"][code] = parsed_price
        else:
            file_data["type"] = "text_only"
            file_data["data"] = set(re.findall(r'\b[A-Z0-9]{7,15}\b', df.to_string().upper()))
    return file_data

# --- UI SETUP ---
st.set_page_config(page_title="Seirai Auto-Matcher", page_icon="🟢", layout="wide")
st.title("🟢 Seirai Group: Direct Reservation Matcher")
st.markdown(f"Scanning **{len(DEFAULT_LINKS)}** spreadsheets sequentially.")

uploaded_file = st.file_uploader("Upload your transaction CSV record", type=["csv"])

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
                st.error("No valid reservation numbers found in your file.")
            else:
                st.success(f"Successfully gathered {len(csv_codes)} reservation numbers from file.")
                st.divider()

                globally_found_codes = set()
                all_price_mismatches = []
                progress_bar = st.progress(0)
                
                green_fmt = cellFormat(backgroundColor=color(0.82, 0.93, 0.85))
                red_fmt = cellFormat(backgroundColor=color(0.97, 0.84, 0.85))
                
                # Step 2: Look through all spreadsheets, one sheet at a time
                for index, link in enumerate(DEFAULT_LINKS):
                    if "docs.google.com" not in link: continue
                        
                    try:
                        sh = client.open_by_url(link)
                        st.subheader(f"📂 Spreadsheet: {sh.title}")
                        
                        for worksheet in sh.worksheets():
                            time.sleep(1.2) # API Rate Limit protection
                            
                            tab_title_upper = worksheet.title.upper()
                            if any(keyword in tab_title_upper for keyword in IGNORED_TAB_KEYWORDS):
                                continue
                                
                            data = worksheet.get_all_values()
                            if not data or len(data) < 1: continue
                            
                            col_idx = None
                            price_idx = None
                            guest_idx = None
                            header_row_num = 0
                            
                            for r_num in range(min(6, len(data))):
                                row_cells = [str(c).strip() for c in data[r_num]]
                                if any(c == TARGET_GSHEET_HEADER for c in row_cells):
                                    col_idx = next(i for i, c in enumerate(row_cells) if c == TARGET_GSHEET_HEADER)
                                    header_row_num = r_num + 1
                                    price_idx = next((i for i, c in enumerate(row_cells) if c == TARGET_GSHEET_PRICE), None)
                                    guest_idx = next((i for i, c in enumerate(row_cells) if any(g in c for g in ["宿泊者", "ゲスト", "名前", "Guest", "Name"])), None)
                                    break
                            
                            if col_idx is None: continue 

                            formatting_requests = []
                            tab_green_matches = 0
                            tab_red_mismatches = 0
                            
                            for current_row, row_data in enumerate(data[header_row_num:], start=header_row_num + 1):
                                if len(row_data) <= col_idx: continue
                                gsheet_val = str(row_data[col_idx]).strip().upper()
                                
                                # Step 3: When found a match based on reservation number...
                                if gsheet_val in csv_codes:
                                    globally_found_codes.add(gsheet_val)
                                    guest_name = str(row_data[guest_idx]).strip() if (guest_idx is not None and guest_idx < len(row_data)) else "Unknown"
                                    
                                    # Step 4: Check if the revenue is correct
                                    if is_csv_mode and price_idx is not None and price_idx < len(row_data):
                                        csv_price = parsed_file_package["data"][gsheet_val]
                                        gsheet_price = clean_numeric_string(row_data[price_idx])
                                        
                                        # If empty string or unentered, let it default green instead of forcing an error flag
                                        if gsheet_price is None or gsheet_price == csv_price:
                                            formatting_requests.append((f"A{current_row}:Z{current_row}", green_fmt))
                                            tab_green_matches += 1
                                        else:
                                            formatting_requests.append((f"A{current_row}:Z{current_row}", red_fmt))
                                            tab_red_mismatches += 1
                                            all_price_mismatches.append({
                                                "Spreadsheet": sh.title,
                                                "Tab Name": worksheet.title,
                                                "Row #": current_row,
                                                "Guest Name": guest_name,
                                                "Code": gsheet_val,
                                                "CSV Price": f"¥{csv_price:,}" if csv_price else "¥0",
                                                "GSheet Price": f"¥{gsheet_price:,}" if gsheet_price else "¥0"
                                            })
                                    else:
                                        formatting_requests.append((f"A{current_row}:Z{current_row}", green_fmt))
                                        tab_green_matches += 1

                            if formatting_requests:
                                format_cell_ranges(worksheet, formatting_requests)
                                if tab_green_matches > 0:
                                    st.success(f"✅ Tab '{worksheet.title}': Found and colored {tab_green_matches} rows.")
                                if tab_red_mismatches > 0:
                                    st.error(f"❌ Tab '{worksheet.title}': Flagged {tab_red_mismatches} price conflicts.")

                    except Exception as e:
                        st.error(f"Error accessing sheet {index + 1}: {e}")
                    
                    progress_bar.progress((index + 1) / len(DEFAULT_LINKS))

                # --- SUMMARY REPORT LOGS ---
                st.divider()
                if is_csv_mode and all_price_mismatches:
                    st.error("🚨 PRICE DISCREPANCIES DETECTED (Reservation code matched, but revenue amount did not)")
                    st.dataframe(pd.DataFrame(all_price_mismatches), use_container_width=True)
                
                # Final check: For reservation numbers that can't find a match, put it in the missing section
                missing_codes = csv_codes - globally_found_codes
                if missing_codes:
                    st.warning(f"⚠️ {len(missing_codes)} Reservations Completely Missing From Spreadsheets")
                    missing_df = pd.DataFrame(list(missing_codes), columns=["Missing Reservation Code"])
                    st.dataframe(missing_df, use_container_width=True)
                
                if not all_price_mismatches and not missing_codes:
                    st.balloons()
                    st.success("Reconciliation successful! Perfect 100% match achieved across all records.")

        except Exception as auth_e:
            st.error(f"Authentication Failure: {auth_e}")
