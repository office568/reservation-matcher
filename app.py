import streamlit as st
import pandas as pd
import pdfplumber
import re

# --- CONFIGURATION ---
TARGET_HEADER = "予約番号"
UPLOAD_HEADERS = ["Confirmation code", "Reference number", "reservation number"]

def extract_ids_from_text(text):
    # Matches: Uppercase letters and numbers, length 7 to 15
    pattern = r'\b[A-Z0-9]{7,15}\b'
    # Use findall and then filter out purely alphabetic strings if needed, 
    # but usually, reservation codes are fine with this.
    return set(re.findall(pattern, text))

def get_file_ids(uploaded_file):
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
        # Check for any of our target headers
        for col in UPLOAD_HEADERS:
            if col in df.columns:
                return set(df[col].astype(str).str.strip().unique())
        
        # Fallback: if no header matches, search the whole CSV textually
        st.warning("Exact header not found in CSV. Searching all cells for patterns...")
        all_text = df.to_string()
        return extract_ids_from_text(all_text)
    else:
        # PDF Extraction
        with pdfplumber.open(uploaded_file) as pdf:
            text = " ".join([page.extract_text() for page in pdf.pages if page.extract_text()])
        return extract_ids_from_text(text)

def highlight_rows(row, matched_list):
    val = str(row[TARGET_HEADER]).strip()
    if val in matched_list:
        return ['background-color: #d4edda'] * len(row)
    return [''] * len(row)

# --- UI SETUP ---
st.set_page_config(page_title="Reservation Matcher", layout="wide")
st.title("📑 Smart Reservation Matcher")
st.markdown("Match CSV/PDF data against multiple Google Sheets automatically.")

with st.sidebar:
    st.header("1. Data Sources")
    links_input = st.text_area("Paste Google Sheet Links (one per line)", 
                               placeholder="https://docs.google.com/spreadsheets/d/...",
                               height=200)
    
    st.header("2. Upload File")
    uploaded_file = st.file_uploader("Upload CSV or PDF", type=["csv", "pdf"])
    
    st.divider()
    st.info(f"Targeting GSheet Column: **{TARGET_HEADER}**")

# --- EXECUTION ---
if st.button("🔍 Run Comparison") and uploaded_file and links_input:
    # 1. Extract IDs
    with st.spinner("Extracting codes from file..."):
        extracted_ids = get_file_ids(uploaded_file)
    
    if not extracted_ids:
        st.error("No valid reservation codes found in your file. Check the format.")
    else:
        st.success(f"Extracted {len(extracted_ids)} potential reservation codes.")
        
        sheet_links = [link.strip() for link in links_input.split('\n') if link.strip()]
        
        for i, link in enumerate(sheet_links):
            try:
                # Convert link to CSV export URL
                base_url = link.split('/edit')[0]
                csv_url = f"{base_url}/export?format=csv"
                
                sheet_df = pd.read_csv(csv_url)
                
                # Normalize column names to handle small typos/spaces
                sheet_df.columns = [c.strip() for c in sheet_df.columns]
                
                if TARGET_HEADER not in sheet_df.columns:
                    st.error(f"Sheet {i+1}: Column '{TARGET_HEADER}' not found.")
                    continue

                # Prepare IDs for matching
                sheet_df[TARGET_HEADER] = sheet_df[TARGET_HEADER].astype(str).str.strip()
                sheet_ids = set(sheet_df[TARGET_HEADER].unique())
                
                found_in_sheet = extracted_ids.intersection(sheet_ids)
                missing_from_sheet = extracted_ids - sheet_ids

                # UI Display
                st.subheader(f"Sheet {i+1} Results")
                st.caption(f"URL: {link}")

                tab1, tab2 = st.tabs(["📊 Matched Table", "❌ Missing Codes"])

                with tab1:
                    # Apply the green highlight
                    styled_df = sheet_df.style.apply(highlight_rows, matched_list=found_in_sheet, axis=1)
                    st.dataframe(styled_df, use_container_width=True)

                with tab2:
                    if missing_from_sheet:
                        st.error(f"The following {len(missing_from_sheet)} codes were NOT found in this sheet:")
                        st.write(list(missing_from_sheet))
                    else:
                        st.success("Perfect Match! All codes from your file are in this sheet.")

            except Exception as e:
                st.error(f"Could not read Sheet {i+1}. Ensure it is shared as 'Anyone with the link'. Error: {e}")
