import streamlit as st
import pandas as pd
import pdfplumber
import re
from urllib.parse import urlparse, parse_qs

# 1. SETTINGS & CONFIGURATION
TARGET_HEADER = "予約番号"
UPLOAD_HEADERS = ["Confirmation code", "Reference number", "reservation number"]

# 2. HELPER FUNCTIONS
def extract_ids_from_text(text):
    # Matches uppercase letters and numbers, length 7 to 15
    pattern = r'\b[A-Z0-9]{7,15}\b'
    return set(re.findall(pattern, text))

def get_file_ids(uploaded_file):
    if uploaded_file.name.endswith('.csv'):
        # Try different encodings for CSVs
        try:
            df = pd.read_csv(uploaded_file)
        except:
            df = pd.read_csv(uploaded_file, encoding='shift-jis')
            
        for col in UPLOAD_HEADERS:
            if col in df.columns:
                return set(df[col].astype(str).str.strip().unique())
        return extract_ids_from_text(df.to_string())
    else:
        with pdfplumber.open(uploaded_file) as pdf:
            text = " ".join([page.extract_text() for page in pdf.pages if page.extract_text()])
        return extract_ids_from_text(text)

def highlight_rows(row, matched_list):
    # We find the actual column name used (in case of spaces)
    actual_col = next((c for c in row.index if TARGET_HEADER in str(c)), None)
    if actual_col:
        val = str(row[actual_col]).strip()
        if val in matched_list:
            return ['background-color: #d4edda'] * len(row)
    return [''] * len(row)

# 3. UI LAYOUT
st.set_page_config(page_title="Reservation Matcher", layout="wide")
st.title("📑 Smart Reservation Matcher")

with st.sidebar:
    st.header("Configuration")
    links_input = st.text_area("Google Sheet Links (One per line)", height=200)
    uploaded_file = st.file_uploader("Upload CSV or PDF", type=["csv", "pdf"])
    st.info(f"Scanning for: **{TARGET_HEADER}**")

# 4. MAIN EXECUTION
if st.button("🔍 Run Comparison"):
    if not uploaded_file or not links_input:
        st.error("Please provide both a file and at least one link.")
    else:
        with st.spinner("Processing..."):
            extracted_ids = get_file_ids(uploaded_file)
        
        if not extracted_ids:
            st.warning("No reservation codes found in your file.")
        else:
            st.success(f"Extracted {len(extracted_ids)} codes from your file.")
            sheet_links = [link.strip() for link in links_input.split('\n') if link.strip()]
            
            for i, link in enumerate(sheet_links):
                try:
                    # IMPROVED URL PARSING FOR GID
                    parsed_url = urlparse(link)
                    base_url = link.split('/edit')[0]
                    query_params = parse_qs(parsed_url.query)
                    gid = query_params.get('gid', ['0'])[0] # Default to 0 if no gid
                    
                    csv_url = f"{base_url}/export?format=csv&gid={gid}"
                    sheet_df = pd.read_csv(csv_url)
                    
                    # Normalize headers
                    clean_cols = [str(c).strip() for c in sheet_df.columns]
                    sheet_df.columns = clean_cols
                    
                    # Check if target header is in any column name
                    matched_header = next((c for c in clean_cols if TARGET_HEADER in c), None)
                    
                    if not matched_header:
                        st.warning(f"⚠️ Sheet {i+1}: Could not find '{TARGET_HEADER}'.")
                        with st.expander("Show available columns in this sheet"):
                            st.write(clean_cols)
                        continue
                    
                    sheet_df[matched_header] = sheet_df[matched_header].astype(str).str.strip()
                    sheet_ids = set(sheet_df[matched_header].unique())
                    
                    found_in_sheet = extracted_ids.intersection(sheet_ids)
                    missing_from_sheet = extracted_ids - sheet_ids

                    st.write(f"### Results for Sheet {i+1}")
                    tab1, tab2 = st.tabs(["📊 Table", "❌ Missing"])

                    with tab1:
                        st.dataframe(sheet_df.style.apply(highlight_rows, matched_list=found_in_sheet, axis=1), use_container_width=True)

                    with tab2:
                        if missing_from_sheet:
                            st.error(f"{len(missing_from_sheet)} codes missing:")
                            st.code("\n".join(list(missing_from_sheet)))
                        else:
                            st.success("All codes matched!")
                            
                except Exception as e:
                    st.error(f"❌ Error with Sheet {i+1}: {e}")
