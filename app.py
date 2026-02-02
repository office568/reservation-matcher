import streamlit as st
import pandas as pd
import pdfplumber
import re

# 1. SETTINGS & CONFIGURATION
TARGET_HEADER = "予約番号"
UPLOAD_HEADERS = ["Confirmation code", "Reference number", "reservation number"]

# 2. HELPER FUNCTIONS
def extract_ids_from_text(text):
    # Matches uppercase letters and numbers, length 7 to 15
    pattern = r'\b[A-Z0-9]{7,15}\b'
    return set(re.findall(pattern, text))

def get_file_ids(uploaded_file):
    """Extracts unique IDs from either CSV or PDF."""
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
        # Try to find specific headers first
        for col in UPLOAD_HEADERS:
            if col in df.columns:
                return set(df[col].astype(str).str.strip().unique())
        # Fallback: scan the entire CSV as text
        return extract_ids_from_text(df.to_string())
    else:
        # PDF Extraction logic
        with pdfplumber.open(uploaded_file) as pdf:
            text = " ".join([page.extract_text() for page in pdf.pages if page.extract_text()])
        return extract_ids_from_text(text)

def highlight_rows(row, matched_list):
    """Applies light green background to matched rows."""
    val = str(row[TARGET_HEADER]).strip()
    if val in matched_list:
        return ['background-color: #d4edda'] * len(row)
    return [''] * len(row)

# 3. UI LAYOUT
st.set_page_config(page_title="Reservation Matcher", layout="wide")
st.title("📑 Smart Reservation Matcher")

# Sidebar for inputs
with st.sidebar:
    st.header("Configuration")
    links_input = st.text_area("Google Sheet Links (One per line)", 
                               placeholder="https://docs.google.com/spreadsheets/d/...",
                               height=200)
    uploaded_file = st.file_uploader("Upload CSV or PDF", type=["csv", "pdf"])
    st.info(f"Scanning for: **{TARGET_HEADER}**")

# 4. MAIN EXECUTION LOGIC
if st.button("🔍 Run Comparison"):
    if not uploaded_file:
        st.error("Please upload a CSV or PDF file.")
    elif not links_input:
        st.error("Please paste at least one Google Sheet link.")
    else:
        # Step A: Extract IDs from the uploaded file
        with st.spinner("Extracting codes from your file..."):
            extracted_ids = get_file_ids(uploaded_file)
        
        if not extracted_ids:
            st.warning("No reservation codes (7-15 chars) found in the file.")
        else:
            st.success(f"Found {len(extracted_ids)} codes in your file.")
            
            # Step B: Parse the links
            sheet_links = [link.strip() for link in links_input.split('\n') if link.strip()]
            
            # Step C: Loop through each link
            for i, link in enumerate(sheet_links):
                try:
                    # Construct the direct CSV export link
                    if '/edit' in link:
                        base_url = link.split('/edit')[0]
                    else:
                        base_url = link.rstrip('/')
                    
                    csv_url = f"{base_url}/export?format=csv"
                    
                    # Read the Google Sheet
                    sheet_df = pd.read_csv(csv_url)
                    
                    # Normalize headers (remove extra spaces)
                    sheet_df.columns = [str(c).strip() for c in sheet_df.columns]
                    
                    # --- RESILIENCE CHECK ---
                    if TARGET_HEADER not in sheet_df.columns:
                        st.warning(f"⚠️ Skipping Sheet {i+1}: Column '{TARGET_HEADER}' not found.")
                        continue # Skips this sheet and goes to the next link
                    
                    # Clean data for matching
                    sheet_df[TARGET_HEADER] = sheet_df[TARGET_HEADER].astype(str).str.strip()
                    sheet_ids = set(sheet_df[TARGET_HEADER].unique())
                    
                    found_in_sheet = extracted_ids.intersection(sheet_ids)
                    missing_from_sheet = extracted_ids - sheet_ids

                    # Display Results
                    st.write(f"### Results for Sheet {i+1}")
                    st.caption(f"Source: {link}")

                    tab1, tab2 = st.tabs(["📊 Matched View", "❌ Missing from Sheet"])

                    with tab1:
                        # Display highlighted dataframe
                        st.dataframe(
                            sheet_df.style.apply(highlight_rows, matched_list=found_in_sheet, axis=1),
                            use_container_width=True
                        )

                    with tab2:
                        if missing_from_sheet:
                            st.error(f"These {len(missing_from_sheet)} codes were not found in this sheet:")
                            st.code("\n".join(list(missing_from_sheet)))
                        else:
                            st.success("All codes from your file exist in this sheet!")
                            
                except Exception as e:
                    st.error(f"❌ Error processing Sheet {i+1}: {e}")
                    continue
