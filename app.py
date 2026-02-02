# --- EXECUTION BLOCK ---
# We wrap everything in this 'if' so it only runs when the button is pressed
if st.button("🔍 Run Comparison"):
    if not uploaded_file:
        st.error("Please upload a CSV or PDF file first.")
    elif not links_input:
        st.error("Please paste at least one Google Sheet link.")
    else:
        # 1. Extract IDs from the uploaded file
        with st.spinner("Extracting codes from file..."):
            extracted_ids = get_file_ids(uploaded_file)
        
        if not extracted_ids:
            st.error("No valid reservation codes found in your file. Check the format.")
        else:
            st.success(f"Extracted {len(extracted_ids)} potential reservation codes.")
            
            # 2. Process the links provided in the text area
            sheet_links = [link.strip() for link in links_input.split('\n') if link.strip()]
            
            for i, link in enumerate(sheet_links):
                try:
                    # Convert link to CSV export URL
                    # Handling different link formats safely
                    if '/edit' in link:
                        base_url = link.split('/edit')[0]
                    else:
                        base_url = link.rstrip('/')
                        
                    csv_url = f"{base_url}/export?format=csv"
                    
                    sheet_df = pd.read_csv(csv_url)
                    
                    # Normalize headers
                    sheet_df.columns = [str(c).strip() for c in sheet_df.columns]
                    
                    # Skip sheet if column is missing
                    if TARGET_HEADER not in sheet_df.columns:
                        st.warning(f"⚠️ Skipping Sheet {i+1}: Column '{TARGET_HEADER}' not found.")
                        continue

                    # Prepare IDs for matching
                    sheet_df[TARGET_HEADER] = sheet_df[TARGET_HEADER].astype(str).str.strip()
                    sheet_ids = set(sheet_df[TARGET_HEADER].unique())
                    
                    found_in_sheet = extracted_ids.intersection(sheet_ids)
                    missing_from_sheet = extracted_ids - sheet_ids

                    # UI Display
                    st.write(f"### Results for Sheet {i+1}")
                    st.caption(f"Source: {link}")

                    tab1, tab2 = st.tabs(["📊 Matched Table", "❌ Missing Codes"])

                    with tab1:
                        # Highlight rows and display
                        styled_df = sheet_df.style.apply(highlight_rows, matched_list=found_in_sheet, axis=1)
                        st.dataframe(styled_df, use_container_width=True)

                    with tab2:
                        if missing_from_sheet:
                            st.error(f"{len(missing_from_sheet)} codes from your file were NOT found in this sheet:")
                            st.write(list(missing_from_sheet))
                        else:
                            st.success("Perfect Match! All codes found.")

                except Exception as e:
                    st.error(f"❌ Error with Sheet {i+1}: {e}")
                    continue
