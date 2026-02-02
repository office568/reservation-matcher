for i, link in enumerate(sheet_links):
            try:
                # Convert link to CSV export URL
                base_url = link.split('/edit')[0]
                csv_url = f"{base_url}/export?format=csv"
                
                sheet_df = pd.read_csv(csv_url)
                
                # NORMALIZE HEADERS: Remove spaces and make it easier to match
                sheet_df.columns = [str(c).strip() for c in sheet_df.columns]
                
                # CHECK: If the target header isn't there, skip this sheet gracefully
                if TARGET_HEADER not in sheet_df.columns:
                    st.warning(f"⚠️ Skipping Sheet {i+1}: Column '{TARGET_HEADER}' was not found. Moving to next...")
                    continue # This jumps to the next link in the list

                # Prepare IDs for matching
                sheet_df[TARGET_HEADER] = sheet_df[TARGET_HEADER].astype(str).str.strip()
                sheet_ids = set(sheet_df[TARGET_HEADER].unique())
                
                found_in_sheet = extracted_ids.intersection(sheet_ids)
                missing_from_sheet = extracted_ids - sheet_ids

                # UI Display
                st.subheader(f"✅ Sheet {i+1} Processed")
                st.caption(f"Source: {link}")

                tab1, tab2 = st.tabs(["📊 Matched Table", "❌ Missing Codes"])

                with tab1:
                    styled_df = sheet_df.style.apply(highlight_rows, matched_list=found_in_sheet, axis=1)
                    st.dataframe(styled_df, use_container_width=True)

                with tab2:
                    if missing_from_sheet:
                        st.error(f"{len(missing_from_sheet)} codes from your file are not in this sheet.")
                        st.write(list(missing_from_sheet))
                    else:
                        st.success("Perfect Match!")

            except Exception as e:
                # This catches errors like 'URL not found' or 'Private Sheet'
                st.error(f"❌ Could not access Sheet {i+1}. Skipping. (Error: {e})")
                continue
