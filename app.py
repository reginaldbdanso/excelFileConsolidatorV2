import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime
import io

st.set_page_config(page_title="Excel-JSON Converter", layout="wide")

def excel_to_json(uploaded_files):
    """Convert multiple Excel files to a single JSON with specific columns"""
    combined_data = []
    
    # Columns to capture (0-indexed, so C=2, D=3, E=4, F=5, H=7, I=8, J=9)
    columns_to_capture = [2, 3, 4, 5, 7, 8, 9]
    
    # Define the field names for these columns
    field_names = [
        "Full Name",
        "Phone Number",
        "School Name",
        "Designation - Level",  # Renamed from "Teaching - Level"
        "Email",                # Renamed from "NTC Email"
        "Region",
        "District"
    ]
    
    for file in uploaded_files:
        try:
            # Read Excel file without header inference
            df = pd.read_excel(file, header=None)
            
            # Skip the header row(s) if needed
            # Assuming the header is in the first row and data starts from second row
            if len(df) > 2:
                df = df.iloc[2:]
            
            # Select only the specified columns if they exist
            if max(columns_to_capture) < len(df.columns):
                selected_df = df.iloc[:, columns_to_capture]
                
                # Rename columns to the desired field names
                selected_df.columns = field_names
                
                # Process until we hit an empty row
                valid_records = []
                for index, row in selected_df.iterrows():
                    # Check if row is empty (all values are NaN or empty string)
                    is_empty = row.isna().all() or (row.astype(str).str.strip() == '').all()
                    
                    if is_empty:
                        st.info(f"Empty row detected at row {index} in file {file.name}. Stopping processing.")
                        break
                    
                    # Add valid row to records
                    valid_records.append(row.fillna("").to_dict())
                
                combined_data.extend(valid_records)
                st.success(f"Processed {len(valid_records)} rows from {file.name}")
            else:
                st.warning(f"File {file.name} doesn't have enough columns. Expected at least {max(columns_to_capture)+1} columns.")
            
        except Exception as e:
            st.error(f"Error processing {file.name}: {str(e)}")
            import traceback
            st.error(traceback.format_exc())
            
    return combined_data

def rearrange_json_fields(json_data):
    """Rearrange JSON data fields to move Email before Phone Number"""
    rearranged_data = []
    
    # Define the desired order of fields
    field_order = [
        "Full Name",
        "Email",
        "Phone Number",
        "School Name",
        "Designation - Level",
        "Region",
        "District"
    ]
    
    for record in json_data:
        rearranged_record = {field: record.get(field, "") for field in field_order}
        rearranged_data.append(rearranged_record)
    
    return rearranged_data

def json_to_excel(json_data):
    """Convert JSON data back to Excel file"""
    output = io.BytesIO()
    
    # Convert JSON to DataFrame
    df = pd.DataFrame(json_data)
    
    # Write to Excel
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="Staff Data", index=False)
    
    output.seek(0)
    return output

def main():
    st.title("Staff Data Excel-JSON Converter")
    
    # Create tabs for different functionalities
    tab1, tab2, tab3 = st.tabs(["Excel to JSON", "Rearrange JSON", "JSON to Excel"])
    
    with tab1:
        st.header("Convert Excel files to JSON")
        
        st.info("""
        This tool will process Excel files with the following rules:
        - Only specific columns will be captured: C, D, E, F, H, I, and J
        - These columns will be mapped to: 
          * Full Name (C)
          * Phone Number (D)
          * School Name (E)
          * Designation - Level (F) 
          * Email (H)
          * Region (I)
          * District (J)
        - The tool assumes data starts from the second row (first row is headers)
        - Processing will stop when an empty row is encountered
        """)
        
        uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx", "xls"], accept_multiple_files=True, key="upload_excel")
        
        if uploaded_files:
            st.write(f"{len(uploaded_files)} files uploaded")
            
            if st.button("Convert to JSON"):
                with st.spinner("Converting files..."):
                    # Convert Excel files to JSON
                    combined_data = excel_to_json(uploaded_files)
                    
                    if combined_data:
                        # Display sample of JSON data (first 5 records)
                        st.subheader("Sample of Converted Data (First 5 records)")
                        st.json(combined_data[:5] if len(combined_data) > 5 else combined_data)
                        
                        # Show total records
                        st.success(f"Successfully converted {len(combined_data)} total records")
                        
                        # Provide download option
                        json_string = json.dumps(combined_data, indent=4)
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        
                        st.download_button(
                            label="Download JSON",
                            data=json_string,
                            file_name=f"staff_data_{timestamp}.json",
                            mime="application/json"
                        )
                        
                        # Save for later use in session state
                        st.session_state.json_data = combined_data
                    else:
                        st.error("No data was converted. Please check the files and try again.")

    
    with tab2:
        st.header("Rearrange JSON (Email before Phone Number)")
        
        st.info("""
        This tab allows you to rearrange the JSON data to place Email before Phone Number.
        The field order will be:
        1. Full Name
        2. Email
        3. Phone Number
        4. School Name
        5. Designation - Level
        6. Region
        7. District
        """)
        
        # Option to use previously converted JSON or upload a new JSON file
        json_source = st.radio(
            "JSON Source",
            ["Upload JSON file", "Use previously converted JSON"],
            index=1 if "json_data" in st.session_state else 0,
            key="json_source_rearrange"
        )
        
        json_data = None
        
        if json_source == "Upload JSON file":
            json_file = st.file_uploader("Upload JSON file", type=["json"], key="upload_json_rearrange")
            
            if json_file:
                try:
                    json_data = json.load(json_file)
                    st.success(f"JSON file loaded successfully with {len(json_data)} records")
                except Exception as e:
                    st.error(f"Error loading JSON file: {str(e)}")
        else:
            if "json_data" in st.session_state:
                json_data = st.session_state.json_data
                st.success(f"Using previously converted JSON data with {len(json_data)} records")
            else:
                st.warning("No previously converted JSON data found. Please convert Excel files first or upload a JSON file.")
        
        if json_data:
            if st.button("Rearrange JSON"):
                with st.spinner("Rearranging JSON data..."):
                    # Rearrange JSON fields
                    rearranged_data = rearrange_json_fields(json_data)
                    
                    # Display sample of rearranged data
                    st.subheader("Sample of Rearranged Data (First 5 records)")
                    st.json(rearranged_data[:5] if len(rearranged_data) > 5 else rearranged_data)
                    
                    # Show total records
                    st.success(f"Successfully rearranged {len(rearranged_data)} records")
                    
                    # Provide download option
                    json_string = json.dumps(rearranged_data, indent=4)
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    
                    st.download_button(
                        label="Download Rearranged JSON",
                        data=json_string,
                        file_name=f"rearranged_staff_data_{timestamp}.json",
                        mime="application/json"
                    )
                    
                    # Also provide Excel download option
                    excel_data = json_to_excel(rearranged_data)
                    
                    st.download_button(
                        label="Download Rearranged Excel",
                        data=excel_data,
                        file_name=f"rearranged_staff_data_{timestamp}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    # Save rearranged data
                    st.session_state.rearranged_data = rearranged_data

    with tab3:
        st.header("Convert JSON to Excel")
        
        # Option to use previously converted JSON or upload a new JSON file
        json_source = st.radio(
            "JSON Source",
            ["Upload JSON file", "Use previously converted JSON"],
            index=1 if "json_data" in st.session_state else 0,
            key="json_source_excel"
        )
        
        json_data = None
        
        if json_source == "Upload JSON file":
            json_file = st.file_uploader("Upload JSON file", type=["json"], key="upload_json_excel")
            
            if json_file:
                try:
                    json_data = json.load(json_file)
                    st.success(f"JSON file loaded successfully with {len(json_data)} records")
                except Exception as e:
                    st.error(f"Error loading JSON file: {str(e)}")
        else:
            if "json_data" in st.session_state:
                json_data = st.session_state.json_data
                st.success(f"Using previously converted JSON data with {len(json_data)} records")
            else:
                st.warning("No previously converted JSON data found. Please convert Excel files first or upload a JSON file.")
        
        if json_data:
            # Show preview of JSON structure (first record)
            if len(json_data) > 0:
                st.subheader("Sample Record Preview")
                st.json(json_data[0])
            
            if st.button("Convert to Excel"):
                with st.spinner("Converting to Excel..."):
                    # Convert JSON to Excel
                    excel_data = json_to_excel(json_data)
                    
                    # Provide download option
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    
                    st.download_button(
                        label="Download Excel",
                        data=excel_data,
                        file_name=f"staff_data_{timestamp}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

if __name__ == "__main__":
    main()
    