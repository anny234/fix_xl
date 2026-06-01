import streamlit as st
import os

st.title("File Processor App")

# 1. File Upload Component
uploaded_file = st.file_uploader("Choose a file to process")

if uploaded_file is not None:
    # Read the file data
    file_bytes = uploaded_file.read()
    
    st.info("Processing your file...")
    
    # 2. YOUR PYTHON SCRIPT LOGIC GOES HERE
    # (Example: Just adding uppercase to text data, replace with your script)
    output_data = file_bytes.upper() 
    output_filename = f"processed_{uploaded_file.name}"
    
    st.success("Done!")
    
    # 3. File Download Component
    st.download_button(
        label="Download Processed File",
        data=output_data,
        file_name=output_filename
    )
