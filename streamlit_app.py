import streamlit as st
import pandas as pd
import os
import re
import zipfile
import io
import tempfile

def parse_question_updated(col):
    match = re.match(r'(\d+)\)\s*Q(\d+)(?:\.([a-zA-Z]))?', col)
    return (int(match.group(2)), match.group(3) if match.group(3) else '', int(match.group(1))) if match else (float('inf'), 'z', float('inf'))

def process_file(input_file, output_file):
    df = pd.read_excel(input_file)
   
    question_pattern = r'\d+\)\s*Q\d+(\.[a-zA-Z])?'
    sorted_question_columns = sorted(df.filter(regex=question_pattern).columns, key=parse_question_updated)
   
    value_columns = list(sorted_question_columns)
    if 'TotalObtainedScore' in df.columns:
        value_columns.append('TotalObtainedScore')
   
    id_columns = [col for col in df.columns if col not in value_columns]
   
    melted_df = pd.melt(df, id_vars=id_columns, value_vars=value_columns, var_name='Attribute', value_name='Value')
   
    def custom_sort(x):
        if x.name == 'Attribute':
            return pd.Series([parse_question_updated(val) for val in x])
        return x
    melted_df = melted_df.sort_values(id_columns + ['Attribute'], key=custom_sort)
   
    melted_df.to_excel(output_file, index=False)

def main():
    st.title("Exam File Processor")
    
    uploaded_files = st.file_uploader("Upload Excel files", type="xlsx", accept_multiple_files=True)
    
    if uploaded_files:
        progress_bar = st.progress(0)
        
        with tempfile.TemporaryDirectory() as tmpdirname:
            input_dir = os.path.join(tmpdirname, 'input')
            output_dir = os.path.join(tmpdirname, 'output')
            os.makedirs(input_dir, exist_ok=True)
            os.makedirs(output_dir, exist_ok=True)
            
            # Save uploaded files to input directory
            for file in uploaded_files:
                with open(os.path.join(input_dir, file.name), 'wb') as f:
                    f.write(file.getbuffer())
            
            # Process files
            for i, filename in enumerate(os.listdir(input_dir)):
                if filename.endswith('.xlsx'):
                    input_file = os.path.join(input_dir, filename)
                    output_file = os.path.join(output_dir, filename)
                    process_file(input_file, output_file)
                progress_bar.progress((i + 1) / len(uploaded_files))
            
            # Create zip file
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for root, dirs, files in os.walk(output_dir):
                    for file in files:
                        zip_file.write(os.path.join(root, file), file)
            
        st.success("All files processed successfully!")
        
        st.download_button(
            label="Download processed files",
            data=zip_buffer.getvalue(),
            file_name="processed_files.zip",
            mime="application/zip"
        )

if __name__ == "__main__":
    main()