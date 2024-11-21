import shutil
import streamlit as st
from doc_splitter import split_document
import os
import tempfile

def main():
    st.set_page_config(
        page_title="Document Splitter",
        page_icon="📄",
        layout="centered"
    )

    # Header
    st.title("📄 Document Splitter")
    st.markdown("Split your document into individual student files")
    
    # File upload
    uploaded_file = st.file_uploader(
        "Choose a Word document",
        type=['docx'],
        help="Upload the Word document you want to split"
    )

    if uploaded_file:
        st.markdown("#### File Details")
        st.write(f"Filename: {uploaded_file.name}")
        st.write(f"Size: {uploaded_file.size/1024:.2f} KB")

        # Process button
        if st.button("Split Document", type="primary"):
            try:            
                # Create progress bar
                progress_bar = st.progress(0)
                
                # Create temporary directory for processing
                with tempfile.TemporaryDirectory() as temp_dir:
                    
                    # Save uploaded file
                    temp_input = os.path.join(temp_dir, uploaded_file.name)
                    with open(temp_input, 'wb') as f:
                        f.write(uploaded_file.getvalue())
                    
                    # Create output directory
                    output_dir = os.path.join(temp_dir, 'split_files')
                    os.makedirs(output_dir, exist_ok=True)
                    
                    progress_bar.progress(25)
                    
                    # Process document with custom logging function
                    num_students = split_document(temp_input, output_dir)
                    progress_bar.progress(75)
                    
                    # Verify files were created
                    created_files = os.listdir(output_dir)
                    
                    # Create zip file
                    zip_path = os.path.join(temp_dir, 'split_documents.zip')
                    shutil.make_archive(
                        os.path.join(temp_dir, 'split_documents'),
                        'zip',
                        output_dir
                    )
                    
                    # Read zip file for download
                    with open(zip_path, 'rb') as f:
                        zip_data = f.read()
                    
                    progress_bar.progress(100)
                    
                    # Success message and download button
                    if num_students > 0:
                        st.success(f"Successfully split into {num_students} student documents!")
                        st.download_button(
                            label="📥 Download Split Documents",
                            data=zip_data,
                            file_name="split_documents.zip",
                            mime="application/zip"
                        )
                    else:
                        st.warning("No student documents were created. Please check if the document contains multiple pages.")

            except Exception as e:
                st.error(f"An error occurred: {str(e)}")
                import traceback

    # Instructions
    with st.expander("ℹ️ How to use"):
        st.markdown("""
        1. Upload a Word document (.docx) containing multiple pages
        2. Click 'Split Document' to process
        3. Download the ZIP file containing individual documents
        4. Each document will include:
           - The general overview from the first page
           - Each student's information on separate pages
        """)

if __name__ == "__main__":
    main() 