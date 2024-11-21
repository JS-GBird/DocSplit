import shutil
import streamlit as st
from doc_splitter import split_document
import os
import tempfile
from pathlib import Path

def main():
    st.set_page_config(
        page_title="Document Splitter",
        page_icon="📄",
        layout="centered"
    )

    # Header
    st.title("📄 Document Splitter")
    st.markdown("Split your document into individual student files")
    
    # Create a container for logs that won't be overwritten
    log_area = st.container()
    with log_area:
        log_placeholder = st.empty()
        logs = []
        
        def add_log(message, type="info"):
            logs.append({"message": message, "type": type})
            log_text = "\n".join([f"{'❌' if log['type']=='error' else '✅' if log['type']=='success' else 'ℹ️'} {log['message']}" for log in logs])
            log_placeholder.text_area("Logs", log_text, height=200)
    
    # File upload
    uploaded_file = st.file_uploader(
        "Choose a Word document",
        type=['docx'],
        help="Upload the Word document you want to split"
    )

    if uploaded_file:
        st.write(f"Filename: {uploaded_file.name}")
        st.write(f"Size: {uploaded_file.size/1024:.2f} KB")

        # Process button
        if st.button("Split Document", type="primary"):
            try:
                # Clear previous logs
                logs.clear()
                add_log("Starting document processing...")
                
                # Create progress bar
                progress_bar = st.progress(0)
                
                # Create temporary directory for processing
                with tempfile.TemporaryDirectory() as temp_dir:
                    add_log(f"Created temporary directory: {temp_dir}")
                    
                    # Save uploaded file
                    temp_input = os.path.join(temp_dir, uploaded_file.name)
                    with open(temp_input, 'wb') as f:
                        f.write(uploaded_file.getvalue())
                    add_log(f"Saved uploaded file to: {temp_input}")
                    
                    # Create output directory
                    output_dir = os.path.join(temp_dir, 'split_files')
                    os.makedirs(output_dir, exist_ok=True)
                    
                    progress_bar.progress(25)
                    add_log("Starting document split...")
                    
                    # Process document with custom logging function
                    num_students = split_document(temp_input, output_dir, add_log)
                    add_log(f"Split completed. Found {num_students} student documents.")
                    progress_bar.progress(75)
                    
                    # Verify files were created
                    created_files = os.listdir(output_dir)
                    add_log(f"Files in output directory: {', '.join(created_files)}")
                    
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
                        add_log(f"Successfully split into {num_students} student documents!", "success")
                        st.success(f"Successfully split into {num_students} student documents!")
                        st.download_button(
                            label="📥 Download Split Documents",
                            data=zip_data,
                            file_name="split_documents.zip",
                            mime="application/zip"
                        )
                    else:
                        add_log("No student documents were created. Check if document has page breaks.", "warning")
                        st.warning("No student documents were created. Please check if the document contains multiple pages.")

            except Exception as e:
                add_log(f"Error occurred: {str(e)}", "error")
                st.error(f"An error occurred: {str(e)}")
                import traceback
                add_log(f"Full error: {traceback.format_exc()}", "error")

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