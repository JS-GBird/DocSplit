import shutil
import streamlit as st
from doc_splitter import split_document
import os
import tempfile
from pathlib import Path
import logging
from io import StringIO
import sys

# Modify the logging configuration to ensure it's working
log_stream = StringIO()
logging.basicConfig(
    level=logging.DEBUG,  # Changed to DEBUG to show all logs
    format='%(asctime)s - %(levelname)s - %(message)s',
    force=True  # Force reconfiguration of the logger
)

# Add a file handler to write logs to a file
file_handler = logging.FileHandler('app.log')
file_handler.setLevel(logging.DEBUG)
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

# Add a stream handler for the StringIO buffer
stream_handler = logging.StreamHandler(log_stream)
stream_handler.setLevel(logging.DEBUG)
stream_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

# Get the logger and add the handlers
logger = logging.getLogger(__name__)
logger.addHandler(file_handler)
logger.addHandler(stream_handler)

def main():
    logger.info("Application started")
    st.set_page_config(
        page_title="Document Splitter",
        page_icon="📄",
        layout="centered"
    )

    # Custom CSS
    st.markdown("""
        <style>
        .stApp {
            max-width: 800px;
            margin: 0 auto;
        }
        .main {
            padding: 2rem;
        }
        .uploadedFile {
            margin: 2rem 0;
        }
        .success {
            color: #28a745;
        }
        .error {
            color: #dc3545;
        }
        </style>
    """, unsafe_allow_html=True)

    # Header
    st.title("📄 Document Splitter")
    st.markdown("Split your Word document into individual files")
    
    # Add this before the file upload section
    with st.expander("📋 View Logs", expanded=False):
        if st.button("Clear Logs"):
            log_stream.truncate(0)
            log_stream.seek(0)
        
        # Display logs
        logs = log_stream.getvalue()
        if logs:
            st.text_area("Application Logs", logs, height=300)
        else:
            st.info("No logs yet.")

    # File upload
    uploaded_file = st.file_uploader(
        "Choose a Word document",
        type=['docx'],
        help="Upload the Word document you want to split"
    )

    if uploaded_file:
        logger.info("File uploaded successfully")
        st.markdown("### File Details")
        st.write(f"Filename: {uploaded_file.name}")
        st.write(f"Size: {uploaded_file.size/1024:.2f} KB")

        # Process button
        if st.button("Split Document", type="primary"):
            try:
                logger.info(f"Starting processing for file: {uploaded_file.name}")
                
                # Create progress bar
                progress_bar = st.progress(0)
                status = st.empty()
                
                # Create temporary directory for processing
                with tempfile.TemporaryDirectory() as temp_dir:
                    logger.info(f"Created temporary directory: {temp_dir}")
                    
                    # Save uploaded file
                    temp_input = os.path.join(temp_dir, uploaded_file.name)
                    with open(temp_input, 'wb') as f:
                        f.write(uploaded_file.getvalue())
                    logger.info(f"Saved uploaded file to: {temp_input}")
                    
                    # Create output directory
                    output_dir = os.path.join(temp_dir, 'split_files')
                    os.makedirs(output_dir, exist_ok=True)
                    logger.info(f"Created output directory: {output_dir}")
                    
                    # Update status
                    status.info("Processing document...")
                    progress_bar.progress(25)
                    
                    # Process document
                    logger.info("Starting document splitting...")
                    num_students = split_document(temp_input, output_dir)
                    logger.info(f"Document split into {num_students} student files")
                    progress_bar.progress(75)
                    
                    # Create zip file
                    logger.info("Creating zip file...")
                    zip_path = os.path.join(temp_dir, 'split_documents.zip')
                    shutil.make_archive(
                        os.path.join(temp_dir, 'split_documents'),
                        'zip',
                        output_dir
                    )
                    
                    # Read zip file for download
                    with open(zip_path, 'rb') as f:
                        zip_data = f.read()
                    logger.info("Zip file created successfully")
                    
                    progress_bar.progress(100)
                    
                    # Success message and download button
                    st.success(f"Successfully split into {num_students} student documents!")
                    st.download_button(
                        label="📥 Download Split Documents",
                        data=zip_data,
                        file_name="split_documents.zip",
                        mime="application/zip"
                    )

            except Exception as e:
                logger.error(f"Error processing document: {str(e)}", exc_info=True)
                st.error(f"An error occurred: {str(e)}")
                
    # Instructions
    with st.expander("ℹ️ How to use"):
        st.markdown("""
        1. Upload a Word document (.docx) containing multiple pages
        2. Click 'Split Document' to process
        3. Download the ZIP file containing individual documents
        4. Each document will include, separately:
           - The general overview from the first page
           - Each page that follows
        """)

if __name__ == "__main__":
    main() 