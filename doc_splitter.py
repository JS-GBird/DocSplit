import logging
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_BREAK
import os

def split_document(input_file_path, output_directory='split_documents', log_function=None):
    def update_status(message, type="info"):
        """Display log message in the UI and console"""
        print(f"[{type.upper()}] {message}")  # Always print to console
        if log_function:
            log_function(message, type)
    
    # Verify input file exists
    if not os.path.exists(input_file_path):
        error_msg = f"Input file not found: {input_file_path}"
        update_status(error_msg, "error")
        raise FileNotFoundError(error_msg)
    
    update_status(f"Found input file: {input_file_path}")
    update_status(f"File size: {os.path.getsize(input_file_path)/1024:.2f} KB")
    
    # Create output directory
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
        update_status(f"Created output directory: {output_directory}")
    
    try:
        # Load the input document
        update_status("Loading document...")
        doc = Document(input_file_path)
        
        # Verify document content
        update_status(f"Document has {len(doc.paragraphs)} paragraphs")
        
        # Initialize variables
        all_pages = [[]]
        current_page_idx = 0
        
        update_status("Analyzing document structure...")
        
        # Iterate through paragraphs to detect page breaks
        for para in doc.paragraphs:
            # Add paragraph to current page
            all_pages[current_page_idx].append(para)
            
            # Check for page breaks
            for run in para.runs:
                if run.element.br_lst:  # Check for break elements
                    for br in run.element.br_lst:
                        if br.type_val == WD_BREAK.PAGE:  # If it's a page break
                            update_status(f"Page break detected after paragraph: {para.text[:50]}...")
                            current_page_idx += 1
                            all_pages.append([])  # Start new page
                            break
        
        update_status(f"Document split into {len(all_pages)} pages")
        
        # Remove empty pages
        all_pages = [page for page in all_pages if page]
        
        if not all_pages:
            update_status("No content found in document", type="error")
            raise ValueError("Document appears to be empty")
            
        update_status(f"Found {len(all_pages)} non-empty pages")
        
        # Create overview document (first page)
        update_status("Creating overview document...")
        overview_doc = Document()
        for para in all_pages[0]:
            new_para = overview_doc.add_paragraph()
            new_para.text = para.text
        overview_path = os.path.join(output_directory, 'Overview.docx')
        overview_doc.save(overview_path)
        update_status(f"Saved overview document: {overview_path}", type="success")
        
        # Create individual student documents (remaining pages)
        student_count = 0
        for idx, page in enumerate(all_pages[1:], 1):
            # Skip empty pages
            if not page or all(not p.text.strip() for p in page):
                update_status(f"Skipping empty page {idx}")
                continue
                
            update_status(f"Processing student document {idx}...")
            student_doc = Document()
            
            # Add overview content
            update_status("Adding overview to student document...")
            for para in all_pages[0]:
                new_para = student_doc.add_paragraph()
                new_para.text = para.text
            
            # Add page break after overview
            student_doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
            
            # Add student content
            update_status("Adding student content...")
            for para in page:
                if para.text.strip():  # Only add non-empty paragraphs
                    new_para = student_doc.add_paragraph()
                    new_para.text = para.text
            
            # Save student document
            student_count += 1
            output_path = os.path.join(output_directory, f'Student_{student_count}.docx')
            student_doc.save(output_path)
            update_status(f"Saved student document: {output_path}", type="success")
        
        # Add verification at the end
        if student_count == 0:
            update_status("No student documents were created. Check if document has page breaks.", "warning")
        else:
            update_status(f"Created {student_count} student documents", "success")
            # List all created files
            created_files = os.listdir(output_directory)
            update_status(f"Files in output directory: {', '.join(created_files)}")
        
        return student_count
        
    except Exception as e:
        import traceback
        error_msg = f"Error processing document: {str(e)}\n{traceback.format_exc()}"
        update_status(error_msg, "error")
        raise

# Add a test function
def test_split_document(input_path):
    """Test function to verify document splitting"""
    print(f"Testing document split for: {input_path}")
    try:
        result = split_document(input_path, 'test_output')
        print(f"Test completed successfully. Created {result} documents")
        return True
    except Exception as e:
        print(f"Test failed: {str(e)}")
        return False