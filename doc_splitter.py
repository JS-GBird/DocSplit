import logging
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_BREAK
import os
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from copy import deepcopy

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
        
        # Initialize variables
        all_pages = [[]]
        current_page_idx = 0
        
        update_status("Analyzing document structure...")
        
        # Get all elements while preserving exact structure
        current_elements = []
        
        # Iterate through all elements in document body
        for element in doc._body._body:
            current_elements.append(element)
            
            # Check for page breaks
            if element.tag.endswith('p'):
                for child in element.iter():
                    if child.tag.endswith('br') and child.get(qn('w:type')) == 'page':
                        # Store current page elements
                        all_pages[current_page_idx] = current_elements.copy()
                        current_page_idx += 1
                        all_pages.append([])
                        current_elements = []
                        update_status(f"Page break detected - Page {current_page_idx}")
                        break
        
        # Add remaining elements to the last page
        if current_elements:
            all_pages[current_page_idx] = current_elements
        
        update_status(f"Document split into {len(all_pages)} pages")
        
        # Create overview document (first page)
        update_status("Creating overview document...")
        overview_doc = Document()
        overview_doc._body._body.clear()  # Clear default content
        
        # Copy section properties from original document
        overview_doc._body._body.append(doc._body._sectPr)
        
        # Copy first page content
        for element in all_pages[0]:
            overview_doc._body._body.append(element)
        
        overview_path = os.path.join(output_directory, 'Overview.docx')
        overview_doc.save(overview_path)
        update_status(f"Saved overview document: {overview_path}", type="success")
        
        # Create individual student documents
        student_count = 0
        for idx, page_elements in enumerate(all_pages[1:], 1):
            if not page_elements:
                update_status(f"Skipping empty page {idx}")
                continue
            
            update_status(f"Processing student document {idx}...")
            student_doc = Document()
            student_doc._body._body.clear()  # Clear default content
            
            # Copy section properties
            student_doc._body._body.append(doc._body._sectPr)
            
            # Copy overview content
            for element in all_pages[0]:
                student_doc._body._body.append(element)
            
            # Add page break
            page_break = OxmlElement('w:p')
            r = OxmlElement('w:r')
            br = OxmlElement('w:br')
            br.set(qn('w:type'), 'page')
            r.append(br)
            page_break.append(r)
            student_doc._body._body.append(page_break)
            
            # Add student content
            for element in page_elements:
                student_doc._body._body.append(element)
            
            # Save student document
            student_count += 1
            output_path = os.path.join(output_directory, f'Student_{student_count}.docx')
            student_doc.save(output_path)
            update_status(f"Saved student document: {output_path}", type="success")
        
        if student_count == 0:
            update_status("No student documents were created. Check if document has page breaks.", "warning")
        else:
            update_status(f"Created {student_count} student documents", "success")
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