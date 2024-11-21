import logging
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_BREAK
import os
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from copy import deepcopy
from lxml import etree

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
        source_doc = Document(input_file_path)
        
        # Initialize variables for page collection
        all_pages = []
        current_page = []
        
        update_status("Analyzing document structure...")
        
        # Collect all elements while preserving their exact structure
        for element in source_doc._element.body:
            if element.tag.endswith('sectPr'):
                continue
                
            # Check for page breaks
            has_page_break = False
            if element.tag.endswith('p'):
                for child in element.iter():
                    if child.tag.endswith('br') and child.get(qn('w:type')) == 'page':
                        has_page_break = True
                        break
            
            # Add element to current page
            current_page.append(element)
            
            # If page break found, start new page
            if has_page_break:
                # Remove the page break element from the current page
                if current_page:
                    current_page.pop()
                all_pages.append(list(current_page))
                current_page = []
                update_status(f"Page break detected - Page {len(all_pages)}")
        
        # Add last page if it has content
        if current_page:
            all_pages.append(current_page)
        
        update_status(f"Document split into {len(all_pages)} pages")
        
        if not all_pages:
            update_status("No content found in document", "error")
            raise ValueError("Document appears to be empty")
        
        # Create overview document (first page)
        update_status("Creating overview document...")
        overview_doc = Document(input_file_path)  # Start with a copy of original
        overview_doc._element.body.clear_content()
        
        # Copy first page content
        for element in all_pages[0]:
            new_element = deepcopy(element)
            overview_doc._element.body.append(new_element)
        
        overview_path = os.path.join(output_directory, 'Overview.docx')
        overview_doc.save(overview_path)
        update_status(f"Saved overview document: {overview_path}", type="success")
        
        # Create individual student documents
        student_count = 0
        for idx, page_elements in enumerate(all_pages[1:], 1):
            update_status(f"Processing student document {idx}...")
            
            # Create new document from original to preserve styles
            student_doc = Document(input_file_path)
            student_doc._element.body.clear_content()
            
            # Copy overview (first page)
            for element in all_pages[0]:
                new_element = deepcopy(element)
                student_doc._element.body.append(new_element)
            
            # Add single page break
            student_doc._element.body.append(
                OxmlElement('w:p')
            ).append(
                OxmlElement('w:r')
            ).append(
                OxmlElement('w:br', {'w:type': 'page'})
            )
            
            # Add student content
            for element in page_elements:
                new_element = deepcopy(element)
                student_doc._element.body.append(new_element)
            
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