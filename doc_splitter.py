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
    
    def extract_student_name(elements):
        """Extract student name from the first table in the elements"""
        for element in elements:
            if element.tag.endswith('tbl'):  # Found a table
                # Get all text elements in the table
                all_texts = element.findall('.//w:t', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                
                # First try to find "Student:" label
                for i, text_elem in enumerate(all_texts):
                    text = text_elem.text if text_elem.text else ""
                    if text.strip().lower() == "student:" or text.strip().lower() == "student":
                        # Get the next text element(s) which should contain the name
                        name_parts = []
                        j = i + 1
                        while j < len(all_texts) and not all_texts[j].text.strip().lower().startswith(("company", "course", "date")):
                            name_parts.append(all_texts[j].text.strip())
                            j += 1
                        if name_parts:
                            return " ".join(name_parts).strip()
                
                # If not found, try to find in combined cell text
                for row in element.findall('.//w:tr', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                    row_text = ""
                    for text_elem in row.findall('.//w:t', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                        if text_elem.text:
                            row_text += text_elem.text
                    
                    row_text = row_text.strip()
                    if "Student:" in row_text or "Student" in row_text:
                        # Extract name after "Student:" or "Student"
                        name = row_text.split(':', 1)[1].strip() if ':' in row_text else row_text.replace('Student', '').strip()
                        if name:
                            return name
        
        return None
    
    try:
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
                if current_page:
                    current_page.pop()  # Remove page break
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
        
        # Create individual student documents
        student_count = 0
        overview_content = all_pages[0]  # Store overview for reuse
        
        for idx, page_elements in enumerate(all_pages[1:], 1):
            # Extract student name from the page
            student_name = extract_student_name(page_elements)
            if not student_name:
                update_status(f"Warning: Could not find student name on page {idx}, using default name", "warning")
                student_name = f"Unknown_Student_{idx}"
            else:
                update_status(f"Found student: {student_name}")
            
            # Create new document from original to preserve styles
            student_doc = Document(input_file_path)
            student_doc._element.body.clear_content()
            
            # Copy overview (first page)
            for element in overview_content:
                new_element = deepcopy(element)
                student_doc._element.body.append(new_element)
            
            # Add page break
            p = OxmlElement('w:p')
            r = OxmlElement('w:r')
            br = OxmlElement('w:br')
            br.set(qn('w:type'), 'page')
            r.append(br)
            p.append(r)
            student_doc._element.body.append(p)
            
            # Add student content
            for element in page_elements:
                new_element = deepcopy(element)
                student_doc._element.body.append(new_element)
            
            # Save student document with their name
            student_count += 1
            safe_name = "".join(c for c in student_name if c.isalnum() or c in (' ', '_', '-')).strip()
            output_path = os.path.join(output_directory, f'{safe_name}.docx')
            student_doc.save(output_path)
            update_status(f"Saved document for: {student_name}", type="success")
        
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