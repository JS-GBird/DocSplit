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
        
        # Find all section breaks (page breaks)
        page_breaks = []
        current_length = 0
        
        update_status("Analyzing document structure...")
        
        # Iterate through paragraphs to find page breaks
        for para in doc.paragraphs:
            current_length += 1
            for run in para.runs:
                element = run._element
                for br in element.findall('.//w:br', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                    if br.get(qn('w:type')) == 'page':
                        page_breaks.append(current_length)
                        update_status(f"Found page break at position {current_length}")
        
        if not page_breaks:
            update_status("No page breaks found in document", "warning")
            return 0
            
        update_status(f"Found {len(page_breaks)} page breaks")
        
        # Create overview document (first page)
        update_status("Creating overview document...")
        overview_doc = Document()
        
        # Copy document styles by copying the styles XML
        overview_doc._element.styles = deepcopy(doc._element.styles)
        overview_doc._element.numbering = deepcopy(doc._element.numbering)
        
        # Copy content up to first page break
        end_idx = page_breaks[0] if page_breaks else len(doc.paragraphs)
        for element in doc.element.body[:end_idx]:
            overview_doc.element.body.append(deepcopy(element))
        
        overview_path = os.path.join(output_directory, 'Overview.docx')
        overview_doc.save(overview_path)
        update_status(f"Saved overview document: {overview_path}", type="success")
        
        # Create student documents
        student_count = 0
        start_idx = 0
        
        for i, break_idx in enumerate(page_breaks):
            # Skip if this would create an empty document
            if break_idx - start_idx <= 1:
                start_idx = break_idx
                continue
                
            student_count += 1
            update_status(f"Processing student document {student_count}...")
            student_doc = Document()
            
            # Copy document styles by copying the styles XML
            student_doc._element.styles = deepcopy(doc._element.styles)
            student_doc._element.numbering = deepcopy(doc._element.numbering)
            
            # Copy overview content
            update_status("Adding overview content...")
            for element in doc.element.body[:page_breaks[0]]:
                student_doc.element.body.append(deepcopy(element))
            
            # Add page break after overview
            student_doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
            
            # Copy student content
            update_status("Adding student content...")
            for element in doc.element.body[start_idx:break_idx]:
                student_doc.element.body.append(deepcopy(element))
            
            # Save student document
            output_path = os.path.join(output_directory, f'Student_{student_count}.docx')
            student_doc.save(output_path)
            update_status(f"Saved student document: {output_path}", type="success")
            
            start_idx = break_idx
        
        # Process the last section if it exists
        if start_idx < len(doc.element.body) - 1:
            student_count += 1
            update_status(f"Processing final student document {student_count}...")
            student_doc = Document()
            
            # Copy document styles by copying the styles XML
            student_doc._element.styles = deepcopy(doc._element.styles)
            student_doc._element.numbering = deepcopy(doc._element.numbering)
            
            # Copy overview content
            for element in doc.element.body[:page_breaks[0]]:
                student_doc.element.body.append(deepcopy(element))
            
            # Add page break after overview
            student_doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
            
            # Copy final student content
            for element in doc.element.body[start_idx:]:
                student_doc.element.body.append(deepcopy(element))
            
            output_path = os.path.join(output_directory, f'Student_{student_count}.docx')
            student_doc.save(output_path)
            update_status(f"Saved student document: {output_path}", type="success")
        
        # Add verification at the end
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