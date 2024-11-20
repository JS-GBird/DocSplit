import logging
from docx import Document
from docx.enum.text import WD_BREAK
import os
import shutil
from docx.oxml import OxmlElement

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
    
    # Create output directory
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
        update_status(f"Created output directory: {output_directory}")
    
    try:
        # Load the original document
        doc = Document(input_file_path)
        
        # Find all page breaks
        page_breaks = []
        for i, para in enumerate(doc.paragraphs):
            for run in para.runs:
                if any(br.type == WD_BREAK.PAGE for br in run._element.br_lst):
                    page_breaks.append(i)
                    update_status(f"Found page break at paragraph {i}")
        
        if not page_breaks:
            update_status("No page breaks found in document", "warning")
            return 0
        
        update_status(f"Found {len(page_breaks)} page breaks")
        
        # Get the overview content (first page)
        overview_end = page_breaks[0]
        
        # Create student documents
        student_count = 0
        
        # Process each student page
        for i in range(1, len(page_breaks), 2):  # Skip every other page break (student name pages)
            student_count += 1
            update_status(f"Processing student document {student_count}...")
            
            # Create new document for this student
            student_path = os.path.join(output_directory, f'Student_{student_count}.docx')
            shutil.copy2(input_file_path, student_path)
            
            # Open the copy and modify it
            student_doc = Document(student_path)
            
            # Keep only the overview and the student's page
            start_idx = page_breaks[i] if i < len(page_breaks) else None
            end_idx = page_breaks[i + 1] if i + 1 < len(page_breaks) else None
            
            # Remove everything except overview and student page
            paragraphs_to_keep = list(range(0, overview_end + 1))  # Overview
            if start_idx is not None and end_idx is not None:
                paragraphs_to_keep.extend(range(start_idx, end_idx + 1))  # Student page
            elif start_idx is not None:
                paragraphs_to_keep.extend(range(start_idx, len(student_doc.paragraphs)))  # Until end
            
            # Remove paragraphs not in our keep list
            for idx in range(len(student_doc.paragraphs) - 1, -1, -1):
                if idx not in paragraphs_to_keep:
                    p = student_doc.paragraphs[idx]._element
                    p.getparent().remove(p)
            
            student_doc.save(student_path)
            update_status(f"Saved student document {student_count}")
        
        # Create overview document
        overview_path = os.path.join(output_directory, 'Overview.docx')
        shutil.copy2(input_file_path, overview_path)
        overview_doc = Document(overview_path)
        
        # Keep only the overview page
        for idx in range(len(overview_doc.paragraphs) - 1, overview_end, -1):
            p = overview_doc.paragraphs[idx]._element
            p.getparent().remove(p)
        
        overview_doc.save(overview_path)
        update_status("Saved overview document")
        
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