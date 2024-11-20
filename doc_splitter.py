import logging
from docx import Document
from docx.enum.text import WD_BREAK
import os
import shutil

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
        # First, create the overview document by copying the original
        overview_path = os.path.join(output_directory, 'Overview.docx')
        shutil.copy2(input_file_path, overview_path)
        update_status("Created overview document")
        
        # Open the overview document and remove everything after first page break
        doc = Document(overview_path)
        
        # Find the first page break
        first_break_idx = None
        for i, para in enumerate(doc.paragraphs):
            for run in para.runs:
                if any(br.type == WD_BREAK.PAGE for br in run._element.br_lst):
                    first_break_idx = i
                    break
            if first_break_idx is not None:
                break
        
        # Remove everything after the first page break
        if first_break_idx is not None:
            for _ in range(len(doc.paragraphs) - first_break_idx - 1):
                p = doc.paragraphs[first_break_idx + 1]._element
                p.getparent().remove(p)
            doc.save(overview_path)
            update_status("Saved overview document with first page content")
        
        # Now create student documents
        student_count = 0
        current_student_doc = None
        current_paras = []
        
        # Open original document
        original_doc = Document(input_file_path)
        
        # Process paragraphs
        for i, para in enumerate(original_doc.paragraphs):
            # Check for page break
            has_page_break = False
            for run in para.runs:
                if any(br.type == WD_BREAK.PAGE for br in run._element.br_lst):
                    has_page_break = True
                    break
            
            if has_page_break or i == len(original_doc.paragraphs) - 1:
                # Save current student document if we have content
                if current_paras:
                    student_count += 1
                    student_path = os.path.join(output_directory, f'Student_{student_count}.docx')
                    # Copy original document as template
                    shutil.copy2(input_file_path, student_path)
                    update_status(f"Creating student document {student_count}...")
                    
                    # Open and modify the copy
                    student_doc = Document(student_path)
                    
                    # Keep only relevant paragraphs
                    while len(student_doc.paragraphs) > 0:
                        p = student_doc.paragraphs[0]._element
                        p.getparent().remove(p)
                    
                    # Add content
                    for p in current_paras:
                        student_doc.add_paragraph(p.text)
                    
                    student_doc.save(student_path)
                    update_status(f"Saved student document {student_count}")
                
                # Start new document
                current_paras = []
            else:
                current_paras.append(para)
        
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