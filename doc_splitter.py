import logging
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_BREAK
import os
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def split_document(input_file_path, output_directory='split_documents', log_function=None):
    def update_status(message, type="info"):
        """Display log message in the UI and console"""
        print(f"[{type.upper()}] {message}")  # Always print to console
        if log_function:
            log_function(message, type)
    
    def get_page_content(page):
        """Get all content from a page including tables"""
        content = []
        for element in page:
            if hasattr(element, 'text'):  # Regular paragraph
                content.append(element.text.strip())
            elif hasattr(element, 'tables'):  # Table container
                for table in element.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            content.append(cell.text.strip())
            elif hasattr(element, 'rows'):  # Direct table
                for row in element.rows:
                    for cell in row.cells:
                        content.append(cell.text.strip())
        return '\n'.join(filter(None, content))
    
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
        update_status(f"Document has {len(doc.paragraphs)} paragraphs and {len(doc.tables)} tables")
        
        # Initialize variables
        all_pages = [[]]
        current_page_idx = 0
        
        update_status("Analyzing document structure...")
        
        # Process both paragraphs and tables
        elements = []
        current_table = None
        
        # Combine paragraphs and tables in order
        for element in doc.element.body:
            if element.tag.endswith('p'):
                if current_table:
                    elements.append(current_table)
                    current_table = None
                elements.append(doc.paragraphs[len([e for e in elements if hasattr(e, 'text')])])
            elif element.tag.endswith('tbl'):
                current_table = doc.tables[len([e for e in elements if hasattr(e, 'rows')])]
                
        # Add final table if exists
        if current_table:
            elements.append(current_table)
        
        # Process all elements
        for element in elements:
            # Add element to current page
            all_pages[current_page_idx].append(element)
            
            # Check for page breaks in paragraphs
            if hasattr(element, 'runs'):
                for run in element.runs:
                    element = run._element
                    for br in element.findall('.//w:br', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                        if br.get(qn('w:type')) == 'page':
                            update_status(f"Page break detected...")
                            current_page_idx += 1
                            all_pages.append([])
                            break
        
        update_status(f"Document split into {len(all_pages)} pages")
        
        # Debug information about pages including table content
        for idx, page in enumerate(all_pages):
            content = get_page_content(page)
            update_status(f"Page {idx + 1} content length: {len(content)} characters")
            if len(content) > 0:
                update_status(f"Page {idx + 1} preview: {content[:100]}...")
        
        # Remove empty pages - check both text and table content
        all_pages = [page for page in all_pages if get_page_content(page)]
        
        if not all_pages:
            update_status("No content found in document", type="error")
            raise ValueError("Document appears to be empty")
            
        update_status(f"Found {len(all_pages)} non-empty pages")
        
        # Create overview document (first page)
        update_status("Creating overview document...")
        overview_doc = Document()
        for element in all_pages[0]:
            if hasattr(element, 'text'):  # Paragraph
                if element.text.strip():
                    new_para = overview_doc.add_paragraph()
                    new_para.text = element.text
            elif hasattr(element, 'rows'):  # Table
                table = overview_doc.add_table(rows=len(element.rows), cols=len(element.rows[0].cells))
                for i, row in enumerate(element.rows):
                    for j, cell in enumerate(row.cells):
                        table.rows[i].cells[j].text = cell.text
        
        overview_path = os.path.join(output_directory, 'Overview.docx')
        overview_doc.save(overview_path)
        update_status(f"Saved overview document: {overview_path}", type="success")
        
        # Create individual student documents (remaining pages)
        student_count = 0
        for idx, page in enumerate(all_pages[1:], 1):
            # Skip truly empty pages
            if not get_page_content(page):
                update_status(f"Skipping empty page {idx}")
                continue
                
            student_count += 1
            update_status(f"Processing student document {student_count}...")
            student_doc = Document()
            
            # Add overview content
            update_status("Adding overview to student document...")
            for element in all_pages[0]:
                if hasattr(element, 'text'):  # Paragraph
                    if element.text.strip():
                        new_para = student_doc.add_paragraph()
                        new_para.text = element.text
                elif hasattr(element, 'rows'):  # Table
                    table = student_doc.add_table(rows=len(element.rows), cols=len(element.rows[0].cells))
                    for i, row in enumerate(element.rows):
                        for j, cell in enumerate(row.cells):
                            table.rows[i].cells[j].text = cell.text
            
            # Add page break after overview
            student_doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
            
            # Add student content
            update_status("Adding student content...")
            for element in page:
                if hasattr(element, 'text'):  # Paragraph
                    if element.text.strip():
                        new_para = student_doc.add_paragraph()
                        new_para.text = element.text
                elif hasattr(element, 'rows'):  # Table
                    table = student_doc.add_table(rows=len(element.rows), cols=len(element.rows[0].cells))
                    for i, row in enumerate(element.rows):
                        for j, cell in enumerate(row.cells):
                            table.rows[i].cells[j].text = cell.text
            
            # Save student document
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