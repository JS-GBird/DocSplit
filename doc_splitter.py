from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import os

def split_document(input_file_path, output_directory='split_documents'):
    # Create output directory if it doesn't exist
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
    
    # Load the input document
    doc = Document(input_file_path)
    
    # Extract the overview (assuming it's the first page)
    overview_paragraphs = []
    current_page_length = 0
    
    for paragraph in doc.paragraphs:
        # Roughly estimate if we're still on the first page
        # (approximate calculation based on line length)
        current_page_length += len(paragraph.text) / 500
        if current_page_length > 1:
            break
        overview_paragraphs.append(paragraph)
    
    # Split remaining content into student sections
    student_sections = []
    current_section = []
    
    # Skip the overview paragraphs
    remaining_paragraphs = doc.paragraphs[len(overview_paragraphs):]
    
    for paragraph in remaining_paragraphs:
        # Assuming each student section starts with "Student" or similar identifier
        if paragraph.text.strip().lower().startswith(('student', 'name:', 'student:')):
            if current_section:
                student_sections.append(current_section)
            current_section = [paragraph]
        else:
            current_section.append(paragraph)
    
    # Add the last section
    if current_section:
        student_sections.append(current_section)
    
    # Create individual documents for each student
    for idx, student_section in enumerate(student_sections, 1):
        # Create new document
        new_doc = Document()
        
        # Add overview
        for para in overview_paragraphs:
            new_para = new_doc.add_paragraph()
            new_para.alignment = para.alignment
            new_para.style = para.style
            new_para.text = para.text
        
        # Add page break between overview and student section
        new_doc.add_page_break()
        
        # Add student section
        for para in student_section:
            new_para = new_doc.add_paragraph()
            new_para.alignment = para.alignment
            new_para.style = para.style
            new_para.text = para.text
        
        # Save the document
        output_path = os.path.join(output_directory, f'Student_{idx}.docx')
        new_doc.save(output_path)

def main():
    # Replace with your input file path
    input_file = 'input_document.docx'
    
    try:
        split_document(input_file)
        print("Documents have been successfully split!")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main() 