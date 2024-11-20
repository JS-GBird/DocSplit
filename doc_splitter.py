import logging
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import os

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def split_document(input_file_path, output_directory='split_documents'):
    logger.info(f"Starting document split for: {input_file_path}")
    
    # Create output directory if it doesn't exist
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
        logger.info(f"Created output directory: {output_directory}")
    
    try:
        # Load the input document
        logger.info("Loading document...")
        doc = Document(input_file_path)
        
        # Get all paragraphs
        paragraphs = doc.paragraphs
        logger.info(f"Total paragraphs found: {len(paragraphs)}")
        
        # Initialize variables for page tracking
        current_page = []
        all_pages = []
        estimated_chars_per_page = 3000  # Approximate characters per page
        current_chars = 0
        
        # Split into pages based on character count
        logger.info("Splitting document into pages...")
        for para in paragraphs:
            current_chars += len(para.text)
            current_page.append(para)
            
            # Check if we've reached a page break or enough characters for a page
            if current_chars >= estimated_chars_per_page or \
               any(run.break_type for run in para.runs):
                all_pages.append(current_page)
                current_page = []
                current_chars = 0
                logger.info(f"Page break detected, total pages so far: {len(all_pages)}")
        
        # Add any remaining paragraphs as the last page
        if current_page:
            all_pages.append(current_page)
        
        logger.info(f"Document split into {len(all_pages)} pages")
        
        # Create overview document (first page)
        if all_pages:
            logger.info("Creating overview document...")
            overview_doc = Document()
            for para in all_pages[0]:
                new_para = overview_doc.add_paragraph()
                new_para.text = para.text
            overview_path = os.path.join(output_directory, 'Overview.docx')
            overview_doc.save(overview_path)
            logger.info(f"Saved overview document: {overview_path}")
        
        # Create individual student documents (remaining pages)
        for idx, page in enumerate(all_pages[1:], 1):
            logger.info(f"Processing student document {idx}...")
            student_doc = Document()
            
            # Add overview content
            logger.info("Adding overview to student document...")
            for para in all_pages[0]:  # Add overview
                new_para = student_doc.add_paragraph()
                new_para.text = para.text
            
            # Add page break after overview
            student_doc.add_page_break()
            
            # Add student content
            logger.info("Adding student content...")
            for para in page:
                new_para = student_doc.add_paragraph()
                new_para.text = para.text
            
            # Save student document
            output_path = os.path.join(output_directory, f'Student_{idx}.docx')
            student_doc.save(output_path)
            logger.info(f"Saved student document: {output_path}")
        
        return len(all_pages) - 1  # Return number of student documents created
        
    except Exception as e:
        logger.error(f"Error processing document: {str(e)}", exc_info=True)
        raise

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