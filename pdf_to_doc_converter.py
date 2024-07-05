import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches

def extract_text_from_page(page):
    # Extract text from the PDF page
    text = page.get_text()
    
    # Handle any encoding issues if present
    return text.encode('utf-8').decode('utf-8')

def pdf_to_docx(pdf_file, docx_file):
    try:
        pdf_document = fitz.open(pdf_file)
        
        docx_document = Document()
        
        for page_number in range(len(pdf_document)):
            page = pdf_document.load_page(page_number)
            
            page_text = extract_text_from_page(page)
            
            docx_document.add_paragraph(page_text)
            
            if page_number < len(pdf_document) - 1:
                docx_document.add_page_break()
        
        docx_document.save(docx_file)
        print(f"PDF '{pdf_file}' converted to DOCX '{docx_file}' successfully.")
    
    except Exception as e:
        print(f"Error converting PDF to DOCX: {e}")
    
    finally:
        try:
            pdf_document.close()
        except:
            pass


pdf_file = "/home/weber/Desktop/pdf_to_docx_converter/internshipLetter.pdf"  
docx_file = "/home/weber/Desktop/pdf_to_docx_converter/internshipLetter.docx"  # Replace with desired output DOCX file path

pdf_to_docx(pdf_file, docx_file)
