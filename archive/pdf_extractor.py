import pdfplumber
import PyPDF2
from typing import List, Tuple
import logging

class PDFExtractor:
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def extract_text_with_pdfplumber(self, pdf_path: str) -> List[Tuple[int, str]]:
        """Extract text from PDF using pdfplumber (better for complex layouts)"""
        pages_text = []
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for i, page in enumerate(pdf.pages):
                    text = page.extract_text()
                    if text:
                        pages_text.append((i + 1, text.strip()))
                    else:
                        self.logger.warning(f"No text extracted from page {i + 1}")
                        pages_text.append((i + 1, ""))
        except Exception as e:
            self.logger.error(f"Error extracting text with pdfplumber: {str(e)}")
            raise
        
        return pages_text
    
    def extract_text_with_pypdf2(self, pdf_path: str) -> List[Tuple[int, str]]:
        """Extract text from PDF using PyPDF2 (fallback method)"""
        pages_text = []
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                for i, page in enumerate(pdf_reader.pages):
                    text = page.extract_text()
                    pages_text.append((i + 1, text.strip() if text else ""))
        except Exception as e:
            self.logger.error(f"Error extracting text with PyPDF2: {str(e)}")
            raise
        
        return pages_text
    
    def extract_text(self, pdf_path: str, method: str = "pdfplumber") -> List[Tuple[int, str]]:
        """Extract text from PDF using specified method"""
        if method == "pdfplumber":
            return self.extract_text_with_pdfplumber(pdf_path)
        elif method == "pypdf2":
            return self.extract_text_with_pypdf2(pdf_path)
        else:
            raise ValueError("Method must be 'pdfplumber' or 'pypdf2'")
    
    def get_pdf_info(self, pdf_path: str) -> dict:
        """Get basic information about the PDF"""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                return {
                    "total_pages": len(pdf.pages),
                    "metadata": pdf.metadata
                }
        except Exception as e:
            self.logger.error(f"Error getting PDF info: {str(e)}")
            raise