import logging
import os
from datetime import datetime

def setup_logging(log_level: str = "INFO", log_file: str = None) -> None:
    """Setup logging configuration for the PDF translation workflow"""
    
    # Create logs directory if it doesn't exist
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    # Generate log filename if not provided
    if log_file is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = os.path.join(log_dir, f"pdf_translation_{timestamp}.log")
    
    # Configure logging
    logging.basicConfig(
        level=getattr(logging, log_level.upper()),
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()  # Also log to console
        ]
    )
    
    # Set specific log levels for external libraries
    logging.getLogger('googletrans').setLevel(logging.WARNING)
    logging.getLogger('urllib3').setLevel(logging.WARNING)
    logging.getLogger('requests').setLevel(logging.WARNING)
    
    logger = logging.getLogger(__name__)
    logger.info(f"Logging setup complete. Log file: {log_file}")

class TranslationLogger:
    def __init__(self, name: str):
        self.logger = logging.getLogger(name)
    
    def log_extraction_start(self, pdf_path: str):
        self.logger.info(f"Starting text extraction from: {pdf_path}")
    
    def log_extraction_complete(self, page_count: int):
        self.logger.info(f"Text extraction complete. Total pages: {page_count}")
    
    def log_splitting_start(self, page_count: int):
        self.logger.info(f"Starting page splitting for {page_count} pages")
    
    def log_splitting_complete(self, chunk_count: int):
        self.logger.info(f"Page splitting complete. Created {chunk_count} chunks")
    
    def log_translation_start(self, chunk_count: int, service: str):
        self.logger.info(f"Starting translation of {chunk_count} chunks using {service}")
    
    def log_translation_progress(self, current: int, total: int, pages: str):
        self.logger.info(f"Translating chunk {current}/{total} (Pages: {pages})")
    
    def log_translation_complete(self, chunk_count: int):
        self.logger.info(f"Translation complete for {chunk_count} chunks")
    
    def log_merge_start(self, chunk_count: int):
        self.logger.info(f"Starting merge of {chunk_count} translated chunks")
    
    def log_merge_complete(self, page_count: int):
        self.logger.info(f"Merge complete. Final document has {page_count} pages")
    
    def log_word_generation_start(self, output_path: str):
        self.logger.info(f"Starting Word document generation: {output_path}")
    
    def log_word_generation_complete(self, output_path: str, file_size: int = None):
        size_info = f" ({file_size} bytes)" if file_size else ""
        self.logger.info(f"Word document generation complete: {output_path}{size_info}")
    
    def log_error(self, operation: str, error: Exception):
        self.logger.error(f"Error in {operation}: {str(error)}", exc_info=True)
    
    def log_warning(self, message: str):
        self.logger.warning(message)
    
    def log_workflow_start(self, pdf_path: str, settings: dict):
        self.logger.info("="*60)
        self.logger.info("PDF TRANSLATION WORKFLOW STARTED")
        self.logger.info("="*60)
        self.logger.info(f"Input PDF: {pdf_path}")
        self.logger.info(f"Translation service: {settings.get('translation_service', 'google')}")
        self.logger.info(f"Max chars per chunk: {settings.get('max_chars_per_chunk', 4000)}")
        self.logger.info(f"Max pages per chunk: {settings.get('max_pages_per_chunk', 5)}")
    
    def log_workflow_complete(self, output_path: str, duration: float):
        self.logger.info("="*60)
        self.logger.info("PDF TRANSLATION WORKFLOW COMPLETED")
        self.logger.info("="*60)
        self.logger.info(f"Output document: {output_path}")
        self.logger.info(f"Total duration: {duration:.2f} seconds")
        self.logger.info("="*60)