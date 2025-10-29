from typing import List, Tuple, Dict
import math
import logging

class PageSplitter:
    def __init__(self, max_chars_per_chunk: int = 4000, max_pages_per_chunk: int = 5):
        self.max_chars_per_chunk = max_chars_per_chunk
        self.max_pages_per_chunk = max_pages_per_chunk
        self.logger = logging.getLogger(__name__)
    
    def split_by_character_count(self, pages_text: List[Tuple[int, str]]) -> List[Dict]:
        """Split pages based on character count to optimize translation API calls"""
        chunks = []
        current_chunk = {
            "pages": [],
            "total_chars": 0,
            "start_page": None,
            "end_page": None
        }
        
        for page_num, text in pages_text:
            text_length = len(text)
            
            # If adding this page exceeds the character limit, save current chunk
            if (current_chunk["total_chars"] + text_length > self.max_chars_per_chunk and 
                current_chunk["pages"]):
                
                current_chunk["end_page"] = current_chunk["pages"][-1][0]
                chunks.append(current_chunk)
                
                # Start new chunk
                current_chunk = {
                    "pages": [(page_num, text)],
                    "total_chars": text_length,
                    "start_page": page_num,
                    "end_page": page_num
                }
            else:
                # Add page to current chunk
                if not current_chunk["pages"]:
                    current_chunk["start_page"] = page_num
                
                current_chunk["pages"].append((page_num, text))
                current_chunk["total_chars"] += text_length
        
        # Don't forget the last chunk
        if current_chunk["pages"]:
            current_chunk["end_page"] = current_chunk["pages"][-1][0]
            chunks.append(current_chunk)
        
        self.logger.info(f"Split document into {len(chunks)} chunks")
        return chunks
    
    def split_by_page_count(self, pages_text: List[Tuple[int, str]]) -> List[Dict]:
        """Split pages based on maximum pages per chunk"""
        chunks = []
        total_pages = len(pages_text)
        
        for i in range(0, total_pages, self.max_pages_per_chunk):
            chunk_pages = pages_text[i:i + self.max_pages_per_chunk]
            chunk = {
                "pages": chunk_pages,
                "total_chars": sum(len(text) for _, text in chunk_pages),
                "start_page": chunk_pages[0][0],
                "end_page": chunk_pages[-1][0]
            }
            chunks.append(chunk)
        
        self.logger.info(f"Split document into {len(chunks)} chunks by page count")
        return chunks
    
    def split_intelligently(self, pages_text: List[Tuple[int, str]]) -> List[Dict]:
        """Smart splitting that considers both character count and page boundaries"""
        # First, try character-based splitting
        char_chunks = self.split_by_character_count(pages_text)
        
        # If any chunk has too many pages, further split by page count
        final_chunks = []
        for chunk in char_chunks:
            if len(chunk["pages"]) > self.max_pages_per_chunk:
                # Split this chunk further by page count
                page_based_chunks = self.split_by_page_count(chunk["pages"])
                final_chunks.extend(page_based_chunks)
            else:
                final_chunks.append(chunk)
        
        return final_chunks
    
    def get_chunk_summary(self, chunks: List[Dict]) -> str:
        """Generate a summary of the chunks"""
        summary = f"Document split into {len(chunks)} chunks:\n"
        for i, chunk in enumerate(chunks, 1):
            page_range = f"Pages {chunk['start_page']}-{chunk['end_page']}" if chunk['start_page'] != chunk['end_page'] else f"Page {chunk['start_page']}"
            summary += f"  Chunk {i}: {page_range} ({chunk['total_chars']} chars, {len(chunk['pages'])} pages)\n"
        return summary