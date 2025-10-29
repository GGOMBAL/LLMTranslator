from typing import List, Dict, Tuple
import logging

class DocumentMerger:
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def merge_translated_chunks(self, translated_chunks: List[Dict]) -> List[Tuple[int, str]]:
        """Merge translated chunks back into a single document structure"""
        merged_pages = []
        
        try:
            # Sort chunks by start page to ensure correct order
            sorted_chunks = sorted(translated_chunks, key=lambda x: x['start_page'])
            
            for chunk in sorted_chunks:
                # Sort pages within each chunk
                sorted_pages = sorted(chunk['pages'], key=lambda x: x[0])
                merged_pages.extend(sorted_pages)
            
            # Verify page continuity
            self._verify_page_continuity(merged_pages)
            
            self.logger.info(f"Successfully merged {len(translated_chunks)} chunks into {len(merged_pages)} pages")
            return merged_pages
            
        except Exception as e:
            self.logger.error(f"Error merging translated chunks: {str(e)}")
            raise
    
    def merge_original_chunks(self, original_chunks: List[Dict]) -> List[Tuple[int, str]]:
        """Merge original chunks back into a single document structure"""
        merged_pages = []
        
        try:
            # Sort chunks by start page
            sorted_chunks = sorted(original_chunks, key=lambda x: x['start_page'])
            
            for chunk in sorted_chunks:
                # Sort pages within each chunk
                sorted_pages = sorted(chunk['pages'], key=lambda x: x[0])
                merged_pages.extend(sorted_pages)
            
            # Verify page continuity
            self._verify_page_continuity(merged_pages)
            
            self.logger.info(f"Successfully merged {len(original_chunks)} original chunks into {len(merged_pages)} pages")
            return merged_pages
            
        except Exception as e:
            self.logger.error(f"Error merging original chunks: {str(e)}")
            raise
    
    def _verify_page_continuity(self, pages: List[Tuple[int, str]]) -> None:
        """Verify that pages are in correct order and no pages are missing"""
        if not pages:
            return
        
        page_numbers = [page_num for page_num, _ in pages]
        
        # Check for duplicates
        if len(page_numbers) != len(set(page_numbers)):
            duplicates = [num for num in set(page_numbers) if page_numbers.count(num) > 1]
            self.logger.warning(f"Duplicate pages found: {duplicates}")
        
        # Check for missing pages
        min_page = min(page_numbers)
        max_page = max(page_numbers)
        expected_pages = set(range(min_page, max_page + 1))
        actual_pages = set(page_numbers)
        missing_pages = expected_pages - actual_pages
        
        if missing_pages:
            self.logger.warning(f"Missing pages: {sorted(missing_pages)}")
        
        # Check if pages are in order
        if page_numbers != sorted(page_numbers):
            self.logger.warning("Pages are not in sequential order")
    
    def create_translation_comparison(self, original_chunks: List[Dict], 
                                     translated_chunks: List[Dict]) -> List[Dict]:
        """Create a comparison structure with original and translated text side by side"""
        comparison_data = []
        
        try:
            # Ensure both lists have the same length
            if len(original_chunks) != len(translated_chunks):
                raise ValueError("Original and translated chunks must have the same length")
            
            for orig_chunk, trans_chunk in zip(original_chunks, translated_chunks):
                # Verify chunk alignment
                if (orig_chunk['start_page'] != trans_chunk['start_page'] or 
                    orig_chunk['end_page'] != trans_chunk['end_page']):
                    self.logger.warning(f"Chunk alignment mismatch: Original {orig_chunk['start_page']}-{orig_chunk['end_page']} vs Translated {trans_chunk['start_page']}-{trans_chunk['end_page']}")
                
                chunk_comparison = {
                    'start_page': orig_chunk['start_page'],
                    'end_page': orig_chunk['end_page'],
                    'pages': []
                }
                
                # Align pages within chunks
                orig_pages_dict = {page_num: text for page_num, text in orig_chunk['pages']}
                trans_pages_dict = {page_num: text for page_num, text in trans_chunk['pages']}
                
                all_page_nums = sorted(set(orig_pages_dict.keys()) | set(trans_pages_dict.keys()))
                
                for page_num in all_page_nums:
                    page_comparison = {
                        'page_number': page_num,
                        'original': orig_pages_dict.get(page_num, ""),
                        'translated': trans_pages_dict.get(page_num, "")
                    }
                    chunk_comparison['pages'].append(page_comparison)
                
                comparison_data.append(chunk_comparison)
            
            self.logger.info(f"Created comparison data for {len(comparison_data)} chunks")
            return comparison_data
            
        except Exception as e:
            self.logger.error(f"Error creating translation comparison: {str(e)}")
            raise
    
    def get_merge_statistics(self, chunks: List[Dict]) -> Dict:
        """Get statistics about the merged document"""
        if not chunks:
            return {}
        
        total_pages = sum(len(chunk['pages']) for chunk in chunks)
        total_chars = sum(
            sum(len(text) for _, text in chunk['pages']) 
            for chunk in chunks
        )
        
        page_numbers = []
        for chunk in chunks:
            page_numbers.extend([page_num for page_num, _ in chunk['pages']])
        
        return {
            'total_chunks': len(chunks),
            'total_pages': total_pages,
            'total_characters': total_chars,
            'page_range': f"{min(page_numbers)}-{max(page_numbers)}" if page_numbers else "0-0",
            'average_chars_per_page': total_chars // total_pages if total_pages > 0 else 0
        }