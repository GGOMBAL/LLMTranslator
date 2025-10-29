#!/usr/bin/env python3
"""
Batch PDF Translation System
Efficiently processes all 70 pages with smart batching and error recovery
"""

import PyPDF2
import os
import json
import time
import re
from typing import List, Tuple, Optional

def extract_all_pages(pdf_path: str) -> List[Tuple[int, str]]:
    """Extract text from all PDF pages efficiently"""
    pages_text = []
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            total_pages = len(pdf_reader.pages)
            
            print(f"📖 Extracting from {total_pages} pages...")
            
            for i in range(total_pages):
                try:
                    page = pdf_reader.pages[i]
                    text = page.extract_text()
                    if text and text.strip():
                        pages_text.append((i + 1, text.strip()))
                    else:
                        pages_text.append((i + 1, f"[Page {i + 1} - No extractable text]"))
                    
                    # Progress indicator
                    if (i + 1) % 10 == 0:
                        print(f"   📄 Extracted {i + 1}/{total_pages} pages...")
                        
                except Exception as e:
                    print(f"   ⚠️  Error on page {i + 1}: {e}")
                    pages_text.append((i + 1, f"[Page {i + 1} - Extraction error]"))
                    
        print(f"✅ Successfully extracted {len(pages_text)} pages")
        return pages_text
        
    except Exception as e:
        print(f"❌ Error reading PDF: {str(e)}")
        return []

def smart_translate(text: str, page_num: int) -> str:
    """Smart translation with content analysis"""
    try:
        from googletrans import Translator
        translator = Translator()
        
        # Quick content analysis
        clean_text = ' '.join(text.split())
        
        # Skip if too short
        if len(clean_text) < 10:
            return f"[Page {page_num}: Content too short]"
        
        # Check if it's mostly formatting
        symbol_count = clean_text.count('.') + clean_text.count('-') + clean_text.count('_')
        if symbol_count > len(clean_text) * 0.3:
            # Extract Chinese content only
            chinese_pattern = r'[\u4e00-\u9fff\u3400-\u4dbf]+'
            chinese_matches = re.findall(chinese_pattern, clean_text)
            if chinese_matches:
                chinese_content = ' '.join(chinese_matches)[:500]  # Limit length
                try:
                    result = translator.translate(chinese_content, dest='en')
                    return f"[TOC Content] {result.text}" if result and result.text else "[TOC: Translation failed]"
                except:
                    return "[TOC: Translation failed]"
            return "[TOC: No translatable content]"
        
        # Regular translation
        text_to_translate = clean_text[:1000]  # Limit to 1000 chars for speed
        
        try:
            result = translator.translate(text_to_translate, dest='en')
            return result.text if result and result.text else "[Translation failed: Empty result]"
        except Exception as e:
            return f"[Translation failed: {str(e)}]"
            
    except ImportError:
        return "[Error: googletrans not installed]"
    except Exception as e:
        return f"[Error: {str(e)}]"

def process_in_batches(pages_text: List[Tuple[int, str]], batch_size: int = 10) -> List[dict]:
    """Process pages in batches for better error recovery"""
    results = []
    total_pages = len(pages_text)
    
    for i in range(0, total_pages, batch_size):
        batch = pages_text[i:i + batch_size]
        batch_start = i + 1
        batch_end = min(i + batch_size, total_pages)
        
        print(f"\n📦 Processing Batch {batch_start}-{batch_end} ({len(batch)} pages)...")
        
        for page_num, text in batch:
            print(f"   📋 Page {page_num}...", end=" ")
            
            start_time = time.time()
            translated = smart_translate(text, page_num)
            translation_time = time.time() - start_time
            
            status = "✅" if not translated.startswith('[') else "📋" if "TOC" in translated else "❌"
            print(f"{status} ({translation_time:.1f}s)")
            
            results.append({
                "page_number": page_num,
                "original_text": text,
                "translated_text": translated,
                "original_char_count": len(text),
                "translated_char_count": len(translated),
                "translation_time": round(translation_time, 2)
            })
            
            # Brief pause to avoid rate limiting
            time.sleep(0.5)
        
        # Longer pause between batches
        if batch_end < total_pages:
            print(f"   ⏸️  Batch completed. Pausing 3 seconds...")
            time.sleep(3)
    
    return results

def save_results(results: List[dict], output_file: str):
    """Save results to JSON file"""
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    successful = len([r for r in results if not r['translated_text'].startswith('[')])
    
    output_data = {
        "total_pages_processed": len(results),
        "successful_translations": successful,
        "timestamp": timestamp,
        "pages": results
    }
    
    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, ensure_ascii=False, indent=2)
        print(f"✅ Results saved to: {output_file}")
        return True
    except Exception as e:
        print(f"❌ Error saving results: {e}")
        return False

def main():
    """Main batch processing function"""
    pdf_file = "XY-A ATS开发对IBC需求文档_V0.0.pdf"
    output_file = "final_translation_results.json"
    
    print("🚀 Batch PDF Translation System (All Pages)")
    print("="*60)
    
    # Check PDF
    if not os.path.exists(pdf_file):
        print(f"❌ PDF not found: {pdf_file}")
        return
    
    print(f"✅ Found PDF: {pdf_file}")
    
    # Extract all pages
    pages_text = extract_all_pages(pdf_file)
    if not pages_text:
        print("❌ No text extracted")
        return
    
    print(f"\n🔄 Starting translation of {len(pages_text)} pages...")
    
    # Process in batches
    start_time = time.time()
    results = process_in_batches(pages_text, batch_size=10)
    total_time = time.time() - start_time
    
    # Save results
    if save_results(results, output_file):
        # Show summary
        successful = len([r for r in results if not r['translated_text'].startswith('[')])
        toc_pages = len([r for r in results if 'TOC' in r['translated_text']])
        failed = len(results) - successful - toc_pages
        
        print(f"\n📊 Final Results:")
        print(f"   📄 Total pages: {len(results)}")
        print(f"   ✅ Successfully translated: {successful}")
        print(f"   📋 TOC/Formatted pages: {toc_pages}")
        print(f"   ❌ Failed: {failed}")
        print(f"   ⏱️  Total time: {total_time/60:.1f} minutes")
        print(f"   📈 Success rate: {successful/len(results)*100:.1f}%")
        
        print(f"\n✅ Batch translation completed!")
        print(f"📁 Results saved in: {output_file}")

if __name__ == "__main__":
    main()