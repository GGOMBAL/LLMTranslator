#!/usr/bin/env python3
"""
Final Production PDF Translation System
Robust Chinese to English PDF translator with comprehensive error handling
"""

import PyPDF2
import os
import json
import time
import re
import csv
from typing import List, Tuple, Optional

def extract_text_from_pdf(pdf_path: str, max_pages: Optional[int] = None) -> List[Tuple[int, str]]:
    """Extract text from PDF using PyPDF2"""
    pages_text = []
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            total_pages = len(pdf_reader.pages)
            
            if max_pages:
                total_pages = min(total_pages, max_pages)
            
            print(f"📖 Extracting from {total_pages} pages...")
            
            for i in range(total_pages):
                try:
                    page = pdf_reader.pages[i]
                    text = page.extract_text()
                    if text and text.strip():
                        pages_text.append((i + 1, text.strip()))
                    else:
                        pages_text.append((i + 1, f"[Page {i + 1} - No extractable text]"))
                except Exception as e:
                    print(f"   ⚠️  Error on page {i + 1}: {e}")
                    pages_text.append((i + 1, f"[Page {i + 1} - Extraction error: {str(e)}]"))
                    
        print(f"✅ Successfully processed {len(pages_text)} pages")
        return pages_text
        
    except Exception as e:
        print(f"❌ Error reading PDF: {str(e)}")
        return []

def safe_text_cleaning(text: str) -> str:
    """Safely clean text with comprehensive error handling"""
    if not text or not isinstance(text, str):
        return ""
    
    try:
        # Handle potential encoding issues
        if isinstance(text, bytes):
            text = text.decode('utf-8', errors='ignore')
        
        # Split and filter out None/empty values
        clean_parts = []
        for part in str(text).split():
            if part and isinstance(part, str) and part.strip():
                clean_parts.append(part.strip())
        
        if not clean_parts:
            return ""
        
        # Join safely
        clean_text = ' '.join(clean_parts)
        
        # Remove excessive formatting
        clean_text = re.sub(r'\.{3,}', '...', clean_text)
        clean_text = re.sub(r'\s+', ' ', clean_text)
        clean_text = clean_text.strip()
        
        return clean_text
        
    except Exception as e:
        print(f"      Text cleaning error: {e}")
        return ""

def extract_chinese_content(text: str) -> str:
    """Extract only Chinese characters and related content"""
    try:
        # Pattern for Chinese characters, parentheses, and common punctuation
        chinese_pattern = r'[\u4e00-\u9fff\u3400-\u4dbf\（\）\uff08\uff09A-Za-z0-9\s]+'
        matches = re.findall(chinese_pattern, text)
        
        if matches:
            chinese_text = ' '.join(matches)
            # Clean up excessive spaces
            chinese_text = re.sub(r'\s+', ' ', chinese_text).strip()
            return chinese_text
        return ""
    except Exception:
        return ""

def translate_with_google_robust(text: str) -> str:
    """Robust translation with multiple fallback strategies"""
    try:
        from googletrans import Translator
        translator = Translator()
        
        # Initial validation
        if not text or not isinstance(text, str):
            return "[Error: Invalid input]"
        
        # Clean text safely
        clean_text = safe_text_cleaning(text)
        if not clean_text:
            return "[Error: No valid text after cleaning]"
        
        # Check text length
        if len(clean_text) < 5:
            return f"[Skipped: Too short - '{clean_text}']"
        
        # Handle heavily formatted text (TOC pages)
        symbol_count = clean_text.count('.') + clean_text.count('-') + clean_text.count('_')
        symbol_ratio = symbol_count / len(clean_text) if len(clean_text) > 0 else 0
        
        if symbol_ratio > 0.25:  # More than 25% symbols
            print(f"      Detected TOC/formatted content...")
            chinese_content = extract_chinese_content(clean_text)
            if chinese_content and len(chinese_content) > 10:
                print(f"      Translating Chinese content: {chinese_content[:50]}...")
                try:
                    result = translator.translate(chinese_content, dest='en')
                    if result and hasattr(result, 'text') and result.text:
                        return f"[TOC] {result.text}"
                except Exception as e:
                    print(f"      TOC translation failed: {e}")
                    return f"[TOC - Translation failed: {str(e)}]"
            return f"[Skipped: Mostly formatting - {clean_text[:50]}...]"
        
        # Regular translation for normal content
        if len(clean_text) > 1500:
            clean_text = clean_text[:1500] + "..."
        
        print(f"      Translating: {clean_text[:50]}...")
        
        # Multiple translation attempts
        for attempt in range(2):
            try:
                result = translator.translate(clean_text, dest='en')
                if result and hasattr(result, 'text') and result.text:
                    return result.text
                else:
                    print(f"      Attempt {attempt + 1}: Empty result")
            except Exception as e:
                print(f"      Attempt {attempt + 1} failed: {e}")
                if attempt == 0:
                    time.sleep(1)  # Brief pause before retry
        
        return f"[Translation failed after 2 attempts]"
        
    except ImportError:
        return "[Error: googletrans not installed]"
    except Exception as e:
        return f"[Translation error: {str(e)}]"

def create_outputs(results: List[dict], base_name: str = "final_translation"):
    """Create both JSON and CSV outputs"""
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    
    # JSON output
    json_data = {
        "total_pages_processed": len(results),
        "timestamp": timestamp,
        "successful_translations": len([r for r in results if not r['translated_text'].startswith('[')]),
        "pages": results
    }
    
    json_file = f"{base_name}.json"
    try:
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(json_data, f, ensure_ascii=False, indent=2)
        print(f"✅ JSON saved: {json_file}")
    except Exception as e:
        print(f"❌ JSON save error: {e}")
    
    # CSV output
    csv_file = f"{base_name}.csv"
    try:
        with open(csv_file, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['Page', 'Original_Sample', 'Translation', 'Original_Length', 'Translation_Length', 'Status']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            writer.writeheader()
            for result in results:
                status = "Success" if not result['translated_text'].startswith('[') else "Processed"
                writer.writerow({
                    'Page': result['page_number'],
                    'Original_Sample': result['original_text'][:200] + '...' if len(result['original_text']) > 200 else result['original_text'],
                    'Translation': result['translated_text'][:200] + '...' if len(result['translated_text']) > 200 else result['translated_text'],
                    'Original_Length': result['original_char_count'],
                    'Translation_Length': result['translated_char_count'],
                    'Status': status
                })
        print(f"✅ CSV saved: {csv_file}")
    except Exception as e:
        print(f"❌ CSV save error: {e}")

def main():
    """Main translation workflow"""
    pdf_file = "XY-A ATS开发对IBC需求文档_V0.0.pdf"
    
    print("🚀 Final PDF Translation System")
    print("="*60)
    
    # Check PDF existence
    if not os.path.exists(pdf_file):
        print(f"❌ PDF not found: {pdf_file}")
        pdf_files = [f for f in os.listdir('.') if f.endswith('.pdf')]
        if pdf_files:
            print("📁 Available PDF files:")
            for f in pdf_files:
                print(f"   - {f}")
        return
    
    print(f"✅ Found PDF: {pdf_file}")
    
    # Extract text from all pages
    pages_text = extract_text_from_pdf(pdf_file, max_pages=None)
    
    if not pages_text:
        print("❌ No text extracted from PDF")
        return
    
    # Process translations
    results = []
    total_pages = len(pages_text)
    
    for i, (page_num, text) in enumerate(pages_text, 1):
        print(f"\n📋 Processing Page {page_num} ({i}/{total_pages})...")
        
        # Show sample of original
        sample = safe_text_cleaning(text)[:100]
        print(f"   📝 Original sample: {sample}...")
        
        # Translate
        print("   🔄 Translating...")
        start_time = time.time()
        translated_text = translate_with_google_robust(text)
        translation_time = time.time() - start_time
        
        # Show result sample
        result_sample = translated_text[:100]
        print(f"   ✅ Result: {result_sample}...")
        print(f"   ⏱️  Time: {translation_time:.1f}s")
        
        results.append({
            "page_number": page_num,
            "original_text": text,
            "translated_text": translated_text,
            "original_char_count": len(text),
            "translated_char_count": len(translated_text),
            "translation_time": round(translation_time, 2)
        })
        
        # Rate limiting
        if i < total_pages:  # Don't wait after last page
            time.sleep(1.5)
    
    # Save results
    print(f"\n💾 Saving results...")
    create_outputs(results, "final_translation_results")
    
    # Summary
    successful = len([r for r in results if not r['translated_text'].startswith('[')])
    processed = len([r for r in results if r['translated_text'].startswith('[TOC]')])
    failed = len(results) - successful - processed
    
    print(f"\n📊 Final Summary:")
    print(f"   📄 Total pages: {len(results)}")
    print(f"   ✅ Successful translations: {successful}")
    print(f"   📋 Processed (TOC/formatted): {processed}")
    print(f"   ❌ Failed/skipped: {failed}")
    
    # Show best samples
    if successful > 0:
        best_results = [r for r in results if not r['translated_text'].startswith('[')]
        if best_results:
            sample = best_results[0]
            print(f"\n🏆 Sample Translation (Page {sample['page_number']}):")
            print("🇨🇳 Original:")
            print("  " + sample['original_text'][:150].replace('\n', ' ') + "...")
            print("🇺🇸 Translation:")
            print("  " + sample['translated_text'][:150] + "...")
    
    print("\n✅ Translation completed successfully!")
    print("📁 Check the output files for complete results.")

if __name__ == "__main__":
    main()