#!/usr/bin/env python3
"""
개선된 PDF 번역기 V2 - 1단계 개선사항 적용
- 재시도 횟수: 2회 → 5회
- 타임아웃: 동적 조정 (텍스트 길이 기반)
- None 값 체크 강화
- 지수 백오프 적용
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
    """Safely clean text with comprehensive error handling - IMPROVED"""
    if not text or not isinstance(text, str):
        return ""

    try:
        # Handle potential encoding issues
        if isinstance(text, bytes):
            text = text.decode('utf-8', errors='ignore')

        # 🆕 IMPROVED: More robust None checking
        if text is None or str(text).strip() == 'None':
            return ""

        # Split and filter out None/empty values
        clean_parts = []
        for part in str(text).split():
            # 🆕 IMPROVED: Stricter validation
            if part is not None and isinstance(part, str) and part.strip() and part.strip() != 'None':
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
    """Extract only Chinese characters and related content - IMPROVED"""
    try:
        # 🆕 IMPROVED: None check
        if text is None or not isinstance(text, str):
            return ""

        # Pattern for Chinese characters, parentheses, and common punctuation
        chinese_pattern = r'[\u4e00-\u9fff\u3400-\u4dbf\（\）\uff08\uff09A-Za-z0-9\s]+'
        matches = re.findall(chinese_pattern, text)

        if matches:
            # 🆕 IMPROVED: Filter None values
            matches = [m for m in matches if m is not None and str(m).strip()]
            if not matches:
                return ""

            chinese_text = ' '.join(matches)
            # Clean up excessive spaces
            chinese_text = re.sub(r'\s+', ' ', chinese_text).strip()
            return chinese_text
        return ""
    except Exception as e:
        print(f"      Chinese extraction error: {e}")
        return ""

def calculate_dynamic_timeout(text_length: int) -> int:
    """🆕 NEW: Calculate timeout based on text length"""
    # Base timeout: 30 seconds
    # Add 10 seconds per 1000 characters
    # Maximum: 180 seconds (3 minutes)
    timeout = min(30 + (text_length // 100), 180)
    return timeout

def translate_with_google_robust(text: str) -> str:
    """🆕 IMPROVED: Robust translation with enhanced retry mechanism"""
    try:
        from googletrans import Translator

        # Initial validation
        if not text or not isinstance(text, str):
            return "[Error: Invalid input]"

        # 🆕 IMPROVED: Stronger None check
        if text is None or str(text).strip() == 'None':
            return "[Error: None value detected]"

        # Clean text safely
        clean_text = safe_text_cleaning(text)
        if not clean_text:
            return "[Error: No valid text after cleaning]"

        # Check text length
        if len(clean_text) < 5:
            return f"[Skipped: Too short - '{clean_text}']"

        # 🆕 NEW: Calculate dynamic timeout
        text_length = len(clean_text)
        timeout = calculate_dynamic_timeout(text_length)
        print(f"      Text length: {text_length} chars, Timeout: {timeout}s")

        # Handle heavily formatted text (TOC pages)
        symbol_count = clean_text.count('.') + clean_text.count('-') + clean_text.count('_')
        symbol_ratio = symbol_count / len(clean_text) if len(clean_text) > 0 else 0

        if symbol_ratio > 0.25:  # More than 25% symbols
            print(f"      Detected TOC/formatted content...")
            chinese_content = extract_chinese_content(clean_text)
            if chinese_content and len(chinese_content) > 10:
                print(f"      Translating Chinese content: {chinese_content[:50]}...")

                # 🆕 IMPROVED: 5 attempts with exponential backoff for TOC
                for attempt in range(5):
                    try:
                        translator = Translator()
                        result = translator.translate(chinese_content, dest='en')

                        # 🆕 IMPROVED: Stronger validation
                        if result and hasattr(result, 'text') and result.text and result.text is not None:
                            return f"[TOC] {result.text}"
                        else:
                            print(f"      TOC attempt {attempt + 1}: Empty result")
                    except Exception as e:
                        print(f"      TOC attempt {attempt + 1} failed: {e}")
                        if attempt < 4:  # Don't wait after last attempt
                            wait_time = 2 ** attempt  # Exponential backoff: 1, 2, 4, 8, 16 seconds
                            print(f"      Waiting {wait_time}s before retry...")
                            time.sleep(wait_time)

                return f"[TOC - Translation failed after 5 attempts]"
            return f"[Skipped: Mostly formatting - {clean_text[:50]}...]"

        # Regular translation for normal content
        if len(clean_text) > 2000:  # 🆕 IMPROVED: Increased limit from 1500 to 2000
            print(f"      Text too long ({len(clean_text)} chars), truncating to 2000...")
            clean_text = clean_text[:2000] + "..."

        print(f"      Translating: {clean_text[:50]}...")

        # 🆕 IMPROVED: 5 attempts with exponential backoff
        for attempt in range(5):
            try:
                # Create new translator instance each time
                translator = Translator()
                result = translator.translate(clean_text, dest='en')

                # 🆕 IMPROVED: Stronger result validation
                if result and hasattr(result, 'text') and result.text and result.text is not None and str(result.text).strip():
                    print(f"      ✅ Success on attempt {attempt + 1}")
                    return result.text
                else:
                    print(f"      Attempt {attempt + 1}: Empty or None result")
            except Exception as e:
                print(f"      Attempt {attempt + 1} failed: {e}")

            # 🆕 IMPROVED: Exponential backoff
            if attempt < 4:  # Don't wait after last attempt
                wait_time = 2 ** attempt  # 1, 2, 4, 8, 16 seconds
                print(f"      Waiting {wait_time}s before retry...")
                time.sleep(wait_time)

        return f"[Translation failed after 5 attempts]"

    except ImportError:
        return "[Error: googletrans not installed]"
    except Exception as e:
        return f"[Translation error: {str(e)}]"

def load_failed_pages_from_json(json_path: str) -> List[int]:
    """Load list of failed page numbers from previous translation"""
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        failed_pages = []
        for page in data['pages']:
            translated = page['translated_text']
            if '[Translation failed' in translated or '[TOC -' in translated or 'NoneType' in translated:
                failed_pages.append(page['page_number'])

        return sorted(failed_pages)
    except Exception as e:
        print(f"Error loading failed pages: {e}")
        return []

def create_outputs(results: List[dict], base_name: str = "improved_translation_v2"):
    """Create both JSON and CSV outputs"""
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")

    # JSON output
    json_data = {
        "total_pages_processed": len(results),
        "timestamp": timestamp,
        "successful_translations": len([r for r in results if not r['translated_text'].startswith('[')]),
        "pages": results
    }

    json_file = f"output/{base_name}.json"
    try:
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(json_data, f, ensure_ascii=False, indent=2)
        print(f"✅ JSON saved: {json_file}")
    except Exception as e:
        print(f"❌ JSON save error: {e}")

    # CSV output
    csv_file = f"output/{base_name}.csv"
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
    """Main translation workflow - Re-translate only failed pages"""
    pdf_file = "input/XY-A ATS开发对IBC需求文档_V0.0.pdf"
    previous_results_json = "output/final_translation_results.json"

    print("🚀 개선된 PDF 번역기 V2 - 1단계 개선사항 적용")
    print("="*80)
    print("✨ 개선사항:")
    print("   1. 재시도 횟수: 2회 → 5회")
    print("   2. 지수 백오프 적용: 1s, 2s, 4s, 8s, 16s")
    print("   3. 타임아웃 동적 조정: 30-180초 (텍스트 길이 기반)")
    print("   4. None 값 체크 강화")
    print("   5. 텍스트 길이 제한: 1500자 → 2000자")
    print("="*80)

    # Check PDF existence
    if not os.path.exists(pdf_file):
        print(f"❌ PDF not found: {pdf_file}")
        return

    print(f"\n✅ Found PDF: {pdf_file}")

    # Load failed pages from previous translation
    print(f"\n📂 Loading previous translation results...")
    failed_pages = load_failed_pages_from_json(previous_results_json)

    if not failed_pages:
        print("❌ No failed pages found or couldn't load previous results")
        print("   Will translate all pages instead...")
        failed_pages = None
    else:
        print(f"✅ Found {len(failed_pages)} failed pages to retry:")
        print(f"   Pages: {', '.join(map(str, failed_pages[:10]))}" +
              (f" ... (and {len(failed_pages)-10} more)" if len(failed_pages) > 10 else ""))

    # Extract text from PDF
    pages_text = extract_text_from_pdf(pdf_file, max_pages=None)

    if not pages_text:
        print("❌ No text extracted from PDF")
        return

    # Filter to only failed pages if available
    if failed_pages:
        pages_to_translate = [(num, text) for num, text in pages_text if num in failed_pages]
        print(f"\n📋 Retrying {len(pages_to_translate)} failed pages...")
    else:
        pages_to_translate = pages_text
        print(f"\n📋 Translating all {len(pages_to_translate)} pages...")

    # Process translations
    results = []
    total_pages = len(pages_to_translate)

    for i, (page_num, text) in enumerate(pages_to_translate, 1):
        print(f"\n{'='*80}")
        print(f"📋 Processing Page {page_num} ({i}/{total_pages})...")

        # Show sample of original
        sample = safe_text_cleaning(text)[:100]
        print(f"   📝 Original sample: {sample}...")

        # Translate
        print("   🔄 Translating with improved retry mechanism...")
        start_time = time.time()
        translated_text = translate_with_google_robust(text)
        translation_time = time.time() - start_time

        # Show result sample
        result_sample = translated_text[:100]
        print(f"   ✅ Result: {result_sample}...")
        print(f"   ⏱️  Time: {translation_time:.1f}s")

        # Determine status
        if translated_text.startswith('['):
            print(f"   ⚠️  Status: FAILED/PROCESSED")
        else:
            print(f"   ✅ Status: SUCCESS")

        results.append({
            "page_number": page_num,
            "original_text": text,
            "translated_text": translated_text,
            "original_char_count": len(text),
            "translated_char_count": len(translated_text),
            "translation_time": round(translation_time, 2)
        })

        # Rate limiting - longer wait between pages
        if i < total_pages:
            wait_time = 2  # 🆕 IMPROVED: Increased from 1.5s to 2s
            print(f"   ⏳ Waiting {wait_time}s before next page...")
            time.sleep(wait_time)

    # Save results
    print(f"\n{'='*80}")
    print(f"💾 Saving results...")
    create_outputs(results, "improved_translation_v2_results")

    # Summary
    successful = len([r for r in results if not r['translated_text'].startswith('[')])
    processed = len([r for r in results if r['translated_text'].startswith('[TOC]')])
    failed = len(results) - successful - processed

    print(f"\n📊 Final Summary:")
    print(f"   📄 Total pages processed: {len(results)}")
    print(f"   ✅ Successful translations: {successful} ({successful/len(results)*100:.1f}%)")
    print(f"   📋 Processed (TOC/formatted): {processed}")
    print(f"   ❌ Failed/skipped: {failed} ({failed/len(results)*100:.1f}%)")

    # Improvement comparison
    if failed_pages:
        previous_failed = len(failed_pages)
        new_successful = successful
        improvement = new_successful
        print(f"\n📈 Improvement Analysis:")
        print(f"   🔴 Previously failed: {previous_failed} pages")
        print(f"   🟢 Now successful: {new_successful} pages ({new_successful/previous_failed*100:.1f}%)")
        print(f"   📊 Improvement: {improvement} pages recovered")
        if failed > 0:
            print(f"   🔴 Still failing: {failed} pages ({failed/previous_failed*100:.1f}%)")

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
