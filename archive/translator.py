import time
import logging
from typing import List, Dict, Optional
from googletrans import Translator
import openai
import requests
from langdetect import detect
import json

class TranslationService:
    def __init__(self, service_type: str = "google", api_key: Optional[str] = None, 
                 claude_router_url: Optional[str] = None):
        self.service_type = service_type
        self.api_key = api_key
        self.claude_router_url = claude_router_url or "http://localhost:3000/v1/chat/completions"
        self.logger = logging.getLogger(__name__)
        
        if service_type == "google":
            self.translator = Translator()
        elif service_type == "openai" and api_key:
            openai.api_key = api_key
        elif service_type == "claude" and api_key:
            # Claude router setup - API key will be used for authentication
            pass
    
    def detect_language(self, text: str) -> str:
        """Detect the language of the text"""
        try:
            return detect(text)
        except:
            return "unknown"
    
    def translate_with_google(self, text: str, src_lang: str = "zh-cn", dest_lang: str = "en") -> str:
        """Translate text using Google Translate"""
        try:
            result = self.translator.translate(text, src=src_lang, dest=dest_lang)
            return result.text
        except Exception as e:
            self.logger.error(f"Google translation error: {str(e)}")
            raise
    
    def translate_with_openai(self, text: str, src_lang: str = "Chinese", dest_lang: str = "English") -> str:
        """Translate text using OpenAI GPT"""
        try:
            prompt = f"""Translate the following {src_lang} text to {dest_lang}. 
            Maintain the original formatting and structure as much as possible.
            Keep technical terms and proper nouns appropriately translated.
            
            Text to translate:
            {text}
            
            Translation:"""
            
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": prompt}],
                max_tokens=4000,
                temperature=0.3
            )
            
            return response.choices[0].message.content.strip()
        except Exception as e:
            self.logger.error(f"OpenAI translation error: {str(e)}")
            raise
    
    def translate_with_claude(self, text: str, src_lang: str = "Chinese", dest_lang: str = "English") -> str:
        """Translate text using Claude Sonnet 4.5 via claude-code-router"""
        try:
            prompt = f"""You are a professional translator specializing in technical documents. Translate the following {src_lang} text to {dest_lang} with these requirements:

1. Maintain the original formatting and structure exactly
2. Preserve technical terms and translate them appropriately
3. Keep proper nouns, company names, and product names as appropriate
4. Ensure the translation is natural and professional
5. Maintain any numbering, bullet points, or special formatting

Text to translate:
{text}

Provide only the translated text without any explanations or additional comments."""
            
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.api_key}" if self.api_key else None,
                "x-api-key": self.api_key if self.api_key else None
            }
            
            # Remove None headers
            headers = {k: v for k, v in headers.items() if v is not None}
            
            payload = {
                "model": "claude-sonnet-4.5",  # This will be routed by claude-code-router
                "messages": [
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                "max_tokens": 4000,
                "temperature": 0.1
            }
            
            response = requests.post(
                self.claude_router_url,
                headers=headers,
                json=payload,
                timeout=60
            )
            
            if response.status_code != 200:
                raise Exception(f"Claude API error: {response.status_code} - {response.text}")
            
            result = response.json()
            
            if "choices" in result and len(result["choices"]) > 0:
                return result["choices"][0]["message"]["content"].strip()
            else:
                raise Exception(f"Unexpected Claude API response format: {result}")
                
        except Exception as e:
            self.logger.error(f"Claude translation error: {str(e)}")
            raise
    
    def translate_chunk(self, chunk: Dict, src_lang: str = "zh-cn", dest_lang: str = "en") -> Dict:
        """Translate a chunk of pages"""
        translated_chunk = {
            "start_page": chunk["start_page"],
            "end_page": chunk["end_page"],
            "pages": []
        }
        
        for page_num, text in chunk["pages"]:
            if not text.strip():
                translated_chunk["pages"].append((page_num, ""))
                continue
            
            try:
                if self.service_type == "google":
                    # Split large text into smaller parts for Google Translate
                    if len(text) > 5000:
                        translated_text = self._translate_large_text_google(text, src_lang, dest_lang)
                    else:
                        translated_text = self.translate_with_google(text, src_lang, dest_lang)
                elif self.service_type == "openai":
                    translated_text = self.translate_with_openai(text, "Chinese", "English")
                elif self.service_type == "claude":
                    # Split large text for Claude if needed
                    if len(text) > 8000:
                        translated_text = self._translate_large_text_claude(text, "Chinese", "English")
                    else:
                        translated_text = self.translate_with_claude(text, "Chinese", "English")
                else:
                    raise ValueError(f"Unsupported translation service: {self.service_type}")
                
                translated_chunk["pages"].append((page_num, translated_text))
                
                # Add delay to avoid rate limiting
                time.sleep(1)
                
            except Exception as e:
                self.logger.error(f"Error translating page {page_num}: {str(e)}")
                translated_chunk["pages"].append((page_num, f"[Translation Error: {str(e)}]"))
        
        return translated_chunk
    
    def _translate_large_text_google(self, text: str, src_lang: str, dest_lang: str) -> str:
        """Split large text and translate in parts"""
        sentences = text.split('\n')
        translated_parts = []
        current_batch = []
        current_length = 0
        
        for sentence in sentences:
            if current_length + len(sentence) > 4000 and current_batch:
                # Translate current batch
                batch_text = '\n'.join(current_batch)
                translated_batch = self.translate_with_google(batch_text, src_lang, dest_lang)
                translated_parts.append(translated_batch)
                
                # Reset for next batch
                current_batch = [sentence]
                current_length = len(sentence)
                time.sleep(1)  # Rate limiting
            else:
                current_batch.append(sentence)
                current_length += len(sentence)
        
        # Don't forget the last batch
        if current_batch:
            batch_text = '\n'.join(current_batch)
            translated_batch = self.translate_with_google(batch_text, src_lang, dest_lang)
            translated_parts.append(translated_batch)
        
        return '\n'.join(translated_parts)
    
    def _translate_large_text_claude(self, text: str, src_lang: str, dest_lang: str) -> str:
        """Split large text and translate in parts for Claude"""
        sentences = text.split('\n')
        translated_parts = []
        current_batch = []
        current_length = 0
        
        for sentence in sentences:
            if current_length + len(sentence) > 7000 and current_batch:
                # Translate current batch
                batch_text = '\n'.join(current_batch)
                translated_batch = self.translate_with_claude(batch_text, src_lang, dest_lang)
                translated_parts.append(translated_batch)
                
                # Reset for next batch
                current_batch = [sentence]
                current_length = len(sentence)
                time.sleep(2)  # Rate limiting for Claude
            else:
                current_batch.append(sentence)
                current_length += len(sentence)
        
        # Don't forget the last batch
        if current_batch:
            batch_text = '\n'.join(current_batch)
            translated_batch = self.translate_with_claude(batch_text, src_lang, dest_lang)
            translated_parts.append(translated_batch)
        
        return '\n'.join(translated_parts)
    
    def translate_document(self, chunks: List[Dict], src_lang: str = "zh-cn", dest_lang: str = "en") -> List[Dict]:
        """Translate all chunks of a document"""
        translated_chunks = []
        
        for i, chunk in enumerate(chunks, 1):
            self.logger.info(f"Translating chunk {i}/{len(chunks)} (Pages {chunk['start_page']}-{chunk['end_page']})")
            translated_chunk = self.translate_chunk(chunk, src_lang, dest_lang)
            translated_chunks.append(translated_chunk)
        
        return translated_chunks