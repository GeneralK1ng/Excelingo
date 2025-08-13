import asyncio
import aiohttp
from typing import List, Dict, Any
from dotenv import load_dotenv
import os

load_dotenv()

class DeepSeekClient:
    def __init__(self):
        self.api_key = os.getenv("DEEPSEEK_API_KEY")
        self.base_url = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com/v1")
        self.session = None
        
    async def __aenter__(self):
        self.session = aiohttp.ClientSession()
        return self
        
    async def __aexit__(self, exc_type, exc_val, exc_tb):
        if self.session:
            await self.session.close()
    
    async def translate_text(self, text: str, target_lang: str) -> str:
        if not text.strip():
            return text
            
        prompt = f"Please translate the following text to {target_lang}, keep the original format, only return the translation result:\n{text}"
        
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        
        data = {
            "model": "deepseek-chat",
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.3
        }
        
        async with self.session.post(f"{self.base_url}/chat/completions", 
                                   headers=headers, json=data) as response:
            result = await response.json()
            return result["choices"][0]["message"]["content"].strip()
    
    async def batch_translate(self, texts: List[str], target_lang: str, 
                            max_concurrent: int = 50, progress_callback=None) -> List[str]:
        semaphore = asyncio.Semaphore(max_concurrent)
        completed = 0
        
        async def translate_with_semaphore(text):
            nonlocal completed
            async with semaphore:
                result = await self.translate_text(text, target_lang)
                completed += 1
                if progress_callback:
                    progress_callback(completed)
                return result
        
        tasks = [translate_with_semaphore(text) for text in texts]
        return await asyncio.gather(*tasks)