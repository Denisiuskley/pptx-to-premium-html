import os
import json
import hashlib
from typing import Optional
from ..config import DESIGN_CONFIG

class LLMCache:
    """Кэширует ответы ИИ для экономии токенов и ускорения работы."""
    def __init__(self, cache_dir: Optional[str] = None):
        self.cache_dir = cache_dir or DESIGN_CONFIG['paths']['cache_dir']
        if not os.path.exists(self.cache_dir):
            os.makedirs(self.cache_dir, exist_ok=True)

    def _get_hash(self, prompt: str) -> str:
        return hashlib.md5(prompt.encode('utf-8')).hexdigest()

    def get(self, prompt: str) -> Optional[str]:
        cache_file = os.path.join(self.cache_dir, f"{self._get_hash(prompt)}.json")
        if os.path.exists(cache_file):
            try:
                with open(cache_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    return data.get('response')
            except Exception:
                return None
        return None

    def set(self, prompt: str, response: str):
        cache_file = os.path.join(self.cache_dir, f"{self._get_hash(prompt)}.json")
        try:
            with open(cache_file, 'w', encoding='utf-8') as f:
                json.dump({'prompt': prompt, 'response': response}, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[Cache] Error saving: {e}")
