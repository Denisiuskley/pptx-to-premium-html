import os
import requests
from dotenv import load_dotenv
from ..config import DESIGN_CONFIG
from .cache import LLMCache

class AIProcessor:
    """Отвечает за взаимодействие с нейросетями и трансформацию текста."""
    def __init__(self):
        load_dotenv(DESIGN_CONFIG['paths']['env_path'])
        self.api_key = os.getenv("OPENROUTER_API_KEY")
        self.cache = LLMCache()
        self.model = DESIGN_CONFIG['ai_config']['default_model']

    def call_ai(self, system_prompt: str, user_prompt: str, use_cache: bool = True) -> str:
        """Делает запрос к LLM через OpenRouter."""
        full_prompt = f"SYSTEM: {system_prompt}\nUSER: {user_prompt}"
        
        if use_cache:
            cached = self.cache.get(full_prompt)
            if cached:
                return cached

        if not self.api_key:
            print("[AI] Warning: OPENROUTER_API_KEY not found in .env")
            return ""

        try:
            response = requests.post(
                url="https://openrouter.ai/api/v1/chat/completions",
                headers={
                    "Authorization": f"Bearer {self.api_key}",
                    "Content-Type": "application/json",
                },
                json={
                    "model": self.model,
                    "messages": [
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ]
                },
                timeout=30
            )
            response.raise_for_status()
            result = response.json()['choices'][0]['message']['content']
            
            if use_cache:
                self.cache.set(full_prompt, result)
            return result
        except Exception as e:
            print(f"[AI] Error: {e}")
            return ""

    def process_txt_files(self, folder_path: str) -> dict:
        """Читает дополнительные текстовые файлы для обогащения контента слайдов."""
        data = {}
        # Список файлов, которые мы ищем
        targets = ["Выводы.txt", "Направление дальнейших исследований.txt"]
        
        for name in targets:
            path = os.path.join(folder_path, name)
            if os.path.exists(path):
                try:
                    with open(path, 'r', encoding='utf-8') as f:
                        data[name.split('.')[0]] = f.read().strip()
                except Exception as e:
                    print(f"[AI] Error reading {name}: {e}")
        return data
