import json
import hashlib
import time
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).parent.parent.parent.resolve()

class LLMCache:
    """Простой кэш для ответов LLM на основе файловой системы."""

    def __init__(self, cache_dir: str = None):
        self.cache_dir = Path(cache_dir) if cache_dir else BASE_DIR / ".cache" / "llm"
        self.cache_dir.mkdir(parents=True, exist_ok=True)

    def get(self, prompt: str, model: str) -> str | None:
        key = hashlib.sha256(f"{model}:{prompt}".encode()).hexdigest()
        path = self.cache_dir / f"{key}.json"
        if path.exists():
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    age_days = (time.time() - data.get("ts", 0)) / 86400
                    if age_days < 30:
                        return data.get("response")
            except Exception as e:
                logger.warning(f"Ошибка чтения кэша LLM: {e}")
        return None

    def set(self, prompt: str, model: str, response: str):
        key = hashlib.sha256(f"{model}:{prompt}".encode()).hexdigest()
        path = self.cache_dir / f"{key}.json"
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(
                    {
                        "prompt": prompt,
                        "model": model,
                        "response": response,
                        "ts": time.time(),
                    },
                    f,
                    ensure_ascii=False,
                    indent=2,
                )
        except Exception as e:
            logger.error(f"Ошибка сохранения кэша LLM: {e}")