import os
import base64
import re
import logging
from typing import Optional

logger = logging.getLogger(__name__)

def get_data_uri(file_path: str) -> Optional[str]:
    """Конвертирует файл в Data URI (Base64)."""
    if not os.path.exists(file_path):
        logger.warning(f"Файл для инлайнинга не найден: {file_path}")
        return None
    
    ext = os.path.splitext(file_path)[1].lower()
    mime_types = {
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".gif": "image/gif",
        ".svg": "image/svg+xml",
        ".woff": "font/woff",
        ".woff2": "font/woff2",
        ".ttf": "font/ttf",
        ".js": "text/javascript",
        ".css": "text/css"
    }
    
    mime = mime_types.get(ext, "application/octet-stream")
    
    try:
        with open(file_path, "rb") as f:
            data = f.read()
            encoded = base64.b64encode(data).decode('utf-8')
            return f"data:{mime};base64,{encoded}"
    except Exception as e:
        logger.error(f"Ошибка при чтении файла {file_path} для Base64: {e}")
        return None

def inline_css_resources(css_content: str, base_dir: str) -> str:
    """Заменяет url(...) в CSS на Data URI."""
    def replacer(match):
        url = match.group(1).strip("'\"")
        if url.startswith("data:"):
            return match.group(0)
        
        full_path = os.path.normpath(os.path.join(base_dir, url))
        data_uri = get_data_uri(full_path)
        if data_uri:
            return f"url('{data_uri}')"
        return match.group(0)

    return re.sub(r"url\((.*?)\)", replacer, css_content)

def get_file_content(file_path: str) -> str:
    """Читает текстовый файл (для вшивания JS/CSS)."""
    if not os.path.exists(file_path):
        return ""
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            return f.read()
    except Exception as e:
        logger.error(f"Ошибка чтения текстового контента {file_path}: {e}")
        return ""
