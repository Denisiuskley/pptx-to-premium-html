import logging
import os
import sys
from pathlib import Path
from typing import Optional

# Импорт модулей проекта из src/
from src.core.parser import PPTParser
from src.converters.html_generator import HTMLGenerator
from src.config import DESIGN_CONFIG, DEFAULT_SPEAKER

# Настройка логирования
logging.basicConfig(
    level=logging.INFO, 
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger("main")

def run_conversion(pptx_path: str):
    """Основные этапы конвертации PPTX в Premium HTML."""
    if not os.path.exists(pptx_path):
        logger.error(f"Файл не найден: {pptx_path}")
        sys.exit(1)

    logger.info(f"Начало обработки: {pptx_path}")
    
    # 1. Извлечение контента
    parser = PPTParser(pptx_path)
    slides_data = parser.extract_content()
    parser.process_txt_files()
    
    # 2. Генерация HTML (Модульная архитектура)
    generator = HTMLGenerator()
    html_content = generator.generate_full_html(parser.slides_data, parser.stats)
    
    # 3. Сохранение (Имя по докладчику для идентичности с эталоном)
    import re
    speaker_name = slides_data[0].get("speaker_name", "") if slides_data else ""
    clean_speaker = re.sub(r"\[\[.*?\]\]", "", speaker_name)
    
    if clean_speaker:
        safe_name = re.sub(r'[\\/*?:"<>|]', '', clean_speaker).strip()
        output_name = f"{safe_name} (ПНИПУ).html" if safe_name else "presentation_output.html"
    else:
        output_name = "presentation_output.html"
    
    with open(output_name, "w", encoding="utf-8") as f:
        f.write(html_content)
    
    logger.info("=" * 40)
    logger.info(" МОДУЛЬНАЯ КОНВЕРТАЦИЯ ЗАВЕРШЕНА")
    logger.info(f" Выход: {output_name}")
    logger.info("=" * 40)

if __name__ == "__main__":
    # По умолчанию используем тестовый файл
    target_pptx = "Промежуточная.pptx"
    if len(sys.argv) > 1:
        target_pptx = sys.argv[1]
    
    run_conversion(target_pptx)
