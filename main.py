import os
import sys
import logging
from pathlib import Path
from src.core.parser import PPTParser
from src.converters.html_generator import HTMLGenerator
from src.config import DESIGN_CONFIG

# Настройка логирования в стиле оригинала
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger(__name__)

def main():
    """Единая точка входа для конвертации PPTX -> HTML."""
    
    # 1. Определение путей
    pptx_path = "Промежуточная.pptx"
    if not os.path.exists(pptx_path):
        # Ищем любой pptx в текущей папке
        pptx_files = list(Path('.').glob('*.pptx'))
        if pptx_files:
            pptx_path = str(pptx_files[0])
        else:
            logger.error(f"Файл {pptx_path} не найден и нет других .pptx файлов.")
            return

    base_name = Path(pptx_path).stem
    output_path = f"{base_name}.html"

    logger.info("=" * 50)
    logger.info(f"STARTING CONVERSION: {pptx_path}")
    logger.info("=" * 50)

    try:
        # 2. Инициализация компонентов
        parser = PPTParser(pptx_path)
        generator = HTMLGenerator()

        # 3. Парсинг (Извлечение контента)
        logger.info("[*] Phase 1: Parsing PPTX and extracting math/images...")
        slides_data = parser.extract_content()
        stats = parser.stats
        
        logger.info(f"[+] Extracted {stats['total_slides']} slides")
        logger.info(f"[+] Formulas found: {stats['formulas']}")
        logger.info(f"[+] Tables found: {stats['tables']}")
        logger.info(f"[+] Images saved: {stats['images_ok']} (fails: {stats['images_fail']})")

        # 4. Генерация (Рендеринг HTML)
        logger.info("[*] Phase 2: Rendering HTML with adaptive layout...")
        full_html = generator.generate_full_html(slides_data, stats)

        # 5. Сохранение
        logger.info(f"[*] Phase 3: Saving output to {output_path}")
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(full_html)

        logger.info("=" * 50)
        logger.info(f"SUCCESS: {output_path} generated!")
        logger.info("=" * 50)

    except Exception as e:
        logger.exception(f"CRITICAL ERROR during conversion: {e}")

if __name__ == "__main__":
    main()
