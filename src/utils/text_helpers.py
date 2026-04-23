import html
import re

def esc(text: str) -> str:
    """Экранирует HTML символы."""
    if not text:
        return ""
    return html.escape(str(text))

def clean_text(text: str) -> str:
    """Нормализует текст: заменяет неразрывные пробелы, убирает лишние пробелы."""
    if not text:
        return ""
    # Заменяем неразрывный пробел на обычный
    text = text.replace("\xa0", " ").replace('\v', '\n')
    # Заменяем Wingdings bullets (U+F03E и подобные) на обычный bullet или удаляем
    text = re.sub(r"[\uf0b0-\uf0ff]", "", text)  # Private Use Area bullets
    # Collapse multiple spaces
    text = re.sub(r" +", " ", text)
    return text.strip()

def split_text(text: str) -> list:
    """Разбивает текст на абзацы по символу новой строки."""
    if not text:
        return []
    return [p.strip() for p in text.split('\n') if p.strip()]

def _is_slide_number(text: str) -> bool:
    """Проверяет, является ли текст номером слайда (число или 'стр. N')."""
    t = text.strip().lower()
    if not t:
        return False
    # Просто число
    if t.isdigit():
        return True
    # Формат 'стр. 10' или '2/25'
    if re.match(r'^(стр\.|c\.|page|p\.)?\s*\d+$', t):
        return True
    if re.match(r'^\d+\s*/\s*\d+$', t):
        return True
    return False

def _is_roman_numeral(text: str) -> bool:
    """Проверяет, является ли текст римской цифрой."""
    roman_regex = r'^M{0,4}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$'
    return bool(re.match(roman_regex, text.strip().upper())) and text.strip() != ""
