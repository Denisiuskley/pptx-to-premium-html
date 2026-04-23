import pytest
import os
from src.core.parser import PPTParser

def test_parser_init():
    """Проверка инициализации парсера."""
    parser = PPTParser("data/Промежуточная.pptx")
    assert parser.pptx_path == "data/Промежуточная.pptx"
    assert hasattr(parser, 'stats')

def test_clean_text_utility():
    """Проверка утилиты очистки текста."""
    from src.utils.text_helpers import clean_text
    raw = "Текст\xa0с\nпереносом    и пробелами"
    cleaned = clean_text(raw)
    assert cleaned == "Текст с переносом и пробелами"

def test_omml_to_mathml_fails_gracefully():
    """Проверка, что конвертер формул не падает на пустых данных."""
    from src.converters.math_converter import omml_to_mathml
    result = omml_to_mathml(None)
    assert result == ""
