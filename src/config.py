import os
from pathlib import Path

# Базовые пути
BASE_DIR = Path(__file__).parent.parent.resolve()

# Данные по умолчанию
DEFAULT_SPEAKER = "Усольцева В.В."
DEFAULT_ORG = "ПНИПУ"

# Константы порогов для динамической типографии и layout
TEXT_LEN_HEAVY = 500  # Если текст длиннее, делаем левую колонку шире
TEXT_LEN_LIGHT = 200  # Если текст короче, делаем правую колонку (изображение) шире
TEXT_LEN_CONDENSED = 800  # Порог для перехода на сжатый шрифт
TEXT_LEN_TIGHT = 1200  # Порог для перехода на очень сжатый шрифт

ASPECT_WIDE = 1.8  # Соотношение сторон для широких изображений
ASPECT_TALL = 0.7  # Соотношение сторон для высоких изображений

CAPTION_H_DIST_FACTOR = 0.8
CAPTION_V_GAP_MIN = -5000
CAPTION_V_GAP_MAX = 30000

# Глобальные пространства имен XML
OMML_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
MATHML_NS = "{http://www.w3.org/1998/Math/MathML}"

DESIGN_CONFIG = {
    "icon_size": "1.6rem",
    "bullet_lists": {
        "enabled": True,
        "icon_map": {
            "•": "diamond",
            "◦": "diamond",
            "▪": "diamond",
            "-": "diamond",
            "–": "diamond",
            "—": "diamond",
            "*": "diamond",
            "": "diamond",
            "·": "diamond",
            "": "diamond",
            "ü": "diamond",
        },
        "icon_size": "1.6rem",
        "indent": "2rem",
        "border_left": "2px solid rgba(0, 242, 255, 0.15)",
        "background": "rgba(0, 242, 255, 0.03)",
    },
    "grid": {
        "max_columns": 4,
        "col_span_threshold": 1.2,
        "row_span_threshold": 0.833,
        "fallback_min_col_width": "200px",
    },
    "caption_search": {
        "vertical_range_mm": 25,
        "max_gap_mm": 20,
        "horizontal_overlap_ratio": 0.6,
        "overlap_tolerance": 0.4,
        "priority": "above",
    },
    "formula": {
        "mathjax_path": "libs/mathjax/tex-mml-svg.js",
        "fallback_font": "Roboto Mono",
        "fallback_font_size": "0.9em",
        "padding": "1.5rem",
        "background": "rgba(255, 255, 255, 0.01)",
        "border_color": "rgba(0, 242, 255, 0.15)",
    },
    "layout": {
        "caption_height_px": 50,
        "text_panel_ratio": 0.35,
    },
    "paths": {
        "logo_white": "logo/white.png",
        "media_output": "media",
        "media_output_full": str(BASE_DIR / "media"),
        "env_path": str(BASE_DIR / ".env"),
        "cache_dir": str(BASE_DIR / ".cache" / "llm"),
    },
    "icon_mapping": {
        "activity": ["динамика", "поле", "процесс", "геодинамика"],
        "droplet": ["нефть", "жидкость", "поток", "вода"],
        "layers": ["стратиграфия", "пласт", "разрез", "толща", "литология", "фондоформ"],
        "bar-chart": ["результат", "статистика", "данные", "анализ", "экономика", "итог"],
        "compass": ["направление", "азимут", "ориентация", "σhmax", "нmax", "тренд"],
        "target": ["цел", "перспектив", "направлен"],
        "list-todo": ["задач", "план", "постановк", "задан", "roadmap"],
        "database": ["модел", "сетк", "данн", "3d", "ячеек"],
        "cpu": ["автоматизац", "алгоритм", "расчет", "abaqus", "внедрен"],
        "alert-triangle": ["риск", "проблем", "опасност", "вниман", "предупрежден"],
        "refresh-ccw": ["ппд", "эффективност", "обработк"],
        "sliders": ["оптимизац", "параметр", "настройк"],
        "maximize": ["разм", "диаметр", "толщ", "глубин", "высот", "длин", "ширин"],
        "map-pin": ["регион", "месторожден", "район", "западн", "сибир", "участ"],
        "tower-control": ["скважин", "скв", "забой", "усть", "ствол"],
        "test-tube-2": ["испытан", "образц", "эксперимент", "лаборат"],
        "wrench": ["установк", "инструмент", "аппарат", "датчик", "прибор"],
        "zap": ["чувствительн", "влиян", "отклик", "эффект", "фактор"],
        "git-merge": ["нормирова", "приведен", "коррекц", "сопоставл"],
        "file-check": ["отчет", "регламент", "утвержден", "формат"],
        "info": ["информ", "инфо", "описан", "сведен", "справоч", "примечан"],
        "box": ["модель", "коробка", "box"],
    },
    "namespaces": {
        "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "a14": "http://schemas.microsoft.com/office/drawing/2014/main",
    },
    "STATIC_ASSETS_EMBED": True,
}
