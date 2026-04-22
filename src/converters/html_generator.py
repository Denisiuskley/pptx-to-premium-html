import re
import html
import logging
from typing import List, Dict, Any
from ..style import BASE_HTML_TEMPLATE, BASE_HTML_TAIL
from ..config import DESIGN_CONFIG, DEFAULT_SPEAKER, DEFAULT_ORG
from ..utils.text_helpers import esc, clean_text

logger = logging.getLogger(__name__)

class HTMLGenerator:
    """Генерирует итоговый HTML, используя сложную логику иконок и форматирования."""
    
    def __init__(self):
        self.design = DESIGN_CONFIG
        self.icon_mapping = self.design.get("icon_mapping", {})

    def generate_full_html(self, slides_data: List[Dict[str, Any]], stats: Dict[str, Any]) -> str:
        """Сборка всего документа."""
        # Для интро-слайда достаем данные из первого слайда
        intro = slides_data[0] if slides_data else {}
        speaker = intro.get("speaker_name", DEFAULT_SPEAKER)
        
        head = BASE_HTML_TEMPLATE.replace("{speaker_name}", esc(speaker))
        
        slides_content = []
        for slide in slides_data:
            slides_content.append(self.render_slide(slide, len(slides_data)))
            
        return head + "\n".join(slides_content) + BASE_HTML_TAIL

    def render_slide(self, slide: Dict[str, Any], total_slides: int) -> str:
        """Рендеринг одного слайда в зависимости от его типа."""
        layout = slide.get("layout_type", "default")
        
        if layout == "intro":
            return self._render_intro_slide(slide)
        
        # Получаем подготовленную HTML-панель текста
        text_panel_html = self._format_text_panel(slide)
        
        # Визуальные элементы (сетка изображений)
        visuals_html = self._render_visuals_grid(slide)
        
        # Определение классов разметки (из оригинала)
        has_text = len(slide.get("content_items", [])) > 0
        layout_class = ""
        grid_split = "1.2fr 1.8fr"
        
        if not slide.get("visuals"):
            grid_split = "1fr"
        
        return f"""
        <section class="slide">
            <div class="logo-container"><img src="{self.design['paths']['logo_white']}" alt="Лого" class="header-logo" onerror="this.style.display='none';"></div>
            <div class="slide-header">
                <div class="slide-number">{slide['slide_num']:02d} / {total_slides:02d}</div>
                <div class="slide-title">{esc(slide['title'])}</div>
            </div>
            <div class="slide-split {layout_class}" style="grid-template-columns: {grid_split}; height: calc(100% - 145px); min-height: 0;">
                <div class="analytical-panel">{text_panel_html}</div>
                <div class="img-stack animate-up" style="display: grid; gap: var(--gap-main);">{visuals_html}</div>
            </div>
        </section>
        """

    def _render_intro_slide(self, slide: Dict[str, Any]) -> str:
        return f"""
        <section class="slide hide-title">
            <div class="logo-container"><img src="{self.design['paths']['logo_white']}" alt="Лого" class="header-logo"></div>
            <div class="slide-content-title">
                <h1 class="main-heading animate-up">{esc(slide['title'])}</h1>
                <div class="presenter-card animate-up">
                    <span class="presenter-label">Докладчик</span>
                    <div class="presenter-name">{esc(slide.get('speaker_name', ''))}</div>
                    <div class="presenter-info">{esc(slide.get('speaker_info', ''))}</div>
                </div>
            </div>
        </section>
        """

    def _render_visuals_grid(self, slide: Dict[str, Any]) -> str:
        """Рендерит сетку изображений."""
        # Упрощенная версия, в следующем этапе подключим LayoutEngine для Grid 1-8
        html_parts = []
        for vis in slide.get("visuals", []):
            caption = f'<div class="viz-caption">{esc(vis["caption"])}</div>' if vis.get("caption") else ""
            html_parts.append(f"""
                <div class="viz-item">
                    {caption}
                    <div class="viz-box"><img src="{vis['src']}" alt="Visual"></div>
                </div>
            """)
        return "".join(html_parts)

    def _format_text_panel(self, slide: Dict[str, Any]) -> str:
        """Превращает content_items в HTML с иконками и форматированием."""
        parts = []
        
        # Определяем иконку для всего слайда на основе заголовка
        slide_icon = self._get_icon_for_text(slide["title"])
        
        for item in slide.get("content_items", []):
            if item["type"] == "text":
                for para_segments in item["data"]:
                    # Слияние сегментов и распаковка маркеров
                    full_p_text = self._unpack_markers("".join(para_segments))
                    
                    # Определение иконки для параграфа
                    icon = self._get_icon_for_text(full_p_text)
                    if not icon: icon = slide_icon if slide_icon else "chevron-right"
                    
                    # Класс для "Выводов/Заключений"
                    item_class = "list-item"
                    if any(w in full_p_text.lower() for w in ["вывод", "заключен", "результат", "итог"]):
                        item_class += " list-item-conclusion"
                        icon = "rocket" # Ракеты для выводов

                    parts.append(f"""
                        <div class="{item_class}">
                            <i data-lucide="{icon}"></i>
                            <div class="list-text">{full_p_text}</div>
                        </div>
                    """)
            
            elif item["type"] == "table":
                parts.append(item["data"])
                
            elif item["type"] == "formula":
                parts.append(f'<div class="formula-block">{item["data"]}</div>')

        return "".join(parts)

    def _unpack_markers(self, text: str) -> str:
        """Распаковывает временные маркеры форматирования и формул в HTML."""
        # 1. Формулы
        text = text.replace("[[[MML_START]]]", '<span class="formula-container">')
        text = text.replace("[[[MML_END]]]", '</span>')
        text = text.replace("[[[MML_FB_START]]]", '<span class="formula-fallback">')
        text = text.replace("[[[MML_FB_END]]]", '</span>')
        
        # 2. Индексы и жирность (маркеры из парсера)
        text = text.replace("[[SUB_S]]", "<sub>").replace("[[SUB_E]]", "</sub>")
        text = text.replace("[[SUP_S]]", "<sup>").replace("[[SUP_E]]", "<sup>")
        text = text.replace("[[B_S]]", "<strong>").replace("[[B_E]]", "</strong>")
        text = text.replace("[[I_S]]", "<em>").replace("[[I_E]]", "</em>")
        
        return text

    def _get_icon_for_text(self, text: str) -> str:
        """Подбирает иконку Lucide на основе ключевых слов в тексте."""
        text_lower = text.lower()
        for icon, keywords in self.icon_mapping.items():
            if any(kw in text_lower for kw in keywords):
                return icon
        return ""
