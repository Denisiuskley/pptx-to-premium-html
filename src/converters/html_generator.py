import mimetypes
import base64
import os
import re
import html
import subprocess
import logging
from typing import List, Dict, Any, Tuple
logger = logging.getLogger(__name__)
from pathlib import Path
from src.style import BASE_HTML_TEMPLATE, BASE_HTML_TAIL
from src.config import DESIGN_CONFIG, DEFAULT_SPEAKER, BASE_DIR
from src.utils.text_helpers import esc, clean_text
from src.utils.resource_helpers import inline_css_resources, get_file_content, get_data_uri
from src.utils.layout_engine import get_best_layout

class HTMLGenerator:

    def __init__(self):
        self.output_html = None
        self.stats = {'total_slides': 0, 'images_ok': 0, 'tables': 0, 'formulas': 0, 'images_fail': 0}

    def _format_text_panel(self, slide_info: dict) -> str:
        """Формирует HTML-панель с чередованием текста, таблиц и формул."""
        parts = []
        bullet_config = DESIGN_CONFIG.get('bullet_lists', {})
        bullet_icon_map = bullet_config.get('icon_map', {})
        bullet_icon_size = DESIGN_CONFIG['icon_size']
        bullet_indent = bullet_config.get('indent', '2rem')
        bullet_border = bullet_config.get('border_left', '2px solid rgba(0,242,255,0.15)')
        bullet_bg = bullet_config.get('background', 'rgba(0,242,255,0.03)')
        for item_block in slide_info.get('content_items', []):
            item_type = item_block['type']
            data = item_block['data']
            if item_type == 'text':
                for p in data:
                    segments_to_join = []
                    formula_store = {}
                    if isinstance(p, list):
                        for idx, seg in enumerate(p):
                            if seg.startswith('[[[MML_START]]]'):
                                token = f'[[_F{idx}_]]'
                                mathml = seg[len('[[[MML_START]]]'):-len('[[[MML_END]]]')]
                                formula_store[token] = f'<span class="formula-container">{mathml}</span>'
                                segments_to_join.append(token)
                            elif seg.startswith('[[[MML_FB_START]]]'):
                                token = f'[[_FB{idx}_]]'
                                fb = seg[len('[[[MML_FB_START]]]'):-len('[[[MML_FB_END]]]')]
                                formula_store[token] = f'<span class="formula-fallback">{esc(fb)}</span>'
                                segments_to_join.append(token)
                            else:
                                segments_to_join.append(seg)
                    else:
                        segments_to_join.append(str(p))
                    full_content_with_tokens = ''.join(segments_to_join)
                    if not full_content_with_tokens.strip():
                        continue
                    clean_search_text = re.sub('<[^>]+>', '', full_content_with_tokens)
                    items = self._split_text_into_items(clean_search_text)
                    for item in items:
                        raw_text = item['text'].strip()
                        is_conclusion = raw_text.lower().startswith('вывод:')
                        if is_conclusion:
                            item_text = re.sub('^[Вв]ывод[:\\s]+', '', raw_text).strip()
                        else:
                            item_text = item['text']
                        display_text = html.escape(item_text)
                        display_text = display_text.replace('[[SUB_S]]', '<sub>').replace('[[SUB_E]]', '</sub>')
                        display_text = display_text.replace('[[SUP_S]]', '<sup>').replace('[[SUP_E]]', '</sup>')
                        display_text = display_text.replace('[[B_S]]', '<strong>').replace('[[B_E]]', '</strong>')
                        display_text = display_text.replace('[[I_S]]', '<em>').replace('[[I_E]]', '</em>')
                        for token, real_html in formula_store.items():
                            display_text = display_text.replace(token, real_html)
                        if is_conclusion:
                            display_text = f'<strong>{display_text}</strong>'
                        if item.get('is_bullet') or is_conclusion:
                            if is_conclusion:
                                icon = 'check-check'
                            else:
                                keyword_icon = self._get_icon_for_text(item['text'])
                                default_bullet = 'diamond'
                                icon = keyword_icon if keyword_icon != 'chevron-right' else default_bullet
                            marker_class = 'animate-marker'
                            if icon == 'diamond':
                                marker_class += ' marker-stretched'
                            wrapper_class = 'list-item-bullet animate-up'
                            wrapper_style = f'padding-left: {bullet_indent};'
                            if is_conclusion:
                                wrapper_class += ' list-item-conclusion'
                            else:
                                wrapper_style += f' border-left: {bullet_border}; background: {bullet_bg};'
                            parts.append(f'<div class="{wrapper_class}" style="{wrapper_style}"><i data-lucide="{icon}" class="{marker_class}" style="width: {bullet_icon_size}; height: {bullet_icon_size}; flex-shrink: 0;"></i><div class="list-text">{display_text}</div></div>')
                        else:
                            keyword_icon = self._get_icon_for_text(item['text'])
                            icon = keyword_icon if keyword_icon != 'chevron-right' else 'diamond'
                            marker_class = 'animate-marker'
                            if icon == 'diamond':
                                marker_class += ' marker-stretched'
                            parts.append(f'''<div class="list-item animate-up"><i data-lucide="{icon}" class="{marker_class}" style="width: {DESIGN_CONFIG['icon_size']}; height: {DESIGN_CONFIG['icon_size']}; flex-shrink: 0;"></i><div class="list-text">{display_text}</div></div>''')
            elif item_type == 'table':
                parts.append(str(data))
            elif item_type == 'formula':
                parts.append(f'<div class="formula-block animate-up">{data}</div>')
        return ''.join(parts)

    def _split_text_into_items(self, text: str) -> list:
        """Разбирает многострочный текст на отдельные пункты.
        Возвращает список dict: [{"text": str, "is_bullet": bool, "bullet_char": str|None, "level": int}]
        Префиксы списков (цифры, bullets, дефисы) удаляются из текста, но тип маркера запоминается.
        """
        if not text:
            return []
        lines = text.split('\n')
        items = []
        bullet_chars = set(DESIGN_CONFIG.get('bullet_lists', {}).get('icon_map', {}).keys())
        prefix_pattern = '^\\s*(\\d+[\\.\\)]\\s*|[a-zа-я][\\.\\)]\\s*|[•◦▪\\-–—*·\uf0d8ü\uf0be]\\s*)+\\s*'
        for line in lines:
            stripped = line.strip()
            if not stripped:
                continue
            first_char = stripped[0]
            is_bullet = first_char in bullet_chars or bool(re.match('^\\s*(\\d+[\\.\\)]|[a-zA-Zа-яА-Я][\\.\\)]|\\([\\d\\w]\\))\\s*', stripped)) or bool(re.match('^\\s*[IVXLCDMivxlcdm]+\\.\\s*', stripped))
            cleaned = re.sub(prefix_pattern, '', stripped).strip()
            if cleaned:
                items.append({'text': cleaned, 'is_bullet': is_bullet, 'bullet_char': first_char if is_bullet else None, 'level': 0})
        return items

    def _get_icon_for_text(self, text: str) -> str:
        """Возвращает идентификатор иконки на основе текста (по ключевым словам)."""
        text = text.lower()
        for icon, keywords in DESIGN_CONFIG['icon_mapping'].items():
            if any((kw in text for kw in keywords)):
                return icon
        return 'chevron-right'

    def _generate_section_tag(self, data: dict) -> str:
        """Генерирует тег секции (например, "Данные", "Результаты") на основе текста."""
        all_texts = []
        for block in data.get('content_items', []):
            if block['type'] == 'text':
                for p in block['data']:
                    all_texts.append(' '.join((s for s in p if not s.startswith('[[['))))
        text_combined = ' '.join(all_texts).lower()
        if any((kw in text_combined for kw in ['данные', 'таблиц', 'исходн'])):
            return 'Данные'
        if any((kw in text_combined for kw in ['результат', 'вывод', 'итог'])):
            return 'Результаты'
        if any((kw in text_combined for kw in ['описание', 'метод', 'подход'])):
            return 'Описание'
        if any((kw in text_combined for kw in ['анализ', 'исследовани'])):
            return 'Анализ'
        return 'Описание'

    def generate_full_html(self, slides_data, stats) -> str:
        """Генерирует итоговый HTML-файл на основе встроенного шаблона и данных слайдов."""
        self.stats = stats  # Синхронизируем статы
        self.stats['total_slides'] = len(slides_data)
        logger.info('Рендеринг HTML...')
        speaker_name = slides_data[0].get('speaker_name', '') if slides_data else ''
        clean_speaker = re.sub('\\[\\[.*?\\]\\]', '', speaker_name)
        if clean_speaker:
            safe_name = re.sub('[\\\\/*?:"<>|]', '', clean_speaker).strip()
            if safe_name:
                self.output_html = f'{safe_name} (ПНИПУ).html'
        if not self.output_html:
            self.output_html = 'presentation_output.html'
        is_standalone = DESIGN_CONFIG.get('STATIC_ASSETS_EMBED', False)
        head_part = BASE_HTML_TEMPLATE
        tail_part = BASE_HTML_TAIL
        s_title = clean_speaker if clean_speaker else 'Доклад'
        head_part = head_part.replace('{speaker_name}', s_title)
        logo_rel = DESIGN_CONFIG['paths']['logo_white']
        v_logo_path = BASE_DIR / 'web_demo' / logo_rel
        if not v_logo_path.exists():
            v_logo_path = BASE_DIR / logo_rel
        if is_standalone:
            logger.info('Подготовка автономного (standalone) файла (все ресурсы вшиваются)...')
            logo = self._get_data_uri(v_logo_path)
            fonts_css_path = BASE_DIR / 'libs' / 'fonts' / 'fonts.css'
            fonts_inlined = self._inline_css_fonts(fonts_css_path)
            head_part = re.sub(r'<link rel="stylesheet" href="libs/fonts/fonts.css">', f'<style>{fonts_inlined}</style>', head_part)
            scripts_to_inline = [
                ('libs/gsap/gsap.min.js', r'<script src="libs/gsap/gsap.min.js" defer></script>'),
                ('libs/lucide/lucide.min.js', r'<script src="libs/lucide/lucide.min.js" defer></script>'),
                ('libs/mathjax/tex-mml-svg.js', r'<script src="libs/mathjax/tex-mml-svg.js" defer></script>')
            ]
            for rel_path, pattern in scripts_to_inline:
                s_path = BASE_DIR / rel_path
                s_content = self._get_file_content(s_path)
                head_part = re.sub(pattern, lambda m, c=s_content: f'<script>{c}</script>', head_part, flags=re.DOTALL)
        else:
            logo = DESIGN_CONFIG['paths']['logo_white']
        slides_content = ''
        total = len(slides_data)
        for idx, data in enumerate(slides_data):
            num = idx + 1
            title = data['title']
            section_tag = self._generate_section_tag(data)
            text_panel_html = self._format_text_panel(data)
            total_text_chars = 0
            for item in data.get('content_items', []):
                if item['type'] == 'text':
                    for p in item['data']:
                        for s in p:
                            if isinstance(s, str):
                                f_count = s.count('[[[MML_START]]]') + s.count('[[[MML_FB_START]]]')
                                if f_count > 0:
                                    clean_s = re.sub(r'\[\[\[MML_START\]\].*?\[\[\[MML_END\]\]\]', '', s, flags=re.DOTALL)
                                    clean_s = re.sub(r'\[\[\[MML_FB_START\]\].*?\[\[\[MML_FB_END\]\]\]', '', clean_s, flags=re.DOTALL)
                                    clean_s = re.sub(r'\[\[[A-Z0-9_]+\]\]', '', clean_s)
                                    total_text_chars += len(clean_s)
                                    total_text_chars += f_count * 20
                                else:
                                    clean_s = re.sub(r'\[\[[A-Z0-9_]+\]\]', '', s)
                                    total_text_chars += len(clean_s)
                elif item['type'] == 'table':
                    total_text_chars += 300
                elif item['type'] == 'formula':
                    total_text_chars += 40
            panel_class = 'analytical-panel animate-up'
            if is_standalone:
                embedded_count = 0
                for vis in data.get('visuals', []):
                    if vis.get('src') and (not vis['src'].startswith('data:')):
                        v_path = BASE_DIR / 'web_demo' / vis['src']
                        if v_path.exists():
                            vis['src'] = get_data_uri(v_path)
                            embedded_count += 1
                        else:
                            logger.warning(f'Медиа не найдено для эмбеддинга: {v_path}')
                if embedded_count > 0:
                    logger.debug(f'Слайд {idx + 1}: эмбедировано {embedded_count} изображений')
            if data['layout_type'] == 'intro':
                s_name = data.get('speaker_name', '')
                s_info = data.get('speaker_info', '')
                # Очистка имен и инфо от маркеров для паритета
                s_name = re.sub(r'\[\[.*?\]\]', '', s_name)
                s_info = re.sub(r'\[\[.*?\]\]', '', s_info)
                slides_content += f"""
        <section class="slide hide-title">
            <div class="logo-container"><img src="{logo}" alt="Логотип" class="header-logo" onerror="this.style.display='none'; this.onerror=null;"></div>
            <div class="slide-header">
                <div class="slide-number">{num:02d} / {total:02d}</div>
            </div>
            <div class="slide-content-title">
                <h1 class="main-heading animate-up">{esc(title)}</h1>
            </div>
            <div class="presenter-card animate-up">
                <span class="presenter-label">Докладчик</span>
                <div class="presenter-name">{esc(s_name)}</div>
                <div class="presenter-info">{esc(s_info)}</div>
            </div>
        </section>"""
            elif data['layout_type'] == 'conclusions_dual':
                slides_content += f"""
        <section class="slide">
            <div class="logo-container"><img src="{logo}" alt="Логотип" class="header-logo" onerror="this.style.display='none'; this.onerror=null;"></div>
            <div class="slide-header"><div class="slide-number">{num:02d} / {total:02d}</div><div class="slide-title">{esc(title)}</div></div>
            <div class="slide-split" style="grid-template-columns: 1fr 1fr; height: calc(100% - 145px);">
                <div class="{panel_class}">
                    {data['left_html']}
                </div>
                <div class="{panel_class}">
                    {data['right_html']}
                </div>
            </div>
        </section>"""
            elif data['layout_type'] == 'research_roadmap':
                slides_content += f"""
        <section class="slide">
            <div class="logo-container"><img src="{logo}" alt="Логотип" class="header-logo" onerror="this.style.display='none'; this.onerror=null;"></div>
            <div class="slide-header"><div class="slide-number">{num:02d} / {total:02d}</div><div class="slide-title">{esc(title)}</div></div>
            <div class="slide-split" style="grid-template-columns: 1.4fr 0.8fr; height: calc(100% - 145px);">
                <div class="{panel_class}">
                    {data['content_html']}
                </div>
                <div class="viz-card animate-up" style="display: flex; flex-direction: column; justify-content: center; align-items: center; background: radial-gradient(circle, var(--accent-soft) 0%, transparent 80%); border-radius: 3rem; padding: 3rem; border: 1px solid var(--glass-border); position: relative; overflow: hidden; height: 100%;">
                     <div class="rocket-glow"></div>
                     <i data-lucide="rocket" style="width: 120px; height: 120px; color: var(--accent); margin-bottom: 2rem; filter: drop-shadow(0 0 30px var(--accent)); transform: rotate(-45deg); animation: pulse-rocket 2s infinite ease-in-out;"></i>
                     <h4 style="font-family: 'Outfit'; font-size: var(--fs-research-year); margin: 0; font-weight: 800; background: linear-gradient(135deg, white 0%, var(--accent) 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent;">ROADMAP</h4>
                </div>
            </div>
        </section>"""
            elif data['layout_type'] == 'full_text':
                slides_content += f"""
        <section class="slide">
            <div class="logo-container"><img src="{logo}" alt="Логотип" class="header-logo" onerror="this.style.display='none'; this.onerror=null;"></div>
            <div class="slide-header"><div class="slide-number">{num:02d} / {total:02d}</div><div class="slide-title">{esc(title)}</div></div>
            <div class="slide-content-full animate-up" style="width: 85%; margin: 2rem auto; height: calc(100% - 145px); min-height: 0;">
                <div class="{panel_class}" style="height: 100%;">
                    {text_panel_html}
                </div>
            </div>
        </section>"""
            else:
                data['visuals'] = self._spatial_sort(data['visuals'])
                has_text = bool(text_panel_html.strip())
                base_ratio = DESIGN_CONFIG['layout']['text_panel_ratio']
                layout_class = 'slide-split'
                pre_results = get_best_layout(data['visuals'], 1200, 940)
                rows_pre, cols_pre = (pre_results[0], pre_results[1])
                if has_text:
                    if cols_pre == 1:
                        grid_split = '1fr auto'
                        layout_class += ' layout-auto-width'
                        cont_w = 1200
                    else:
                        d_factor = max(0.65, min(1.2, total_text_chars / 1200.0 + 0.5))
                        dynamic_ratio = base_ratio * d_factor
                        dynamic_ratio = max(0.25, min(0.42, dynamic_ratio))
                        grid_split = f'{dynamic_ratio:.3f}fr {1.0 - dynamic_ratio:.3f}fr'
                        cont_w = 1728 * (1.0 - dynamic_ratio)
                else:
                    cont_w = 1728
                    grid_split = '1fr'
                cont_h = 940
                results = get_best_layout(data['visuals'], cont_w, cont_h)
                rows, cols, grid_styles, row_tmpl, col_tmpl = results
                visuals_items_html = ''
                for idx, vis in enumerate(data['visuals']):
                    style = grid_styles[idx] if idx < len(grid_styles) else ''
                    caption_html = f"""<div class="viz-caption">{esc(vis.get('caption', ''))}</div>""" if vis.get('caption') else ''
                    if vis.get('src'):
                        alt_text = esc(vis['caption']) if vis.get('caption') else 'Изображение'
                        content_html = f'''<div class="viz-box"><img src="{vis['src']}" alt="{alt_text}"></div>'''
                    else:
                        content_html = '<div class="error-box">Нет изображения</div>'
                    visuals_items_html += f'<div class="viz-item" style="{style}">{caption_html}{content_html}</div>'
                if has_text:
                    slides_content += f"""
        <section class="slide">
            <div class="logo-container"><img src="{logo}" alt="Логотип" class="header-logo" onerror="this.style.display='none'; this.onerror=null;"></div>
            <div class="slide-header"><div class="slide-number">{num:02d} / {total:02d}</div><div class="slide-title">{esc(title)}</div></div>
            <div class="slide-split {layout_class}" style="grid-template-columns: {grid_split}; height: calc(100% - 145px); min-height: 0;">
                <div class="{panel_class}">{text_panel_html}</div>
                <div class="img-stack animate-up" style="display: grid; grid-template-columns: {col_tmpl}; grid-template-rows: {row_tmpl}; gap: var(--gap-main);">{visuals_items_html}</div>
            </div>
        </section>"""
                else:
                    # Images only: full-width grid
                    slides_content += f"""
        <section class="slide">
            <div class="logo-container"><img src="{logo}" alt="Логотип" class="header-logo" onerror="this.style.display='none'; this.onerror=null;"></div>
            <div class="slide-header"><div class="slide-number">{num:02d} / {total:02d}</div><div class="slide-title">{esc(title)}</div></div>
            <div class="slide-content-full animate-up" style="width: 100%; margin: 0; height: calc(100% - 145px); display: grid; grid-template-columns: {col_tmpl}; grid-template-rows: {row_tmpl}; gap: var(--gap-main); padding: 0;">
                {visuals_items_html}
            </div>
        </section>"""
        final_content = head_part + slides_content + tail_part
        logger.info(f'Сохранение в {self.output_html}...')
        with open(self.output_html, 'w', encoding='utf-8') as f:
            f.write(final_content)
        
        logger.info('\n' + '=' * 40)
        logger.info(' ИТОГОВЫЙ ОТЧЕТ КОНВЕРТАЦИИ')
        logger.info('=' * 40)
        logger.info(f" Всего слайдов:   {stats['total_slides']}")
        logger.info(f" Успешных фото:   {stats['images_ok']}")
        logger.info(f" Таблиц:          {stats['tables']}")
        logger.info(f" Формул:          {stats['formulas']}")
        logger.info(f" Пропущено:       {stats['images_fail']} (см. лог выше)")
        logger.info('=' * 40)
        return final_content

    def _spatial_sort(self, items: list, threshold_mm: float=20.0) -> list:
        """
        Сортировка изображений с использованием 'жадных строк' (greedy rows).
        Для элементов с pos = (left, top, width, height).
        """
        if not items:
            return []
        threshold = threshold_mm * 36000
        sorted_by_y = sorted(items, key=lambda x: x.get('pos', (0, 0, 0, 0))[1])
        rows = []
        if sorted_by_y:
            current_row = [sorted_by_y[0]]
            last_y = sorted_by_y[0].get('pos', (0, 0, 0, 0))[1]
            for item in sorted_by_y[1:]:
                curr_y = item.get('pos', (0, 0, 0, 0))[1]
                if abs(curr_y - last_y) < threshold:
                    current_row.append(item)
                else:
                    rows.append(sorted(current_row, key=lambda x: x.get('pos', (0, 0, 0, 0))[0]))
                    current_row = [item]
                    last_y = curr_y
            rows.append(sorted(current_row, key=lambda x: x.get('pos', (0, 0, 0, 0))[0]))
        final_sorted = []
        for row in rows:
            final_sorted.extend(row)
        return final_sorted

    def _spatial_sort_strict(self, items: list) -> list:
        """
        Строгая вертикальная сортировка элементов контента.
        Используется для соблюдения хронологического порядка (сверху вниз)
        в текстовой панели (текст, таблицы, формулы).
        """
        if not items:
            return []
        return sorted(items, key=lambda x: (x.get('pos', (0, 0))[0], x.get('pos', (0, 0))[1]))

    def _inline_css_fonts(self, css_path: Path) -> str:
        """Читает CSS и вшивает шрифты в него через Data URI."""
        content = self._get_file_content(css_path)
        if not content:
            return ''

        def fill_url(match):
            url_path = match.group(1).strip('\'"')
            full_path = (css_path.parent / url_path).resolve()
            data_uri = self._get_data_uri(full_path)
            return f"url('{data_uri}')"
        return re.sub('url\\((.*?)\\)', fill_url, content)

    def _get_data_uri(self, file_path: Path) -> str:
        """Превращает файл в Base64 Data URI."""
        if not file_path.exists():
            logger.warning(f'Файл не найден для эмбеддинга: {file_path}')
            return ''
        mime_type, _ = mimetypes.guess_type(str(file_path))
        if not mime_type:
            mime_type = 'application/octet-stream'
        with open(file_path, 'rb') as f:
            data = f.read()
            b64 = base64.b64encode(data).decode('utf-8')
            return f'data:{mime_type};base64,{b64}'

    def _get_file_content(self, file_path: Path) -> str:
        """Читает текстовый файл."""
        if not file_path.exists():
            logger.warning(f'Текстовый файл не найден: {file_path}')
            return ''
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()