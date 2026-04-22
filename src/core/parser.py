import os
import re
import html
import logging
from pathlib import Path
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from lxml import etree
from typing import List, Dict, Any, Optional

from ..config import DESIGN_CONFIG, OMML_NS, MATHML_NS
from ..utils.text_helpers import clean_text, _is_slide_number, esc
from ..utils.image_helpers import _save_image_with_white_bg
from ..converters.math_converter import omml_to_mathml

logger = logging.getLogger(__name__)

class PPTParser:
    """Извлекает структурированный контент из файлов PowerPoint, сохраняя иерархию и формулы."""
    
    def __init__(self, pptx_path: str):
        self.pptx_path = pptx_path
        self.presentation = Presentation(pptx_path)
        self.slide_width = self.presentation.slide_width
        self.slide_height = self.presentation.slide_height
        self.stats = {
            "total_slides": 0,
            "images_ok": 0,
            "images_fail": 0,
            "tables": 0,
            "formulas": 0
        }
        self.slides_data = []

    def extract_content(self) -> List[Dict[str, Any]]:
        """Главный цикл обхода слайдов и извлечения данных."""
        logger.info(f"Начало извлечения из: {self.pptx_path}")
        
        media_output_full = DESIGN_CONFIG["paths"]["media_output_full"]
        os.makedirs(media_output_full, exist_ok=True)

        for i, slide in enumerate(self.presentation.slides):
            slide_info = {
                "slide_num": i + 1,
                "title": "",
                "layout_type": "default",
                "content_items": [],  # Текст, таблицы, формулы в порядке появления
                "visuals": [],        # Изображения для правой панели или сетки
                "speaker_name": "",
                "speaker_info": ""
            }
            
            if slide.shapes.title:
                slide_info["title"] = clean_text(slide.shapes.title.text)

            # Перебор всех фигур (включая скрытые в XML)
            for shape, elem in self._iter_slide_shapes(slide):
                pos = (0, 0)
                if shape:
                    try:
                        pos = (shape.top, shape.left)
                    except: pass
                
                if shape == slide.shapes.title:
                    continue

                # 1. ТЕКСТ (с формулами внутри)
                if shape and shape.has_text_frame:
                    try:
                        shape_txt = clean_text(shape.text)
                        if shape_txt.lower() == slide_info["title"].lower():
                            continue

                        paragraphs = self._extract_math_segments_from_textframe(shape.text_frame)
                        if paragraphs:
                            slide_info["content_items"].append({
                                "type": "text",
                                "data": paragraphs,
                                "pos": pos
                            })
                            # Статистика формул
                            for para in paragraphs:
                                f_count = sum(1 for s in para if isinstance(s, str) and ("[[[MML_START]]]" in s or "[[[MML_FB_START]]]" in s))
                                self.stats["formulas"] += f_count
                    except Exception as e:
                        logger.warning(f"Ошибка парсинга текста на слайде {i+1}: {e}")

                # 2. ТАБЛИЦЫ
                elif shape and shape.has_table:
                    try:
                        table_html = self._table_to_html(shape.table)
                        slide_info["content_items"].append({
                            "type": "table",
                            "data": table_html,
                            "pos": pos
                        })
                        self.stats["tables"] += 1
                    except Exception as e:
                        logger.error(f"Ошибка таблицы на слайде {i+1}: {e}")

                # 3. ИЗОБРАЖЕНИЯ
                if shape and shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    try:
                        self._process_image(shape, i, slide_info)
                    except Exception as e:
                        logger.error(f"Ошибка изображения на слайде {i+1}: {e}")

            # Привязка подписей
            if slide_info["visuals"]:
                self._extract_captions_from_shapes(slide, slide_info)

            # Сортировка контента по вертикали
            slide_info["content_items"].sort(key=lambda x: x["pos"][0])
            slide_info["visuals"].sort(key=lambda x: (x["pos"][1], x["pos"][0])) # Sort by top, then left

            # Логика типа разметки
            if i == 0:
                slide_info["layout_type"] = "intro"
                self._enrich_intro_slide(slide_info)
            elif len(slide_info["visuals"]) == 0:
                has_table = any(it["type"] == "table" for it in slide_info["content_items"])
                if not has_table:
                    slide_info["layout_type"] = "full_text"

            self.slides_data.append(slide_info)

        self.stats["total_slides"] = len(self.slides_data)
        return self.slides_data

    def _iter_slide_shapes(self, slide):
        """Итерируется по всем фигурам слайда, включая скрытые в AlternateContent."""
        NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
              "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006"}
        
        standard_shapes = {s.shape_id: s for s in slide.shapes if hasattr(s, 'shape_id')}
        spTree = slide.shapes._spTree
        
        def process_container(container):
            for child in container:
                tag = child.tag.split('}')[-1]
                if tag in ('sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp'):
                    cNvPr = child.find('.//p:cNvPr', {"p": NS["p"]})
                    s_id = int(cNvPr.get('id')) if cNvPr is not None and cNvPr.get('id') else None
                    if s_id in standard_shapes:
                        yield standard_shapes[s_id], child
                    else:
                        yield None, child
                elif tag == 'AlternateContent':
                    choice = child.find(f".//{{{NS['mc']}}}Choice")
                    if choice is not None: yield from process_container(choice)
                elif tag == 'grpSp':
                    yield from process_container(child)

        yield from process_container(spTree)

    def _extract_math_segments_from_textframe(self, text_frame) -> list:
        """Извлекает сегменты текста и MathML маркеры из текстового фрейма."""
        NS = {
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
            "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
            "a14": "http://schemas.microsoft.com/office/drawing/2014/main",
        }
        all_paragraphs = []
        for para in text_frame.paragraphs:
            is_bullet = para.level > 0 or (para._element.pPr is not None and any(c.tag.endswith(('buChar', 'buAutoNum', 'buBlip', 'buBullet')) for c in para._element.pPr))
            
            p_elem = para._element
            segments = []
            current_text = []

            for child in p_elem:
                local = child.tag.split("}")[-1]
                if local == "r":
                    # Поиск формул внутри runs (обычное дело для Inline)
                    omath = child.find(".//m:oMath", NS) or child.find(".//mc:AlternateContent/mc:Choice//m:oMath", NS)
                    if omath is not None:
                        if current_text: segments.append("".join(current_text)); current_text = []
                        segments.append(self._get_mml_token(omath))
                        continue
                        
                    t_elems = child.findall(".//a:t", NS)
                    txt = "".join(t.text or "" for t in t_elems)
                    if txt:
                        # Обработка форматирования (B, I, Sub, Sup)
                        rPr = child.find(".//a:rPr", NS)
                        if rPr is not None:
                            is_sub = rPr.get("baseline") and int(rPr.get("baseline")) < 0 or rPr.get("subscript") in ("1", "true")
                            is_sup = rPr.get("baseline") and int(rPr.get("baseline")) > 0 or rPr.get("superscript") in ("1", "true")
                            if is_sub: txt = f"[[SUB_S]]{txt}[[SUB_E]]"
                            elif is_sup: txt = f"[[SUP_S]]{txt}[[SUP_E]]"
                            if rPr.get("b") in ("1", "true"): txt = f"[[B_S]]{txt}[[B_E]]"
                            if rPr.get("i") in ("1", "true"): txt = f"[[I_S]]{txt}[[I_E]]"
                        current_text.append(txt)
                elif local == "br":
                    current_text.append("\n")
                elif local in("oMath", "oMathPara"):
                    # Формулы как отдельные элементы параграфа (Display Math)
                    if current_text: segments.append("".join(current_text)); current_text = []
                    segments.append(self._get_mml_token(child))

            if current_text: segments.append("".join(current_text))
            if segments:
                if is_bullet and segments[0] and not segments[0].startswith("[[["):
                    segments[0] = "• " + segments[0]
                all_paragraphs.append(segments)
        return all_paragraphs

    def _get_mml_token(self, omath_elem) -> str:
        mathml = omml_to_mathml(omath_elem)
        if mathml: return f"[[[MML_START]]]{mathml}[[[MML_END]]]"
        texts = [t.text for t in omath_elem.findall(f'.//{OMML_NS}t') if t.text]
        return f"[[[MML_FB_START]]]{''.join(texts)}[[[MML_FB_END]]]" if texts else ""

    def _omml_to_mathml(self, omath_elem) -> Optional[str]:
        """Обертка над функцией из math_converter."""
        return omml_to_mathml(omath_elem)

    def _table_to_html(self, table) -> str:
        rows = []
        for r_idx, row in enumerate(table.rows):
            cells = []
            for cell in row.cells:
                # В таблицах тоже могут быть формулы, но для простоты пока текст
                cell_html = html.escape(cell.text)
                tag = "th" if r_idx == 0 else "td"
                cells.append(f"<{tag}>{cell_html}</{tag}>")
            rows.append(f"<tr>{''.join(cells)}</tr>")
        return f'<table class="data-table"><tbody>{"".join(rows)}</tbody></table>'

    def _process_image(self, shape, slide_idx, slide_info):
        media_output_full = DESIGN_CONFIG["paths"]["media_output_full"]
        img_name = f"slide_{slide_idx + 1}_img_{len(slide_info['visuals']) + 1}.png"
        img_path_full = str(Path(media_output_full) / img_name)
        img_path_rel = f"{DESIGN_CONFIG['paths']['media_output']}/{img_name}"

        try:
            blob = shape.image.blob
            if _save_image_with_white_bg(blob, img_path_full):
                slide_info["visuals"].append({
                    "src": img_path_rel,
                    "caption": "",
                    "pos": (shape.top, shape.left, shape.width, shape.height)
                })
                self.stats["images_ok"] += 1
        except Exception as e:
            logger.warning(f"Slide {slide_idx+1}: Image fail: {e}")
            self.stats["images_fail"] += 1

    def _extract_captions_from_shapes(self, slide, slide_info):
        """Интеллектуальный поиск подписей к изображениям (Spatial Matching)."""
        cfg = DESIGN_CONFIG["caption_search"]
        vert_range_emu = cfg["vertical_range_mm"] * 36000
        max_gap_emu = cfg["max_gap_mm"] * 36000
        
        text_shapes = []
        for shape in slide.shapes:
            if shape.has_text_frame and shape != slide.shapes.title:
                txt = clean_text(shape.text)
                if txt and not _is_slide_number(txt) and len(txt) < 300:
                    text_shapes.append({
                        "text": txt,
                        "left": shape.left, "top": shape.top, 
                        "width": shape.width, "height": shape.height
                    })

        used_texts = set()
        for vis in slide_info["visuals"]:
            v_l, v_t, v_w, v_h = vis["pos"]
            v_cx = v_l + v_w / 2
            
            best_match = None
            best_score = 0.0

            # Поиск преимущественно СНИЗУ (как в монолите)
            for ts in text_shapes:
                if ts["text"] in used_texts: continue
                
                ts_cx = ts["left"] + ts["width"] / 2
                ts_t = ts["top"]
                
                # Расстояние от низа картинки до верха текста
                gap = ts_t - (v_t + v_h)
                
                if 0 <= gap <= max_gap_emu:
                    # Горизонтальное соответствие
                    h_offset = abs(ts_cx - v_cx)
                    if h_offset < v_w * cfg["horizontal_overlap_ratio"]:
                        score = 1.0 - (gap / max_gap_emu)
                        if score > best_score:
                            best_score = score
                            best_match = ts

            if best_match:
                vis["caption"] = best_match["text"]
                used_texts.add(best_match["text"])
                # Удаляем текст из основного контента, чтобы не дублировать
                slide_info["content_items"] = [it for it in slide_info["content_items"] 
                    if it["type"] != "text" or clean_text(" ".join([" ".join(p) for p in it["data"]])) != best_match["text"]]

    def _enrich_intro_slide(self, slide_info):
        """Извлекает имя докладчика."""
        texts = []
        for it in slide_info["content_items"]:
            if it["type"] == "text":
                for p in it["data"]: texts.append(" ".join(s for s in p if not s.startswith("[[[") ))
        full_text = "\n".join(texts)
        lines = [l.strip() for l in full_text.split("\n") if l.strip()]
        if len(lines) > 1:
            slide_info["speaker_name"] = lines[0]
            slide_info["speaker_info"] = "\n".join(lines[1:])
        else:
            slide_info["speaker_name"] = "Докладчик"
            slide_info["speaker_info"] = ""
