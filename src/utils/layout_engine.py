import math
import logging
from typing import List, Dict, Any, Tuple

from ..config import DESIGN_CONFIG

logger = logging.getLogger(__name__)

def sort_shapes_spatially(items: list, threshold_mm: float = 20.0) -> list:
    """
    Сортировка объектов слайда с использованием 'жадных строк' (greedy rows).
    Для элементов с pos = (top, left, width, height) или объектов python-pptx.
    """
    if not items:
        return []
    
    threshold = threshold_mm * 36000 # 1 mm ~ 36000 EMU
    
    # Пытаемся извлечь верхнюю координату
    def get_top(x):
        if hasattr(x, 'top'): return x.top
        if isinstance(x, dict) and "pos" in x: return x["pos"][0]
        return 0

    def get_left(x):
        if hasattr(x, 'left'): return x.left
        if isinstance(x, dict) and "pos" in x: return x["pos"][1]
        return 0

    # Сортируем сначала по Y (top)
    sorted_by_y = sorted(items, key=get_top)
    
    rows = []
    if sorted_by_y:
        current_row = [sorted_by_y[0]]
        last_y = get_top(sorted_by_y[0])
        
        for item in sorted_by_y[1:]:
            curr_y = get_top(item)
            if abs(curr_y - last_y) < threshold:
                current_row.append(item)
            else:
                rows.append(sorted(current_row, key=get_left))
                current_row = [item]
                last_y = curr_y
        rows.append(sorted(current_row, key=get_left))
        
    final_sorted = []
    for row in rows:
        final_sorted.extend(row)
    return final_sorted

def sort_shapes_spatially_strict(items: list) -> list:
    """
    Строгая вертикальная сортировка элементов контента.
    """
    if not items:
        return []
        
    def get_pos(x):
        if hasattr(x, 'top'): return (x.top, x.left)
        if isinstance(x, dict) and "pos" in x: return (x["pos"][0], x["pos"][1])
        return (0, 0)

    return sorted(items, key=get_pos)

def get_best_layout(visuals: list, container_w: float = 1000, container_h: float = 600) -> Tuple[int, int, List[str], str, str]:
    """
    Рассчитывает оптимальную сетку CSS Grid для набора изображений.
    Минимизирует пустое пространство, подбирает веса колонок.
    
    Возвращает: (rows, cols, grid_styles, row_template, col_template)
    """
    n = len(visuals)
    if n == 0: 
        return 1, 1, [], "1fr", "1fr"
    
    caption_h = DESIGN_CONFIG["layout"]["caption_height_px"]
    
    # 6-8 рисунков: строгая сетка в 2 строки
    if n >= 6:
        cols = math.ceil(n / 2)
        grid_styles = []
        for i in range(n):
            r, c = i // cols, i % cols
            grid_styles.append(f"grid-row: {r+1} / span 1; grid-column: {c+1} / span 1;")
        col_tmpl = " ".join(["1fr"] * cols)
        return 2, cols, grid_styles, "1fr 1fr", col_tmpl

    # Вспомогательная логика для 1-5 рисунков
    aspects = []
    for v in visuals:
        # pos в EMU
        _, _, vw, vh = v.get("pos", (0, 0, 1, 1))
        asp = vw / vh if vh > 0 else 1.33
        aspects.append(asp)

    templates_n = {
        1: [(1, 1, [{'r':0, 'c':0, 'rs':1, 'cs':1}])],
        2: [
            (1, 2, [{'r':0, 'c':0, 'rs':1, 'cs':1}, {'r':0, 'c':1, 'rs':1, 'cs':1}]),
            (2, 1, [{'r':0, 'c':0, 'rs':1, 'cs':1}, {'r':1, 'c':0, 'rs':1, 'cs':1}]),
        ],
        3: [
            (1, 3, [{'r':0, 'c':0, 'rs':1, 'cs':1}, {'r':0, 'c':1, 'rs':1, 'cs':1}, {'r':0, 'c':2, 'rs':1, 'cs':1}]),
            (2, 2, [{'r':0, 'c':0, 'rs':2, 'cs':1}, {'r':0, 'c':1, 'rs':1, 'cs':1}, {'r':1, 'c':1, 'rs':1, 'cs':1}]),
            (2, 2, [{'r':0, 'c':1, 'rs':2, 'cs':1}, {'r':0, 'c':0, 'rs':1, 'cs':1}, {'r':1, 'c':0, 'rs':1, 'cs':1}]),
            (2, 2, [{'r':0, 'c':0, 'rs':1, 'cs':2}, {'r':1, 'c':0, 'rs':1, 'cs':1}, {'r':1, 'c':1, 'rs':1, 'cs':1}]),
        ],
        4: [
            (2, 2, [{'r':0, 'c':0, 'rs':1, 'cs':1}, {'r':0, 'c':1, 'rs':1, 'cs':1}, {'r':1, 'c':0, 'rs':1, 'cs':1}, {'r':1, 'c':1, 'rs':1, 'cs':1}]),
            (1, 4, [{'r':0, 'c':0, 'rs':1, 'cs':1}, {'r':0, 'c':1, 'rs':1, 'cs':1}, {'r':0, 'c':2, 'rs':1, 'cs':1}, {'r':0, 'c':3, 'rs':1, 'cs':1}]),
        ],
        5: [
            (2, 3, [{'r':0, 'c':0, 'rs':2, 'cs':1}, {'r':0, 'c':1, 'rs':1, 'cs':1}, {'r':0, 'c':2, 'rs':1, 'cs':1}, {'r':1, 'c':1, 'rs':1, 'cs':1}, {'r':1, 'c':2, 'rs':1, 'cs':1}]),
        ]
    }
    
    candidates = templates_n.get(n, [])
    best_score = -1
    best_layout = (1, n, [{'r':0, 'c':i, 'rs':1, 'cs':1} for i in range(n)])
    
    for rows, cols, items in candidates:
        cell_w = container_w / cols
        cell_h = container_h / rows
        current_total_area = 0
        
        for i, item in enumerate(items):
            cw = cell_w * item['cs']
            ch = cell_h * item['rs']
            img_ch = ch - caption_h
            if img_ch <= 0: continue
            
            asp = aspects[i]
            scale = min(cw / asp, img_ch)
            current_total_area += (scale * asp) * scale
        
        if current_total_area > best_score:
            best_score = current_total_area
            best_layout = (rows, cols, items)
    
    r_res, c_res, i_res = best_layout
    grid_styles = []
    for item in i_res:
        style = f"grid-row: {item['r']+1} / span {item['rs']}; grid-column: {item['c']+1} / span {item['cs']};"
        grid_styles.append(style)
        
    col_weights = ["1fr"] * c_res
    row_weights = ["1fr"] * r_res
    
    # Тонкая настройка весов для 2-3 элементов (из оригинала)
    if n == 2 and c_res == 2:
        a1, a2 = aspects[0], aspects[1]
        if a1 < 0.7 or a2 < 0.7:
            w1 = max(0.6, min(1.4, a1))
            w2 = max(0.6, min(1.4, a2))
            col_weights = [f"{w1:.1f}fr", f"{w2:.1f}fr"]
    elif n == 3 and c_res == 3:
        w = [max(0.7, min(1.3, a)) for a in aspects]
        col_weights = [f"{val:.1f}fr" for val in w]
        
    return r_res, c_res, grid_styles, " ".join(row_weights), " ".join(col_weights)
