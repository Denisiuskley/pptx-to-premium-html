from ..utils.text_helpers import esc, clean_text

def table_to_html(table) -> str:
    """Преобразует объект таблицы python-pptx в HTML строку."""
    html_rows = []
    
    # Определим, есть ли заголовок (обычно первая строка)
    has_header = True # В презентациях почти всегда так
    
    for i, row in enumerate(table.rows):
        cells = []
        tag = "th" if (i == 0 and has_header) else "td"
        
        for cell in row.cells:
            text = clean_text(cell.text_frame.text)
            # Обработка объединения ячеек (colspan/rowspan)
            # В python-pptx это хранится в _tc.gridSpan и т.д.
            attrs = ""
            if cell.is_merge_origin:
                # В текущей версии упростим, но это точка расширения
                pass
            
            cells.append(f"<{tag}{attrs}>{esc(text)}</{tag}>")
            
        html_rows.append(f"<tr>{''.join(cells)}</tr>")
    
    return f'<table class="data-table"><tbody>{"".join(html_rows)}</tbody></table>'
