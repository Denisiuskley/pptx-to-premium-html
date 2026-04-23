import html
import logging
from typing import Optional
from ..config import OMML_NS
from ..utils.text_helpers import esc

logger = logging.getLogger(__name__)

def token_type(txt: Optional[str]) -> str:
    """Определяет тип токена MathML (mn - число, mo - оператор, mi - идентификатор)."""
    if not txt:
        return "mi"
    txt_str = str(txt)
    if txt_str.isdigit() or txt_str.replace('.', '').isdigit():
        return "mn"
    ops = set("+-*/=<>≤≥≈≡∑∏∫∂∇±×÷∈∉⊂⊃∪∩∧∨¬∞∂")
    if txt_str in ops:
        return "mo"
    return "mi"

def omml_to_mathml(elem) -> str:
    """Рекурсивно конвертирует OMML (Office Math ML) в стандартный MathML."""
    if elem is None:
        return ""
    
    def convert(e):
        if e is None:
            return ""
        
        # Получаем тег без пространства имен
        tag = e.tag.replace(OMML_NS, '')
        
        if tag == 'oMath' or tag == 'oMathPara':
            children = "".join(convert(c) for c in e)
            return f'<math xmlns="http://www.w3.org/1998/Math/MathML">{children}</math>'
            
        elif tag == 'r':
            texts = []
            for t in e.findall(f'.//{OMML_NS}t'):
                if t.text:
                    texts.append(t.text)
            txt = ''.join(texts)
            if txt:
                tt = token_type(txt)
                return f'<{tt}>{html.escape(txt)}</{tt}>'
            return ''
            
        elif tag == 't':
            return html.escape(e.text or '')
            
        elif tag == 'sSub':
            base = convert(e.find(f'{OMML_NS}e'))
            sub = convert(e.find(f'{OMML_NS}sub'))
            return f'<msub><mrow>{base}</mrow><mrow>{sub}</mrow></msub>'
            
        elif tag == 'sSup':
            base = convert(e.find(f'{OMML_NS}e'))
            sup = convert(e.find(f'{OMML_NS}sup'))
            return f'<msup><mrow>{base}</mrow><mrow>{sup}</mrow></msup>'
            
        elif tag == 'sSubSup':
            base = convert(e.find(f'{OMML_NS}e'))
            sub = convert(e.find(f'{OMML_NS}sub'))
            sup = convert(e.find(f'{OMML_NS}sup'))
            return f'<msubsup><mrow>{base}</mrow><mrow>{sub}</mrow><mrow>{sup}</mrow></msubsup>'
            
        elif tag == 'f':
            num = convert(e.find(f'{OMML_NS}num'))
            den = convert(e.find(f'{OMML_NS}den'))
            return f'<mfrac><mrow>{num}</mrow><mrow>{den}</mrow></mfrac>'
            
        elif tag == 'rad':
            deg = e.find(f'{OMML_NS}deg')
            expr = convert(e.find(f'{OMML_NS}e'))
            if deg is not None:
                deg_val = convert(deg)
                return f'<mroot><mrow>{expr}</mrow><mrow>{deg_val}</mrow></mroot>'
            else:
                return f'<msqrt><mrow>{expr}</mrow></msqrt>'
                
        elif tag == 'd':
            # Delimiters (parentheses, brackets)
            dPr = e.find(f'{OMML_NS}dPr')
            beg, end = '(', ')'
            if dPr is not None:
                begChr = dPr.find(f'{OMML_NS}begChr')
                endChr = dPr.find(f'{OMML_NS}endChr')
                if begChr is not None and begChr.get('val'):
                    beg = begChr.get('val')
                if endChr is not None and endChr.get('val'):
                    end = endChr.get('val')
            content = convert(e.find(f'{OMML_NS}e'))
            return f'<mrow><mo>{html.escape(beg)}</mo><mrow>{content}</mrow><mo>{html.escape(end)}</mo></mrow>'
            
        elif tag == 'eqArr':
            rows = []
            for sub_e in e.findall(f'{OMML_NS}e'):
                rows.append(convert(sub_e))
            return f'<mtable>{"".join(f"<mtr><mtd><mrow>{r}</mrow></mtd></mtr>" for r in rows)}</mtable>'
            
        elif tag == 'nary':
            chr_elem = e.find(f'{OMML_NS}chr')
            op = chr_elem.get('val') if chr_elem is not None else '∑'
            sub = convert(e.find(f'{OMML_NS}sub'))
            sup = convert(e.find(f'{OMML_NS}sup'))
            expr = convert(e.find(f'{OMML_NS}e'))
            if sub or sup:
                s = sub if sub else '<mrow></mrow>'
                sp = sup if sup else '<mrow></mrow>'
                return f'<munderover><mrow><mo>{html.escape(op)}</mo></mrow><mrow>{s}</mrow><mrow>{sp}</mrow></munderover><mrow>{expr}</mrow>'
            else:
                return f'<mo>{html.escape(op)}</mo><mrow>{expr}</mrow>'
                
        elif tag == 'acc':
            chr_elem = e.find(f'{OMML_NS}chr')
            acc_val = chr_elem.get('val') if chr_elem is not None else '̂'
            base = convert(e.find(f'{OMML_NS}e'))
            return f'<mover><mrow>{base}</mrow><mrow><mo>{html.escape(acc_val)}</mo></mrow></mover>'
            
        elif tag == 'box':
            return f'<mrow>{convert(e.find(f"{OMML_NS}e"))}</mrow>'
            
        elif tag == 'e':
            return "".join(convert(c) for c in e)
        
        else:
            return "".join(convert(c) for c in e)

    try:
        return convert(elem)
    except Exception as ex:
        logger.warning(f"OMML conversion failed: {ex}")
        texts = [t.text for t in elem.findall(f'.//{OMML_NS}t') if t.text]
        if texts:
            return f'<span class="formula-fallback">{esc("".join(texts))}</span>'
        return ""
