import io
import base64
from PIL import Image
try:
    import cairosvg
except ImportError:
    cairosvg = None

def _save_image_with_white_bg(image_bytes: bytes, output_path: str) -> bool:
    """Обеспечивает белый фон для прозрачных изображений и сохраняет их по указанному пути."""
    try:
        img = Image.open(io.BytesIO(image_bytes))
        if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
            background = Image.new("RGB", img.size, (255, 255, 255))
            if img.mode == 'P':
                img = img.convert('RGBA')
            background.paste(img, mask=img.split()[3])
            background.save(output_path, "PNG")
        else:
            with open(output_path, "wb") as f:
                f.write(image_bytes)
        return True
    except Exception as e:
        print(f"[ImageHelper] Error: Failed to process image background: {e}")
        return False

def _get_data_uri(image_bytes: bytes, ext: str) -> str:
    """Преобразует байты изображения в Data URI для встраивания в HTML."""
    mime_map = {
        '.png': 'image/png',
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.gif': 'image/gif',
        '.svg': 'image/svg+xml',
        '.emf': 'image/x-emf',
        '.wmf': 'image/x-wmf'
    }
    mime = mime_map.get(ext.lower(), 'application/octet-stream')
    b64 = base64.b64encode(image_bytes).decode('utf-8')
    return f"data:{mime};base64,{b64}"
