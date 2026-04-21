import requests
import re
import os
from pathlib import Path

LIBS_DIR = Path("libs")
FONTS_DIR = LIBS_DIR / "fonts"

def download_file(url, target_path):
    print(f"Downloading {url} to {target_path}...")
    response = requests.get(url, stream=True)
    response.raise_for_status()
    os.makedirs(os.path.dirname(target_path), exist_ok=True)
    with open(target_path, "wb") as f:
        for chunk in response.iter_content(chunk_size=8192):
            f.write(chunk)
    print("Done.")

def localize_google_fonts(css_url):
    print(f"Localizing fonts from {css_url}...")
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    response = requests.get(css_url, headers=headers)
    response.raise_for_status()
    css_content = response.text

    # Find all font URLs
    # format: url(https://fonts.gstatic.com/s/inter/v12/...)
    font_urls = re.findall(r'url\((https?://fonts\.gstatic\.com/[^)]+)\)', css_content)
    
    localized_css = css_content
    for url in font_urls:
        # Determine family from URL
        family = "unknown"
        if "inter" in url.lower(): family = "inter"
        elif "outfit" in url.lower(): family = "outfit"
        elif "roboto" in url.lower(): family = "roboto-mono"
        
        filename = os.path.basename(url)
        target_dir = FONTS_DIR / family
        target_path = target_dir / filename
        
        download_file(url, target_path)
        
        # Replace in CSS with relative path
        rel_path = f"{family}/{filename}"
        localized_css = localized_css.replace(url, rel_path)
    
    with open(FONTS_DIR / "fonts.css", "w", encoding="utf-8") as f:
        f.write(localized_css)
    print("Fonts localization complete.")

if __name__ == "__main__":
    # 1. GSAP
    download_file("https://cdnjs.cloudflare.com/ajax/libs/gsap/3.12.2/gsap.min.js", LIBS_DIR / "gsap" / "gsap.min.js")
    
    # 2. Lucide
    # Link from official documentation: https://unpkg.com/lucide@latest
    # We use the UMD build for easy inclusion via script tag
    download_file("https://unpkg.com/lucide@latest/dist/umd/lucide.min.js", LIBS_DIR / "lucide" / "lucide.min.js")
    
    # 3. Google Fonts
    google_fonts_url = "https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;800&family=Outfit:wght@500;700&family=Roboto+Mono&display=swap"
    localize_google_fonts(google_fonts_url)
