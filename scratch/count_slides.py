import re
import os

def count_matching(filename, pattern):
    if not os.path.exists(filename):
        return "Not found"
    with open(filename, 'r', encoding='utf-8') as f:
        content = f.read()
    return len(re.findall(pattern, content))

print(f"Monolith:")
print(f"  Slides:   {count_matching('monolith.html', r'<section class=\"slide')}")
print(f"  Images:   {count_matching('monolith.html', r'<img ')}")
print(f"  Formulas: {count_matching('monolith.html', r'class=\"formula-container')}")
print(f"  Tables:   {count_matching('monolith.html', r'<table ')}")

print(f"Modular:")
print(f"  Slides:   {count_matching('modular.html', r'<section class=\"slide')}")
print(f"  Images:   {count_matching('modular.html', r'<img ')}")
print(f"  Formulas: {count_matching('modular.html', r'class=\"formula-container')}")
print(f"  Tables:   {count_matching('modular.html', r'<table ')}")
