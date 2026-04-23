import os

def get_head(filename, bytes_count=2000):
    if not os.path.exists(filename):
        return "Not found"
    with open(filename, 'r', encoding='utf-8') as f:
        return f.read(bytes_count)

f1 = 'monolith.html'
f2 = 'Шустов Денис Владимирович, (ПНИПУ).html'

print(f"--- {f1} ---")
print(get_head(f1))
print("\n" + "="*80 + "\n")
print(f"--- {f2} ---")
print(get_head(f2))
