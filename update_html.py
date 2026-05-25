"""
Replace the var DB = ...; line in the HTML with the new DB from db_output.js
"""
import os

# Rutas relativas al script para que corra en cualquier PC
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
HTML_FILE = os.path.join(SCRIPT_DIR, "Tablero_Director_FUENTE.html")
DB_FILE = os.path.join(SCRIPT_DIR, "db_output.js")
OUTPUT_FILE = HTML_FILE  # se sobrescribe in-place

print("Leyendo HTML...")
with open(HTML_FILE, 'r', encoding='utf-8') as f:
    lines = f.readlines()

print(f"  {len(lines)} líneas")

print("Leyendo nuevo DB...")
with open(DB_FILE, 'r', encoding='utf-8') as f:
    new_db = f.read()

print(f"  {len(new_db)} caracteres")

# Find and replace the DB line (line 409, 0-indexed = 408)
db_line_idx = None
for i, line in enumerate(lines):
    if line.strip().startswith('var DB') or line.strip().startswith('var DB='):
        db_line_idx = i
        print(f"  Encontrada línea DB en índice {i} (línea {i+1})")
        break

if db_line_idx is None:
    print("ERROR: No se encontró la línea var DB")
    exit(1)

# Replace the line
lines[db_line_idx] = new_db + '\n'

print(f"Escribiendo {OUTPUT_FILE}...")
with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
    f.writelines(lines)

size_mb = os.path.getsize(OUTPUT_FILE) / (1024*1024)
print(f"  Tamaño: {size_mb:.1f} MB")
print("Done!")
