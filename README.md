# Tablero Director Industria

Tablero ejecutivo (versión Director) consumido por Jaime Echeverría y Juan Dávila.

URL pública: `https://hivimarindustrial.github.io/Tablero-Director/`

## Acceso

El HTML publicado en `docs/index.html` está **cifrado con AES-256-GCM**. Al abrir
la URL, se pide usuario y clave; sin la clave correcta el navegador no puede
descifrar los datos.

- Cada usuario tiene su propia clave (definida en `claves.json`, fuera del repo).
- La clave maestra se deriva con PBKDF2-SHA256 (200.000 iteraciones).
- Nada sensible viaja en texto claro al repo público.

## Pipeline (uso interno)

```
update_tablero.py        → extrae Qlik+Odoo, enriquece, regenera HTMLs
cifrar_y_publicar.py     → cifra Tablero_Director_Industria.html → docs/index.html
git add docs/ ; git commit ; git push   (publica en GitHub Pages)
```

## Estructura del repo

```
automatizacion/          # pipeline ETL Python (Qlik + Odoo + enriquecimiento)
docs/                    # ← lo único público: HTML cifrado para GitHub Pages
cifrar_y_publicar.py     # script de cifrado + inyección de login
claves.json.template     # plantilla de usuarios (copiar a claves.json)
.gitignore               # excluye Excel, salida_raw, HTMLs en claro, claves
```

## Rotar claves

1. Editar `claves.json` con las nuevas claves.
2. `python cifrar_y_publicar.py`
3. `git add docs/index.html && git commit -m "rotar claves" && git push`
