# Guía de edición de diseño — Tablero Industrial Hivimar

Esta guía explica cómo hacer cambios **estéticos** al archivo
`Tablero_Director_FUENTE.html` sin romper el pipeline automático
que lo actualiza con datos de Qlik y Odoo.

Pensada para:
- Trabajar desde otro PC con Claude Code
- Cambiar logo, colores, fuentes, textos visibles, emojis
- Probar ajustes visuales antes de dejarlos definitivos

---

## 🚦 Regla de oro

> **Toca solo CSS, texto visible y el logo. Nunca toques el JavaScript
> ni los `id=""` de elementos HTML ni la línea `var DB = ...;`**

Si sigues esta regla, tus cambios van a sobrevivir a todas las corridas
del pipeline y el tablero seguirá funcionando perfecto.

---

## 📋 Prompt listo para pegarle a Claude Code

Cuando abras Claude Code en otro PC con este archivo, copia y pega este
prompt **al inicio de la conversación**:

```
Voy a trabajar en cambios de DISEÑO VISUAL sobre el archivo
"Tablero_Director_FUENTE.html". Antes de tocar nada, lee la
guía "GUIA_EDICION_DISENO.md" que está en la misma carpeta para
entender qué puedes y qué no puedes modificar.

Reglas innegociables:
1. NO modifiques la línea que empieza con "var DB =" (≈línea 460).
2. NO modifiques ninguna función JavaScript (function ...).
3. NO cambies los atributos id="..." de ningún elemento HTML.
4. SÍ puedes cambiar: el bloque <style> al inicio, el logo
   (<img id="logo">), textos visibles, emojis de botones, textos
   de títulos de tarjetas (<div class="ct">...</div>) y headers de
   tablas (<th>...</th>).
5. Haz un backup del archivo antes de empezar (copialo como
   "Tablero_Director_FUENTE.preview_YYYYMMDD.html").
6. Cuando termines, dime exactamente qué cambiaste para poder
   verificar que no haya efectos colaterales.

Los cambios que quiero son: [AQUÍ DESCRIBES LO QUE QUIERES CAMBIAR]
```

---

## ✅ Zonas 100 % seguras para editar

### 1. Bloque `<style>` al inicio del `<head>`

Ubicación aproximada: líneas 9 – 180.

Ahí vive todo el CSS. Cambios típicos sin riesgo:

```css
:root {
  --navy: #1B2F6E;       /* Color principal del header — cámbialo */
  --navy2: #2a4494;      /* Navy secundario */
  --red: #8B2020;        /* Color de alertas */
  --accent: #0071a7;     /* Color de acento */
  --green: #16a34a;      /* Verde de cumplimiento */
  --amber: #d97706;      /* Ámbar de atención */
  --danger: #dc2626;     /* Rojo de crítico */
  --bg: #eef0f6;         /* Fondo general */
  --surface: #fff;       /* Fondo de tarjetas */
  --text: ...;           /* Color de texto */
  ...
}

body {
  font-family: 'DM Sans', sans-serif;  /* Cámbiala si quieres otra fuente */
}
```

### 2. Logo

Ubicación: `<img id="logo" src="..." alt="Hivimar">` — buscar en el HTML.

El logo es una **imagen embebida en base64** dentro del atributo `src="data:image/...;base64,...."`.

Para cambiarlo:
- Codifica tu nueva imagen a base64 (hay conversores online gratuitos).
- Reemplaza el valor completo del atributo `src`.
- Mantén el atributo `id="logo"` **sin cambios**.

**Regla**: NO cambies el `id="logo"`, solo el contenido del `src`.

### 3. Textos visibles en el header

```html
<div class="hdr-t">Tablero Comercial
  <div class="hdr-sub">Hivimar Industrial</div>
</div>
```

Puedes cambiar "Tablero Comercial" o "Hivimar Industrial" por otros
textos. No toques las clases `hdr-t` ni `hdr-sub`.

### 4. Nombres de pestañas (tabs)

```html
<button class="sb on" data-sec="g" data-id="gv">📊 Ventas</button>
<button class="sb" data-sec="g" data-id="gp">🎯 Presupuesto</button>
```

Puedes cambiar:
- El emoji (📊 → 💼)
- El texto visible ("Ventas" → "Ingresos")

**NO cambies** ni `data-sec`, ni `data-id`, ni `class="sb"`.

### 5. Títulos de tarjetas

```html
<div class="ct">Venta por vendedor
  <span class="ctag">vs presupuesto comercial</span>
</div>
```

Puedes cambiar los textos visibles. No toques las clases.

### 6. Headers de columnas

```html
<thead>
  <tr>
    <th>Vendedor</th><th>Supervisor</th><th>Venta neta</th>
    <th>%MB</th><th>Ppto Com.</th><th>Cumpl.</th>
  </tr>
</thead>
```

Puedes cambiar los textos entre `<th>...</th>`. No agregues ni quites columnas
(rompería la tabla porque el JS pone datos en un número específico de columnas).

---

## ⚠️ Zonas que se pueden tocar con cuidado

### Textos dentro de funciones JavaScript

A veces hay textos visibles dentro del JS, como:

```javascript
kpi('Venta neta', fK(tot), 'período seleccionado', 'kpi kn')
kpi('Margen bruto', pct(mg), fK(tot-totC), 'kpi kn')
```

Puedes cambiar `'Venta neta'` por `'Ventas'` o el texto descriptivo
`'período seleccionado'` por otra cosa. Pero **respeta las comillas**
y **no toques** las variables (`tot`, `mg`, `fK`, `pct`, etc.).

**Recomendación**: si no te sientes cómodo tocando JS, pídele a Claude
Code que lo haga por ti indicando el texto exacto que quieres cambiar.

---

## 🚫 Zonas que NUNCA se deben tocar

| Zona | Ubicación aprox. | Por qué |
|------|------------------|---------|
| Línea `var DB = {...};` | Línea 460 | El pipeline la reemplaza en cada corrida. Si la quitas o renombras, el tablero deja de actualizarse. |
| Función `render()` | Línea 1893 | Orquesta qué vista se muestra. |
| Funciones `rGV`, `rGP`, `rSV`, `rVV`, `rPP`, etc. | Líneas 908-1890 | Son las que llenan de datos todas las tablas y gráficos. |
| Función `DOMContentLoaded` al final | Líneas 1960-2100 | Inicia todo el tablero. |
| Línea que actualiza `srcLbl` | Línea 2091 | Muestra la fecha de actualización y los volúmenes. |
| Todos los atributos `id="..."` | Varios | Son referencias que el JavaScript busca. |
| Etiquetas `<canvas id="...">` | Varios | Los gráficos se renderizan ahí. |

---

## 🎨 Ejemplos de cambios seguros

### Cambiar el color principal de azul a verde corporativo

En `<style>`:

```css
:root {
  --navy: #0b6b3a;   /* antes #1B2F6E */
  --navy2: #0e8b4b;  /* antes #2a4494 */
}
```

Un solo cambio, se refleja en todo el header y acentos.

### Cambiar la fuente

```css
body {
  font-family: 'Inter', sans-serif;   /* antes 'DM Sans' */
}
```

Asegúrate de que la fuente nueva esté importada arriba:

```html
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
```

### Cambiar el logo por uno nuevo

Buscar `<img id="logo" src="data:image/...">`, reemplazar el valor
del `src` por el base64 del nuevo logo.

### Cambiar texto del header

```html
<div class="hdr-t">Dashboard Comercial
  <div class="hdr-sub">Hivimar Industrial — Edición 2026</div>
</div>
```

### Cambiar emojis de los botones

```html
<button class="sb on" data-sec="g" data-id="gv">💼 Ventas</button>
<button class="sb" data-sec="g" data-id="gp">💵 Presupuesto</button>
```

---

## 🛡️ Protocolo obligatorio antes de editar

### 1. Backup

Antes de cualquier cambio, haz una copia:

```powershell
# En PowerShell
$fecha = Get-Date -Format "yyyyMMdd_HHmmss"
Copy-Item "Tablero_Director_FUENTE.html" "Tablero_Director_FUENTE.preview_$fecha.html"
```

O simplemente: clic derecho → Copiar → Pegar → renombra el pegado.

### 2. Verifica que OneDrive esté sincronizado

- Ícono de nube en la barra de tareas **verde/check**: estás actualizado.
- Ícono **azul rotando**: espera a que termine, o pausa la sync antes de
  editar (clic derecho en el ícono → "Pausar sincronización" → "2 horas").

### 3. Edita

Usa Claude Code con el prompt de arriba o edita manualmente en un
editor de texto que respete encoding UTF-8 (VS Code, Sublime, Notepad++,
etc. — **NO uses Bloc de notas de Windows clásico** porque puede romper caracteres).

### 4. Verifica que funciona

Abre el HTML editado en el navegador:
1. **Debe cargar el tablero completo** (si queda "cargando" infinitamente, algo se rompió).
2. **Navega todas las pestañas** (Gerencia, Supervisores, Vendedores, CRM, Perspectiva Producto) y verifica que las tablas y gráficos se vean.
3. **Verifica que arriba a la derecha aparezca**: `● Actualizado: DD/MM/YYYY HH:MM · XX cotiz | XX opp | XX cartera`. Si dice solo "—", el JS no está ejecutándose.
4. **Abre la consola del navegador** (F12 → pestaña "Consola"). Si hay errores en rojo, algo se rompió.

### 5. Si algo falla

Cierra el archivo SIN guardar. Si ya lo guardaste, restaura el backup:

```powershell
Copy-Item "Tablero_Director_FUENTE.preview_YYYYMMDD_HHMMSS.html" "Tablero_Director_FUENTE.html" -Force
```

---

## ⏰ Timing ideal para editar

El pipeline corre automáticamente (Task Scheduler) y modifica el HTML.
Para evitar conflictos de OneDrive:

- **Horario seguro para editar**: entre corridas del pipeline. Si corre
  a las 6 AM, edita después de las 6:15 AM (cuando ya sincronizó).
- **Nunca edites mientras OneDrive sincroniza activamente**: espera a
  que termine.
- **Si vas a editar mucho rato**: pausa OneDrive en tu PC mientras
  editas, termina, guarda, y luego reanuda para que suba todo junto.

---

## 🆘 Qué hacer si rompes algo

### Escenario 1: el tablero no carga (pantalla en blanco o loader infinito)

1. Abre consola del navegador (F12).
2. Busca el error en rojo. Los más comunes:
   - `Uncaught SyntaxError: ...` → rompiste la sintaxis JS. Restaura backup.
   - `Cannot read property 'X' of undefined` → eliminaste algo que el JS busca. Restaura backup.
3. Si no lo puedes arreglar, **restaura el backup** y pide ayuda.

### Escenario 2: el tablero carga pero falta información

- Tablas vacías → seguramente cambiaste un `id=""`. Restaura backup.
- Gráficos no aparecen → tocaste un `<canvas id="...">`. Restaura backup.
- Fecha no aparece → modificaste la línea 2091 o la línea 183 del HTML.

### Escenario 3: OneDrive pide resolver conflicto

Si aparece `Tablero_Director_FUENTE-NombrePC.html`:

1. **No borres ni uno ni otro** todavía.
2. Abre ambos en un editor de texto.
3. El que tiene los datos más recientes (mira la línea `var DB = ...`
   y busca `last_update`) es el que el pipeline actualizó.
4. Copia tus cambios de diseño al archivo con datos recientes.
5. Guarda como `Tablero_Director_FUENTE.html` (el original).
6. Borra el archivo `-NombrePC.html`.

En caso de duda, **siempre prioriza el que tiene datos más recientes**
(el del pipeline) y vuelve a aplicar tus cambios de diseño encima.

---

## 🧪 Cambios que siempre deberías probar antes de dejar definitivos

Abre el HTML ya editado y verifica:

- [ ] Carga el tablero sin errores en consola
- [ ] Se ven los 5 menús superiores (Gerencia, Supervisores, Vendedores, CRM, Perspectiva Producto)
- [ ] La fecha de actualización aparece arriba a la derecha
- [ ] Los KPIs (venta, margen, pipeline, cotizaciones, tasa de cierre) muestran números
- [ ] Los gráficos se renderizan (barras, donut, líneas)
- [ ] Las tablas tienen filas con datos
- [ ] Los filtros (año, mes, supervisor, vendedor) funcionan
- [ ] Click en "Supervisores" → pestaña "Ventas" muestra la tabla "Cumplimiento por grupo de artículo"
- [ ] Click en "Vendedores" → pestaña "Ventas" muestra la misma tabla
- [ ] El botón "Oscuro" alterna tema claro/oscuro
- [ ] Se ve bien en celular (abrir desde el móvil)

Si todos los checks pasan, tus cambios son seguros.

---

## 📞 Contacto

Si después de leer esta guía algo no queda claro o rompes el tablero:
- Responsable del pipeline: **Eduardo Espino**
- Correo: `consultorindustrial@hivimar.com`
- Backups del HTML: hay `Tablero_Director_FUENTE.backup_YYYYMMDD_HHMMSS.html`
  en la misma carpeta (el pipeline los genera automáticamente).

---

## 📁 Archivos relacionados

En la misma carpeta donde está esta guía:

- `Tablero_Director_FUENTE.html` ← el archivo a editar
- `Tablero_Director_FUENTE.backup_*.html` ← backups automáticos del pipeline
- `automatizacion/DOCUMENTACION_EXTRACCION.md` ← cómo extraer datos (para otro proyecto)
- `db_output.js` ← datos generados por el pipeline (no editar)
- `regenerar_db.py` ← script que genera `db_output.js` (no editar desde otro PC)
- `update_html.py` ← script que inyecta el DB en el HTML (no editar desde otro PC)

---

**Última actualización de esta guía**: 2026-04-18
