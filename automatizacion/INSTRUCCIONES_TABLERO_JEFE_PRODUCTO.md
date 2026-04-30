# Instrucciones para el tablero del Jefe de Producto

Este documento explica cómo el **tablero del Jefe de Producto y sus Vendedores
Especializados** puede reutilizar la infraestructura del **Tablero Industrial**
para obtener todos los datos que necesita **sin duplicar credenciales ni
conexiones de red**.

Todos los proyectos viven en la misma PC (`consultorindustrial`), así que el
keyring de Windows, Python, la VPN y el Task Scheduler ya están configurados.

---

## 🎯 Perspectiva del Jefe de Producto

A diferencia del Tablero Industrial (dirigido a la gerencia, todos los grupos)
y del Tablero de la Jefa Administrativa Comercial (enfoque solo en cumplimiento
de venta), este tablero se enfoca en **un grupo de artículo específico** (el
que maneja el Jefe de Producto) y **sus Vendedores Especializados** de ese grupo.

Métricas típicas que interesan al Jefe de Producto:

| Área                          | Indicadores                                     |
|-------------------------------|-------------------------------------------------|
| **Venta del grupo**           | Venta neta mes/año, tendencia, top clientes     |
| **Desempeño de vendedores**   | Venta por vendedor especializado, ranking       |
| **Cumplimiento**              | Vendido vs presupuesto por vendedor y mes       |
| **Oportunidades abiertas**    | Pipeline del grupo, monto esperado              |
| **Cotizaciones**              | Vigentes, vencidas, tasa de cierre              |
| **Inventario y rotación**     | Stock, rotación, cobertura, ABC del grupo       |
| **Cartera**                   | Cartera vencida de clientes que compran el grupo|
| **Actividad comercial**       | Visitas y actividades de los vendedores         |

---

## 🟢 Opción recomendada — Consumir los CSVs ya generados

### Por qué esta opción

- Cero credenciales, cero VPN, cero código de red en este proyecto
- Los datos ya vienen **enriquecidos** con las dimensiones normalizadas
  (`jefe_producto_std`, `grupo_articulo_std`, `vendedor_principal_std`,
  `cliente_std`)
- Filtrar por jefe de producto es **una sola línea de pandas**
- Si el pipeline principal falla, ya hay notificación — no tienes que
  construir tu propio sistema de alertas

### Carpeta de origen

```
C:\Users\consultorindustrial\OneDrive - Hivimar\Documentos\HIVIMAR\Proyectos\SGI 2.0\Indicadores Tablero\salida_raw\
```

Todos los archivos son CSV UTF-8, separador `,`, con encabezado. Se actualizan
a diario a las 06:00 AM.

### Archivos y sus usos para el Jefe de Producto

| Archivo                          | ¿Para qué lo usa el Jefe de Producto?              |
|----------------------------------|----------------------------------------------------|
| `ventas_enriquecido.csv`         | Base de todo: venta real del grupo y por vendedor  |
| `oportunidades_enriquecido.csv`  | Pipeline abierto del grupo                         |
| `cotizaciones_enriquecido.csv`   | Cotizaciones vigentes y tasa de cierre             |
| `actividades_enriquecido.csv`    | Actividades de los vendedores especializados       |
| `visitas_enriquecido.csv`        | Visitas realizadas por los vendedores              |
| `cartera_enriquecido.csv`        | Cartera vencida de clientes del grupo              |
| `inventario_enriquecido.csv`     | Stock valorizado del grupo                         |
| `rotacion_profiler.csv`          | Rotación, cobertura, ABC — clave para el JP        |

---

## 🔑 Columnas clave que debes conocer

Todas las columnas `_std` son normalizadas (sin duplicados por tildes,
mayúsculas, variantes). Úsalas siempre para agrupar y filtrar.

### En `ventas_enriquecido.csv`

| Columna                       | Descripción                                  |
|-------------------------------|----------------------------------------------|
| `FECHA_FACTURA`               | Fecha YYYY-MM-DD                             |
| `ANIO_FACTURA`, `MES_FACTURA` | Año y mes                                    |
| `VALOR_ANTES_IVA`             | Monto neto (esta es la venta)                |
| `CANTIDAD`                    | Unidades vendidas                            |
| `CODIGO_MATERIAL`             | SKU                                          |
| `NOMBRE_MATERIAL`             | Descripción del producto                     |
| `COD_CLASE_PEDIDO`            | Excluir `Z23` (NC) para ventas netas         |
| **`jefe_producto_std`**       | ✨ Filtro principal del tablero              |
| **`grupo_articulo_std`**      | ✨ Grupo (MCO, MCT, SR, etc.)                |
| **`tipo_grupo_articulo_std`** | ✨ Tipo (Industria, Construcción, etc.)      |
| **`vendedor_principal_std`**  | ✨ Vendedor (especializado o no)             |
| **`supervisor_std`**          | ✨ Supervisor del vendedor                   |
| **`cliente_std`**             | ✨ Cliente normalizado                       |

### En `oportunidades_enriquecido.csv`

| Columna                      | Descripción                                 |
|------------------------------|---------------------------------------------|
| `id`                         | ID Odoo                                     |
| `name`                       | Título                                      |
| `stage_id`                   | Etapa pipeline                              |
| `expected_revenue`           | Monto estimado                              |
| `probability`                | % cierre                                    |
| `date_deadline`              | Fecha cierre estimada                       |
| `create_date`                | Fecha creación                              |
| **`jefe_producto_std`**      | ✨ Filtro principal                         |
| **`grupo_articulo_std`**     | ✨ Grupo                                    |
| **`vendedor_principal_std`** | ✨ Vendedor                                 |
| **`cliente_std`**            | ✨ Cliente                                  |

### En `rotacion_profiler.csv`

| Columna                  | Descripción                                  |
|--------------------------|----------------------------------------------|
| `CODIGO_MATERIAL`        | SKU                                          |
| `NOMBRE_MATERIAL`        | Descripción                                  |
| `venta_12m`              | Venta últimos 12 meses                       |
| `cantidad_12m`           | Unidades 12 meses                            |
| `stock_valorizado`       | Stock actual en $                            |
| `stock_unidades`         | Stock actual en unidades                     |
| `rotacion`               | Veces al año (venta/stock)                   |
| `cobertura_meses`        | Meses de stock al ritmo actual               |
| `dias_inventario`        | Días de stock                                |
| `clasificacion_abc`      | A / B / C según Pareto de venta              |
| **`jefe_producto_std`**  | ✨ Filtro principal                          |
| **`grupo_articulo_std`** | ✨ Grupo                                     |

---

## 📏 Definiciones estándar que debes respetar

Para que este tablero sea **consistente con el Tablero Industrial** (o sea,
que la Gerencia no vea un número distinto al que ve el JP), usa las mismas
reglas:

### Ventas netas reales

```python
ventas_netas = ventas[
    (ventas['COD_CLASE_PEDIDO'] != 'Z23') &          # sin notas de crédito
    (ventas['CODIGO_MATERIAL'].astype(str) != '6000120')  # SKU excluido
]
```

### Filtro por Jefe de Producto

```python
MI_JEFE = 'DAVID FLORES'  # Jefe de Producto titular del tablero

mis_ventas = ventas_netas[ventas_netas['jefe_producto_std'] == MI_JEFE]
```

> 💡 Si el nombre `'DAVID FLORES'` no coincide exactamente con la versión
> normalizada en los CSVs (por espacios, tildes o variantes), verifica los
> valores reales con:
> ```python
> print(sorted(ventas['jefe_producto_std'].dropna().unique()))
> ```
> y ajusta la constante `MI_JEFE` al valor exacto que aparezca.

### Vendedores Especializados

No hay una columna `es_especializado` en los CSVs. Los "Vendedores
Especializados" del JP son típicamente los que aparecen en su grupo con
venta significativa. Puedes definir tu criterio:

**Criterio A — Los que venden > X% del grupo**
```python
venta_por_vend = (mis_ventas
    .groupby('vendedor_principal_std')['VALOR_ANTES_IVA'].sum()
    .sort_values(ascending=False))
especializados = venta_por_vend[venta_por_vend > 0.05 * venta_por_vend.sum()]
```

**Criterio B — Lista manual en un Excel tuyo**

Mantén en tu Excel pequeño una hoja `vendedores_especializados` con
columnas `vendedor_principal_std` y `jefe_producto_std`. Es lo más explícito
y controlable por el JP.

---

## 🧩 Ejemplos completos — código listo para copiar

### Setup común

```python
import pandas as pd
from datetime import date

RUTA = r"C:\Users\consultorindustrial\OneDrive - Hivimar\Documentos\HIVIMAR\Proyectos\SGI 2.0\Indicadores Tablero\salida_raw"

MI_JEFE = 'DAVID FLORES'  # Jefe de Producto titular del tablero
ANIO   = date.today().year

def leer(archivo):
    return pd.read_csv(f"{RUTA}\\{archivo}", encoding='utf-8')
```

### 1. Venta del grupo — YTD vs año anterior

```python
v = leer('ventas_enriquecido.csv')
v = v[(v['COD_CLASE_PEDIDO']!='Z23') & (v['CODIGO_MATERIAL'].astype(str)!='6000120')]
v = v[v['jefe_producto_std'] == MI_JEFE]

ytd_actual   = v[v['ANIO_FACTURA']==ANIO    ]['VALOR_ANTES_IVA'].sum()
ytd_anterior = v[v['ANIO_FACTURA']==ANIO-1  ]['VALOR_ANTES_IVA'].sum()

variacion = (ytd_actual/ytd_anterior - 1) * 100 if ytd_anterior else 0
print(f"YTD {ANIO}: ${ytd_actual:,.0f}  ({variacion:+.1f}% vs {ANIO-1})")
```

### 2. Ranking de vendedores del grupo

```python
ranking = (v[v['ANIO_FACTURA']==ANIO]
    .groupby('vendedor_principal_std', as_index=False)['VALOR_ANTES_IVA'].sum()
    .sort_values('VALOR_ANTES_IVA', ascending=False)
    .rename(columns={'VALOR_ANTES_IVA':'vendido_ytd'}))
print(ranking)
```

### 3. Top 10 clientes del grupo

```python
top = (v[v['ANIO_FACTURA']==ANIO]
    .groupby('cliente_std', as_index=False)['VALOR_ANTES_IVA'].sum()
    .nlargest(10, 'VALOR_ANTES_IVA'))
print(top)
```

### 4. Pipeline abierto (oportunidades)

```python
o = leer('oportunidades_enriquecido.csv')
o = o[o['jefe_producto_std'] == MI_JEFE]

# Filtrar etapas no cerradas (ajusta según tus etapas reales de Odoo)
abiertas = o[~o['stage_id'].str.contains('Ganada|Perdida', case=False, na=False)]

pipeline_por_vendedor = (abiertas
    .groupby('vendedor_principal_std', as_index=False)
    .agg(cantidad=('id','count'),
         monto_esperado=('expected_revenue','sum'),
         monto_ponderado=('expected_revenue', lambda x: (x * abiertas.loc[x.index,'probability']/100).sum())))
print(pipeline_por_vendedor)
```

### 5. Rotación e inventario del grupo

```python
r = leer('rotacion_profiler.csv')
r = r[r['jefe_producto_std'] == MI_JEFE]

# KPIs agregados
stock_total     = r['stock_valorizado'].sum()
venta_12m_total = r['venta_12m'].sum()
rotacion_global = venta_12m_total / stock_total if stock_total else 0
cobertura_prom  = r['cobertura_meses'].mean()

print(f"Stock total: ${stock_total:,.0f}")
print(f"Rotación global: {rotacion_global:.2f} veces/año")
print(f"Cobertura promedio: {cobertura_prom:.1f} meses")

# SKUs críticos: cobertura alta (exceso) y baja (posible quiebre)
exceso  = r[r['cobertura_meses'] > 12].sort_values('stock_valorizado', ascending=False).head(20)
quiebre = r[(r['cobertura_meses'] < 1) & (r['clasificacion_abc']=='A')]
```

### 6. Cumplimiento vs presupuesto (si tienes tabla propia)

Si mantienes un Excel con presupuestos por vendedor × grupo × mes:

```python
v_mes = (v[v['ANIO_FACTURA']==ANIO]
    .groupby(['vendedor_principal_std','MES_FACTURA'], as_index=False)['VALOR_ANTES_IVA'].sum()
    .rename(columns={'VALOR_ANTES_IVA':'vendido'}))

ppto = pd.read_excel(r"<tu_excel>.xlsx", sheet_name='presupuesto_vendedor_mes')
# Columnas esperadas: vendedor_principal_std, mes, presupuesto

cump = ppto.merge(v_mes,
    left_on=['vendedor_principal_std','mes'],
    right_on=['vendedor_principal_std','MES_FACTURA'],
    how='left').fillna(0)
cump['cumplimiento_pct'] = (cump['vendido']/cump['presupuesto']*100).round(1)
```

---

## 🛡️ Guardarraíl — detectar si el pipeline principal falló

Ponlo al inicio de tu script para no armar un tablero con datos viejos:

```python
import os
from datetime import datetime, timedelta

BASE = r"C:\Users\consultorindustrial\OneDrive - Hivimar\Documentos\HIVIMAR\Proyectos\SGI 2.0\Indicadores Tablero"
ERR  = os.path.join(BASE, 'ULTIMO_ERROR.txt')
VENTAS_CSV = os.path.join(BASE, 'salida_raw', 'ventas_enriquecido.csv')

# Si hay error reciente, abortar
if os.path.exists(ERR):
    raise RuntimeError(f"Pipeline principal falló. Revisar {ERR}")

# Si el CSV tiene más de 36 horas, también abortar (por si no corrió)
edad = datetime.now() - datetime.fromtimestamp(os.path.getmtime(VENTAS_CSV))
if edad > timedelta(hours=36):
    raise RuntimeError(f"CSVs demasiado viejos ({edad}). Correr pipeline principal.")
```

---

## ⏰ Automatización (Task Scheduler)

El pipeline principal termina ~06:02 AM. Programa tu tablero a las **06:30 AM**.

Duplica `configurar_scheduler.ps1` del Tablero Industrial y ajusta:

- `$TaskName = "Tablero Jefe Producto - Actualizar"`
- `$PipelineScript = ...\<tu_script_principal>.py`
- `$HoraDiaria = "06:30"`

Registrar:
```powershell
cd <carpeta_de_tu_proyecto>
powershell -ExecutionPolicy Bypass -File configurar_scheduler.ps1
```

---

## 🟡 Alternativa — Extracción propia (si necesitas medidas custom)

Solo úsala si los CSVs enriquecidos no te alcanzan (por ejemplo, si necesitas
cortes de fechas que no están, o medidas específicas de Qlik que no exportamos).

### Archivos a copiar desde `automatizacion/` del Tablero Industrial

1. `qlik_client.py` — cliente Qlik (NTLM + WebSocket)
2. `odoo_client.py` — cliente Odoo (JSON-RPC)
3. `enriquecer_datos.py` (opcional) — para aplicar VLOOKUPs de dimensiones

### Credenciales ya están en keyring (no reconfigurar)

| Servicio                  | Usuario                             |
|---------------------------|-------------------------------------|
| `hivimar-tablero-qlik`    | `Industria`                         |
| `hivimar-tablero-odoo`    | `consultorindustrial@hivimar.com`   |

Lectura:
```python
import keyring
p = keyring.get_password('hivimar-tablero-odoo', 'consultorindustrial@hivimar.com')
```

### Ejemplo de extracción propia

```python
import sys
sys.path.insert(0, r"C:\...\Indicadores Tablero\automatizacion")

from qlik_client import fetch_ventas
from odoo_client import fetch_oportunidades
import keyring

# Qlik
qu = 'Industria'
qp = keyring.get_password('hivimar-tablero-qlik', qu)
ventas = fetch_ventas(qu, qp)

# Odoo
ou = 'consultorindustrial@hivimar.com'
op = keyring.get_password('hivimar-tablero-odoo', ou)
oport = fetch_oportunidades(ou, op)
```

### VPN Sophos

Si haces extracción propia, necesitas la VPN arriba. Dos caminos:

- Confiar en que el pipeline principal (6 AM) ya la levantó — corre a las 6:30
- Copiar la función `ensure_vpn()` de `update_tablero.py` e invocarla

---

## 📋 Checklist de implementación

**Fase 1 — Estructura mínima (1-2 horas)**
- [ ] Crear carpeta del proyecto
- [ ] Identificar nombre normalizado del Jefe de Producto
  (ejecutar `print(ventas['jefe_producto_std'].unique())`)
- [ ] Definir lista de Vendedores Especializados (manual o por criterio)
- [ ] Script Python que lea CSVs, filtre por `jefe_producto_std`, y genere el
  objeto `DB` (igual patrón que el Tablero Industrial)
- [ ] HTML que consuma `var DB = {...}`

**Fase 2 — Métricas progresivas**
- [ ] Venta del grupo (YTD, mes, vs año anterior)
- [ ] Ranking de vendedores especializados
- [ ] Top clientes
- [ ] Pipeline (oportunidades abiertas)
- [ ] Rotación e inventario (SKUs críticos)
- [ ] Cumplimiento vs presupuesto (si aplica)

**Fase 3 — Automatización (30 min)**
- [ ] Duplicar `configurar_scheduler.ps1`
- [ ] Programar a las 06:30 AM
- [ ] Probar con `Start-ScheduledTask`
- [ ] Validar logs

**Fase 4 — Publicación**
- [ ] Decidir si se sube a SharePoint/OneDrive igual que el Tablero Industrial
- [ ] Dar permisos solo al JP y a quien corresponda

---

## 🔄 Coordinación con los otros tableros

Los tres tableros conviven así:

| Tablero            | Enfoque                              | Horario  |
|--------------------|--------------------------------------|----------|
| Industrial         | Gerencia — todos los grupos          | 06:00 AM |
| Jefa Admin. Com.   | Supervisores y Vendedores (venta)    | 06:30 AM |
| Jefe de Producto   | Un grupo + sus vendedores            | 06:30 AM |

Pueden correr los tres a 06:30 sin problema (leen archivos, no se pisan entre sí).

Los tres usan la **misma fuente de datos** (los CSVs enriquecidos), así que
los números **coinciden** — no hay "venta del MCO según Gerencia" distinta a
"venta del MCO según el Jefe de Producto".

---

**Última actualización**: 2026-04-21
**Responsable**: consultorindustrial@hivimar.com
