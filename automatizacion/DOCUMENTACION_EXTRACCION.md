# Extracción de datos — Tablero Industrial Hivimar

Documentación para que otros proyectos de indicadores (internos de Hivimar)
puedan consumir los mismos datos que alimentan el Tablero Industrial, o
reproducir la extracción de manera independiente.

**Última actualización**: 2026-04-18

---

## Índice

1. [Arquitectura del pipeline actual](#1-arquitectura-del-pipeline-actual)
2. [Cómo consumir los datos YA listos (CSVs)](#2-cómo-consumir-los-datos-ya-listos-csvs)
3. [Cómo reproducir la extracción desde cero](#3-cómo-reproducir-la-extracción-desde-cero)
4. [Fuentes de datos originales](#4-fuentes-de-datos-originales)
5. [Credenciales y autenticación](#5-credenciales-y-autenticación)
6. [Filtros de negocio aplicados](#6-filtros-de-negocio-aplicados)
7. [Estructura de cada dataset](#7-estructura-de-cada-dataset)
8. [Zona horaria y formatos de fecha](#8-zona-horaria-y-formatos-de-fecha)
9. [Diccionario de correcciones de nombres](#9-diccionario-de-correcciones-de-nombres)
10. [Buenas prácticas de consumo](#10-buenas-prácticas-de-consumo)

---

## 1. Arquitectura del pipeline actual

```
Sophos VPN  →  Qlik Sense (hivms76.hivimar.com)  ─┐
                                                   │
                   Odoo (hivimar-crm.odoo.com)  ───┤
                                                   │
                                                   ▼
                            Pipeline Python (automatizacion/)
                            ├─ descarga y filtra
                            ├─ aplica correcciones de nombres
                            ├─ exporta CSVs a salida_raw/     ◄─── aquí lees tú
                            ├─ escribe Excel raw_*
                            ├─ regenera db_output.js
                            └─ inyecta HTML del tablero
```

El orquestador es `automatizacion/update_tablero.py`. Corre en la
máquina `consultorindustrial` (Windows 11) bajo Task Scheduler.

---

## 2. Cómo consumir los datos YA listos (CSVs)

**Esta es la forma más simple.** Si tu proyecto puede leer archivos CSV, usa esto.

### 2.1 Ubicación

Después de cada corrida del pipeline, los datos quedan en:

```
C:\Users\consultorindustrial\OneDrive - Hivimar\Documentos\HIVIMAR\
Proyectos\SGI 2.0\Indicadores Tablero\salida_raw\
    ├── ventas.csv
    ├── cartera.csv
    ├── inventario.csv
    ├── cotizaciones.csv
    ├── oportunidades.csv
    ├── actividades.csv
    ├── visitas.csv
    └── ultima_actualizacion.txt
```

Si sincronizas el OneDrive, los tienes también en tu PC.

### 2.2 Formato

- Encoding UTF-8
- Separador: `,` (coma)
- Primera fila: headers
- Fechas: `YYYY-MM-DD HH:MM:SS` (hora local Ecuador UTC-5)
- Números: punto decimal, sin separador de miles

### 2.3 Ejemplo de lectura en Python

```python
import pandas as pd
df = pd.read_csv(
    r'C:\Users\consultorindustrial\OneDrive - Hivimar\...\salida_raw\ventas.csv',
    encoding='utf-8'
)
print(df.shape, df.columns.tolist())
```

### 2.4 Frecuencia de actualización

Los CSVs se regeneran en cada corrida del pipeline. Si el pipeline está
programado para correr diariamente a las 6 AM, los CSVs están frescos
cada mañana.

Verifica `ultima_actualizacion.txt` para saber el timestamp exacto.

---

## 3. Cómo reproducir la extracción desde cero

Si tu proyecto necesita extraer directamente (sin depender de este
pipeline), puedes reutilizar los módulos ya construidos.

### 3.1 Dependencias Python

Instala con el Python del usuario (sin admin):

```bash
pip install keyring playwright websocket-client requests openpyxl urllib3
python -m playwright install chrome
```

### 3.2 Módulos reutilizables

Los archivos principales en `automatizacion/`:

| Archivo                  | Qué hace                                                                           |
|--------------------------|------------------------------------------------------------------------------------|
| `qlik_client.py`         | Login NTLM + descarga hypercubes de Qlik con filtros. Métodos: `fetch_all()`, `fetch_table()`.  |
| `odoo_client.py`         | Auth + `search_read` de los 4 modelos (sale.order, crm.lead, mail.activity, crm.visit). Método `fetch_all()`. |
| `conectar_vpn.py`        | Detecta SSID / estado VPN Sophos. Levanta conexión si hace falta.                   |
| `update_tablero.py`      | Orquestador. Llama a VPN → Qlik → Odoo → Excel → HTML → CSVs.                      |

### 3.3 Uso programático mínimo

```python
# Ejemplo: solo necesitas ventas y cartera de Qlik
from conectar_vpn import ensure_vpn
from qlik_client import login_and_get_cookie, fetch_table, APPS, _resolve_selections

ensure_vpn()  # asegura VPN arriba si no estamos en HIVICORP
cookie = login_and_get_cookie()

# Descargar ventas
cfg = APPS['ventas']
sel = _resolve_selections(cfg['selections'])
headers, rows = fetch_table(cfg['app_id'], cfg['obj_id'], cookie,
                             expected_cols=cfg['n_cols'], selections=sel)
print(f'Ventas: {len(rows)} filas, {len(headers)} cols')
```

### 3.4 Solo Odoo

```python
from odoo_client import OdooClient, extract_oportunidades

cli = OdooClient().authenticate()
data = extract_oportunidades(cli)  # dict con headers y rows
```

---

## 4. Fuentes de datos originales

### 4.1 Qlik Sense

- **Servidor**: `https://hivms76.hivimar.com`
- **Auth**: NTLM (Virtual Proxy tiene UDC; user corporativo sin dominio)
- **Protocolo**: JSON-RPC sobre WebSocket a `wss://.../app/<app_id>`
- **Apps usadas**:

| Dataset      | App nombre                          | App GUID                                 | Objeto/tabla |
|--------------|-------------------------------------|------------------------------------------|--------------|
| ventas       | BI - Nuevos canales ventas          | `a433d7f6-40d6-4afd-ba1f-5986058613ab`   | `WLZKpFa`    |
| cartera      | BI - Cartera Corriente y Futura     | `4153c866-cd5a-4db2-8d45-0f5e82ab5f2d`   | `KGTrxF`     |
| inventario   | BI - Tablero Logistica Ind          | `25755181-0d50-49fc-a670-d80e322cf606`   | `Jmzsjm`     |

### 4.2 Odoo (CRM + ERP)

- **URL**: `https://hivimar-crm.odoo.com`
- **DB**: `elitum-crm-production-11628642`
- **Versión**: Odoo 16.0+e (enterprise)
- **Auth**: JSON-RPC con `web/session/authenticate`
- **Modelos usados**:

| Dataset        | Modelo Odoo                                     | Notas                     |
|----------------|-------------------------------------------------|---------------------------|
| cotizaciones   | `sale.order` + `sale.order.line` (join por id)   | Export a nivel de línea   |
| oportunidades  | `crm.lead` con `active=False` (incluye archivadas) | Pipeline + histórico     |
| actividades    | `mail.activity`                                 | Pendientes por defecto    |
| visitas        | `crm.visit` (módulo custom `p022_visits`)        |                           |

---

## 5. Credenciales y autenticación

**Nunca hardcodear credenciales en código.** Siempre usar Windows
Credential Manager vía la librería `keyring`.

### 5.1 Servicios keyring requeridos

| Servicio                  | Usuario                              | Uso                        |
|---------------------------|--------------------------------------|----------------------------|
| `hivimar-tablero-qlik`    | `Industria`                          | Login Qlik NTLM            |
| `hivimar-tablero-odoo`    | `consultorindustrial@hivimar.com`    | Login Odoo JSON-RPC        |
| `hivimar-tablero-vpn`     | (opcional, fallback a qlik)          | Sophos Connect             |

### 5.2 Lectura en código

```python
import keyring
cred = keyring.get_credential('hivimar-tablero-qlik', None)
user = cred.username
password = cred.password
```

### 5.3 Guardado de credenciales nuevas

Hay scripts helpers en `automatizacion/`:

- `guardar_credenciales_qlik.py`
- (Puede crearse uno equivalente para Odoo si hace falta.)

---

## 6. Filtros de negocio aplicados

Estos filtros están documentados para que el otro proyecto aplique
exactamente los mismos criterios.

### 6.1 Qlik

| Dataset     | Filtro                                                          |
|-------------|------------------------------------------------------------------|
| ventas      | `AÑO_N IN [2024, 2025, <año_actual>]` — rango móvil desde 2024 |
| cartera     | `jefe_ventas IN ['JUAN DAVILA', 'JUAN BLADIMIR DAVILA CHACON']` + `AÑO = <año_actual>` + `MES = <mes_actual>` |
| inventario  | `AÑO = <año_actual>` + `MES = <mes_actual>` (snapshot del mes)  |

### 6.2 Odoo

- **Ninguna filtración por "Industria"**: la BD Odoo entera ya es de
  Industria (instancia dedicada).
- `crm.lead`: usar `active_test=False` en el contexto para incluir
  archivadas. De lo contrario, se pierde el histórico.
- `sale.order.line`: extracción completa. Lo join con `sale.order`
  permite enriquecer con campos del pedido.
- `mail.activity`: todas las pendientes (Odoo las borra al completarse
  automáticamente).
- `crm.visit`: todas las visitas, el estado indica Realizado/Planificado.

---

## 7. Estructura de cada dataset

### 7.1 Qlik → CSV

#### `ventas.csv` (11 cols)
```
AÑO, MES, CANAL, grupo_articulo, marca, NOMBRE_AGENTE,
NOMBRE_CLIENTE, COD_CLIENTE, Venta neta, Costo Interno, Unidades Vendidas
```

#### `cartera.csv` (32 cols)
```
LIBRO_MAYOR, AÑO, MES, jefe_ventas, supervisor, NOMBRE_AGENTE,
CODIGO_AGENTE, COD_CANAL, nombre_cliente, CODIGO_CLIENTE,
COD_OFICIAL_CRED, NOMBRE_OFICIAL, COD_RECAUDADOR, NOMBRE_RECAUDADOR,
FECHA_BASE, FECHA_CONTABLE, FECHA_DOCUMENTO, FECHA_PAGO, VP,
VENCIMIENTO, REFERENCIA, DOCUMENTO_SRI, COD_CONTABLE, DEMORA, DIAS,
COD_COMPENSACION, TEXTO, COD_TIPO_DOCUMENTO,
Importe, Ppto Final sin Ref, Ppto Final Ref., Ppto Final
```

#### `inventario.csv` (4 cols)
```
grupo_articulo, marca, $ Stock Total, Und Stock Total
```

### 7.2 Odoo → CSV

#### `cotizaciones.csv` (19 cols)
Basado en `sale.order + sale.order.line`. Ver mapeo completo en
`odoo_mapeo_definitivo.json`.

Campos clave: Referencia del pedido, Creado, Código SAP Cliente,
Cliente, Description, Grupo Artículo, Comercial, Cantidad, Subtotal,
Total, Estado del pedido.

#### `oportunidades.csv` (24 cols)
Basado en `crm.lead`. Campos clave: Creado, Name Sequence, Canal,
Código SAP Cliente, Cliente, Etapa, Oportunidad, Comercial,
Vendedor Especialista, Ingreso esperado, Fecha de cierre.

#### `actividades.csv` (7 cols)
Basado en `mail.activity`. Campos: Nombre del documento, Código SAP
Asignado, Asignada a, Tipo de actividad, Modelo de documento,
Resumen, Fecha de vencimiento.

#### `visitas.csv` (12 cols)
Basado en `crm.visit`. Campos: Nombre, Código SAP Cliente, Cliente,
Código SAP Responsable, Responsable, Geocerca relacionada,
Fecha planificada, Fecha efectiva, Fecha de entrada, Fecha de salida,
Duracion entrada/salida, Estado.

---

## 8. Zona horaria y formatos de fecha

- **Qlik**: fechas vienen en hora local del servidor (Ecuador UTC-5).
  No requieren conversión.
- **Odoo**: devuelve datetimes en **UTC**. El pipeline los convierte a
  `America/Guayaquil` (UTC-5 sin DST) antes de exportar.
- Formatos finales en CSV:
  - Fecha: `YYYY-MM-DD`
  - Fecha con hora: `YYYY-MM-DD HH:MM:SS`

---

## 9. Diccionario de correcciones de nombres

El pipeline aplica un diccionario de correcciones para uniformizar
nombres de vendedores/clientes que puedan venir mal escritos desde
los sistemas origen. Actualmente:

```python
CORRECCIONES_NOMBRES = {
    'EDUAROD CHILAN': 'EDUARDO CHILAN',
}
```

**Si tu proyecto encuentra otros typos**, coordina con el responsable
del pipeline (`consultorindustrial@hivimar.com`) para agregarlos al
diccionario, así ambos proyectos los manejan igual.

---

## 10. Buenas prácticas de consumo

### 10.1 Verificar frescura

Lee `ultima_actualizacion.txt` antes de procesar:

```python
from datetime import datetime, timedelta
import os
p = os.path.join(SALIDA_RAW, 'ultima_actualizacion.txt')
with open(p) as f:
    ts = datetime.fromisoformat(f.read().strip())
if datetime.now() - ts > timedelta(hours=24):
    print("ALERTA: datos tienen más de 24h de antigüedad")
```

### 10.2 Joins con las dim

Las hojas `dim_*` siguen viviendo en `Base Tablero Industria.xlsx`.
Si necesitas enriquecer los datos raw con vendedor_principal_std,
segmento, supervisor, etc., léelas en modo read-only:

```python
import pandas as pd
dim = pd.read_excel(EXCEL, sheet_name='dim_vendedores', engine='openpyxl')
```

### 10.3 No tocar el Excel

**No escribir** al archivo `Base Tablero Industria.xlsx` desde otro
proyecto. El pipeline principal lo maneja. Escribir con openpyxl puede
destruir los objetos Tabla del Excel.

Si tu proyecto necesita datos derivados (agregaciones, cálculos),
genera tu propio archivo de salida aparte.

### 10.4 Manejo de errores

Si los CSVs no están presentes o están corruptos:
- Verifica `ultima_actualizacion.txt`
- Revisa los logs del pipeline en `automatizacion/logs/`
- Si urge, corre el pipeline manualmente:
  ```
  python automatizacion/update_tablero.py --backup
  ```

---

## Contacto

Para preguntas sobre este pipeline o pedir extensiones:
- **Responsable**: Eduardo Espino
- **Correo**: `consultorindustrial@hivimar.com`
- **Repositorio**: (indicar Git/OneDrive path)
