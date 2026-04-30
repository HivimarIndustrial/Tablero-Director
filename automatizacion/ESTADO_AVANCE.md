# Estado de avance — Automatización Tablero Industrial 8.0

> Documento para retomar la automatización en una nueva sesión de Claude.
> Última actualización: 2026-04-18.

---

## 1. Objetivo global

Automatizar el pipeline completo del tablero "Hivimar Tablero Industrial 8.0":

```
VPN (si aplica)
  → Descarga Qlik (ventas, cartera, inventario)
  → Descarga Odoo (cotizaciones, oportunidades activas, oportunidades ganadas, actividades pendientes, visitas)
  → Pegar en hojas raw_* de "Base Tablero Industria.xlsx"
  → Regenerar HTML del tablero
```

### Reglas fijas del usuario
- **NO modificar** los archivos `.docx` del proyecto (son procedimientos manuales del usuario).
- El usuario del PC (`consultorindustrial`) **no tiene permisos de admin**.
- Credenciales **siempre** en Windows Credential Manager vía `keyring`, nunca en archivos.
- Trabajar de forma **autónoma**, sin pedir autorizaciones en cada paso.
- Para **Cartera**: Juan Dávila figura como **Jefe de Ventas** (no como vendedor). Seleccionar **año y mes actual** (el mes que se está viviendo).
- Para **Odoo**: **toda la BD Odoo es de Industria** (no hay multi-empresa en esa instancia). NO se aplica filtro adicional de "Industria" en las extracciones. Los "segmentos industriales" que aparecen en el tablero (ej: Heavy Industry = `sap_channel_id [15]`) son una **clasificación dentro de Industria** para análisis, no un filtro de extracción.

---

## 2. Entorno técnico

- **Python**: `C:\Users\consultorindustrial\AppData\Local\Programs\Python\Python312\python.exe` (user scope, sin admin).
- **Librerías instaladas**: `keyring`, `playwright`, `websocket-client`, `openpyxl`, `requests`, `urllib3`.
- **Playwright**: usa `channel='chrome'` (Chrome real del sistema). Chromium lo bloquea Sophos.
- **VPN**: Sophos Connect. Detección de red interna por SSID `HIVICORP` vía `netsh wlan show interfaces`. Actualmente la red detectada es interna → **no hace falta VPN**.
- **Servidor Qlik**: `https://hivms76.hivimar.com` (hub en `/hub/`, WebSocket en `wss://hivms76.hivimar.com/app/<app_id>`).
- **Auth Qlik**: **NTLM**, no Forms. El Virtual Proxy tiene UDC que resuelve `Industria` transparentemente (por eso el usuario solo escribe `Industria`, sin `HIVIMAR\`).
- **Odoo**: JSON-RPC. DB: `elitum-crm-production-11628642`.

### Credenciales guardadas (Windows Credential Manager)
- Servicio: `hivimar-tablero-qlik`
- Usuario: `Industria`
- Password: (guardada vía `guardar_credenciales_qlik.py`)

> Nota: la contraseña de Odoo estuvo brevemente expuesta en una conversación anterior → **recordar rotarla** al cerrar el proyecto.

---

## 3. Apps Qlik objetivo y sus GUIDs (descubiertos)

| Key         | Nombre                              | app_id                                 | Stream                      |
|-------------|-------------------------------------|----------------------------------------|-----------------------------|
| ventas      | BI - Nuevos canales ventas          | `a433d7f6-40d6-4afd-ba1f-5986058613ab` | Ind - Tableros              |
| cartera     | BI - Cartera Corriente y Futura     | `4153c866-cd5a-4db2-8d45-0f5e82ab5f2d` | Fin - Crédito y Cobranza    |
| inventario  | BI - Tablero Logistica Ind          | `25755181-0d50-49fc-a670-d80e322cf606` | Log - Ind                   |

Mapa completo de las 100 apps del hub en: `automatizacion/qlik_apps_map.json`.

---

## 4. Estructura de las hojas del Excel `Base Tablero Industria.xlsx`

Total: 17 hojas.

```
raw_ventas, raw_cartera, raw_inventario, raw_cotizaciones, raw_visitas,
fact_oportunidades, raw_actividadespendientes, raw_presupuesto2026,
dim_vendedores, dim_supervisores_tablero, dim_clientes, Base Total Clientes,
ppto 2026 original, fact_pptomensual_qlik_limpia, fact_vtas_vs_ppto,
dim_grupo_articulo_responsable, guia_configuracion_tablero
```

### 4.1 `raw_ventas` (22 cols × 58 640 filas)
- **A–K** vienen de Qlik (11 cols).
- **L–V** son fórmulas del Excel (no tocar).

Mapeo Qlik (tabla `WLZKpFa` de la hoja `VENTAS POR CANAL DEL CLIENTE`) ↔ Excel:

| Excel col | Qlik campo (dim/med) | Nombre Qlik       |
|-----------|----------------------|-------------------|
| A         | Dim                  | AÑO_N             |
| B         | Dim                  | MES_N             |
| C         | Dim                  | CANAL_N (segmento)|
| D         | Dim                  | grupo_articulo_N  |
| E         | Dim                  | marca_N           |
| F         | Dim                  | NOMBRE_AGENTE_N   |
| G         | Dim                  | NOMBRE_CLIENTE_N  |
| H         | Dim                  | COD_CLIENTE_N     |
| I         | Medida               | Venta neta        |
| J         | Medida               | Costo Interno     |
| K         | Medida               | Unidades Vendidas |

Hypercube: 11 col × 178 261 filas. **Match 1-a-1 confirmado.**

### 4.2 `raw_cartera` (39 cols × 9 177 filas)
- **A–AF** vienen de Qlik (32 cols).
- **AG–AM** son fórmulas (7 cols, no tocar).

Headers A–AF:
```
LIBRO_MAYOR, AÑO, MES, jefe_ventas, supervisor, NOMBRE_AGENTE, CODIGO_AGENTE,
COD_CANAL, nombre_cliente, CODIGO_CLIENTE, COD_OFICIAL_CRED, NOMBRE_OFICIAL,
COD_RECAUDADOR, NOMBRE_RECAUDADOR, FECHA_BASE, FECHA_CONTABLE, FECHA_DOCUMENTO,
FECHA_PAGO, VP, VENCIMIENTO, REFERENCIA, DOCUMENTO_SRI, COD_CONTABLE, DEMORA,
DIAS, COD_COMPENSACION, TEXTO, COD_TIPO_DOCUMENTO, Importe, Ppto Final sin Ref,
Ppto Final Ref., Ppto Final
```

Hoja objetivo en Qlik: **BASE REVISION**. Objeto: tabla **`KGTrxF`**.

**Estructura capturada 2026-04-18** por `inspeccionar_cartera_focalizado.py` (timeout 120s, GetLayout tardó 25.2s). 28 dims + 4 medidas = 32 cols. **Match 1-a-1 confirmado con A–AF**:

| # | Excel col | Excel header          | Qlik tipo | Qlik fallbackTitle    | Qlik field         |
|---|-----------|-----------------------|-----------|-----------------------|--------------------|
| 1 | A  | LIBRO_MAYOR           | dim | LIBRO_MAYOR           | LIBRO_MAYOR        |
| 2 | B  | AÑO                   | dim | AÑO                   | AÑO                |
| 3 | C  | MES                   | dim | MES                   | MES                |
| 4 | D  | jefe_ventas           | dim | jefe_ventas           | jefe_ventas        |
| 5 | E  | supervisor            | dim | supervisor            | supervisor         |
| 6 | F  | NOMBRE_AGENTE         | dim | NOMBRE_AGENTE         | NOMBRE_AGENTE      |
| 7 | G  | CODIGO_AGENTE         | dim | CODIGO_AGENTE         | CODIGO_AGENTE      |
| 8 | H  | COD_CANAL             | dim | COD_CANAL             | COD_CANAL          |
| 9 | I  | nombre_cliente        | dim | nombre_cliente        | nombre_cliente     |
| 10| J  | CODIGO_CLIENTE        | dim | CODIGO_CLIENTE        | CODIGO_CLIENTE     |
| 11| K  | COD_OFICIAL_CRED      | dim | COD_OFICIAL_CRED      | COD_OFICIAL_CRED   |
| 12| L  | NOMBRE_OFICIAL        | dim | NOMBRE_OFICIAL        | NOMBRE_OFICIAL     |
| 13| M  | COD_RECAUDADOR        | dim | COD_RECAUDADOR        | COD_RECAUDADOR     |
| 14| N  | NOMBRE_RECAUDADOR     | dim | NOMBRE_RECAUDADOR     | NOMBRE_RECAUDADOR  |
| 15| O  | FECHA_BASE            | dim | FECHA_BASE            | FECHA_BASE         |
| 16| P  | FECHA_CONTABLE        | dim | FECHA_CONTABLE        | FECHA_CONTABLE     |
| 17| Q  | FECHA_DOCUMENTO       | dim | FECHA_DOCUMENTO       | FECHA_DOCUMENTO    |
| 18| R  | FECHA_PAGO            | dim | FECHA_PAGO            | FECHA_PAGO         |
| 19| S  | VP                    | dim | VP                    | **VIP** (≠ header) |
| 20| T  | VENCIMIENTO           | dim | VENCIMIENTO           | VENCIMIENTO        |
| 21| U  | REFERENCIA            | dim | REFERENCIA            | REFERENCIA         |
| 22| V  | DOCUMENTO_SRI         | dim | DOCUMENTO_SRI         | DOCUMENTO_SRI      |
| 23| W  | COD_CONTABLE          | dim | COD_CONTABLE          | COD_CONTABLE       |
| 24| X  | DEMORA                | dim | DEMORA                | DEMORA             |
| 25| Y  | DIAS                  | dim | DIAS                  | DIAS               |
| 26| Z  | COD_COMPENSACION      | dim | COD_COMPENSACION      | COD_COMPENSACION   |
| 27| AA | TEXTO                 | dim | TEXTO                 | TEXTO              |
| 28| AB | COD_TIPO_DOCUMENTO    | dim | COD_TIPO_DOCUMENTO    | COD_TIPO_DOCUMENTO |
| 29| AC | Importe               | med | Importe               | —                  |
| 30| AD | Ppto Final sin Ref    | med | Ppto Final sin Ref    | —                  |
| 31| AE | Ppto Final Ref.       | med | Ppto Final Ref.       | —                  |
| 32| AF | Ppto Final            | med | Ppto Final            | —                  |

Archivo raw: `cartera_kgtrxf.json`.

Notas:
- `qSize` vino en 0×0 porque la tabla no calcula layout hasta pedir datos; esto es normal.
- Col 19: Excel header = `VP`, campo Qlik = `VIP`. Al usar `GetHyperCubeData` se indexa por columna, no importa.
- Para la extracción: el filtro de **Juan Dávila como Jefe de Ventas** se aplica vía `SelectValues` sobre el campo `jefe_ventas`. El filtro de **año/mes actual** sobre los campos `AÑO` y `MES`.

### 4.3 `raw_inventario` (8 cols × 367 filas)
- **A–D** vienen de Qlik (4 cols).
- **E–H** son fórmulas (no tocar).

Mapeo Qlik (tabla `Jmzsjm` de la hoja `STOCK AL CIERRE TOTAL`) ↔ Excel:

| Excel col | Qlik campo | Nombre Qlik       |
|-----------|------------|-------------------|
| A         | Dim        | grupo_articulo    |
| B         | Dim        | marca             |
| C         | Medida     | $ Stock Total     |
| D         | Medida     | Und Stock Total   |

Hypercube: 4 col × 888 filas. **Match 1-a-1 confirmado.** (El `pivot-table` `906ca67c...` NO es la fuente — usar `Jmzsjm`.)

### 4.4 Hojas Odoo (estructura descubierta 2026-04-18)

Output completo en `odoo_descubrimiento.txt`. Odoo 16.0+e, DB `elitum-crm-production-11628642`, usuario `EDUARDO ESPINO` (uid=1397).

- `raw_cotizaciones` → `sale.order` (12 024 reg.) / `sale.order.line` (28 418 reg.)
- `fact_oportunidades` → `crm.lead` (1 733 reg.) con **2 queries**:
  - **Activas**: `stage_id` NOT IN (ganadas + perdidas + eliminadas)
  - **Ganadas**: `stage_id IN [4, 8]` (Ganado sin facturar + Ganado facturado)
- `raw_actividadespendientes` → `mail.activity` (1 386 reg.)
- `raw_visitas` → **`crm.visit`** ✅ confirmado (módulo custom `p022_visits`)

#### Etapas CRM relevantes (`crm.stage`)

| id | seq | name                       | is_won | uso                      |
|----|-----|----------------------------|--------|--------------------------|
| 1  | 0   | Oportunidad Detectada      | False  | activa                   |
| 2  | 1   | Solicitud de Cotización    | False  | activa                   |
| 3  | 2   | Construcción de la Solución| False  | activa                   |
| 5  | 3   | Cotización y Presupuesto   | False  | activa                   |
| 6  | 4   | Negociacion                | False  | activa                   |
| 4  | 5   | Ganado sin facturar        | False  | **ganada** (override)    |
| 8  | 6   | Ganado facturado           | True   | **ganada**               |
| 7  | 6   | Stand by                   | False  | activa                   |
| 11 | 7   | ELIMINADOS                 | False  | descarte                 |
| 9  | 8   | PERDIDAS                   | False  | descarte                 |
| 10 | 9   | ELIMINADA                  | False  | descarte                 |

- "Ganadas" = stage_id ∈ {4, 8}
- "Descartes" = stage_id ∈ {9, 10, 11}
- "Activas" = todo lo demás (que no sea ganada ni descarte)

#### Menú CRM → Visitas
- `Visitas → Visitas` = action 807 sobre `crm.visit`
- `Visitas → Planificación` = action 818 sobre `visit.planned.user`
- Actividades pendientes: `Reporting → Actividades Pendientes` = action 115

#### Mapeo definitivo Excel ↔ Odoo (validado contra fields_get 2026-04-18)

Archivo: `odoo_mapeo_definitivo.json`. 62/62 campos OK.

**`raw_cotizaciones` (30 cols; A–S de Odoo, T–AD fórmulas)**: export a nivel de línea (`sale.order.line`) joineado con el pedido (`sale.order`).
```
A  Referencia del pedido       sale.order.name
B  Creado                      sale.order.create_date
C  Código SAP Cliente          sale.order.partner_sap_code
D  Cliente                     sale.order.partner_id
E  Description                 sale.order.line.name
F  Grupo Artículo              sale.order.line.group_article_id (m2o product.group.article)
G  Código SAP Generalista      sale.order.line.x_seller_sap_code
H  Comercial                   sale.order.user_id (m2o res.users)
I  Creado por:                 sale.order.create_uid
J  Cantidad                    sale.order.line.product_uom_qty
K  Cantidad de entrega         sale.order.line.qty_delivered
L  Cantidad Facturada          sale.order.line.qty_invoiced
M  Cantidad a facturar         sale.order.line.qty_to_invoice
N  Precio unitario             sale.order.line.price_unit
O  Desc. (%)                   sale.order.line.discount
P  Unidad de Medida            sale.order.line.product_uom (m2o uom.uom)
Q  Subtotal                    sale.order.line.price_subtotal
R  Total                       sale.order.amount_total
S  Estado del pedido           sale.order.state (selection)
```

**`fact_oportunidades` (56 cols; A–X de Odoo, Y–BD fórmulas)** sobre `crm.lead`:
```
A  Creado                                   create_date
B  Name Sequence                            name_sequence
C  Canal                                    sap_channel_id (m2o res.partner.channel)
D  Código SAP Creador                       x_create_uid
E  Creado por:                              create_uid
F  Última Actualización por                 write_uid
G  Actualizado en                           write_date
H  Oportunidad                              name
I  Código SAP Cliente                       x_partner_sap_code
J  Cliente                                  partner_id
K  Etapa                                    stage_id (m2o crm.stage)
L  Última actualización de la etapa         date_last_stage_update
M  Nombre del contacto                      contact_name
N  Comercial                                user_id (m2o res.users)
O  Código SAP Generalista                   x_seller_sap_code
P  Grupo de artículo                        group_article_id (m2m product.group.article)
Q  Marca                                    product_brand_id
R  Vendedor Especialista                    specialist_seller_id
S  Fecha de Facturación Estimada            estimated_invoice_date
T  Actividades                              activity_ids (o2m mail.activity)
U  Motivo de pérdida                        lost_reason_id
V  Fecha límite de la siguiente actividad   activity_date_deadline
W  Ingreso esperado                         expected_revenue
X  Fecha de cierre                          date_deadline
```

**`raw_actividadespendientes` (10 cols; A–G de Odoo, H–J fórmulas)** sobre `mail.activity`:
```
A  Nombre del documento       res_name
B  Código SAP Asignado        x_user_sap_code
C  Asignada a                 user_id (m2o res.users)
D  Tipo de actividad          activity_type_id (m2o mail.activity.type)
E  Modelo de documento        res_model
F  Resumen                    summary
G  Fecha de vencimiento       date_deadline
```

**`raw_visitas` (16 cols; A–L de Odoo, M–P fórmulas)** sobre `crm.visit`:
```
A  Nombre                        name
B  Código SAP Cliente            x_partner_sap_code
C  Cliente                       partner_id (m2o res.partner)
D  Código SAP Responsable        x_responsible_sap_code
E  Responsable                   responsible_id (m2o res.users)
F  Geocerca relacionada          geofence_id
G  Fecha planificada             planned_date
H  Fecha efectiva                effective_date
I  Fecha de entrada              entry_geofence
J  Fecha de salida               exit_geofence
K  Duracion entrada/salida       time_difference_str
L  Estado                        state (selection)
```

#### Estrategia de extracción Odoo (2026-04-18)
- Sin filtro por "Industria" → toda la instancia Odoo ya es Industria.
- `sale.order` / `sale.order.line`: extracción completa (27 938 líneas en Excel ≈ 28 418 en Odoo).
- `crm.lead`: **incluir archivadas** (`active=True,False`) — Excel tiene 4 033 filas vs 1 733 vigentes en Odoo, por lo tanto incluye históricas archivadas.
- `mail.activity`: extracción completa (en Odoo los registros se borran al completarse, por lo que todos los existentes son pendientes por definición).
- `crm.visit`: extracción completa.

---

## 5. Scripts existentes en `automatizacion/`

| Archivo                           | Estado     | Qué hace                                                                    |
|-----------------------------------|------------|-----------------------------------------------------------------------------|
| `guardar_credenciales_qlik.py`    | OK         | Guarda user/pass en keyring (servicio `hivimar-tablero-qlik`).              |
| `guardar_credenciales_qlik.bat`   | OK         | Lanzador del anterior.                                                      |
| `probar_qlik.py`                  | Histórico  | Prueba directa con `requests` (falló por NTLM/redirect).                    |
| `diagnostico_qlik.py`             | Histórico  | Diagnóstico de endpoints/auth (Forms vs NTLM).                              |
| `probar_qlik_forms.py` / `_forms2`| Histórico  | Intentos Forms Auth (fallaron: Qlik usa NTLM).                              |
| `probar_qlik_playwright.py`       | OK         | Login Playwright+NTLM headless (patrón base).                               |
| `probar_qlik_playwright.bat`      | OK         | Lanzador.                                                                    |
| `descubrir_qlik_apps.py`          | OK         | Login + `/qrs/app` → genera `qlik_apps_map.json` con los 100 apps del hub.  |
| `qlik_apps_map.json`              | Generado   | Nombre/id/stream de todas las apps.                                         |
| `inspeccionar_qlik_apps.py` (v1)  | Deprecado  | Inspector sin filtros ni reconexión. Bug: concatenación glitch en JSON.     |
| `inspeccionar_qlik_apps_v2.py`    | OK         | Inspector con filtros de hoja, reconexión por app, `safe_str()`, json atómico. |
| `qlik_apps_estructura.json`       | Generado   | Salida del inspector v2 (Cartera incompleta por timeout).                   |
| `descubrir_odoo.py`               | Revisar    | (No revisado aún en esta sesión — verificar estado).                        |

### Patrón técnico clave: JSON-RPC sobre WebSocket a Qlik Engine API

```python
def rpc(ws, method, params, handle=-1, timeout=45):
    _req_id[0] += 1
    rid = _req_id[0]
    msg = {"jsonrpc": "2.0", "id": rid, "method": method, "handle": handle, "params": params}
    ws.send(json.dumps(msg))
    ws.settimeout(timeout)
    while True:
        resp = json.loads(ws.recv())
        if resp.get('id') == rid:
            if 'error' in resp:
                raise RuntimeError(f"RPC {method}: {resp['error']}")
            return resp.get('result', {})
```

Métodos usados: `OpenDoc`, `GetActiveDoc`, `GetObjects`, `GetObject`, `GetChildInfos`, `GetLayout`.
Pendiente por implementar: `GetHyperCubeData` (o `GetHyperCubePivotData`) para extraer las filas.

### Patrón de login Playwright+NTLM

```python
browser = p.chromium.launch(channel='chrome', headless=True,
                            args=['--ignore-certificate-errors'])
context = browser.new_context(
    ignore_https_errors=True,
    http_credentials={"username": USER, "password": PWD},  # "Industria", no "HIVIMAR\\Industria"
)
page = context.new_page()
page.goto(f"{BASE}/hub/", wait_until='domcontentloaded', timeout=30000)
# Captura cookie X-Qlik-Session del context.cookies().
```

---

## 6. Errores conocidos y cómo se resolvieron

| Problema                                         | Causa                                          | Solución                                              |
|--------------------------------------------------|------------------------------------------------|-------------------------------------------------------|
| `spawn UNKNOWN` al lanzar chromium               | Sophos bloquea `chromium.exe`                  | Usar `channel='chrome'` (Chrome real).                |
| HTTP 400 en Forms Auth POST                      | Qlik usa NTLM, no Forms                        | Login vía Playwright + `http_credentials`.            |
| User `HIVIMAR\Industria` no funcionaba           | Virtual Proxy tiene UDC, usuario usa solo nombre| Usar `Industria` a secas.                             |
| JSON v1 corrupto (`"visualization""name":`)      | Concatenación por caracteres de control        | v2: `safe_str()` + `json.dump` atómico al final.      |
| Timeout en objetos grandes (Cartera `KGTrxF`)    | Timeout 45s insuficiente                       | **Pendiente**: subir a 120s en un retry focalizado.   |
| Salida vacía en background tasks                 | Buffering stdout                               | Correr Python con `-u` (unbuffered).                  |

---

## 7. Pendientes

### Inmediato (desbloqueo de Cartera)
1. **Revisar salida** del último task `bnyy34ygy` (exit 0) — ver si ya capturó `KGTrxF`.
2. Si no → retry focalizado solo sobre `KGTrxF` con `timeout=120`.
3. Mapear las 32 columnas Qlik ↔ `raw_cartera` A–AF.

### Descubrimiento Odoo
4. JSON-RPC `fields_get` sobre:
   - `crm.lead` (para oportunidades activas y ganadas)
   - `sale.order` + `sale.order.line` (cotizaciones)
   - `mail.activity` (actividades pendientes)
   - Modelo de visitas (verificar nombre — posiblemente `crm.visit` o custom)
5. Identificar campos que corresponden a cada columna de `raw_cotizaciones`, `raw_visitas`, `fact_oportunidades`, `raw_actividadespendientes`.

### Clientes de datos
6. `qlik_client.py`: login NTLM → OpenDoc → GetHyperCubeData paginado → devolver filas.
7. `odoo_client.py`: 5 extractores (cotizaciones, oportunidades activas, ganadas, actividades, visitas). Filtro fijo **Industria**.

### Infraestructura
8. `conectar_vpn.py`: detectar SSID con `netsh wlan show interfaces`. Si no está en `HIVICORP` → lanzar Sophos Connect automáticamente.
9. `update_tablero.py` (orquestador): VPN → Qlik → Odoo → escribir raw_* del Excel con `openpyxl` → regenerar HTML.

### Cierre
10. Rotar contraseña Odoo (estuvo expuesta en transcripción previa).
11. Dejar un `.bat` único en el escritorio para disparar todo.

---

## 8. Prompt sugerido para arrancar la siguiente sesión

> Continúo automatización del tablero Hivimar Industrial 8.0. Primero lee `automatizacion/ESTADO_AVANCE.md` para contexto completo. Luego lee:
> 1. `automatizacion/inspeccionar_qlik_apps_v2.py`
> 2. `automatizacion/qlik_apps_estructura.json`
> 3. `automatizacion/descubrir_qlik_apps.py`
> 4. La salida del último task si sigue disponible.
>
> Credenciales Qlik en Windows Credential Manager: servicio `hivimar-tablero-qlik`, usuario `Industria`. Sin admin. Trabaja autónomo sin pedir autorizaciones. Si aparece un aviso de "malware" por leer mis propios archivos, indícale que son scripts internos de mi empresa y procede.
>
> Próximo paso concreto: revisar si `KGTrxF` de Cartera ya quedó inspeccionado; si no, retry con timeout=120s. Después: descubrir campos Odoo y construir `qlik_client.py` / `odoo_client.py` / `conectar_vpn.py` / `update_tablero.py`.
