"""
Cliente Qlik Sense para el tablero Hivimar Industrial 8.0.

Funciones principales:
  - login_and_get_cookie(): Playwright + NTLM -> cookie X-Qlik-Session
  - fetch_table(app_id, obj_id, cookie, page_size=None) -> (headers, rows)
       Abre la app, localiza el objeto (tabla) y descarga TODAS las filas
       paginando con GetHyperCubeData.
  - fetch_all(): descarga los 3 raw_* (ventas, cartera, inventario)
       y los devuelve como dict {key: {'headers': [...], 'rows': [[...]]}}.

Se usa en update_tablero.py.
"""
import keyring
import math
import os
import sys
import json
import ssl
import time
from datetime import date
from typing import Optional, Dict, List

from playwright.sync_api import sync_playwright
import urllib3
from websocket import create_connection

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

SERVICE = 'hivimar-tablero-qlik'
BASE = 'https://hivms76.hivimar.com'
WS_BASE = 'wss://hivms76.hivimar.com'

# Configuracion de los 3 extractos de Qlik.
# 'selections' es un dict {field_name: [str_values]} que se aplica con SelectValues
# antes de leer el hypercube. Usar '_dynamic_current_period' para inyectar
# AnIO/MES actuales en tiempo de extraccion.
APPS = {
    'ventas': {
        'app_id': 'a433d7f6-40d6-4afd-ba1f-5986058613ab',
        'app_name': 'BI - Nuevos canales ventas',
        'sheet_title_contains': 'VENTAS POR CANAL',
        'obj_id': 'WLZKpFa',
        # 2026-05-19: sistemas anadio columna NOMBRE_DIGITAR_N (etiqueta UI
        # 'nombre_digitador') a la tabla del sheet -> ahora son 12 columnas.
        # La columna nueva se propaga a ventas_enriquecido.csv para otros
        # proyectos que la consumen.
        'n_cols': 12,
        'selections': {
            # Desde 2024 hasta el anio actual
            'AÑO_N': '_dynamic_years_from_2024',
        },
    },
    'cartera': {
        'app_id': '4153c866-cd5a-4db2-8d45-0f5e82ab5f2d',
        'app_name': 'BI - Cartera Corriente y Futura',
        'sheet_title_contains': 'BASE REVISION',
        'obj_id': 'KGTrxF',
        'n_cols': 32,
        'selections': {
            'jefe_ventas': ['JUAN DAVILA', 'JUAN BLADIMIR DAVILA CHACON'],
            'AÑO': '_dynamic_current_year',
            'MES': '_dynamic_current_month',
        },
    },
    'inventario': {
        'app_id': '25755181-0d50-49fc-a670-d80e322cf606',
        'app_name': 'BI - Tablero Logistica Ind',
        'sheet_title_contains': 'STOCK AL CIERRE TOTAL',
        'obj_id': 'Jmzsjm',
        'n_cols': 4,
        'selections': {
            'AÑO': '_dynamic_current_year',
            'MES': '_dynamic_current_month',
        },
    },
}


def _resolve_selections(raw: dict) -> Dict[str, List[str]]:
    """Expande tokens _dynamic_* a listas concretas de valores (strings)."""
    today = date.today()
    out = {}
    for field, val in raw.items():
        if val == '_dynamic_years_from_2024':
            out[field] = [str(y) for y in range(2024, today.year + 1)]
        elif val == '_dynamic_current_year':
            out[field] = [str(today.year)]
        elif val == '_dynamic_current_month':
            out[field] = [str(today.month)]
        elif isinstance(val, list):
            out[field] = [str(v) for v in val]
        else:
            out[field] = [str(val)]
    return out

MAX_CELLS_PER_REQ = 10_000  # limite Qlik Engine API
DEFAULT_LOGIN_TIMEOUT_MS = 90_000
DEFAULT_RPC_TIMEOUT = 120


def _cell_value(cell: dict):
    """
    Devuelve el valor a escribir en Excel para una celda de Qlik HyperCube.
    Regla: si qNum es un numero finito, devolver qNum (float / int); si no, qText.
    Qlik usa NaN (float) para celdas text-only o vacias.
    """
    q = cell.get('qNum')
    if isinstance(q, (int, float)) and not (isinstance(q, float) and math.isnan(q)):
        # Si es entero representable, devolver int para Excel
        if isinstance(q, float) and q.is_integer() and abs(q) < 1e15:
            return int(q)
        return q
    return cell.get('qText', '')


def _get_credential():
    cred = keyring.get_credential(SERVICE, None)
    if not cred:
        raise SystemExit(f"ERROR: credenciales Qlik no guardadas en keyring (servicio '{SERVICE}')")
    return cred.username, cred.password


def login_and_get_cookie(verbose: bool = True, max_retries: int = 3) -> str:
    """Login headless via Playwright+NTLM. Devuelve el valor de X-Qlik-Session.
    Reintenta hasta max_retries veces en caso de Chrome timeout intermitente."""
    user, pwd = _get_credential()
    last_err = None
    for attempt in range(1, max_retries + 1):
        if verbose:
            print(f"[qlik] Login NTLM headless a {BASE} (usuario={user}) "
                  f"intento {attempt}/{max_retries}...")
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(
                    channel='chrome', headless=True,
                    args=[
                        '--ignore-certificate-errors',
                        '--no-proxy-server',
                        '--proxy-server=direct://',
                        '--proxy-bypass-list=*',
                        '--disable-extensions',
                    ],
                )
                context = browser.new_context(
                    ignore_https_errors=True,
                    http_credentials={"username": user, "password": pwd},
                )
                page = context.new_page()
                page.goto(f"{BASE}/hub/", wait_until='domcontentloaded',
                          timeout=DEFAULT_LOGIN_TIMEOUT_MS)
                if 'authentication' in page.url.lower():
                    browser.close()
                    raise SystemExit("[qlik] Login NTLM fallo (quedamos en pagina auth)")
                cookies = context.cookies()
                browser.close()
            for c in cookies:
                if c['name'] == 'X-Qlik-Session':
                    if verbose:
                        print(f"[qlik] cookie X-Qlik-Session obtenida")
                    return c['value']
            raise RuntimeError("no se encontro cookie X-Qlik-Session")
        except Exception as e:
            last_err = e
            if verbose:
                print(f"[qlik] fallo intento {attempt}: {type(e).__name__}: {str(e)[:150]}")
            if attempt < max_retries:
                time.sleep(3)
    raise SystemExit(f"[qlik] login fallo tras {max_retries} intentos: {last_err}")


class _QlikSession:
    """Maneja una conexion WebSocket a una app, con un contador de request id."""
    def __init__(self, app_id: str, session_cookie: str,
                 timeout: int = DEFAULT_RPC_TIMEOUT,
                 max_retries: int = 2):
        self.app_id = app_id
        self.cookie = session_cookie
        self.timeout = timeout
        self._req_id = 0
        url = f"{WS_BASE}/app/{app_id}"
        headers = [f"Cookie: X-Qlik-Session={session_cookie}"]
        ssl_opt = {"cert_reqs": ssl.CERT_NONE}
        last_err = None
        for attempt in range(max_retries + 1):
            try:
                self.ws = create_connection(url, header=headers, sslopt=ssl_opt,
                                             timeout=timeout)
                # Consumir el OnConnected notification (obligatorio, en algunos casos
                # el server cierra la conexion si no se consume antes de enviar OpenDoc).
                try:
                    self.ws.settimeout(10)
                    _ = self.ws.recv()
                except Exception:
                    pass
                return
            except Exception as e:
                last_err = e
                if attempt < max_retries:
                    time.sleep(2)
        raise RuntimeError(f"[qlik] no pude abrir WS para app {app_id}: {last_err}")

    def rpc(self, method: str, params, handle: int = -1,
            timeout: Optional[int] = None) -> dict:
        self._req_id += 1
        rid = self._req_id
        msg = {"jsonrpc": "2.0", "id": rid, "method": method,
               "handle": handle, "params": params}
        self.ws.send(json.dumps(msg))
        self.ws.settimeout(timeout or self.timeout)
        while True:
            raw = self.ws.recv()
            if not raw:
                # Mensaje vacio (ping/pong/control frame) - ignorar
                continue
            if isinstance(raw, bytes):
                try:
                    raw = raw.decode('utf-8')
                except Exception:
                    continue
            try:
                resp = json.loads(raw)
            except json.JSONDecodeError:
                # Mensaje no-JSON, ignorar
                continue
            if resp.get('id') == rid:
                if 'error' in resp:
                    raise RuntimeError(f"RPC {method}: {resp['error']}")
                return resp.get('result', {})

    def close(self):
        try:
            self.ws.close()
        except Exception:
            pass


def _apply_selections(sess: '_QlikSession', app_handle: int,
                      selections: Dict[str, List[str]], verbose: bool = True):
    """
    Aplica selecciones {field: [values]} usando Field.Select con busqueda
    numerica/texto explicita. Para multiples valores usa Select con
    expresion OR encerrada en parentesis, que Qlik interpreta como
    busqueda de multiples terminos.
    """
    for field, values in selections.items():
        rf = sess.rpc("GetField", [field], handle=app_handle)
        fh = rf['qReturn']['qHandle']

        # Construir lista de qFieldValue detectando si cada valor es numerico
        qvalues = []
        for v in values:
            sv = str(v)
            try:
                num = float(sv)
                qvalues.append({'qText': sv, 'qIsNumeric': True, 'qNumber': num})
            except (ValueError, TypeError):
                qvalues.append({'qText': sv, 'qIsNumeric': False, 'qNumber': 0})

        # SelectValues(qFieldValues, qToggleMode=False, qSoftLock=False)
        r = sess.rpc("SelectValues", [qvalues, False, False], handle=fh)
        ok = r.get('qReturn', False)
        if verbose:
            kinds = 'num' if all(q['qIsNumeric'] for q in qvalues) else 'txt'
            print(f"[qlik] SelectValues {field}[{kinds}]={values} -> {ok}")
        if not ok:
            raise RuntimeError(f"SelectValues({field}={values}) devolvio False")


def fetch_table(app_id: str, obj_id: str, session_cookie: str,
                expected_cols: Optional[int] = None,
                selections: Optional[Dict[str, List[str]]] = None,
                verbose: bool = True) -> tuple:
    """
    Abre la app, (opcional) aplica selecciones, obtiene el objeto tabla
    y descarga TODAS las filas con GetHyperCubeData paginado.

    Devuelve (headers, rows) donde:
      headers = [nombre_col_1, ...]
      rows    = [[val1, val2, ...], ...]
    """
    sess = _QlikSession(app_id, session_cookie)
    try:
        t0 = time.time()
        res = sess.rpc("OpenDoc", [app_id], handle=-1)
        app_handle = res['qReturn']['qHandle']
        if verbose:
            print(f"[qlik/{app_id[:8]}] OpenDoc en {time.time()-t0:.1f}s")

        if selections:
            _apply_selections(sess, app_handle, selections, verbose=verbose)

        res = sess.rpc("GetObject", [obj_id], handle=app_handle)
        obj_handle = res['qReturn']['qHandle']

        layout = sess.rpc("GetLayout", [], handle=obj_handle)
        hc = layout.get('qLayout', {}).get('qHyperCube', {})

        dim_info = hc.get('qDimensionInfo', [])
        mea_info = hc.get('qMeasureInfo', [])
        all_titles = [d.get('qFallbackTitle', '') for d in dim_info] + \
                     [m.get('qFallbackTitle', '') for m in mea_info]
        # Qlik puede aplicar 'qColumnOrder' para reordenar columnas en la
        # tabla visual; cuando esta presente, el GetHyperCubeData devuelve
        # los datos en ese orden, no en dims-then-measures natural.
        # Por eso alineamos headers al mismo orden que las filas.
        col_order = hc.get('qColumnOrder') or []
        if col_order and len(col_order) == len(all_titles):
            headers = [all_titles[i] for i in col_order]
        else:
            headers = all_titles
        n_cols = len(headers)

        if expected_cols is not None and n_cols != expected_cols:
            raise RuntimeError(
                f"[qlik/{obj_id}] columnas inesperadas: obtenidas={n_cols}, esperadas={expected_cols}"
            )

        # qSize viene del layout; a veces viene 0 si la tabla no se ha calculado.
        # En ese caso pedimos la primera pagina para forzar calculo y leer qSize.
        n_rows_layout = hc.get('qSize', {}).get('qcy', 0)
        if verbose:
            print(f"[qlik/{obj_id}] layout: {n_cols}cols x {n_rows_layout}rows (segun qSize)")

        # Paginacion
        page_height = max(1, MAX_CELLS_PER_REQ // n_cols)
        rows_all = []
        top = 0
        while True:
            pages = [{
                "qLeft": 0, "qTop": top,
                "qWidth": n_cols, "qHeight": page_height,
            }]
            t_page = time.time()
            res = sess.rpc("GetHyperCubeData",
                           ["/qHyperCubeDef", pages],
                           handle=obj_handle)
            data_pages = res.get('qDataPages', [])
            if not data_pages:
                break
            matrix = data_pages[0].get('qMatrix', [])
            if not matrix:
                break
            for r in matrix:
                rows_all.append([_cell_value(cell) for cell in r])
            if verbose:
                print(f"[qlik/{obj_id}] pagina top={top} -> +{len(matrix)} filas "
                      f"(acum {len(rows_all)}) en {time.time()-t_page:.1f}s")
            if len(matrix) < page_height:
                # ultima pagina
                break
            top += page_height

        if verbose:
            print(f"[qlik/{obj_id}] descarga completa: {len(rows_all)} filas en {time.time()-t0:.1f}s")
        return headers, rows_all
    finally:
        sess.close()


def _previous_month(y: int, m: int) -> tuple:
    """Devuelve (anio, mes) del mes anterior con rollover de anio."""
    if m == 1:
        return (y - 1, 12)
    return (y, m - 1)


def fetch_all(session_cookie: Optional[str] = None, verbose: bool = True) -> dict:
    """Descarga las 3 tablas (ventas, cartera, inventario). Devuelve dict con keys.

    Para 'cartera': si la consulta del mes actual devuelve 0 filas (caso comun
    a inicio de mes / dias festivos cuando aun no se carga el cierre), se
    reintenta automaticamente con el mes anterior (con rollover de anio).
    """
    if session_cookie is None:
        session_cookie = login_and_get_cookie(verbose=verbose)
    out = {}
    for key, cfg in APPS.items():
        if verbose:
            print()
            print(f"=== [qlik] {key}: {cfg['app_name']} ===")
        selections = _resolve_selections(cfg.get('selections') or {})
        headers, rows = fetch_table(cfg['app_id'], cfg['obj_id'],
                                    session_cookie,
                                    expected_cols=cfg['n_cols'],
                                    selections=selections,
                                    verbose=verbose)

        # Fallback solo para cartera cuando viene vacia: reintentar con el
        # mes anterior. Las apps de cartera publican el cierre con desfase,
        # asi que los primeros dias del mes / despues de festivos puede no
        # haber datos del mes en curso.
        if key == 'cartera' and not rows and 'AÑO' in selections and 'MES' in selections:
            try:
                y = int(selections['AÑO'][0])
                m = int(selections['MES'][0])
                py, pm = _previous_month(y, m)
                fb = dict(selections)
                fb['AÑO'] = [str(py)]
                fb['MES'] = [str(pm)]
                if verbose:
                    print(f"[qlik/cartera] mes actual vacio -> reintentando "
                          f"con mes anterior AÑO={py} MES={pm}")
                headers, rows = fetch_table(cfg['app_id'], cfg['obj_id'],
                                            session_cookie,
                                            expected_cols=cfg['n_cols'],
                                            selections=fb,
                                            verbose=verbose)
                if verbose:
                    print(f"[qlik/cartera] fallback {py}-{pm:02d}: "
                          f"{len(rows)} filas")
            except Exception as e:
                if verbose:
                    print(f"[qlik/cartera] fallback fallo: {e}")

        out[key] = {'headers': headers, 'rows': rows}
    return out


# ===========================================================================
# Extraccion ROTACION (desde SKU PROFILER - ventas + stock de TODA la empresa)
# ===========================================================================

PROFILER_APP_ID = '7889fe5c-5a65-4a07-ace0-75398aec5ddd'  # BI - SKU PROFILER

# Expresiones set-analysis tomadas del pivot original de la app.
# Excluyen COD_CLASE_PEDIDO=Z23 y CODIGO_MATERIAL=6000120, y años 2017-2018 para stock.
PROFILER_MEASURES = [
    {'label': 'Costo',
     'qDef': "(SUM({<COD_CLASE_PEDIDO-={'Z23'},CODIGO_MATERIAL-= {'6000120'} >}COSTO_INTERNO))"},
    {'label': 'Cantidad',
     'qDef': "(SUM({<COD_CLASE_PEDIDO-={'Z23'},CODIGO_MATERIAL-= {'6000120'} >}CANTIDAD))"},
    {'label': 'Venta',
     'qDef': "(SUM({<COD_CLASE_PEDIDO-={'Z23'},CODIGO_MATERIAL-= {'6000120'} >}VALOR_ANTES_IVA))"},
    {'label': 'Stock',
     'qDef': ("(SUM({<[AÑO]-={2017,2018}>}ECS_LIBRE_UTILICACION) "
              "+ SUM({<[AÑO]-={2017,2018}>}ECS_BLOQUEADO) "
              "+ SUM({<[AÑO]-={2017,2018}>}ECS_TRASLADO) "
              "+ SUM({<[AÑO]-={2017,2018}>} ECS_CALIDAD))")},
]


def _ultimos_13_meses(hoy=None):
    if hoy is None:
        hoy = date.today()
    out = []
    y, m = hoy.year, hoy.month
    for _ in range(13):
        out.append((y, m))
        m -= 1
        if m == 0:
            m = 12
            y -= 1
    return out


def fetch_rotacion(session_cookie: Optional[str] = None,
                   dims: List[str] = None, verbose: bool = True) -> dict:
    """
    Extrae de BI - SKU PROFILER las ventas + stock de toda la empresa para
    los ultimos 13 meses, agrupado por `dims` (default: grupo_articulo, marca,
    AÑO, MES).
    Devuelve: {'headers': [...], 'rows': [...]}
    """
    if session_cookie is None:
        session_cookie = login_and_get_cookie(verbose=verbose)
    if dims is None:
        dims = ['grupo_articulo', 'marca', 'AÑO', 'MES']

    if verbose:
        print(f"[qlik/rotacion] abriendo SKU PROFILER ({PROFILER_APP_ID[:8]}...)")

    sess = _QlikSession(PROFILER_APP_ID, session_cookie)
    try:
        t0 = time.time()
        res = sess.rpc("OpenDoc", [PROFILER_APP_ID], handle=-1)
        ah = res['qReturn']['qHandle']
        if verbose:
            print(f"[qlik/rotacion] OpenDoc en {time.time()-t0:.1f}s")

        # Seleccionar los anios de los ultimos 13 meses
        anios = sorted(set(y for y, _ in _ultimos_13_meses()))
        rf = sess.rpc("GetField", ['AÑO'], handle=ah)
        fh = rf['qReturn']['qHandle']
        qvals = [{'qText': str(a), 'qIsNumeric': True, 'qNumber': float(a)} for a in anios]
        r = sess.rpc("SelectValues", [qvals, False, False], handle=fh)
        if verbose:
            print(f"[qlik/rotacion] SelectValues AÑO={anios} -> {r.get('qReturn')}")

        # Crear session object con las dims pedidas + medidas del profiler
        hc_def = {
            'qInfo': {'qType': 'custom_rotacion'},
            'qHyperCubeDef': {
                'qDimensions': [{'qDef': {'qFieldDefs': [d]}} for d in dims],
                'qMeasures': [{'qDef': {'qDef': m['qDef'], 'qLabel': m['label']}}
                              for m in PROFILER_MEASURES],
                'qInitialDataFetch': [],
                'qSuppressZero': True,
                'qSuppressMissing': True,
            }
        }
        res = sess.rpc("CreateSessionObject", [hc_def], handle=ah, timeout=120)
        oh = res['qReturn']['qHandle']
        lay = sess.rpc("GetLayout", [], handle=oh, timeout=300)
        hc = lay.get('qLayout', {}).get('qHyperCube', {})
        size = hc.get('qSize', {})
        n_cols = size.get('qcx', 0)
        n_rows = size.get('qcy', 0)
        if verbose:
            print(f"[qlik/rotacion] tamano total: {n_cols} col x {n_rows} fil")

        headers = [d.get('qFallbackTitle', '') for d in hc.get('qDimensionInfo', [])] + \
                  [m.get('qFallbackTitle', '') for m in hc.get('qMeasureInfo', [])]

        # Paginacion
        page_height = max(1, MAX_CELLS_PER_REQ // max(1, n_cols))
        rows_all = []
        top = 0
        while top < n_rows:
            pages = [{"qLeft": 0, "qTop": top, "qWidth": n_cols, "qHeight": page_height}]
            t_page = time.time()
            res = sess.rpc("GetHyperCubeData", ["/qHyperCubeDef", pages],
                           handle=oh, timeout=120)
            data_pages = res.get('qDataPages', [])
            if not data_pages:
                break
            matrix = data_pages[0].get('qMatrix', [])
            if not matrix:
                break
            for row in matrix:
                rows_all.append([_cell_value(c) for c in row])
            if verbose:
                print(f"[qlik/rotacion] pagina top={top} -> +{len(matrix)} "
                      f"(acum {len(rows_all)}) en {time.time()-t_page:.1f}s")
            if len(matrix) < page_height:
                break
            top += page_height

        if verbose:
            print(f"[qlik/rotacion] descarga completa: {len(rows_all)} filas en {time.time()-t0:.1f}s")
        return {'headers': headers, 'rows': rows_all}
    finally:
        sess.close()


# ===========================================================================
# Extraccion SEGMENTACION CLIENTES (desde la misma app de ventas)
# ===========================================================================
# La app 'BI - Nuevos canales ventas' tiene el campo SEGMENTO_CLIENTE con
# ~37 valores detallados (Agricola, Agroindustrial, Alimentos/Bebidas,
# Automotriz, Farmaceutica, Forestal/Madera, Mineria/Construccion, Pesquera,
# Petroleo Privado, Plastico/Caucho, Puertos, etc). Lo extraemos por
# cod_cliente para que enriquecer_datos.py pueda agregarlo como columna
# segmento_cliente_detalle a todas las tablas.

VENTAS_APP_ID = 'a433d7f6-40d6-4afd-ba1f-5986058613ab'  # BI - Nuevos canales ventas


def fetch_segmentacion_clientes(session_cookie: Optional[str] = None,
                                 verbose: bool = True) -> dict:
    """
    Extrae de la app de ventas un mapeo cod_cliente -> SEGMENTO_CLIENTE.
    Devuelve: {'headers': ['cod_cliente','SEGMENTO_CLIENTE'], 'rows': [...]}
    """
    if session_cookie is None:
        session_cookie = login_and_get_cookie(verbose=verbose)

    if verbose:
        print(f"[qlik/segmento] abriendo Ventas ({VENTAS_APP_ID[:8]}...)")

    sess = _QlikSession(VENTAS_APP_ID, session_cookie)
    try:
        t0 = time.time()
        res = sess.rpc("OpenDoc", [VENTAS_APP_ID], handle=-1)
        ah = res['qReturn']['qHandle']
        if verbose:
            print(f"[qlik/segmento] OpenDoc en {time.time()-t0:.1f}s")

        # IMPORTANTE: 'cod_cliente' (minuscula) NO es un campo real en la
        # app, solo CODIGO_CLIENTE (mayuscula) lo es. Tambien hace falta
        # una medida tecnica para forzar el join (sin medida Qlik solo
        # devuelve los 37 segmentos sueltos sin cliente).
        dims = ['CODIGO_CLIENTE', 'SEGMENTO_CLIENTE']
        hc_def = {
            'qInfo': {'qType': 'custom_segmento'},
            'qHyperCubeDef': {
                'qDimensions': [{'qDef': {'qFieldDefs': [d]}} for d in dims],
                'qMeasures': [
                    {'qDef': {'qDef': 'Sum(VALOR_ANTES_IVA)',
                              'qLabel': '_venta_aux'}}
                ],
                'qInitialDataFetch': [],
                'qSuppressZero': True,
                'qSuppressMissing': True,
            }
        }
        res = sess.rpc("CreateSessionObject", [hc_def], handle=ah, timeout=120)
        oh = res['qReturn']['qHandle']
        lay = sess.rpc("GetLayout", [], handle=oh, timeout=300)
        hc = lay.get('qLayout', {}).get('qHyperCube', {})
        size = hc.get('qSize', {})
        n_cols = size.get('qcx', 0)
        n_rows = size.get('qcy', 0)
        if verbose:
            print(f"[qlik/segmento] tamano total: {n_cols} col x {n_rows} fil")

        # Headers: solo dims (no incluimos la medida tecnica en el output).
        # Renombramos CODIGO_CLIENTE -> cod_cliente para que enriquecer_datos
        # siga reconociendo la columna sin tener que cambiar nada mas alla.
        headers = [d.get('qFallbackTitle', '') for d in hc.get('qDimensionInfo', [])]
        headers = ['cod_cliente' if h == 'CODIGO_CLIENTE' else h for h in headers]

        page_height = max(1, MAX_CELLS_PER_REQ // max(1, n_cols))
        rows_all = []
        top = 0
        n_dims = len(hc.get('qDimensionInfo', []))
        while top < n_rows:
            pages = [{"qLeft": 0, "qTop": top, "qWidth": n_cols, "qHeight": page_height}]
            res = sess.rpc("GetHyperCubeData", ["/qHyperCubeDef", pages],
                           handle=oh, timeout=120)
            data_pages = res.get('qDataPages', [])
            if not data_pages:
                break
            matrix = data_pages[0].get('qMatrix', [])
            if not matrix:
                break
            for row in matrix:
                # Descartar la medida tecnica final; conservar solo dims.
                rows_all.append([_cell_value(c) for c in row[:n_dims]])
            if len(matrix) < page_height:
                break
            top += page_height

        if verbose:
            print(f"[qlik/segmento] descarga completa: {len(rows_all)} filas en {time.time()-t0:.1f}s")
        return {'headers': headers, 'rows': rows_all}
    finally:
        sess.close()


# ============ CLI de prueba ============
if __name__ == '__main__':
    # Uso: python qlik_client.py [ventas|cartera|inventario|all] [--count-only]
    args = sys.argv[1:]
    target = args[0] if args else 'all'
    count_only = '--count-only' in args

    cookie = login_and_get_cookie()

    if target == 'all':
        data = fetch_all(session_cookie=cookie)
    else:
        if target not in APPS:
            sys.exit(f"target invalido: {target}. Use ventas|cartera|inventario|all")
        cfg = APPS[target]
        selections = _resolve_selections(cfg.get('selections') or {})
        headers, rows = fetch_table(cfg['app_id'], cfg['obj_id'], cookie,
                                     expected_cols=cfg['n_cols'],
                                     selections=selections)
        data = {target: {'headers': headers, 'rows': rows}}

    print()
    for key, d in data.items():
        print(f"[{key}] {len(d['headers'])} cols x {len(d['rows'])} filas")
        if not count_only and d['rows']:
            print(f"  headers: {d['headers']}")
            print(f"  fila[0]: {d['rows'][0]}")
