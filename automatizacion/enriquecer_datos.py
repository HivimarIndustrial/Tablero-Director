"""
enriquecer_datos.py

Reemplaza las formulas VLOOKUP de las hojas raw_* del Excel por codigo Python.
Toma los datos crudos que devuelven qlik_client y odoo_client, y les agrega
las columnas _std (vendedor_principal_std, cliente_nombre_std, supervisor_std,
grupo_articulo_std, fecha_std, vista_tablero, responsable_producto_std,
tipo_grupo_artículo, etc) usando las dim tables del Excel (que sí seguimos
leyendo en modo read-only).

Output: DataFrames pandas listos para escribir como CSV y ser consumidos
por regenerar_db.py sin que este tenga que abrir el Excel para las raw_*.
"""
import os
import re
import pandas as pd
from datetime import datetime, date

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
XLSX = os.path.join(PROJECT_DIR, 'Base Tablero Industria.xlsx')

# Diccionario de correcciones de nombres (typos)
CORRECCIONES_NOMBRES = {
    'EDUAROD CHILAN': 'EDUARDO CHILAN',
}


def aplicar_correccion(v):
    if isinstance(v, str):
        return CORRECCIONES_NOMBRES.get(v.strip(), v.strip()) if v else v
    return v


def _s(v):
    """Safe string with strip and correction."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    s = str(v).strip()
    return aplicar_correccion(s)


def _n(v, default=0):
    """Safe number."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return default
    try:
        return float(v)
    except (ValueError, TypeError):
        return default


def _norm_cliente_id(v):
    """Normaliza COD_CLIENTE / Código SAP Cliente a string sin ceros a la izquierda."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    if isinstance(v, bool):
        return ''
    if isinstance(v, (int, float)):
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v)
    s = str(v).strip()
    if not s:
        return ''
    # El Código SAP viene tipo "0000465394" — lo queremos como "465394"
    if s.isdigit():
        return str(int(s))
    return s


def _extract_code_from_bracket(s):
    """Extrae el código numérico de strings tipo '[100008] - FREDDY JIMENEZ'."""
    if not s:
        return ''
    m = re.match(r'^\s*\[([^\]]+)\]', str(s))
    if m:
        return m.group(1).strip()
    return ''


def _extract_name_from_bracket(s):
    """Extrae el nombre de strings tipo '[100008] - FREDDY JIMENEZ'."""
    if not s:
        return ''
    m = re.match(r'^\s*\[[^\]]+\]\s*-\s*(.+)$', str(s))
    if m:
        return m.group(1).strip()
    return str(s).strip()


def _strip_accents(s):
    """Quita tildes para comparar cabeceras (Codigo SAP vs Código SAP)."""
    if not isinstance(s, str):
        return s
    repl = {'á':'a','é':'e','í':'i','ó':'o','ú':'u','ñ':'n',
            'Á':'A','É':'E','Í':'I','Ó':'O','Ú':'U','Ñ':'N'}
    return ''.join(repl.get(c, c) for c in s)


def _make_get(headers, row):
    """Devuelve un get(key) que acepta header con o sin tilde indistintamente.

    Crítico porque Odoo devuelve los headers SIN tildes (segun
    odoo_mapeo_definitivo.json), pero el código histórico los pide CON tildes
    ('Código SAP Cliente', 'Grupo de artículo', etc).
    """
    idx_lit = {h: i for i, h in enumerate(headers)}
    idx_norm = {_strip_accents(h).lower(): i for i, h in enumerate(headers)}
    def get(k):
        i = idx_lit.get(k)
        if i is None:
            i = idx_norm.get(_strip_accents(k).lower())
        if i is None or i >= len(row):
            return None
        return row[i]
    return get


def _norm_grupo(g):
    if g is None or (isinstance(g, float) and pd.isna(g)):
        return ''
    return str(g).strip().upper()


def _fecha_ym(anio, mes):
    """YYYY-MM-01 desde anio + mes."""
    try:
        return date(int(anio), int(mes), 1).strftime('%Y-%m-%d')
    except (TypeError, ValueError):
        return ''


def _fecha_simple(v):
    """Devuelve YYYY-MM-DD de un datetime/str."""
    if v is None or v == '' or (isinstance(v, float) and pd.isna(v)):
        return ''
    if isinstance(v, (datetime, date)):
        return v.strftime('%Y-%m-%d')
    s = str(v).strip()
    # Acepta "2024-01-08 14:46:07" o "2024-01-08"
    if len(s) >= 10:
        return s[:10]
    return s


# =======================================================================
# CARGA DE DIMs
# =======================================================================

def load_dims(xlsx=XLSX):
    """Carga las dim tables necesarias para el enriquecimiento."""
    dim_clientes = pd.read_excel(xlsx, sheet_name='dim_clientes', engine='openpyxl')
    dim_vendedores = pd.read_excel(xlsx, sheet_name='dim_vendedores', engine='openpyxl')
    dim_grupo = pd.read_excel(xlsx, sheet_name='dim_grupo_articulo_responsable', engine='openpyxl')
    return {'clientes': dim_clientes, 'vendedores': dim_vendedores, 'grupos': dim_grupo}


def build_lookups(dims, segmento_data=None, clientes_vendedor_data=None):
    """Crea diccionarios de lookup rápido.

    segmento_data (opcional): dict {'headers': [...], 'rows': [...]} con
    columnas cod_cliente + SEGMENTO_CLIENTE extraído de Qlik (app ventas).
    Se usa para poblar lookups['seg_cli_detalle']: cod_cliente ->
    segmento detallado (Agrícola, Agroindustrial, Pesquera, etc, 37 valores).

    clientes_vendedor_data (opcional): dict del extractor Odoo
    extract_clientes_vendedor (res.partner.user_id, fuente oficial sincronizada
    con SAP). Pobla lookups['cli_vendedor_odoo']: cod_cliente -> {sap_vendedor,
    vendedor_odoo, equipo_odoo, write_date}. Se usa para 'vendedor_cartera_std'
    con prioridad Odoo > Excel dim_clientes.
    """
    lookups = {}

    # Segmento detallado (37 valores) desde Qlik SEGMENTO_CLIENTE
    lookups['seg_cli_detalle'] = {}
    if segmento_data and segmento_data.get('rows'):
        hh = segmento_data.get('headers', [])
        try:
            i_cli = hh.index('cod_cliente')
            i_seg = hh.index('SEGMENTO_CLIENTE')
        except ValueError:
            i_cli, i_seg = 0, 1
        for r in segmento_data['rows']:
            if len(r) <= max(i_cli, i_seg):
                continue
            cid = _norm_cliente_id(r[i_cli])
            seg = _s(r[i_seg])
            if cid and seg:
                lookups['seg_cli_detalle'][cid] = seg

    # Cliente por id
    lookups['cli'] = {}
    for _, r in dims['clientes'].iterrows():
        cid = _norm_cliente_id(r.get('cliente_id_std'))
        if not cid:
            continue
        lookups['cli'][cid] = {
            'nombre': _s(r.get('cliente_nombre_std', '')),
            'vendedor': _s(r.get('vendedor_principal_std', '')),
            'supervisor': _s(r.get('supervisor_std', '')),
            'segmento': _s(r.get('Segmento', '')),
        }

    # Vendedor por vendedor_std (alineado con dim_vendedores)
    lookups['vend'] = {}
    lookups['vend_por_origen'] = {}
    for _, r in dims['vendedores'].iterrows():
        vs = _s(r.get('vendedor_std'))
        vo = _s(r.get('vendedor_origen'))
        info = {
            'vendedor_std': vs,
            'supervisor': _s(r.get('supervisor_std')),
            'tipo': _s(r.get('tipo_vendedor')) or 'Generalista',
            'vista_tablero': _s(r.get('vista_tablero')),
            'es_jefe_producto': _s(r.get('es_jefe_producto')) == 'SI',
        }
        if vs:
            lookups['vend'][vs] = info
        if vo:
            lookups['vend_por_origen'][vo] = info

    # Vendedor por código SAP (extraído de '[100008] - FREDDY JIMENEZ')
    # Esto requiere mirar dim_vendedores vendedor_origen que suele ser
    # '[100008] - FREDDY JIMENEZ RONQUILLO' o similar.
    lookups['vend_por_sap'] = {}
    for vo, info in lookups['vend_por_origen'].items():
        sap = _extract_code_from_bracket(vo)
        if sap:
            lookups['vend_por_sap'][sap] = info

    # Cliente -> vendedor generalista de Odoo (res.partner.user_id).
    # Fuente FRESCA (sincroniza con SAP, write_date <30dias = 203 cambios).
    # Para fallback, ver 'cli' (dim_clientes Excel).
    lookups['cli_vendedor_odoo'] = {}
    if clientes_vendedor_data and clientes_vendedor_data.get('rows'):
        hh = clientes_vendedor_data.get('headers', [])
        try:
            i_sap = hh.index('cliente_sap')
            i_sapv = hh.index('sap_vendedor')
            i_vend = hh.index('vendedor_odoo')
            i_team = hh.index('equipo_odoo')
            i_wd = hh.index('write_date')
        except ValueError:
            i_sap = i_sapv = i_vend = i_team = i_wd = None
        if i_sap is not None:
            for r in clientes_vendedor_data['rows']:
                cid = _norm_cliente_id(r[i_sap])
                vend = _s(r[i_vend])
                if not cid or not vend:
                    continue
                lookups['cli_vendedor_odoo'][cid] = {
                    'sap_vendedor': _s(r[i_sapv]),
                    'vendedor_odoo': vend,
                    'equipo_odoo': _s(r[i_team]),
                    'write_date': _s(r[i_wd]),
                }

    # Grupo por grupo_articulo_std
    lookups['grupo'] = {}
    for _, r in dims['grupos'].iterrows():
        gs = _s(r.get('grupo_articulo_std'))
        if not gs:
            continue
        lookups['grupo'][_norm_grupo(gs)] = {
            'jefe_producto_std': _s(r.get('jefe_producto_std')) or 'COMERCIO',
            'supervisor_producto_std': _s(r.get('supervisor_producto_std')) or '',
            'tipo': _s(r.get('tipo_grupo_artículo')) or 'OTROS',
        }

    return lookups


# =======================================================================
# HELPERS
# =======================================================================

def resolver_vendedor_cartera(cod_cliente, lookups):
    """Resuelve el vendedor 'dueño' del cliente HOY (asignacion actual fija).

    Estrategia hibrida:
      1) Odoo res.partner.user_id (fuente fresca, sincronizada con SAP).
         El SAP del vendedor se mapea a vendedor_std via lookups['vend_por_sap'].
      2) Fallback: dim_clientes del Excel (cobertura mayor pero menos fresca).

    Returns:
      (vendedor_std, fuente)  donde fuente in {'ODOO', 'DIM_CLIENTES', 'NONE'}.

    Esta es la asignacion que se usa para 'crecimiento de cartera del vendedor'
    (NO para venta directa, que usa el vendedor de la transaccion).
    """
    info_odoo = lookups.get('cli_vendedor_odoo', {}).get(cod_cliente)
    if info_odoo:
        sap_v = info_odoo.get('sap_vendedor', '')
        if sap_v:
            v = lookups.get('vend_por_sap', {}).get(sap_v, {}).get(
                'vendedor_std', '')
            if v:
                return v, 'ODOO'
        # Si Odoo da nombre pero no SAP, intentar match por nombre directo
        nm = info_odoo.get('vendedor_odoo', '')
        if nm:
            v = lookups.get('vend_por_origen', {}).get(nm, {}).get(
                'vendedor_std', '')
            if v:
                return v, 'ODOO'
            # ultimo recurso: nombre crudo
            return nm, 'ODOO_RAW'

    # Fallback Excel
    cli_info = lookups.get('cli', {}).get(cod_cliente, {})
    if cli_info.get('vendedor'):
        return cli_info['vendedor'], 'DIM_CLIENTES'

    return '', 'NONE'


# =======================================================================
# ENRICHERS - uno por cada raw_*
# =======================================================================

def enriquecer_ventas(qlik_data, lookups):
    """
    qlik_data['rows'] viene de qlik_client.fetch_table('ventas').
    Headers esperados: AÑO, MES, segmento, grupo_articulo, marca,
    nombre_agente, nombre_cliente, cod_cliente, Venta neta, Costo Interno,
    Unidades Vendidas.
    """
    headers = qlik_data['headers']
    idx = {h: i for i, h in enumerate(headers)}
    out = []
    for row in qlik_data['rows']:
        get = _make_get(headers, row)
        anio = get('AÑO')
        mes = get('MES')
        seg = _s(get('segmento'))
        grupo = _s(get('grupo_articulo'))
        marca = _s(get('marca'))
        nombre_agente = _s(get('nombre_agente'))
        nombre_cliente = _s(get('nombre_cliente'))
        cod_cliente = _norm_cliente_id(get('cod_cliente'))
        # nombre_digitador: anadido al hypercube por sistemas en 2026-05-19.
        # Se propaga al enriquecido para que otros proyectos puedan consumirlo.
        # Si el campo no esta presente (rollback de sistemas) cae a vacio.
        nombre_digitador = _s(get('nombre_digitador'))
        venta = _n(get('Venta neta'))
        costo = _n(get('Costo Interno'))
        unid = _n(get('Unidades Vendidas'))

        cli_info = lookups['cli'].get(cod_cliente, {})
        vendedor_std = cli_info.get('vendedor', '')
        cliente_nombre = cli_info.get('nombre', '') or nombre_cliente
        grupo_std = _norm_grupo(grupo)

        vend_info = lookups['vend'].get(vendedor_std, {})
        supervisor_std = vend_info.get('supervisor', '') or cli_info.get('supervisor', '')
        vista = vend_info.get('vista_tablero', '')

        grupo_info = lookups['grupo'].get(grupo_std, {})
        resp_prod = grupo_info.get('jefe_producto_std', 'COMERCIO')
        sup_prod = grupo_info.get('supervisor_producto_std', '')
        tipo_g = grupo_info.get('tipo', 'OTROS')

        # Vendedor de cartera (asignacion actual del cliente, hibrida Odoo+Excel)
        vendedor_cartera, fuente_cartera = resolver_vendedor_cartera(
            cod_cliente, lookups)

        out.append({
            'AÑO': anio, 'MES': mes, 'segmento': seg, 'grupo_articulo': grupo,
            'marca': marca, 'nombre_agente': nombre_agente,
            'nombre_cliente': nombre_cliente, 'cod_cliente': cod_cliente,
            'nombre_digitador': nombre_digitador,
            'Venta neta': venta, 'Costo Interno': costo,
            'Unidades Vendidas': unid,
            'vendedor_principal_std': vendedor_std,
            'vendedor_cartera_std': vendedor_cartera,
            'fuente_vendedor_cartera': fuente_cartera,
            'cliente_nombre_std': cliente_nombre,
            'grupo_articulo_std': grupo_std,
            'fecha_std': _fecha_ym(anio, mes),
            'cliente_id_std': cod_cliente,
            'supervisor_std': supervisor_std,
            'fuente_vendedor': 'DIM_CLIENTES' if cli_info else 'REVISAR',
            'vista_tablero': vista,
            'responsable_producto_std': resp_prod,
            'supervisor_producto_std': sup_prod,
            'tipo_grupo_artículo': tipo_g,
            'segmento_cliente_detalle': lookups['seg_cli_detalle'].get(cod_cliente, ''),
        })
    return out


def enriquecer_cartera(qlik_data, lookups):
    """raw_cartera: 32 cols Qlik + 7 _std.

    POST-FILTRO INDUSTRIA: el objeto KGTrxF en Qlik ignora el filtro
    'jefe_ventas' (set expression o alternate state interno). Por eso aunque
    el extractor seleccione Juan Davila + Juan Bladimir Davila Chacon,
    Qlik devuelve TODAS las 490k filas (todos los jefes de Hivimar).
    Filtramos aqui en Python para quedarnos solo con cartera Industria.
    """
    JEFES_INDUSTRIA = {'JUAN DAVILA', 'JUAN BLADIMIR DAVILA CHACON'}
    headers = qlik_data['headers']
    idx = {h: i for i, h in enumerate(headers)}
    idx_jv = idx.get('jefe_ventas')
    out = []
    n_descartadas = 0
    for row in qlik_data['rows']:
        # Post-filtro Industria por jefe_ventas
        if idx_jv is not None:
            jv = (row[idx_jv] if idx_jv < len(row) else '') or ''
            if jv not in JEFES_INDUSTRIA:
                n_descartadas += 1
                continue
        get = _make_get(headers, row)
        record = {h: row[i] if i < len(row) else None for i, h in enumerate(headers)}

        cod_cliente = _norm_cliente_id(get('CODIGO_CLIENTE'))
        cli_info = lookups['cli'].get(cod_cliente, {})

        vendedor_std = cli_info.get('vendedor', '')
        cliente_nombre = cli_info.get('nombre', '') or _s(get('nombre_cliente'))

        vend_info = lookups['vend'].get(vendedor_std, {})
        supervisor_std = vend_info.get('supervisor', '') or cli_info.get('supervisor', '')

        # Fecha base como fecha_std
        fecha_base = get('FECHA_BASE')
        fecha_std = _fecha_simple(fecha_base)

        # Vendedor de cartera (asignacion actual del cliente, hibrida Odoo+Excel)
        vendedor_cartera, fuente_cartera = resolver_vendedor_cartera(
            cod_cliente, lookups)

        record['vendedor_principal_std'] = vendedor_std
        record['vendedor_cartera_std'] = vendedor_cartera
        record['fuente_vendedor_cartera'] = fuente_cartera
        record['cliente_nombre_std'] = cliente_nombre
        record['fecha_std'] = fecha_std
        record['cliente_id_std'] = cod_cliente
        record['supervisor_std'] = supervisor_std
        record['fuente_vendedor'] = 'DIM_CLIENTES' if cli_info else 'REVISAR'
        record['Segmento'] = cli_info.get('segmento', '')
        record['segmento_cliente_detalle'] = lookups['seg_cli_detalle'].get(cod_cliente, '')
        out.append(record)
    print(f"[enriquecer/cartera] post-filtro Industria: kept={len(out):,} "
          f"descartadas (no Industria)={n_descartadas:,}")
    return out


def enriquecer_inventario(qlik_data, lookups):
    """raw_inventario: 4 cols Qlik + 4 _std."""
    headers = qlik_data['headers']
    idx = {h: i for i, h in enumerate(headers)}
    out = []
    for row in qlik_data['rows']:
        get = _make_get(headers, row)
        grupo = _s(get('grupo_articulo'))
        marca = _s(get('marca'))
        stock = _n(get('$ Stock Total'))
        und = _n(get('Und Stock Total'))
        grupo_std = _norm_grupo(grupo)
        grupo_info = lookups['grupo'].get(grupo_std, {})
        out.append({
            'grupo_articulo': grupo, 'marca': marca,
            '$ Stock Total': stock, 'Und Stock Total': und,
            'grupo_articulo_std': grupo_std,
            'jefe_producto_std': grupo_info.get('jefe_producto_std', 'COMERCIO'),
            'supervisor_producto_std': grupo_info.get('supervisor_producto_std', ''),
            'tipo_grupo_artículo': grupo_info.get('tipo', 'OTROS'),
        })
    return out


ESTADO_COT_MAP = {
    'draft': 'Presupuesto', 'sent': 'Enviado',
    'sale': 'Cotización Ganada', 'done': 'Cotización Ganada',
    'cancel': 'Cancelado',
}


def enriquecer_cotizaciones(odoo_data, lookups):
    """raw_cotizaciones desde Odoo: 19 cols base + 11 _std."""
    headers = odoo_data['headers']
    idx = {h: i for i, h in enumerate(headers)}
    out = []
    for row in odoo_data['rows']:
        get = _make_get(headers, row)
        record = {h: row[i] if i < len(row) else None for i, h in enumerate(headers)}

        cod_cli = _norm_cliente_id(get('Código SAP Cliente'))
        sap_gen = _extract_code_from_bracket(get('Comercial')) or _s(get('Código SAP Generalista'))
        cli_info = lookups['cli'].get(cod_cli, {})
        vend_info = lookups['vend_por_sap'].get(sap_gen, {})

        vendedor_std = vend_info.get('vendedor_std', '') or cli_info.get('vendedor', '') or 'REVISAR'
        cliente_nombre = cli_info.get('nombre', '') or _s(get('Cliente'))
        grupo_raw = _s(get('Grupo Artículo'))
        # Odoo trae '[GA022] - Rodamientos Industriales'; dim_grupo tiene
        # solo el nombre uppercased ('RODAMIENTOS INDUSTRIALES'), asi que
        # extraemos la parte despues del bracket antes de normalizar.
        grupo = _extract_name_from_bracket(grupo_raw) if grupo_raw else ''
        grupo_std = _norm_grupo(grupo) or 'SIN_GRUPO'
        grupo_info = lookups['grupo'].get(grupo_std, {})

        supervisor_std = vend_info.get('supervisor', '') or cli_info.get('supervisor', '')
        vista = vend_info.get('vista_tablero', '') or 'REVISAR'

        estado_raw = _s(get('Estado del pedido'))
        estado_std = ESTADO_COT_MAP.get(estado_raw, 'Presupuesto')

        fecha_std = _fecha_simple(get('Creado'))

        # Vendedor de cartera (asignacion actual del cliente, hibrida Odoo+Excel)
        vendedor_cartera, fuente_cartera = resolver_vendedor_cartera(
            cod_cli, lookups)

        record['estado_std'] = estado_std
        record['vendedor_principal_std'] = vendedor_std
        record['vendedor_cartera_std'] = vendedor_cartera
        record['fuente_vendedor_cartera'] = fuente_cartera
        record['cliente_nombre_std'] = cliente_nombre
        record['grupo_articulo_std'] = grupo_std
        record['fecha_std'] = fecha_std
        record['cliente_id_std'] = cod_cli or 'SIN_CODIGO'
        record['supervisor_std'] = supervisor_std
        record['fuente_vendedor'] = 'DIM_VEND' if vend_info else ('DIM_CLI' if cli_info else 'REVISAR')
        record['vista_tablero'] = vista
        record['responsable_producto_std'] = grupo_info.get('jefe_producto_std', 'REVISAR')
        record['supervisor_producto_std'] = grupo_info.get('supervisor_producto_std', 'REVISAR')
        record['segmento_cliente_detalle'] = lookups['seg_cli_detalle'].get(cod_cli, '')
        out.append(record)
    return out


# Etapas CRM -> estado_oportunidad.
# Reglas SSoT (ver INSTRUCCIONES_PIPELINE_OPORTUNIDADES.md):
#   1) Ganado facturado          -> Ganada Facturada
#   2) Eliminada / Eliminados    -> Eliminada
#   3) Perdidas/Perdida  O  Motivo de perdida != ''  -> Perdida
#   4) Resto                     -> Activa
ETAPAS_GANADA_FACT = {'ganado facturado'}
ETAPAS_ELIMINADA   = {'eliminada', 'eliminados'}
ETAPAS_PERDIDA     = {'perdidas', 'perdida'}


def _clasificar_estado(etapa, motivo):
    e = ('' if etapa is None else str(etapa)).strip().lower()
    m = ('' if motivo is None else str(motivo)).strip()
    if e in ETAPAS_GANADA_FACT:
        return 'Ganada Facturada'
    if e in ETAPAS_ELIMINADA:
        return 'Eliminada'
    if e in ETAPAS_PERDIDA or m:
        return 'Perdida'
    return 'Activa'


def enriquecer_oportunidades(odoo_data, lookups):
    """fact_oportunidades desde Odoo. Sólo generamos los campos que regenerar_db.py usa."""
    headers = odoo_data['headers']
    idx = {h: i for i, h in enumerate(headers)}
    out = []
    for row in odoo_data['rows']:
        get = _make_get(headers, row)
        record = {h: row[i] if i < len(row) else None for i, h in enumerate(headers)}

        cod_cli = _norm_cliente_id(get('Código SAP Cliente'))
        sap_gen = _extract_code_from_bracket(get('Comercial')) or _s(get('Código SAP Generalista'))
        sap_esp = _extract_code_from_bracket(get('Vendedor Especialista'))
        cli_info = lookups['cli'].get(cod_cli, {})
        vend_info = lookups['vend_por_sap'].get(sap_gen, {})
        esp_info = lookups['vend_por_sap'].get(sap_esp, {})

        vendedor_std = vend_info.get('vendedor_std', '') or cli_info.get('vendedor', '') or 'REVISAR'
        especialista_std = esp_info.get('vendedor_std', '') or _extract_name_from_bracket(get('Vendedor Especialista'))
        cliente_nombre = cli_info.get('nombre', '') or _s(get('Cliente'))

        grupo_raw = _s(get('Grupo de artículo'))
        # Idem cotizaciones: Odoo entrega '[GAxxx] - Nombre' (m2m con varios
        # grupos a veces). Tomamos el primero y extraemos el nombre puro.
        grupo_first = grupo_raw.split(',')[0].strip() if grupo_raw else ''
        grupo = _extract_name_from_bracket(grupo_first) if grupo_first else ''
        grupo_std = _norm_grupo(grupo) or grupo
        grupo_info = lookups['grupo'].get(grupo_std, {})

        # Supervisor: SOLO para oportunidades, fuente unica = Odoo
        # (crm.lead.team_id -> crm.team.user_id, ya resuelto en odoo_client).
        # Fallback al supervisor del Excel si Odoo no lo trae.
        supervisor_odoo = _s(get('Supervisor Odoo'))
        supervisor_std = (supervisor_odoo
                          or vend_info.get('supervisor', '')
                          or cli_info.get('supervisor', ''))
        vista = vend_info.get('vista_tablero', '') or 'REVISAR'

        etapa = _s(get('Etapa'))
        motivo_perdida = _s(get('Motivo de perdida'))
        estado_oportunidad = _clasificar_estado(etapa, motivo_perdida)
        # active: en odoo_client se preserva como bool. None -> archivada.
        active_raw = get('active')
        is_active = active_raw is True or (
            isinstance(active_raw, str)
            and active_raw.strip().lower() in ('true', '1', 'verdadero')
        )

        # anio/mes de creacion y de cierre
        fecha_creacion = get('Creado')
        fs_creacion = _fecha_simple(fecha_creacion)
        anio_creacion = int(fs_creacion[:4]) if len(fs_creacion) >= 4 else None
        mes_creacion = int(fs_creacion[5:7]) if len(fs_creacion) >= 7 else None

        fecha_cierre = get('Fecha de cierre')
        fs_cierre = _fecha_simple(fecha_cierre)
        anio_cierre = int(fs_cierre[:4]) if len(fs_cierre) >= 4 else None
        mes_cierre = int(fs_cierre[5:7]) if len(fs_cierre) >= 7 else None

        # Vendedor de cartera (asignacion actual del cliente, hibrida Odoo+Excel)
        vendedor_cartera, fuente_cartera = resolver_vendedor_cartera(
            cod_cli, lookups)

        record['vendedor_principal_std'] = vendedor_std
        record['vendedor_cartera_std'] = vendedor_cartera
        record['fuente_vendedor_cartera'] = fuente_cartera
        record['especialista_std'] = especialista_std
        record['cliente_id_std'] = cod_cli or 'SIN_CODIGO'
        record['cliente_nombre_std'] = cliente_nombre
        record['grupo_articulo_std'] = grupo_std
        record['supervisor_std'] = supervisor_std
        record['estado_oportunidad'] = estado_oportunidad
        record['motivo_perdida'] = motivo_perdida
        record['active'] = is_active
        record['vista_tablero'] = vista
        record['responsable_producto_std'] = grupo_info.get('jefe_producto_std', 'REVISAR')
        record['supervisor_producto_std'] = grupo_info.get('supervisor_producto_std', 'REVISAR')
        record['tipo_grupo_artículo'] = grupo_info.get('tipo', 'OTROS')
        record['segmento_cliente'] = cli_info.get('segmento', 'Sin Segmento')
        record['segmento_cliente_detalle'] = lookups['seg_cli_detalle'].get(cod_cli, '')
        record['anio_creacion'] = anio_creacion
        record['mes_creacion'] = mes_creacion
        record['anio_cierre'] = anio_cierre
        record['mes_cierre'] = mes_cierre
        out.append(record)

    # === Reglas SSoT (ver INSTRUCCIONES_PIPELINE_OPORTUNIDADES.md) ===
    # NOTA: en Odoo Hivimar las opps cerradas (Ganadas/Perdidas/Eliminadas)
    # tienen active=False. Si filtramos archivadas aqui, perdemos TODO el
    # historico de cerradas y se rompen los embudos. En su lugar dejamos
    # la columna `active` y cada tablero downstream filtra cuando quiera
    # "solo vivas" (ej: df[df.active & (df.estado_oportunidad=='Activa')]).
    n_archivadas = sum(1 for r in out if not r.get('active', True))
    print(f"[enriquecer/opps] archivadas conservadas (active=False): "
          f"{n_archivadas}/{len(out)}")

    # Deduplicar por Name Sequence (primera ocurrencia gana).
    seen = set()
    dedup = []
    n_pre_dedup = len(out)
    for r in out:
        ns = (r.get('Name Sequence') or '').strip() if r.get('Name Sequence') else ''
        if ns:
            if ns in seen:
                continue
            seen.add(ns)
        dedup.append(r)
    print(f"[enriquecer/opps] duplicados eliminados: "
          f"{n_pre_dedup - len(dedup)}")
    return dedup


def enriquecer_visitas(odoo_data, lookups):
    """raw_visitas desde Odoo. 12 cols base + 4 _std."""
    headers = odoo_data['headers']
    idx = {h: i for i, h in enumerate(headers)}
    out = []
    ESTADO_VIS_MAP = {'done': 'Realizado', 'draft': 'Borrador',
                      'confirmed': 'Planificado', 'cancel': 'Cancelado',
                      'planned': 'Planificado'}
    for row in odoo_data['rows']:
        get = _make_get(headers, row)
        record = {h: row[i] if i < len(row) else None for i, h in enumerate(headers)}

        # Responsable vendedor
        resp_display = _s(get('Responsable'))
        sap_resp = _s(get('Código SAP Responsable')) or _extract_code_from_bracket(resp_display)
        vend_info = lookups['vend_por_sap'].get(sap_resp, {})
        vendedor = vend_info.get('vendedor_std', '') or _extract_name_from_bracket(resp_display)

        # Cliente
        cod_cli = _norm_cliente_id(get('Código SAP Cliente'))
        cli_info = lookups['cli'].get(cod_cli, {})
        cliente_nombre = cli_info.get('nombre', '') or _s(get('Cliente'))

        # Estado español
        estado_raw = _s(get('Estado'))
        estado_es = ESTADO_VIS_MAP.get(estado_raw, estado_raw)

        # Fecha std
        fecha_ref = get('Fecha efectiva') or get('Fecha planificada')
        fecha_std = _fecha_simple(fecha_ref)

        # Vendedor de cartera (asignacion actual del cliente, hibrida Odoo+Excel)
        vendedor_cartera, fuente_cartera = resolver_vendedor_cartera(
            cod_cli, lookups)

        record['Estado'] = estado_es  # override
        record['vendedor_std'] = vendedor
        record['vendedor_cartera_std'] = vendedor_cartera
        record['fuente_vendedor_cartera'] = fuente_cartera
        record['cliente_nombre_std'] = cliente_nombre
        record['fecha_std'] = fecha_std
        record['cliente_id_std'] = cod_cli or ''
        record['segmento_cliente_detalle'] = lookups['seg_cli_detalle'].get(cod_cli, '')
        out.append(record)
    return out


def enriquecer_actividades(odoo_data, lookups):
    """raw_actividadespendientes: 7 cols base + 3 _std (vendedor, cliente, fecha)."""
    headers = odoo_data['headers']
    idx = {h: i for i, h in enumerate(headers)}
    out = []
    for row in odoo_data['rows']:
        get = _make_get(headers, row)
        record = {h: row[i] if i < len(row) else None for i, h in enumerate(headers)}

        asignado = _s(get('Asignada a'))
        sap_usr = _s(get('Código SAP Asignado')) or _extract_code_from_bracket(asignado)
        vend_info = lookups['vend_por_sap'].get(sap_usr, {})
        vendedor = vend_info.get('vendedor_std', '') or asignado.upper()

        fecha_venc = get('Fecha de vencimiento')
        fecha_std = _fecha_simple(fecha_venc)

        record['vendedor_std'] = vendedor
        record['cliente_std'] = ''
        record['fecha_std'] = fecha_std
        out.append(record)
    return out


# =======================================================================
# EXPORT
# =======================================================================

def exportar_csvs(qlik_data, odoo_data, salida_dir, segmento_data=None,
                  clientes_vendedor_data=None, verbose=True):
    """Aplica enrichment y escribe CSVs enriquecidos a salida_dir.

    qlik_data = {'ventas': {headers, rows}, 'cartera': {...}, 'inventario': {...}}
    odoo_data = {'cotizaciones': {...}, 'oportunidades': {...},
                 'actividades': {...}, 'visitas': {...}}
    clientes_vendedor_data = {headers, rows} de odoo_client.extract_clientes_vendedor
        (res.partner.user_id). Pobla vendedor_cartera_std.
    """
    os.makedirs(salida_dir, exist_ok=True)

    if verbose:
        print("[enriquecer] cargando dims del Excel...")
    dims = load_dims()
    lookups = build_lookups(dims, segmento_data=segmento_data,
                            clientes_vendedor_data=clientes_vendedor_data)
    if verbose:
        print(f"  clientes dim: {len(lookups['cli'])}")
        print(f"  vendedores dim: {len(lookups['vend'])}")
        print(f"  vendedores por SAP: {len(lookups['vend_por_sap'])}")
        print(f"  grupos dim: {len(lookups['grupo'])}")
        print(f"  segmento detalle (Qlik): {len(lookups['seg_cli_detalle'])}")
        print(f"  cli vendedor Odoo: {len(lookups['cli_vendedor_odoo'])}")

    enrichers = {
        'ventas':        (qlik_data, enriquecer_ventas),
        'cartera':       (qlik_data, enriquecer_cartera),
        'inventario':    (qlik_data, enriquecer_inventario),
        'cotizaciones':  (odoo_data, enriquecer_cotizaciones),
        'oportunidades': (odoo_data, enriquecer_oportunidades),
        'actividades':   (odoo_data, enriquecer_actividades),
        'visitas':       (odoo_data, enriquecer_visitas),
    }
    counts = {}
    for key, (src, fn) in enrichers.items():
        if src is None or key not in src:
            if verbose:
                print(f"[enriquecer] skip {key} (no data)")
            continue
        enriched = fn(src[key], lookups)
        df = pd.DataFrame(enriched)
        path = os.path.join(salida_dir, f'{key}_enriquecido.csv')
        df.to_csv(path, index=False, encoding='utf-8')
        counts[key] = len(enriched)
        if verbose:
            print(f"[enriquecer] {key}: {len(enriched)} filas -> {os.path.basename(path)}")
    return counts
