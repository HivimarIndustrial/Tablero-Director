"""
Cliente Odoo para el tablero Hivimar Industrial 8.0.

Toda la BD Odoo (elitum-crm-production-11628642) ya es de Industria.
No se aplica filtro adicional.

Extrae 4 hojas raw:
  - raw_cotizaciones           (sale.order + sale.order.line, nivel de linea)
  - fact_oportunidades         (crm.lead, incl. archivadas)
  - raw_actividadespendientes  (mail.activity)
  - raw_visitas                (crm.visit)

Formato de salida (por cada hoja):
  {'headers': [str], 'rows': [[val, ...]]}
donde las columnas corresponden a las primeras N columnas de la hoja Excel
(las columnas `_std` / formulas se dejan intactas).

Uso CLI:
  python odoo_client.py cotizaciones|oportunidades|actividades|visitas|all
  python odoo_client.py <target> --count-only
  python odoo_client.py <target> --limit 50
"""
import keyring
import sys
import os
import json
import time
import re
import requests
from datetime import datetime
from zoneinfo import ZoneInfo
from typing import Optional

TZ_LOCAL = ZoneInfo('America/Guayaquil')  # Ecuador UTC-5, sin horario verano
TZ_UTC = ZoneInfo('UTC')
_DATETIME_RE = re.compile(r'^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$')

URL = 'https://hivimar-crm.odoo.com'
DB = 'elitum-crm-production-11628642'
USER = 'consultorindustrial@hivimar.com'
SERVICE = 'hivimar-tablero-odoo'

BATCH_SIZE = 1000  # registros por lote de search_read

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def _load_mapping() -> dict:
    """Carga el mapeo definitivo col Excel -> (model, field)."""
    path = os.path.join(SCRIPT_DIR, 'odoo_mapeo_definitivo.json')
    with open(path, encoding='utf-8') as f:
        return json.load(f)


class OdooClient:
    def __init__(self, url: str = URL, db: str = DB, user: str = USER,
                 verbose: bool = True):
        self.url = url
        self.db = db
        self.user = user
        self.verbose = verbose
        self.session = requests.Session()
        self.uid = None
        self.user_ctx = None

    def authenticate(self):
        pwd = keyring.get_password(SERVICE, self.user)
        if not pwd:
            raise SystemExit(
                f"ERROR: credenciales Odoo no guardadas en keyring "
                f"(servicio '{SERVICE}', usuario '{self.user}')"
            )
        r = self.session.post(
            f'{self.url}/web/session/authenticate',
            json={'jsonrpc': '2.0',
                  'params': {'db': self.db, 'login': self.user, 'password': pwd}},
            timeout=30,
        )
        res = r.json().get('result') or {}
        if not res.get('uid'):
            raise SystemExit("Odoo auth fallo: " + json.dumps(r.json())[:300])
        self.uid = res['uid']
        self.user_ctx = res.get('user_context', {})
        if self.verbose:
            print(f"[odoo] auth OK: {res.get('name')} uid={self.uid}")
        return self

    def call(self, model: str, method: str,
             args: Optional[list] = None, kwargs: Optional[dict] = None):
        r = self.session.post(
            f'{self.url}/web/dataset/call_kw',
            json={'jsonrpc': '2.0', 'params': {
                'model': model, 'method': method,
                'args': args or [], 'kwargs': kwargs or {}
            }}, timeout=120,
        )
        data = r.json()
        if 'error' in data:
            err = data['error']
            msg = err.get('data', {}).get('message') or err.get('message') or str(err)
            raise RuntimeError(f"Odoo {model}.{method}: {msg}")
        return data['result']

    def search_read_paged(self, model: str, domain: list, fields: list,
                          order: Optional[str] = None,
                          active_test: bool = True) -> list:
        """Trae todos los registros en lotes de BATCH_SIZE."""
        offset = 0
        all_rows = []
        # read() via search_read; para incluir archivadas: contexto active_test=False
        ctx = {}
        if not active_test:
            ctx['active_test'] = False

        while True:
            t_page = time.time()
            batch = self.call(model, 'search_read',
                              [domain, fields],
                              {'limit': BATCH_SIZE, 'offset': offset,
                               'order': order or 'id asc',
                               'context': ctx})
            if not batch:
                break
            all_rows.extend(batch)
            if self.verbose:
                print(f"[odoo/{model}] offset={offset} +{len(batch)} "
                      f"(acum {len(all_rows)}) en {time.time()-t_page:.1f}s")
            if len(batch) < BATCH_SIZE:
                break
            offset += BATCH_SIZE
        return all_rows


# ===== formatters =====

def _as_scalar(v):
    """Convierte valores Odoo a algo que Excel pueda guardar:
       - many2one [id, 'display'] -> 'display'
       - False -> None (celda vacia)
       - listas de ids (o2m/m2m) -> 'id1, id2, ...' o vacio
       - datetime string 'YYYY-MM-DD HH:MM:SS' (UTC) -> datetime object en TZ_LOCAL
    """
    if v is False or v is None:
        return None
    if isinstance(v, list):
        # many2one
        if len(v) == 2 and isinstance(v[0], int) and isinstance(v[1], str):
            return v[1]
        # o2m / m2m: lista de ids
        if all(isinstance(x, int) for x in v):
            return ', '.join(str(x) for x in v) if v else None
    if isinstance(v, str) and _DATETIME_RE.match(v):
        # datetime UTC -> local. Devolvemos datetime object (Excel lo guarda como fecha).
        try:
            dt_utc = datetime.strptime(v, '%Y-%m-%d %H:%M:%S').replace(tzinfo=TZ_UTC)
            dt_local = dt_utc.astimezone(TZ_LOCAL)
            # openpyxl no soporta tz-aware datetimes; devolver naive en hora local.
            return dt_local.replace(tzinfo=None)
        except Exception:
            return v
    return v


# ===== extractores por hoja =====

def extract_cotizaciones(cli: OdooClient, limit: Optional[int] = None) -> dict:
    """
    raw_cotizaciones: nivel de linea (sale.order.line).
    Une campos de la linea con campos del pedido parent.
    """
    mapping = _load_mapping()['raw_cotizaciones']
    headers = [m['header'] for m in mapping]

    # Separar campos por modelo
    so_fields = sorted({m['field'] for m in mapping if m['model'] == 'sale.order'})
    sol_fields = sorted({m['field'] for m in mapping if m['model'] == 'sale.order.line'})
    # Necesitamos order_id en la linea para unir
    if 'order_id' not in sol_fields:
        sol_fields.append('order_id')

    # 1) Leer todas las lineas
    t0 = time.time()
    if cli.verbose:
        print(f"[odoo/cotizaciones] leyendo sale.order.line (fields={sol_fields})")
    domain_sol = []
    kwargs = {'limit': limit} if limit else None
    if limit:
        sol_rows = cli.call('sale.order.line', 'search_read',
                            [domain_sol, sol_fields], kwargs)
    else:
        sol_rows = cli.search_read_paged('sale.order.line', domain_sol, sol_fields)

    # 2) Leer los sale.order correspondientes (solo los que referencian las lineas)
    order_ids = list({r['order_id'][0] for r in sol_rows if r.get('order_id')})
    if cli.verbose:
        print(f"[odoo/cotizaciones] leyendo {len(order_ids)} sale.order unicos")
    so_map = {}
    for i in range(0, len(order_ids), BATCH_SIZE):
        chunk = order_ids[i:i + BATCH_SIZE]
        so_rows = cli.call('sale.order', 'read', [chunk, so_fields])
        for r in so_rows:
            so_map[r['id']] = r

    # 3) Componer filas en el orden de las columnas del Excel
    out_rows = []
    for sol in sol_rows:
        so = so_map.get(sol['order_id'][0], {}) if sol.get('order_id') else {}
        row = []
        for m in mapping:
            if m['model'] == 'sale.order':
                v = so.get(m['field'])
            else:
                v = sol.get(m['field'])
            row.append(_as_scalar(v))
        out_rows.append(row)

    if cli.verbose:
        print(f"[odoo/cotizaciones] listo: {len(out_rows)} filas en {time.time()-t0:.1f}s")
    return {'headers': headers, 'rows': out_rows}


def extract_oportunidades(cli: OdooClient, limit: Optional[int] = None) -> dict:
    """fact_oportunidades: crm.lead, incluyendo archivadas.

    Nota: el campo `group_article_id` en crm.lead es many2many a
    `product.group.article`. Odoo lo entrega como [id1, id2, ...] sin
    display_name. Aqui pre-cargamos el catalogo y resolvemos cada id a
    su display_name (`[GAxxx] - Nombre Grupo`) antes de exportar.
    """
    mapping = _load_mapping()['fact_oportunidades']
    headers = [m['header'] for m in mapping]
    fields = sorted({m['field'] for m in mapping})

    t0 = time.time()
    if cli.verbose:
        print(f"[odoo/oportunidades] leyendo crm.lead (incl. archivadas)")

    # Pre-cache catalogo de grupos (id -> display_name) para resolver m2m
    if cli.verbose:
        print(f"[odoo/oportunidades] precargando catalogo product.group.article")
    grupos_raw = cli.call('product.group.article', 'search_read',
                          [[]], {'fields': ['id', 'display_name']})
    grupo_map = {g['id']: g.get('display_name') or '' for g in grupos_raw}
    if cli.verbose:
        print(f"[odoo/oportunidades]   {len(grupo_map)} grupos cacheados")

    # Pre-cache catalogo de equipos: team_id -> nombre del lider (= supervisor).
    # En Odoo: crm.lead.team_id -> crm.team -> user_id (lider del equipo).
    # El display_name del lider viene como "[SAP] - NOMBRE", limpiamos a "NOMBRE".
    if cli.verbose:
        print(f"[odoo/oportunidades] precargando catalogo crm.team (supervisores)")
    teams_raw = cli.call('crm.team', 'search_read',
                         [[]], {'fields': ['id', 'name', 'user_id']})

    def _clean_supervisor(display: str) -> str:
        """'[100168] - JUAN JARAMILLO' -> 'JUAN JARAMILLO'. Si no matchea, devuelve tal cual."""
        if not display:
            return ''
        s = str(display).strip()
        # patron "[xxx] - NOMBRE"
        if s.startswith('[') and ']' in s:
            try:
                _, rest = s.split(']', 1)
                rest = rest.strip()
                if rest.startswith('-'):
                    rest = rest[1:].strip()
                return rest or s
            except Exception:
                return s
        return s

    team_supervisor_map = {}
    for t in teams_raw:
        tid = t['id']
        leader = t.get('user_id')
        if isinstance(leader, list) and len(leader) >= 2:
            team_supervisor_map[tid] = _clean_supervisor(leader[1])
        else:
            team_supervisor_map[tid] = ''
    if cli.verbose:
        print(f"[odoo/oportunidades]   {len(team_supervisor_map)} equipos cacheados "
              f"(con lider: {sum(1 for v in team_supervisor_map.values() if v)})")

    domain = []
    if limit:
        rows = cli.call('crm.lead', 'search_read',
                        [domain, fields],
                        {'limit': limit, 'context': {'active_test': False}})
    else:
        rows = cli.search_read_paged('crm.lead', domain, fields, active_test=False)

    GROUP_FIELD = 'group_article_id'
    TEAM_FIELD = 'team_id'
    out_rows = []
    supervisores_resueltos = []
    for r in rows:
        # Resolver group_article_id (m2m de IDs) a join de display_names
        gv = r.get(GROUP_FIELD)
        if isinstance(gv, list) and gv and all(isinstance(x, int) for x in gv):
            names = [grupo_map.get(gid, str(gid)) for gid in gv]
            r[GROUP_FIELD] = ', '.join(n for n in names if n)

        # Resolver team_id -> supervisor (lider del equipo en Odoo)
        tv = r.get(TEAM_FIELD)
        sup_odoo = ''
        if isinstance(tv, list) and len(tv) >= 1:
            sup_odoo = team_supervisor_map.get(tv[0], '')
        supervisores_resueltos.append(sup_odoo)

        row = []
        for m in mapping:
            v = r.get(m['field'])
            # boolean: preservar False (caso 'active') porque _as_scalar lo
            # colapsa a None y perderiamos la distincion archivado vs vivo.
            if m.get('type') == 'boolean':
                row.append(bool(v))
            else:
                row.append(_as_scalar(v))
        out_rows.append(row)

    # Agregar columna virtual "Supervisor Odoo" (no esta en el mapeo, se
    # construye en runtime via crm.team -> user_id). La consume
    # enriquecer_oportunidades para sobrescribir supervisor_std.
    headers = list(headers) + ['Supervisor Odoo']
    for i, row in enumerate(out_rows):
        row.append(supervisores_resueltos[i])

    if cli.verbose:
        n_con_sup = sum(1 for s in supervisores_resueltos if s)
        print(f"[odoo/oportunidades] supervisor Odoo resuelto en "
              f"{n_con_sup}/{len(supervisores_resueltos)} opps")
        print(f"[odoo/oportunidades] listo: {len(out_rows)} filas en {time.time()-t0:.1f}s")
    return {'headers': headers, 'rows': out_rows}


def extract_actividades(cli: OdooClient, limit: Optional[int] = None) -> dict:
    """raw_actividadespendientes: mail.activity (en Odoo solo hay pendientes)."""
    mapping = _load_mapping()['raw_actividadespendientes']
    headers = [m['header'] for m in mapping]
    fields = sorted({m['field'] for m in mapping})

    t0 = time.time()
    if cli.verbose:
        print(f"[odoo/actividades] leyendo mail.activity")
    domain = []
    if limit:
        rows = cli.call('mail.activity', 'search_read',
                        [domain, fields], {'limit': limit})
    else:
        rows = cli.search_read_paged('mail.activity', domain, fields)

    out_rows = []
    for r in rows:
        row = [_as_scalar(r.get(m['field'])) for m in mapping]
        out_rows.append(row)
    if cli.verbose:
        print(f"[odoo/actividades] listo: {len(out_rows)} filas en {time.time()-t0:.1f}s")
    return {'headers': headers, 'rows': out_rows}


def extract_visitas(cli: OdooClient, limit: Optional[int] = None) -> dict:
    """raw_visitas: crm.visit."""
    mapping = _load_mapping()['raw_visitas']
    headers = [m['header'] for m in mapping]
    fields = sorted({m['field'] for m in mapping})

    t0 = time.time()
    if cli.verbose:
        print(f"[odoo/visitas] leyendo crm.visit")
    domain = []
    if limit:
        rows = cli.call('crm.visit', 'search_read',
                        [domain, fields], {'limit': limit})
    else:
        rows = cli.search_read_paged('crm.visit', domain, fields)

    out_rows = []
    for r in rows:
        row = [_as_scalar(r.get(m['field'])) for m in mapping]
        out_rows.append(row)
    if cli.verbose:
        print(f"[odoo/visitas] listo: {len(out_rows)} filas en {time.time()-t0:.1f}s")
    return {'headers': headers, 'rows': out_rows}


def extract_clientes_vendedor(cli: OdooClient,
                              limit: Optional[int] = None) -> dict:
    """Asignacion vigente cliente -> vendedor generalista en Odoo.

    Fuente: res.partner.user_id (string='Vendedor Generalista'). Esta es la
    fuente oficial que se sincroniza con SAP y se actualiza diariamente.
    Sirve para responder "DE QUIEN es este cliente HOY?" — usado en analisis
    de crecimiento de cartera del vendedor (asignacion actual fija aplicada
    a todo el periodo de comparacion).

    Headers fijos (no via mapping): cliente_sap, cliente_nombre,
        sap_vendedor, vendedor_odoo, equipo_odoo, write_date.
    """
    t0 = time.time()
    if cli.verbose:
        print(f"[odoo/clientes_vendedor] leyendo res.partner (customer_rank>0)")
    domain = [('customer_rank', '>', 0)]
    fields = ['id', 'name', 'sap_code', 'user_id', 'team_id',
              'specialized_seller_ids', 'write_date']
    if limit:
        rows = cli.call('res.partner', 'search_read',
                        [domain, fields], {'limit': limit})
    else:
        rows = cli.search_read_paged('res.partner', domain, fields)

    headers = ['cliente_sap', 'cliente_nombre', 'sap_vendedor',
               'vendedor_odoo', 'equipo_odoo', 'write_date']
    out_rows = []
    n_con_vend = 0
    for r in rows:
        cliente_sap = (r.get('sap_code') or '').strip()
        cliente_nombre = r.get('name') or ''
        # user_id viene como [60, '[100041] - CARLOS CANTOS']
        u = r.get('user_id')
        sap_vend = ''
        vend_odoo = ''
        if isinstance(u, list) and len(u) >= 2:
            disp = u[1] or ''
            # Extraer SAP del bracket "[100041] - CARLOS CANTOS"
            if disp.startswith('[') and ']' in disp:
                try:
                    sap_vend = disp.split(']', 1)[0].lstrip('[').strip()
                    rest = disp.split(']', 1)[1].strip()
                    if rest.startswith('-'):
                        rest = rest[1:].strip()
                    vend_odoo = rest or disp
                except Exception:
                    vend_odoo = disp
            else:
                vend_odoo = disp
            if vend_odoo:
                n_con_vend += 1
        # team
        t = r.get('team_id')
        equipo = t[1] if isinstance(t, list) and len(t) >= 2 else ''
        # write_date a string (UTC)
        wd = r.get('write_date') or ''
        out_rows.append([cliente_sap, cliente_nombre, sap_vend,
                         vend_odoo, equipo, str(wd)])
    if cli.verbose:
        print(f"[odoo/clientes_vendedor] {len(out_rows)} clientes, "
              f"{n_con_vend} con vendedor asignado, "
              f"en {time.time()-t0:.1f}s")
    return {'headers': headers, 'rows': out_rows}


# Registry
EXTRACTORS = {
    'cotizaciones':       extract_cotizaciones,
    'oportunidades':      extract_oportunidades,
    'actividades':        extract_actividades,
    'visitas':            extract_visitas,
    'clientes_vendedor':  extract_clientes_vendedor,
}


def fetch_all(cli: Optional[OdooClient] = None,
              limit: Optional[int] = None,
              verbose: bool = True) -> dict:
    if cli is None:
        cli = OdooClient(verbose=verbose).authenticate()
    out = {}
    for key, fn in EXTRACTORS.items():
        if verbose:
            print()
            print(f"=== [odoo] {key} ===")
        out[key] = fn(cli, limit=limit)
    return out


# ============ CLI ============
if __name__ == '__main__':
    args = sys.argv[1:]
    if not args:
        print("Uso: odoo_client.py cotizaciones|oportunidades|actividades|visitas|all [--limit N] [--count-only]")
        sys.exit(1)
    target = args[0]
    limit = None
    count_only = '--count-only' in args
    if '--limit' in args:
        i = args.index('--limit')
        limit = int(args[i + 1])

    cli = OdooClient().authenticate()

    if target == 'all':
        data = fetch_all(cli, limit=limit)
    else:
        if target not in EXTRACTORS:
            sys.exit(f"target invalido: {target}")
        data = {target: EXTRACTORS[target](cli, limit=limit)}

    print()
    for key, d in data.items():
        print(f"[{key}] {len(d['headers'])} cols x {len(d['rows'])} filas")
        if not count_only and d['rows']:
            print(f"  headers: {d['headers'][:5]}...")
            print(f"  fila[0]: {d['rows'][0][:6]}...")
