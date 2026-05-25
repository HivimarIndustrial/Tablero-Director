"""
Extrae entregas Industria (lineaNegocio=1100) desde Hivitrack 2.0 con
estrategia INCREMENTAL: solo descarga ventana reciente y hace upsert
contra un CSV historico acumulativo.

Credenciales: keyring servicio 'hivimar-tablero-hivitrack'.
Fallback: .env del proyecto 'Voz del cliente'.

Estrategia:
  - Cada corrida descarga ultimos DIAS_INCREMENTAL dias (default 45).
  - Upsert por COD_FACTURA / numeroEntrega contra CSV historico:
    - Si numeroEntrega existe -> reemplaza con la version fresca
    - Si es nuevo -> append
  - Entregas con fecha de creacion anterior a 2026-01-01 se descartan
    (Hivitrack en pleno uso solo desde Ene 2026).

Salida: salida_raw/entregas_industria.csv
"""
import csv
import os
import sys
import time
import keyring
from datetime import datetime, date, timedelta
from typing import List, Dict, Optional

# Path al cliente Hivitrack del kit
KIT_PATH = (r'C:\Users\consultorindustrial\OneDrive - Hivimar\Documentos\HIVIMAR'
            r'\Proyectos\Hivitrack Starter Kit')
if KIT_PATH not in sys.path:
    sys.path.insert(0, KIT_PATH)

SERVICE = 'hivimar-tablero-hivitrack'
USERNAME_DEFAULT = 'industriahiv'

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
SALIDA_RAW = os.path.join(PROJECT_DIR, 'salida_raw')
DEST_CSV = os.path.join(SALIDA_RAW, 'entregas_industria.csv')

DIAS_INCREMENTAL = 45
ROWS_PER_PAGE = 500
LINEA_NEGOCIO_INDUSTRIA = '1100'
FECHA_INICIO_HIVITRACK = date(2026, 1, 1)

# Columnas a persistir (subset de los 38 que trae la API)
COLUMNAS = [
    'numeroEntrega', 'numeroFactura', 'numeroPedido', 'guia',
    'fechaHoraCreacionPedido', 'fechaHoraLiberacionPedido',
    'fechaHoraCreacionEntrega', 'fechaHoraFacturacion',
    'fechaAsignadoTransporte', 'fechaEnDistribucion',
    'fechaLlegoDestino', 'fechaEntregado', 'promesaEntrega',
    'promisedDays',
    'codigoDestinatario', 'nombreCliente', 'segmentoCliente',
    'categorizacionCliente',
    'codigoAgente', 'nombreAgente',
    'rutaSAP', 'rutaHivitrack', 'transporteSAP', 'transporteHivitrack',
    'tipologiaTransporte', 'estadoGeneral', 'causalEntrega',
    'lineaNegocio', 'sucursal', 'provincia', 'ciudadDestino',
    'cantidad', 'peso',
]


def _cargar_credenciales():
    """Lee credenciales del keyring; fallback al .env de Voz del Cliente."""
    pwd = keyring.get_password(SERVICE, USERNAME_DEFAULT)
    user = USERNAME_DEFAULT
    if not pwd:
        env_path = (r'C:\Users\consultorindustrial\OneDrive - Hivimar\Documentos'
                    r'\HIVIMAR\Proyectos\Voz del cliente y mejora contínua'
                    r'\Tablero Voz del Cliente\.env')
        if os.path.exists(env_path):
            for line in open(env_path, encoding='utf-8'):
                line = line.strip()
                if line.startswith('HIVITRACK_USERNAME='):
                    user = line.split('=', 1)[1].strip()
                elif line.startswith('HIVITRACK_PASSWORD='):
                    pwd = line.split('=', 1)[1].strip()
    if not pwd:
        raise SystemExit(
            f"ERROR: credenciales Hivitrack no encontradas en keyring "
            f"('{SERVICE}') ni en .env de Voz del Cliente")
    return user, pwd


def _setup_env(user, pwd):
    os.environ['HIVITRACK_USERNAME'] = user
    os.environ['HIVITRACK_PASSWORD'] = pwd
    os.environ.setdefault('HIVITRACK_URL', 'https://hivitrack2.hivimar.com:9000')


def descargar_ventana(dias: int = DIAS_INCREMENTAL,
                       verbose: bool = True) -> List[Dict]:
    """Descarga todas las paginas de Hivitrack para `dias` recientes y
    devuelve solo lineaNegocio=1100 (Industria)."""
    user, pwd = _cargar_credenciales()
    _setup_env(user, pwd)
    import hivitrack_client as h
    h.login()
    if verbose:
        print(f"[hivitrack] Login OK (user={user})")

    t0 = time.time()
    if verbose:
        print(f"[hivitrack] Descargando ventana de {dias} dias "
              f"(rows={ROWS_PER_PAGE})...")
    r = h.buscar_entregas(dias=dias, page=1, rows=ROWS_PER_PAGE)
    total_filas = r.get('total_filas', 0)
    total_paginas = r.get('total_paginas', 1)
    if verbose:
        print(f"[hivitrack] Total filas Hivimar: {total_filas:,}  "
              f"paginas: {total_paginas}")

    industria = []
    for e in r.get('entregas', []):
        if e.get('lineaNegocio') == LINEA_NEGOCIO_INDUSTRIA:
            industria.append(e)

    pag = 1
    while pag < total_paginas:
        pag += 1
        r = h.buscar_entregas(dias=dias, page=pag, rows=ROWS_PER_PAGE)
        for e in r.get('entregas', []):
            if e.get('lineaNegocio') == LINEA_NEGOCIO_INDUSTRIA:
                industria.append(e)
        if verbose and pag % 25 == 0:
            print(f"[hivitrack]  pag {pag}/{total_paginas}  "
                  f"Industria acum={len(industria):,}  ({time.time()-t0:.0f}s)")
    if verbose:
        print(f"[hivitrack] Descarga completa: {len(industria):,} entregas "
              f"Industria en {time.time()-t0:.0f}s")
    return industria


def _parse_fecha_creacion(s: str) -> Optional[date]:
    if not s:
        return None
    try:
        return datetime.fromisoformat(str(s).replace('Z', '')).date()
    except Exception:
        return None


def cargar_historico(path: str = DEST_CSV) -> Dict[str, Dict]:
    """Carga el CSV historico en un dict {numeroEntrega -> fila_dict}."""
    if not os.path.exists(path):
        return {}
    out = {}
    with open(path, encoding='utf-8-sig', newline='') as f:
        rd = csv.DictReader(f, delimiter=';')
        for row in rd:
            ne = (row.get('numeroEntrega') or '').strip()
            if ne:
                out[ne] = row
    return out


def upsert(historico: Dict[str, Dict],
           nuevas: List[Dict],
           verbose: bool = True) -> Dict[str, Dict]:
    """Aplica upsert por numeroEntrega. Descarta entregas con creacion
    anterior a FECHA_INICIO_HIVITRACK."""
    n_new = 0
    n_upd = 0
    n_skip_old = 0
    for e in nuevas:
        ne = str(e.get('numeroEntrega') or '').strip()
        if not ne:
            continue
        # Descartar muy viejas (pre-Hivitrack en produccion)
        fc = _parse_fecha_creacion(e.get('fechaHoraCreacionPedido'))
        if fc and fc < FECHA_INICIO_HIVITRACK:
            n_skip_old += 1
            continue
        row = {col: e.get(col, '') for col in COLUMNAS}
        if ne in historico:
            n_upd += 1
        else:
            n_new += 1
        historico[ne] = row
    if verbose:
        print(f"[upsert] nuevas: {n_new:,}  actualizadas: {n_upd:,}  "
              f"viejas descartadas: {n_skip_old:,}  "
              f"total historico: {len(historico):,}")
    return historico


def escribir_csv(historico: Dict[str, Dict], path: str = DEST_CSV,
                 verbose: bool = True):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.DictWriter(f, fieldnames=COLUMNAS, delimiter=';',
                           extrasaction='ignore')
        w.writeheader()
        # Ordenar por fechaHoraCreacionPedido para que el CSV sea estable
        items = sorted(historico.values(),
                       key=lambda r: r.get('fechaHoraCreacionPedido') or '')
        for row in items:
            w.writerow(row)
    if verbose:
        sz = os.path.getsize(path) / 1024 / 1024
        print(f"[csv] escritas {len(historico):,} filas en {path} "
              f"({sz:.1f} MB)")


def main():
    t0 = time.time()
    print(f"=== extraer_entregas_hivitrack.py "
          f"{datetime.now().isoformat(timespec='seconds')} ===")
    print(f"Destino: {DEST_CSV}")
    print(f"Ventana incremental: ultimos {DIAS_INCREMENTAL} dias")
    print(f"Filtro: lineaNegocio={LINEA_NEGOCIO_INDUSTRIA}")
    print()

    nuevas = descargar_ventana(verbose=True)
    historico = cargar_historico()
    print(f"\n[historico] {len(historico):,} entregas previas en CSV")
    historico = upsert(historico, nuevas, verbose=True)
    escribir_csv(historico, verbose=True)
    print(f"\nOK en {time.time()-t0:.0f}s")


if __name__ == '__main__':
    main()
