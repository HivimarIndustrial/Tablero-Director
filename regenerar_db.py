"""
Regenerar window.DB para Hivimar_Tablero_Industrial desde Base Tablero Industria.xlsx
Genera db_output.js con el objeto DB actualizado.
"""
import pandas as pd
import json
import math
import os
from datetime import datetime, date

# Rutas relativas al script: funciona en cualquier PC mientras los archivos estén
# en la misma carpeta que regenerar_db.py
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL = os.path.join(SCRIPT_DIR, "Base Tablero Industria.xlsx")
OUTPUT = os.path.join(SCRIPT_DIR, "db_output.js")

def safe(v):
    """Convert value to JSON-safe type"""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    if pd.isna(v):
        return None
    if isinstance(v, (datetime, date)):
        return v.strftime('%Y-%m-%d')
    return v

def safe_num(v):
    if v is None or (isinstance(v, float) and (math.isnan(v) or math.isinf(v))):
        return 0
    try:
        return round(float(v), 2)
    except:
        return 0

def safe_str(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ''
    return str(v).strip()

def km(anio, mes):
    """Build YYYY-MM key"""
    return f"{int(anio)}-{int(mes):02d}"

print("Leyendo dims del Excel (solo lectura)...")
# Las dim_* y el ppto financiero SÍ viven en Excel (tú las mantienes a mano)
dim_vend = pd.read_excel(EXCEL, sheet_name='dim_vendedores')
dim_sup = pd.read_excel(EXCEL, sheet_name='dim_supervisores_tablero')
dim_cli = pd.read_excel(EXCEL, sheet_name='dim_clientes')
dim_grupo = pd.read_excel(EXCEL, sheet_name='dim_grupo_articulo_responsable')
raw_ppto = pd.read_excel(EXCEL, sheet_name='raw_presupuesto2026')

# Los raw_* (datos transaccionales) pueden venir de dos fuentes:
#   1) CSVs enriquecidos en automatizacion/salida_raw/ (preferido, generados por el pipeline)
#   2) Hojas del Excel (fallback, para ejecuciones manuales)
SALIDA_RAW = os.path.join(SCRIPT_DIR, 'salida_raw')

def _leer_raw(nombre_csv, nombre_excel):
    """Lee el CSV enriquecido. Si no existe, falla con mensaje claro.

    NOTA: el fallback al Excel fue deliberadamente removido porque las
    hojas raw_* del Excel son AHORA solo referencia y pueden estar
    recortadas; cargarlas daría datos incompletos sin avisar.
    Si ves este error, corre el pipeline completo primero:
        python automatizacion/update_tablero.py --skip-vpn
    """
    csv_path = os.path.join(SALIDA_RAW, nombre_csv)
    if not os.path.exists(csv_path):
        raise FileNotFoundError(
            f"No existe el CSV enriquecido: {csv_path}\n"
            f"Corre primero el pipeline con:\n"
            f"    python automatizacion/update_tablero.py --skip-vpn\n"
            f"Eso genera los CSVs en salida_raw/ que este script necesita."
        )
    print(f"  [CSV] {nombre_csv}")
    try:
        return pd.read_csv(csv_path, encoding='utf-8')
    except pd.errors.EmptyDataError:
        # CSV vacio (0 filas Y sin header). Pasa cuando el extractor Qlik
        # devuelve 0 filas (ej. cartera a inicio de mes / festivo, antes
        # del fallback al mes anterior). En vez de tumbar el HTML del
        # tablero, devolvemos un DataFrame vacio y avisamos.
        print(f"  [CSV] AVISO: {nombre_csv} esta vacio -> DataFrame vacio")
        return pd.DataFrame()

print("Leyendo datos transaccionales (CSV enriquecido preferido, Excel fallback)...")
raw_ventas = _leer_raw('ventas_enriquecido.csv',        'raw_ventas')
raw_cartera = _leer_raw('cartera_enriquecido.csv',      'raw_cartera')
raw_cots = _leer_raw('cotizaciones_enriquecido.csv',    'raw_cotizaciones')
raw_vis = _leer_raw('visitas_enriquecido.csv',          'raw_visitas')
raw_acts = _leer_raw('actividades_enriquecido.csv',     'raw_actividadespendientes')
raw_inv = _leer_raw('inventario_enriquecido.csv',       'raw_inventario')
fact_opps = _leer_raw('oportunidades_enriquecido.csv',  'fact_oportunidades')

# NUEVO: lee dim_ppto_vendedor (solo presupuesto limpio). La venta real y el
# cumplimiento se calculan mas abajo desde raw_ventas.
# Mantiene compatibilidad con fact_vtas_vs_ppto si la hoja nueva no existe.
try:
    dim_ppto_vend = pd.read_excel(EXCEL, sheet_name='dim_ppto_vendedor')
    print(f"dim_ppto_vendedor: {len(dim_ppto_vend)} rows (fuente de ppto)")
    _USE_DIM_PPTO = True
except Exception as e:
    print(f"dim_ppto_vendedor no existe, usando fact_vtas_vs_ppto (fallback)")
    dim_ppto_vend = pd.read_excel(EXCEL, sheet_name='fact_vtas_vs_ppto')
    _USE_DIM_PPTO = False
print(f"raw_ventas: {len(raw_ventas)} rows")
print(f"raw_cartera: {len(raw_cartera)} rows")
print(f"raw_cotizaciones: {len(raw_cots)} rows")
print(f"fact_oportunidades: {len(fact_opps)} rows")

# ══ DIMENSION LOOKUPS ══
# Vendedores
vend_sup_map = {}
vend_tipo_map = {}
vend_display = {}
for _, r in dim_vend.iterrows():
    vs = safe_str(r.get('vendedor_std'))
    if not vs:
        continue
    sup = safe_str(r.get('supervisor_std'))
    tipo = safe_str(r.get('tipo_vendedor'))
    vista = safe_str(r.get('vista_tablero'))
    vend_sup_map[vs] = sup
    vend_tipo_map[vs] = tipo if tipo else 'Generalista'
    vend_display[vs] = vs

# Supervisores
sup_comercial = []
sup_jp = []
for _, r in dim_sup.iterrows():
    s = safe_str(r.get('supervisor_std'))
    vista = safe_str(r.get('vista_tablero'))
    if vista == 'COMERCIAL':
        sup_comercial.append(s)
    else:
        sup_jp.append(s)

# Segmentos
SEGS = ['General Industries', 'Heavy Industries', 'Mining', 'Oil and Power', 'Sector Estratégico']

# Grupo -> responsable/tipo
grupo_resp = {}
for _, r in dim_grupo.iterrows():
    g = safe_str(r.get('grupo_articulo_std'))
    if g:
        grupo_resp[g] = {
            'resp': safe_str(r.get('jefe_producto_std')) or 'COMERCIO',
            'sup': safe_str(r.get('supervisor_producto_std')) or '',
            'tipo': safe_str(r.get('tipo_grupo_artículo')) or 'OTROS'
        }

# Cliente -> segmento lookup
cli_seg = {}
for _, r in dim_cli.iterrows():
    cid = safe_str(r.get('cliente_id_std'))
    seg = safe_str(r.get('Segmento'))
    if cid and seg:
        cli_seg[cid] = seg

print("Procesando ventas...")

# ══ VENTAS ══
# Filter: only rows with valid data
rv = raw_ventas.dropna(subset=['AÑO', 'MES', 'Venta neta'])

ventas_mes = {}          # {YYYY-MM: {venta, costo}}
ventas_seg_mes = {}      # {seg: {YYYY-MM: {venta, costo}}}
ventas_seg_mes_all = {}  # same but ALL vendedores
ventas_vend_mes = {}     # {vendedor: {YYYY-MM: {venta, costo}}} - solo Generalistas COMERCIAL
ventas_vend_seg_mes = {} # {vendedor: {seg: {YYYY-MM: {venta, costo}}}} - solo Generalistas
ventas_tg_mes = {}       # {tipo_grupo: {YYYY-MM: {venta, costo}}}
ventas_seg_tg_mes = {}   # {seg: {tipo_grupo: {YYYY-MM: {venta, costo}}}}
ventas_vend_tg_mes = {}  # {vendedor: {tipo_grupo: {YYYY-MM: {venta, costo}}}} - solo Generalistas
ventas_vend_cli = {}     # {vendedor: {cliente: {venta, costo}}}
ventas_grupo_mes = {}    # {grupo: {YYYY-MM: {venta, costo}}}
ventas_resp_mes = {}     # {resp: {YYYY-MM: {venta, costo}}}

for _, r in rv.iterrows():
    k = km(r['AÑO'], r['MES'])
    venta = safe_num(r.get('Venta neta'))
    costo = safe_num(r.get('Costo Interno'))
    vendedor = safe_str(r.get('vendedor_principal_std'))
    seg = safe_str(r.get('segmento'))
    cli = safe_str(r.get('cliente_nombre_std'))
    grupo = safe_str(r.get('grupo_articulo_std'))
    vista = safe_str(r.get('vista_tablero'))
    resp = safe_str(r.get('responsable_producto_std')) or 'COMERCIO'

    # ventas_mes (total)
    if k not in ventas_mes:
        ventas_mes[k] = {'venta': 0, 'costo': 0}
    ventas_mes[k]['venta'] += venta
    ventas_mes[k]['costo'] += costo

    # ventas_seg_mes_all (todos los vendedores)
    if seg:
        if seg not in ventas_seg_mes_all:
            ventas_seg_mes_all[seg] = {}
        if k not in ventas_seg_mes_all[seg]:
            ventas_seg_mes_all[seg][k] = {'venta': 0}
        ventas_seg_mes_all[seg][k]['venta'] += venta

    # ventas_seg_mes (solo COMERCIAL)
    if seg and vista == 'COMERCIAL':
        if seg not in ventas_seg_mes:
            ventas_seg_mes[seg] = {}
        if k not in ventas_seg_mes[seg]:
            ventas_seg_mes[seg][k] = {'venta': 0, 'costo': 0}
        ventas_seg_mes[seg][k]['venta'] += venta
        ventas_seg_mes[seg][k]['costo'] += costo

    # ventas_seg_tg_mes (segmento × tipo_grupo × mes) para proyección
    tg_v = safe_str(r.get('tipo_grupo_artículo')) or grupo_resp.get(grupo, {}).get('tipo', 'OTROS')
    if seg and tg_v:
        if seg not in ventas_seg_tg_mes:
            ventas_seg_tg_mes[seg] = {}
        if tg_v not in ventas_seg_tg_mes[seg]:
            ventas_seg_tg_mes[seg][tg_v] = {}
        if k not in ventas_seg_tg_mes[seg][tg_v]:
            ventas_seg_tg_mes[seg][tg_v][k] = {'venta': 0, 'costo': 0}
        ventas_seg_tg_mes[seg][tg_v][k]['venta'] += venta
        ventas_seg_tg_mes[seg][tg_v][k]['costo'] += costo

    # ventas por vendedor (solo Generalistas COMERCIAL)
    tipo = vend_tipo_map.get(vendedor, '')
    if vendedor and vista == 'COMERCIAL' and tipo != 'Especializado':
        # ventas_vend_mes
        if vendedor not in ventas_vend_mes:
            ventas_vend_mes[vendedor] = {}
        if k not in ventas_vend_mes[vendedor]:
            ventas_vend_mes[vendedor][k] = {'venta': 0, 'costo': 0}
        ventas_vend_mes[vendedor][k]['venta'] += venta
        ventas_vend_mes[vendedor][k]['costo'] += costo

        # NEW: ventas_vend_seg_mes
        if seg:
            if vendedor not in ventas_vend_seg_mes:
                ventas_vend_seg_mes[vendedor] = {}
            if seg not in ventas_vend_seg_mes[vendedor]:
                ventas_vend_seg_mes[vendedor][seg] = {}
            if k not in ventas_vend_seg_mes[vendedor][seg]:
                ventas_vend_seg_mes[vendedor][seg][k] = {'venta': 0, 'costo': 0}
            ventas_vend_seg_mes[vendedor][seg][k]['venta'] += venta
            ventas_vend_seg_mes[vendedor][seg][k]['costo'] += costo

        # ventas_vend_tg_mes
        tg = grupo_resp.get(grupo, {}).get('tipo', 'OTROS') if grupo else 'OTROS'
        if vendedor not in ventas_vend_tg_mes:
            ventas_vend_tg_mes[vendedor] = {}
        if tg not in ventas_vend_tg_mes[vendedor]:
            ventas_vend_tg_mes[vendedor][tg] = {}
        if k not in ventas_vend_tg_mes[vendedor][tg]:
            ventas_vend_tg_mes[vendedor][tg][k] = {'venta': 0, 'costo': 0}
        ventas_vend_tg_mes[vendedor][tg][k]['venta'] += venta
        ventas_vend_tg_mes[vendedor][tg][k]['costo'] += costo

        # ventas_vend_cli
        if cli:
            if vendedor not in ventas_vend_cli:
                ventas_vend_cli[vendedor] = {}
            if cli not in ventas_vend_cli[vendedor]:
                ventas_vend_cli[vendedor][cli] = {'venta': 0, 'costo': 0}
            ventas_vend_cli[vendedor][cli]['venta'] += venta
            ventas_vend_cli[vendedor][cli]['costo'] += costo

    # ventas_tg_mes (total por tipo grupo)
    tg_all = grupo_resp.get(grupo, {}).get('tipo', 'OTROS') if grupo else 'OTROS'
    if tg_all not in ventas_tg_mes:
        ventas_tg_mes[tg_all] = {}
    if k not in ventas_tg_mes[tg_all]:
        ventas_tg_mes[tg_all][k] = {'venta': 0, 'costo': 0}
    ventas_tg_mes[tg_all][k]['venta'] += venta
    ventas_tg_mes[tg_all][k]['costo'] += costo

    # ventas_grupo_mes
    if grupo:
        if grupo not in ventas_grupo_mes:
            ventas_grupo_mes[grupo] = {}
        if k not in ventas_grupo_mes[grupo]:
            ventas_grupo_mes[grupo][k] = {'venta': 0, 'costo': 0}
        ventas_grupo_mes[grupo][k]['venta'] += venta
        ventas_grupo_mes[grupo][k]['costo'] += costo

    # ventas_resp_mes
    if resp:
        if resp not in ventas_resp_mes:
            ventas_resp_mes[resp] = {}
        if k not in ventas_resp_mes[resp]:
            ventas_resp_mes[resp][k] = {'venta': 0, 'costo': 0}
        ventas_resp_mes[resp][k]['venta'] += venta
        ventas_resp_mes[resp][k]['costo'] += costo

# Round all values
for d in [ventas_mes]:
    for k2 in d:
        d[k2]['venta'] = round(d[k2]['venta'], 2)
        d[k2]['costo'] = round(d[k2]['costo'], 2)

# ══ YTD CON FECHA DE CORTE ══
# Los datos son mensuales (fecha_std = 1er día del mes).
# Para comparar correctamente: meses completos + prorrateo del mes parcial.
# La fecha de corte real es HOY (cuando se ejecuta el script).
import calendar
print("Calculando YTD con fecha de corte...")

corte_date = date.today()  # Fecha real de corte = hoy
anio_actual = 2026  # Año del tablero
mes_corte = corte_date.month
dia_corte = corte_date.day

# Acumular meses COMPLETOS (ene hasta mes anterior al corte)
ytd_actual = 0
ytd_anterior = 0
for m in range(1, mes_corte):
    km_act = km(anio_actual, m)
    km_ant = km(anio_actual - 1, m)
    ytd_actual += (ventas_mes.get(km_act, {}).get('venta', 0))
    ytd_anterior += (ventas_mes.get(km_ant, {}).get('venta', 0))

# Mes parcial (mes del corte): el año actual tiene datos hasta día_corte
# Prorrateamos el mismo mes del año anterior
km_parcial_act = km(anio_actual, mes_corte)
km_parcial_ant = km(anio_actual - 1, mes_corte)
venta_mes_parcial_act = ventas_mes.get(km_parcial_act, {}).get('venta', 0)
venta_mes_completo_ant = ventas_mes.get(km_parcial_ant, {}).get('venta', 0)

# Prorrateo: del año anterior, solo contar (dia_corte / dias_del_mes) del total mensual
dias_mes_ant = calendar.monthrange(anio_actual - 1, mes_corte)[1]
factor_prorrateo = dia_corte / dias_mes_ant
venta_mes_prorrateado_ant = venta_mes_completo_ant * factor_prorrateo

ytd_actual += venta_mes_parcial_act
ytd_anterior += venta_mes_prorrateado_ant

ytd_crecimiento = round((ytd_actual - ytd_anterior) / ytd_anterior * 100, 2) if ytd_anterior else 0

ytd_info = {
    'anio': anio_actual,
    'corte': corte_date.strftime('%Y-%m-%d'),
    'dia_corte': dia_corte,
    'mes_corte': mes_corte,
    'ytd_actual': round(ytd_actual, 2),
    'ytd_anterior': round(ytd_anterior, 2),
    'crecimiento': ytd_crecimiento,
    'nota': f'Comparación al {dia_corte}/{mes_corte} (mes ant. prorrateado {factor_prorrateo:.0%})'
}
print(f"  Corte: {corte_date.strftime('%Y-%m-%d')} (día {dia_corte} de {mes_corte})")
print(f"  YTD {anio_actual}: ${ytd_actual:,.0f}")
print(f"  YTD {anio_actual-1} (prorrateado): ${ytd_anterior:,.0f}")
print(f"  Abril {anio_actual}: ${venta_mes_parcial_act:,.0f} (parcial)")
print(f"  Abril {anio_actual-1}: ${venta_mes_completo_ant:,.0f} x {factor_prorrateo:.0%} = ${venta_mes_prorrateado_ant:,.0f}")
print(f"  Crecimiento: {ytd_crecimiento:.1f}%")

print("Procesando presupuestos...")

# ══ PRESUPUESTO 2026 ══
ppto_fin_mes = {}
ppto_fin_seg_mes = {}
ppto_fin_grupo_mes = {}
ppto_fin_resp_mes = {}
ppto_grupo_mes = {}
ppto_resp_mes = {}
ppto_seg_tg_mes = {}  # {seg: {tipo_grupo: {YYYY-MM: valor}}}

# Mapeo de Clasificación Productos del ppto a tipo_grupo del tablero
CLASIF_TO_TG = {'ESTRATEGICAS': 'ESTRATÉGICOS', 'CORE': 'CORE', 'MRO': 'MRO', 'AUTOPARTES': 'OTROS'}

for _, r in raw_ppto.iterrows():
    anio = safe_num(r.get('anio'))
    mes = safe_num(r.get('mes_num'))
    if not anio or not mes:
        continue
    k = km(anio, mes)
    vp = safe_num(r.get('venta_ppto'))
    cp = safe_num(r.get('costo_ppto'))
    mp = safe_num(r.get('margen_ppto'))
    seg = safe_str(r.get('segmento'))
    grupo = safe_str(r.get('grupo_articulo'))
    grupo_std = grupo.upper().strip() if grupo else ''
    clasif = safe_str(r.get('Clasificación Productos'))
    tg_ppto = CLASIF_TO_TG.get(clasif, 'OTROS')

    # ppto_seg_tg_mes
    if seg and tg_ppto:
        if seg not in ppto_seg_tg_mes:
            ppto_seg_tg_mes[seg] = {}
        if tg_ppto not in ppto_seg_tg_mes[seg]:
            ppto_seg_tg_mes[seg][tg_ppto] = {}
        if k not in ppto_seg_tg_mes[seg][tg_ppto]:
            ppto_seg_tg_mes[seg][tg_ppto][k] = 0
        ppto_seg_tg_mes[seg][tg_ppto][k] += vp

    # ppto_fin_mes (total)
    if k not in ppto_fin_mes:
        ppto_fin_mes[k] = {'venta': 0, 'costo': 0, 'margen': 0}
    ppto_fin_mes[k]['venta'] += vp
    ppto_fin_mes[k]['costo'] += cp
    ppto_fin_mes[k]['margen'] += mp

    # ppto_fin_seg_mes
    if seg:
        if seg not in ppto_fin_seg_mes:
            ppto_fin_seg_mes[seg] = {}
        if k not in ppto_fin_seg_mes[seg]:
            ppto_fin_seg_mes[seg][k] = 0
        ppto_fin_seg_mes[seg][k] += vp

    # ppto by grupo (for product view)
    if grupo_std:
        # Try to find the matching grupo in dim_grupo
        resp_info = grupo_resp.get(grupo_std, {})
        resp_name = resp_info.get('resp', 'COMERCIO')

        if grupo_std not in ppto_grupo_mes:
            ppto_grupo_mes[grupo_std] = {}
        if k not in ppto_grupo_mes[grupo_std]:
            ppto_grupo_mes[grupo_std][k] = 0
        ppto_grupo_mes[grupo_std][k] += vp

        if resp_name not in ppto_resp_mes:
            ppto_resp_mes[resp_name] = {}
        if k not in ppto_resp_mes[resp_name]:
            ppto_resp_mes[resp_name][k] = 0
        ppto_resp_mes[resp_name][k] += vp

    # ppto_fin by grupo
    if grupo_std:
        if grupo_std not in ppto_fin_grupo_mes:
            ppto_fin_grupo_mes[grupo_std] = {}
        if k not in ppto_fin_grupo_mes[grupo_std]:
            ppto_fin_grupo_mes[grupo_std][k] = 0
        ppto_fin_grupo_mes[grupo_std][k] += vp

        if resp_name not in ppto_fin_resp_mes:
            ppto_fin_resp_mes[resp_name] = {}
        if k not in ppto_fin_resp_mes[resp_name]:
            ppto_fin_resp_mes[resp_name][k] = 0
        ppto_fin_resp_mes[resp_name][k] += vp

# Round pptos
for d in [ppto_fin_mes]:
    for k2 in d:
        for f in d[k2]:
            d[k2][f] = round(d[k2][f], 2)

print("Procesando presupuesto comercial por vendedor...")

# ══ PRESUPUESTO COMERCIAL POR VENDEDOR ══
# Fuente de ppto: dim_ppto_vendedor (nueva hoja, solo tiene ppto)
# Fuente de venta real: raw_ventas (calculada al vuelo)
# Cumplimiento = venta_real / ppto (se calcula en el HTML o aqui si hace falta)

# ppto_vend_mes: {vendedor: {YYYY-MM: {ppto, real}}} - agregado por vendedor+mes
# ppto_vend_grupo_mes: {vendedor: {grupo: {YYYY-MM: {ppto, real}}}} - detalle por grupo
ppto_vend_mes = {}
ppto_vend_grupo_mes = {}

# 1) Agregar PPTO desde dim_ppto_vendedor
for _, r in dim_ppto_vend.iterrows():
    anio_v = safe_num(r.get('anio'))
    mes_v = safe_num(r.get('mes'))
    if not anio_v or not mes_v:
        continue
    k_v = km(anio_v, mes_v)
    vendedor_v = safe_str(r.get('vendedor_principal_std'))
    grupo_v = safe_str(r.get('grupo_articulo_std'))
    ppto_v = safe_num(r.get('presupuesto'))
    if not vendedor_v:
        continue
    # Filtrar a vendedores COMERCIAL si estamos usando fallback (fact_vtas_vs_ppto)
    if not _USE_DIM_PPTO:
        vista_v = safe_str(r.get('vista_tablero'))
        if vista_v != 'COMERCIAL':
            continue

    # Agregado vendedor x mes
    if vendedor_v not in ppto_vend_mes:
        ppto_vend_mes[vendedor_v] = {}
    if k_v not in ppto_vend_mes[vendedor_v]:
        ppto_vend_mes[vendedor_v][k_v] = {'ppto': 0, 'real': 0}
    ppto_vend_mes[vendedor_v][k_v]['ppto'] += ppto_v

    # Detalle vendedor x grupo x mes
    if grupo_v:
        if vendedor_v not in ppto_vend_grupo_mes:
            ppto_vend_grupo_mes[vendedor_v] = {}
        if grupo_v not in ppto_vend_grupo_mes[vendedor_v]:
            ppto_vend_grupo_mes[vendedor_v][grupo_v] = {}
        if k_v not in ppto_vend_grupo_mes[vendedor_v][grupo_v]:
            ppto_vend_grupo_mes[vendedor_v][grupo_v][k_v] = {'ppto': 0, 'real': 0}
        ppto_vend_grupo_mes[vendedor_v][grupo_v][k_v]['ppto'] += ppto_v

# 2) Agregar VENTA REAL calculada desde raw_ventas
# Reutilizamos las estructuras ya calculadas arriba:
#   - ventas_vend_mes: {vendedor: {YYYY-MM: {venta, costo}}}  -> sumamos a 'real'
#   - ventas_vend_cli: (no se usa aqui)
# Para detalle por grupo iteramos raw_ventas directamente.
for vendedor_v, mdict in ventas_vend_mes.items():
    for k_v, vals in mdict.items():
        if vendedor_v not in ppto_vend_mes:
            ppto_vend_mes[vendedor_v] = {}
        if k_v not in ppto_vend_mes[vendedor_v]:
            ppto_vend_mes[vendedor_v][k_v] = {'ppto': 0, 'real': 0}
        ppto_vend_mes[vendedor_v][k_v]['real'] = vals.get('venta', 0)

# Detalle por vendedor x grupo x mes desde raw_ventas
for _, r in rv.iterrows():
    vendedor_v = safe_str(r.get('vendedor_principal_std'))
    grupo_v = safe_str(r.get('grupo_articulo_std'))
    if not vendedor_v or not grupo_v:
        continue
    vista_v = safe_str(r.get('vista_tablero'))
    tipo_vend = vend_tipo_map.get(vendedor_v, '')
    if vista_v != 'COMERCIAL' or tipo_vend == 'Especializado':
        continue
    try:
        anio_v = int(r.get('AÑO'))
        mes_v = int(r.get('MES'))
    except (TypeError, ValueError):
        continue
    k_v = km(anio_v, mes_v)
    venta_v = safe_num(r.get('Venta neta'))

    if vendedor_v not in ppto_vend_grupo_mes:
        ppto_vend_grupo_mes[vendedor_v] = {}
    if grupo_v not in ppto_vend_grupo_mes[vendedor_v]:
        ppto_vend_grupo_mes[vendedor_v][grupo_v] = {}
    if k_v not in ppto_vend_grupo_mes[vendedor_v][grupo_v]:
        ppto_vend_grupo_mes[vendedor_v][grupo_v][k_v] = {'ppto': 0, 'real': 0}
    ppto_vend_grupo_mes[vendedor_v][grupo_v][k_v]['real'] += venta_v

# Round
for v2 in ppto_vend_mes:
    for k_v2 in ppto_vend_mes[v2]:
        ppto_vend_mes[v2][k_v2]['ppto'] = round(ppto_vend_mes[v2][k_v2]['ppto'], 2)
        ppto_vend_mes[v2][k_v2]['real'] = round(ppto_vend_mes[v2][k_v2]['real'], 2)

for v2 in ppto_vend_grupo_mes:
    for g2 in ppto_vend_grupo_mes[v2]:
        for k_v2 in ppto_vend_grupo_mes[v2][g2]:
            ppto_vend_grupo_mes[v2][g2][k_v2]['ppto'] = round(ppto_vend_grupo_mes[v2][g2][k_v2]['ppto'], 2)
            ppto_vend_grupo_mes[v2][g2][k_v2]['real'] = round(ppto_vend_grupo_mes[v2][g2][k_v2]['real'], 2)

print(f"  {len(ppto_vend_mes)} vendedores con ppto (vendedor x mes)")
print(f"  {len(ppto_vend_grupo_mes)} vendedores con detalle por grupo (vendedor x grupo x mes)")

print("Procesando cotizaciones...")

# ══ COTIZACIONES ══
# Group by Referencia (each quote has multiple lines)
cot_groups = raw_cots.groupby('Referencia del pedido')
cotizaciones = []
for ref, group in cot_groups:
    first = group.iloc[0]
    total = safe_num(group['Total'].sum())
    fecha = safe_str(first.get('fecha_std')) or safe(first.get('Creado'))
    estado = safe_str(first.get('estado_std'))
    # Map estado to short codes
    estado_map = {'Presupuesto': 'draft', 'Enviado': 'sent', 'Cotizacion Ganada': 'sale', 'Cancelado': 'cancel'}
    e_code = estado_map.get(estado, 'draft')

    gr_cot = safe_str(first.get('grupo_articulo_std'))
    tg_cot = grupo_resp.get(gr_cot, {}).get('tipo', 'OTROS')
    cotizaciones.append({
        'r': safe_str(ref),
        'f': safe_str(fecha)[:10] if fecha else '',
        'co': safe_str(first.get('vendedor_principal_std')),
        'su': safe_str(first.get('supervisor_std')),
        'cl': safe_str(first.get('cliente_nombre_std')),
        'gr': gr_cot,
        'resp': safe_str(first.get('responsable_producto_std')) or 'COMERCIO',
        't': round(total, 2),
        'e': e_code,
        'tg': tg_cot
    })

print(f"  {len(cotizaciones)} cotizaciones agrupadas")

print("Procesando oportunidades...")

# ══ OPORTUNIDADES ══
# Clasificación usando estado_oportunidad (fuente de verdad del Excel):
#   Activa           → pipeline activo (incluye ganado sin facturar)
#   Ganada Facturada  → cerrada ganada, ya facturada
#   Perdida/Eliminada → cerrada perdida, NO cuenta en pipeline
opps_activas = []
opps_cerradas = []

for _, r in fact_opps.iterrows():
    estado = safe_str(r.get('estado_oportunidad'))

    # Solo procesar estados válidos
    if estado not in ('Activa', 'Ganada Facturada', 'Perdida', 'Eliminada'):
        continue

    ingreso = safe_num(r.get('Ingreso esperado'))
    etapa = safe_str(r.get('Etapa'))
    vendedor = safe_str(r.get('vendedor_principal_std'))
    supervisor = safe_str(r.get('supervisor_std'))

    grupo_opp = safe_str(r.get('grupo_articulo_std'))
    # Usar columnas directas del Excel (más preciso que lookup)
    tipo_g_opp = safe_str(r.get('tipo_grupo_artículo')) or grupo_resp.get(grupo_opp, {}).get('tipo', 'OTROS')
    seg_opp = safe_str(r.get('segmento_cliente'))
    if seg_opp == 'Sin Segmento':
        seg_opp = ''

    opp = {
        'seq': safe_str(r.get('Name Sequence')),
        'op': safe_str(r.get('Oportunidad')),
        'cl': safe_str(r.get('cliente_nombre_std')),
        'etapa': etapa,
        'co': vendedor,
        'esp': safe_str(r.get('especialista_std')),
        'su': supervisor,
        'gr': grupo_opp,
        'resp': safe_str(r.get('responsable_producto_std')) or 'COMERCIO',
        'canal': safe_str(r.get('Canal')),
        'ingreso': ingreso,
        'tg': tipo_g_opp if tipo_g_opp != 'Sin Tipo' else 'OTROS',
        'seg': seg_opp,
        'vencida_act': False,
        'anio_c': int(safe_num(r.get('anio_creacion'))) if safe_num(r.get('anio_creacion')) else None,
        'mes_c': int(safe_num(r.get('mes_creacion'))) if safe_num(r.get('mes_creacion')) else None,
    }

    if estado == 'Activa':
        # Pipeline activo: todas las etapas activas + ganado sin facturar
        act_fecha = safe(r.get('Fecha limite de la siguiente actividad'))
        if act_fecha and isinstance(act_fecha, str) and act_fecha > '':
            try:
                act_date = datetime.strptime(act_fecha[:10], '%Y-%m-%d').date()
                opp['vencida_act'] = act_date < date.today()
            except:
                pass
        opps_activas.append(opp)
    else:
        # Cerradas: Ganada Facturada o Perdida/Eliminada
        opp['anio'] = int(safe_num(r.get('anio_cierre'))) if safe_num(r.get('anio_cierre')) else None
        opp['mes'] = int(safe_num(r.get('mes_cierre'))) if safe_num(r.get('mes_cierre')) else None
        opp['ganada'] = estado == 'Ganada Facturada'
        opp['perdida'] = estado in ('Perdida', 'Eliminada')
        opp['motivo'] = safe_str(r.get('Motivo de pérdida') or r.get('Motivo de perdida')) if opp['perdida'] else ''
        opps_cerradas.append(opp)

print(f"  {len(opps_activas)} activas (pipeline), {len(opps_cerradas)} cerradas (ganadas+perdidas)")

print("Procesando actividades...")

# ══ ACTIVIDADES ══
actividades = []
today = date.today()
for _, r in raw_acts.iterrows():
    vendedor = safe_str(r.get('vendedor_std'))
    fv = safe(r.get('Fecha de vencimiento'))
    fv_str = safe_str(fv)[:10] if fv else ''
    supervisor = vend_sup_map.get(vendedor, '')

    vencida = False
    dias_venc = 0
    if fv_str:
        try:
            fv_date = datetime.strptime(fv_str, '%Y-%m-%d').date()
            dias_venc = (today - fv_date).days
            vencida = dias_venc > 0
        except:
            pass

    actividades.append({
        'doc': safe_str(r.get('Nombre del documento')),
        'asignado': vendedor,
        'supervisor': supervisor,
        'tipo': safe_str(r.get('Tipo de actividad')),
        'resumen': safe_str(r.get('Resumen')),
        'fecha_venc': fv_str,
        'vencida': vencida,
        'dias_venc': dias_venc
    })

print(f"  {len(actividades)} actividades")

print("Procesando visitas...")

# ══ VISITAS ══
visitas = []
for _, r in raw_vis.iterrows():
    vendedor = safe_str(r.get('vendedor_std'))
    estado = safe_str(r.get('Estado'))
    supervisor = vend_sup_map.get(vendedor, '')
    cliente = safe_str(r.get('cliente_nombre_std'))
    duracion = safe_str(r.get('Duracion entrada/salida'))

    # Parse dates for anio/mes
    fecha = safe(r.get('Fecha planificada')) or safe(r.get('Fecha efectiva'))
    anio = None
    mes = None
    hora_entrada = None
    if fecha:
        try:
            if isinstance(fecha, str):
                fd = datetime.strptime(fecha[:10], '%Y-%m-%d')
            else:
                fd = fecha
            anio = fd.year
            mes = fd.month
        except:
            pass

    # Parse hora_entrada from Fecha de entrada
    fe = safe(r.get('Fecha de entrada'))
    if fe:
        try:
            if isinstance(fe, str) and len(fe) > 10:
                parts = fe.split(' ')
                if len(parts) > 1:
                    tp = parts[1].split(':')
                    hora_entrada = int(tp[0]) + int(tp[1])/60
            elif isinstance(fe, datetime):
                hora_entrada = fe.hour + fe.minute/60
        except:
            pass

    visitas.append({
        'responsable': vendedor,
        'supervisor': supervisor,
        'cliente': cliente or '',
        'anio': anio,
        'mes': mes,
        'estado': estado,
        'duracion': duracion or '0h 0m',
        'hora_entrada': round(hora_entrada, 2) if hora_entrada is not None else None
    })

print(f"  {len(visitas)} visitas")

print("Procesando cartera...")

# ══ CARTERA ══ (NOW WITH SEGMENTO!)
cartera = []
for _, r in raw_cartera.iterrows():
    importe = safe_num(r.get('Importe'))
    vendedor = safe_str(r.get('vendedor_principal_std'))
    supervisor = safe_str(r.get('supervisor_std'))
    cliente = safe_str(r.get('cliente_nombre_std'))
    vencimiento = safe_str(r.get('VENCIMIENTO'))
    segmento = safe_str(r.get('Segmento'))  # NEW: from dim_clientes lookup

    cartera.append({
        'agente': vendedor,
        'supervisor': supervisor,
        'cliente': cliente,
        'vencimiento': vencimiento,
        'importe': importe,
        'segmento': segmento  # NEW FIELD
    })

print(f"  {len(cartera)} registros de cartera")

print("Procesando inventario...")

# ══ INVENTARIO ══
inv_grupo = {}
inv_resp = {}
for _, r in raw_inv.iterrows():
    grupo = safe_str(r.get('grupo_articulo_std'))
    valor = safe_num(r.get('$ Stock Total'))
    uds = safe_num(r.get('Und Stock Total'))
    resp = safe_str(r.get('jefe_producto_std')) or 'COMERCIO'

    # Leer tipo_grupo_articulo directo del raw_inventario (columna nueva)
    tipo_g = safe_str(r.get('tipo_grupo_articulo')) or safe_str(r.get('tipo_grupo_artículo')) or grupo_resp.get(grupo, {}).get('tipo', 'OTROS')

    if grupo:
        if grupo not in inv_grupo:
            inv_grupo[grupo] = {'valor': 0, 'unidades': 0, 'resp': resp, 'tipo': tipo_g}
        inv_grupo[grupo]['valor'] += valor
        inv_grupo[grupo]['unidades'] += uds

        if resp not in inv_resp:
            inv_resp[resp] = {'valor': 0, 'unidades': 0}
        inv_resp[resp]['valor'] += valor
        inv_resp[resp]['unidades'] += uds

# Round
for g in inv_grupo:
    inv_grupo[g]['valor'] = round(inv_grupo[g]['valor'], 2)
    inv_grupo[g]['unidades'] = round(inv_grupo[g]['unidades'], 0)
for rr in inv_resp:
    inv_resp[rr]['valor'] = round(inv_resp[rr]['valor'], 2)
    inv_resp[rr]['unidades'] = round(inv_resp[rr]['unidades'], 0)

print("Procesando rotacion de inventario (SKU Profiler)...")

# ══ ROTACION DE INVENTARIO ══
# Fuente: salida_raw/rotacion_profiler.csv (ventas + stock de toda la empresa
# por grupo_articulo x marca x AÑO x MES durante los ultimos 13 meses).
# Calculo: Rotacion anual, Cobertura (meses), Dias de inventario, ABC.
rotacion_por_grupo = {}
rotacion_total = {}
rotacion_path = os.path.join(SALIDA_RAW, 'rotacion_profiler.csv')
if os.path.exists(rotacion_path):
    df_rot = pd.read_csv(rotacion_path, encoding='utf-8')
    print(f"  [CSV] rotacion_profiler.csv: {len(df_rot)} filas")
    # Normalizar nombres - Qlik devuelve 'AÑO' (puede venir con encoding raro)
    df_rot.columns = [c.strip() for c in df_rot.columns]
    # Buscar columna AÑO con cualquier capitalizacion
    anio_col = next((c for c in df_rot.columns if c.upper() in ('AÑO', 'ANIO', 'ANO')), 'AÑO')
    mes_col = next((c for c in df_rot.columns if c.upper() == 'MES'), 'MES')

    # Determinar mes mas reciente disponible
    df_rot[anio_col] = pd.to_numeric(df_rot[anio_col], errors='coerce')
    df_rot[mes_col] = pd.to_numeric(df_rot[mes_col], errors='coerce')
    df_rot = df_rot.dropna(subset=[anio_col, mes_col])
    df_rot['anio_mes'] = df_rot[anio_col].astype(int) * 100 + df_rot[mes_col].astype(int)
    ultimo_am = int(df_rot['anio_mes'].max())
    ultimo_anio, ultimo_mes = ultimo_am // 100, ultimo_am % 100
    print(f"  mes mas reciente en rotacion: {ultimo_anio}-{ultimo_mes:02d}")

    # Calcular por grupo_articulo
    df_rot['grupo_articulo'] = df_rot['grupo_articulo'].astype(str).str.strip().str.upper()
    # Stock actual = stock del ultimo mes, promediado por grupo (suma de marcas)
    stock_mes_actual = df_rot[df_rot['anio_mes'] == ultimo_am].groupby('grupo_articulo').agg(
        stock_actual=('Stock', 'sum')
    ).reset_index()
    # Salidas 12m = suma de Venta y Cantidad en los ultimos 12 meses (excl. mes actual si esta parcial)
    # Usamos desde hace 12 meses (inclusivo) hasta el ultimo disponible
    meses_12 = sorted(df_rot['anio_mes'].unique())[-12:]
    df_12m = df_rot[df_rot['anio_mes'].isin(meses_12)]
    salidas_12m = df_12m.groupby('grupo_articulo').agg(
        venta_12m=('Venta', 'sum'),
        cantidad_12m=('Cantidad', 'sum'),
        costo_12m=('Costo', 'sum'),
    ).reset_index()
    # Stock promedio = promedio mensual (de los meses presentes en 12m)
    stock_por_mes = df_12m.groupby(['grupo_articulo', 'anio_mes']).agg(
        stock=('Stock', 'sum')).reset_index()
    stock_prom = stock_por_mes.groupby('grupo_articulo').agg(
        stock_promedio=('stock', 'mean')).reset_index()

    # Merge
    rot_grupo = stock_mes_actual.merge(salidas_12m, on='grupo_articulo', how='outer') \
                                 .merge(stock_prom, on='grupo_articulo', how='outer')
    rot_grupo = rot_grupo.fillna(0)

    # Metricas
    rot_grupo['rotacion_anual'] = rot_grupo.apply(
        lambda r: round(r['costo_12m'] / r['stock_promedio'], 2) if r['stock_promedio'] > 0 else 0, axis=1)
    rot_grupo['cobertura_meses'] = rot_grupo.apply(
        lambda r: round(r['stock_actual'] / (r['venta_12m'] / 12), 2) if r['venta_12m'] > 0 else 0, axis=1)
    rot_grupo['dias_inventario'] = rot_grupo['cobertura_meses'] * 30

    # Clasificacion ABC por $ Stock actual (Pareto)
    rot_grupo_sorted_stock = rot_grupo.sort_values('stock_actual', ascending=False).reset_index(drop=True)
    total_stock = rot_grupo_sorted_stock['stock_actual'].sum()
    acum = 0
    abc_stock = {}
    for _, r in rot_grupo_sorted_stock.iterrows():
        acum += r['stock_actual']
        pct = acum / total_stock * 100 if total_stock else 0
        if pct <= 80:
            abc_stock[r['grupo_articulo']] = 'A'
        elif pct <= 95:
            abc_stock[r['grupo_articulo']] = 'B'
        else:
            abc_stock[r['grupo_articulo']] = 'C'
    # ABC por ventas
    rot_grupo_sorted_venta = rot_grupo.sort_values('venta_12m', ascending=False).reset_index(drop=True)
    total_venta = rot_grupo_sorted_venta['venta_12m'].sum()
    acum = 0
    abc_venta = {}
    for _, r in rot_grupo_sorted_venta.iterrows():
        acum += r['venta_12m']
        pct = acum / total_venta * 100 if total_venta else 0
        if pct <= 80:
            abc_venta[r['grupo_articulo']] = 'A'
        elif pct <= 95:
            abc_venta[r['grupo_articulo']] = 'B'
        else:
            abc_venta[r['grupo_articulo']] = 'C'

    # Construir dict de salida
    for _, r in rot_grupo.iterrows():
        g = r['grupo_articulo']
        rotacion_por_grupo[g] = {
            'stock_actual': round(r['stock_actual'], 2),
            'stock_promedio': round(r['stock_promedio'], 2),
            'venta_12m': round(r['venta_12m'], 2),
            'costo_12m': round(r['costo_12m'], 2),
            'cantidad_12m': round(r['cantidad_12m'], 2),
            'rotacion_anual': r['rotacion_anual'],
            'cobertura_meses': r['cobertura_meses'],
            'dias_inventario': round(r['dias_inventario'], 0),
            'abc_stock': abc_stock.get(g, 'C'),
            'abc_venta': abc_venta.get(g, 'C'),
            # Marca "AC" (mucho stock + poco movimiento) = A en stock + C en ventas
            'es_ac': abc_stock.get(g) == 'A' and abc_venta.get(g) == 'C',
        }

    # Totales globales
    rotacion_total = {
        'mes_referencia': f"{ultimo_anio}-{ultimo_mes:02d}",
        'stock_total': round(rot_grupo['stock_actual'].sum(), 2),
        'venta_12m_total': round(rot_grupo['venta_12m'].sum(), 2),
        'costo_12m_total': round(rot_grupo['costo_12m'].sum(), 2),
        'rotacion_anual_prom': round(
            rot_grupo['costo_12m'].sum() / rot_grupo['stock_promedio'].sum(), 2
        ) if rot_grupo['stock_promedio'].sum() > 0 else 0,
        'cobertura_meses_prom': round(
            rot_grupo['stock_actual'].sum() / (rot_grupo['venta_12m'].sum() / 12), 2
        ) if rot_grupo['venta_12m'].sum() > 0 else 0,
        'grupos_con_datos': int(len(rot_grupo)),
        'grupos_ac': int(sum(1 for g in rotacion_por_grupo.values() if g['es_ac'])),
    }
    rotacion_total['dias_inventario_prom'] = round(rotacion_total['cobertura_meses_prom'] * 30, 0)
    print(f"  {len(rotacion_por_grupo)} grupos con rotacion calculada")
    print(f"  rotacion anual promedio: {rotacion_total['rotacion_anual_prom']}")
    print(f"  cobertura promedio: {rotacion_total['cobertura_meses_prom']} meses")
    print(f"  grupos AC (problema): {rotacion_total['grupos_ac']}")
else:
    print(f"  [WARN] no existe {rotacion_path}; rotacion en 0. Corre pipeline Qlik.")

print("Procesando clientes_cartera...")

# ══ CLIENTES_CARTERA (for visit tracking) ══
clientes_cartera = []
cli_seen = set()
for _, r in dim_cli.iterrows():
    cli = safe_str(r.get('cliente_nombre_std'))
    vend = safe_str(r.get('vendedor_principal_std'))
    sup = safe_str(r.get('supervisor_std'))
    if cli and cli not in cli_seen:
        cli_seen.add(cli)
        # Find last visit for this client
        vis_cli = [v for v in visitas if v['cliente'] == cli and v['estado'] == 'Realizado']
        ultima = ''
        if vis_cli:
            fechas = [(v.get('anio', 0) or 0) * 100 + (v.get('mes', 0) or 0) for v in vis_cli]
            max_f = max(fechas)
            if max_f > 0:
                ultima = f"{max_f // 100}-{max_f % 100:02d}"
        clientes_cartera.append({
            'vend': vend,
            'sup': sup,
            'cli': cli,
            'ultima_visita': ultima
        })

print(f"  {len(clientes_cartera)} clientes")

# ══ ENTREGAS HIVITRACK (Industria) ══
# Lee salida_raw/entregas_industria.csv producido por
# extraer_entregas_hivitrack.py (Paso 2g de update_tablero.py) y produce
# las estructuras que consume la pestana 'Entregas' del tablero.
print("Procesando entregas Hivitrack...")
import math
entregas_csv = os.path.join(SALIDA_RAW, 'entregas_industria.csv')
entregas_kpi_global = {'n_total': 0, 'n_estricto': 0, 'cobertura': 0,
                       'lead_total_mediana': 0, 'credito_mediana': 0,
                       'logistica_mediana': 0,
                       'lead_total_prom': 0, 'credito_prom': 0,
                       'logistica_prom': 0}
entregas_por_clasif = {}
entregas_retira_agente = {'n_retira': 0, 'n_total': 0, 'pct': 0,
                          'top_agentes': [], 'evol_mensual': {}}
entregas_evol_mensual = {}

# Mapa codigoDestinatario (con leading zeros) -> Clasificacion AAA/A/B/C/D
# dim_clientes.cliente_id_std viene SIN ceros (ej '465656'); Hivitrack los
# trae con leading zeros (ej '0000465656'). Normalizamos a SIN ceros.
clas_lookup = {}
sup_lookup_cli = {}
for _, r in dim_cli.iterrows():
    cid = str(r.get('cliente_id_std') or '').strip().lstrip('0')
    if cid:
        clas = safe_str(r.get('Clasificación')) or 'D'
        clas_lookup[cid] = clas
        sup_lookup_cli[cid] = safe_str(r.get('supervisor_std'))

if os.path.exists(entregas_csv):
    ent_df = pd.read_csv(entregas_csv, sep=';', encoding='utf-8-sig',
                          dtype=str)
    print(f"  {len(ent_df)} entregas en entregas_industria.csv")

    def _parse_dt(s):
        if not s or s == 'nan' or s == '':
            return None
        try:
            return pd.to_datetime(s, errors='coerce')
        except Exception:
            return None

    ent_df['t_crea'] = ent_df['fechaHoraCreacionPedido'].apply(_parse_dt)
    ent_df['t_lib']  = ent_df['fechaHoraLiberacionPedido'].apply(_parse_dt)
    ent_df['t_ent']  = ent_df['fechaEntregado'].apply(_parse_dt)
    ent_df['mes']    = ent_df['t_crea'].dt.strftime('%Y-%m')

    # Cliente normalizado y clasificacion (default D si no aparece en dim)
    ent_df['cli_norm'] = ent_df['codigoDestinatario'].astype(str).str.lstrip('0')
    ent_df['clasif'] = ent_df['cli_norm'].map(clas_lookup).fillna('D')

    # Tiempos en dias (NaN si falta alguna fecha)
    def _diff_days(a, b):
        if pd.isna(a) or pd.isna(b):
            return None
        d = (b - a).total_seconds() / 86400
        return d if d >= 0 else None
    ent_df['lead_d']   = ent_df.apply(lambda r: _diff_days(r['t_crea'], r['t_ent']), axis=1)
    ent_df['cred_d']   = ent_df.apply(lambda r: _diff_days(r['t_crea'], r['t_lib']), axis=1)
    ent_df['logi_d']   = ent_df.apply(lambda r: _diff_days(r['t_lib'],  r['t_ent']), axis=1)

    # Marcado retira agente
    ruta_sap = ent_df['rutaSAP'].fillna('').astype(str).str.upper()
    trp_sap  = ent_df['transporteSAP'].fillna('').astype(str).str.upper()
    ent_df['retira_agente'] = ruta_sap.str.contains('RETIROS AGENTE') | (trp_sap == 'RETIRA AGENTE')

    estricto = ent_df.dropna(subset=['lead_d', 'cred_d', 'logi_d'])

    # KPI global
    entregas_kpi_global['n_total'] = int(len(ent_df))
    entregas_kpi_global['n_estricto'] = int(len(estricto))
    entregas_kpi_global['cobertura'] = round(
        100.0 * len(estricto) / max(1, len(ent_df)), 1)
    if len(estricto):
        entregas_kpi_global['lead_total_mediana'] = round(float(estricto['lead_d'].median()), 2)
        entregas_kpi_global['credito_mediana']   = round(float(estricto['cred_d'].median()), 2)
        entregas_kpi_global['logistica_mediana'] = round(float(estricto['logi_d'].median()), 2)
        entregas_kpi_global['lead_total_prom']   = round(float(estricto['lead_d'].mean()), 2)
        entregas_kpi_global['credito_prom']      = round(float(estricto['cred_d'].mean()), 2)
        entregas_kpi_global['logistica_prom']    = round(float(estricto['logi_d'].mean()), 2)

    # Por clasificacion
    for clas in ['AAA', 'A', 'B', 'C', 'D']:
        sub_tot = ent_df[ent_df['clasif'] == clas]
        sub_est = estricto[estricto['clasif'] == clas]
        entregas_por_clasif[clas] = {
            'n_total': int(len(sub_tot)),
            'n_estricto': int(len(sub_est)),
            'cobertura': round(100.0 * len(sub_est) / max(1, len(sub_tot)), 1),
            'lead_mediana':  round(float(sub_est['lead_d'].median()),  2) if len(sub_est) else 0,
            'credito_mediana': round(float(sub_est['cred_d'].median()), 2) if len(sub_est) else 0,
            'logistica_mediana': round(float(sub_est['logi_d'].median()), 2) if len(sub_est) else 0,
            'lead_prom':     round(float(sub_est['lead_d'].mean()),    2) if len(sub_est) else 0,
        }

    # Retira agente
    ra_n = int(ent_df['retira_agente'].sum())
    ra_total = int(len(ent_df))
    entregas_retira_agente['n_retira'] = ra_n
    entregas_retira_agente['n_total'] = ra_total
    entregas_retira_agente['pct'] = round(100.0 * ra_n / max(1, ra_total), 2)
    # Top 10 agentes con mas retiros
    top_ag = (ent_df[ent_df['retira_agente']]
              .groupby('nombreAgente').size()
              .sort_values(ascending=False).head(10))
    entregas_retira_agente['top_agentes'] = [
        {'agente': str(k), 'n': int(v)} for k, v in top_ag.items()
    ]
    # Evolucion mensual de retira agente
    evol_ra = {}
    for mes, grp in ent_df.groupby('mes'):
        if not mes or pd.isna(mes):
            continue
        total = len(grp)
        retira = int(grp['retira_agente'].sum())
        evol_ra[str(mes)] = {
            'total': total,
            'retira': retira,
            'pct': round(100.0 * retira / max(1, total), 2),
        }
    entregas_retira_agente['evol_mensual'] = evol_ra

    # Evolucion mensual de tiempos (medianas)
    for mes, grp in estricto.groupby('mes'):
        if not mes or pd.isna(mes):
            continue
        entregas_evol_mensual[str(mes)] = {
            'n': int(len(grp)),
            'lead_mediana':      round(float(grp['lead_d'].median()), 2),
            'credito_mediana':   round(float(grp['cred_d'].median()), 2),
            'logistica_mediana': round(float(grp['logi_d'].median()), 2),
        }

    print(f"  entregas_kpi_global: n_estricto={entregas_kpi_global['n_estricto']}/"
          f"{entregas_kpi_global['n_total']} cob={entregas_kpi_global['cobertura']}%")
    print(f"  retira_agente: {ra_n}/{ra_total} = {entregas_retira_agente['pct']}%")
else:
    print(f"  AVISO: no existe {entregas_csv} (sin datos de entregas)")

# ══ BUILD DB OBJECT ══
print("Construyendo DB...")

DB = {
    'ventas_mes': ventas_mes,
    'ventas_seg_mes': ventas_seg_mes,
    'ventas_seg_mes_all': ventas_seg_mes_all,
    'ventas_vend_mes': ventas_vend_mes,
    'ventas_vend_seg_mes': ventas_vend_seg_mes,
    'ventas_tg_mes': ventas_tg_mes,
    'ventas_seg_tg_mes': ventas_seg_tg_mes,
    'ppto_seg_tg_mes': ppto_seg_tg_mes,
    'ppto_vend_mes': ppto_vend_mes,
    'ppto_vend_grupo_mes': ppto_vend_grupo_mes,
    'ventas_vend_tg_mes': ventas_vend_tg_mes,
    'ventas_vend_cli': ventas_vend_cli,
    'ventas_grupo_mes': ventas_grupo_mes,
    'ventas_resp_mes': ventas_resp_mes,
    'ppto_grupo_mes': ppto_grupo_mes,
    'ppto_resp_mes': ppto_resp_mes,
    'cotizaciones': cotizaciones,
    'opps_activas': opps_activas,
    'opps_cerradas': opps_cerradas,
    'inv_grupo': inv_grupo,
    'inv_resp': inv_resp,
    'ppto_fin_mes': ppto_fin_mes,
    'ppto_fin_seg_mes': ppto_fin_seg_mes,
    'ppto_fin_grupo_mes': ppto_fin_grupo_mes,
    'ppto_fin_resp_mes': ppto_fin_resp_mes,
    'actividades': actividades,
    'cartera': cartera,
    'visitas': visitas,
    'clientes_cartera': clientes_cartera,
    'vendedor_supervisor': vend_sup_map,
    'vend_tipo': vend_tipo_map,
    'display': vend_display,
    'grupo_tipo': {g: info['tipo'] for g, info in grupo_resp.items()},
    'ytd_info': ytd_info,
    'segmentos': SEGS,
    'tipos_grupo': ['CORE', 'ESTRATÉGICOS', 'MRO', 'OTROS'],
    'sup_comercial': sup_comercial,
    'sup_jp': sup_jp,
    'rotacion_grupo': rotacion_por_grupo,
    'rotacion_total': rotacion_total,
    'entregas_kpi_global': entregas_kpi_global,
    'entregas_por_clasif': entregas_por_clasif,
    'entregas_retira_agente': entregas_retira_agente,
    'entregas_evol_mensual': entregas_evol_mensual,
    'last_update': datetime.now().strftime('%d/%m/%Y %H:%M'),
}

print("Escribiendo archivo...")
with open(OUTPUT, 'w', encoding='utf-8') as f:
    f.write('var DB = ')
    json.dump(DB, f, ensure_ascii=False, separators=(',', ':'))
    f.write(';\n')

# Stats
import os
size_mb = os.path.getsize(OUTPUT) / (1024*1024)
print(f"\nArchivo generado: {OUTPUT}")
print(f"Tamaño: {size_mb:.1f} MB")
print(f"\nEstadísticas:")
print(f"  ventas_mes: {len(ventas_mes)} meses")
print(f"  ventas_seg_mes: {len(ventas_seg_mes)} segmentos")
print(f"  ventas_vend_mes: {len(ventas_vend_mes)} vendedores")
print(f"  ventas_vend_seg_mes: {len(ventas_vend_seg_mes)} vendedores (NEW)")
print(f"  cotizaciones: {len(cotizaciones)}")
print(f"  opps_activas: {len(opps_activas)}")
print(f"  opps_cerradas: {len(opps_cerradas)}")
print(f"  actividades: {len(actividades)}")
print(f"  cartera: {len(cartera)} (con segmento)")
print(f"  visitas: {len(visitas)}")
print(f"  inv_grupo: {len(inv_grupo)}")
print(f"  vendedores: {len(vend_sup_map)}")

print("\nDone!")
