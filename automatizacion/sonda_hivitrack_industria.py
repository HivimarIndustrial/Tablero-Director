"""
Sonda Hivitrack para entregas Industria desde 2026-01-01.

Objetivos:
  1) Cuantas entregas trae lineaNegocio=1100 desde Ene 2026
  2) Validar cobertura de fechas (cuantas tienen Pedido+Liberacion+Entregado)
  3) Validar cruce con los 2,985 clientes Industria (por codigoDestinatario)
  4) Validar valores de categorizacionCliente (Hivitrack) vs Clasificacion dim_clientes
  5) Estimar tiempo de descarga + tamano CSV
  6) Inventariar 'Retira Agente' (rutaSAP / transporteSAP / nombres de agentes)
"""
import os
import sys
import time
from datetime import datetime, date
from collections import Counter

# Cargar .env de Voz del Cliente
ENV = (r'C:\Users\consultorindustrial\OneDrive - Hivimar\Documentos\HIVIMAR'
       r'\Proyectos\Voz del cliente y mejora contínua\Tablero Voz del Cliente\.env')
for line in open(ENV, encoding='utf-8'):
    line = line.strip()
    if '=' in line and not line.startswith('#'):
        k, v = line.split('=', 1)
        if k.startswith('HIVITRACK'):
            os.environ[k] = v.strip()

# Path al cliente Hivitrack
KIT = (r'C:\Users\consultorindustrial\OneDrive - Hivimar\Documentos\HIVIMAR'
       r'\Proyectos\Hivitrack Starter Kit')
sys.path.insert(0, KIT)
import hivitrack_client as h

# Lista Industria
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from industria_clientes import obtener_clientes_industria

import pandas as pd


def parse_dt(s):
    if not s: return None
    try:
        return datetime.fromisoformat(str(s).replace('Z', ''))
    except Exception:
        return None


def main():
    h.login()
    print("Login Hivitrack OK\n")

    # Lista Industria (2,985 codigos)
    codigos_ind = set(obtener_clientes_industria(verbose=False))
    print(f"Lista Industria: {len(codigos_ind)} CODIGO_CLIENTE SAP\n")

    # dias desde 2026-01-01 hasta hoy
    hoy = date.today()
    inicio = date(2026, 1, 1)
    dias_atras = (hoy - inicio).days + 5  # margen
    print(f"Ventana: {inicio} -> {hoy} ({dias_atras} dias atras)\n")

    # Pagina pequena para sondear; rows=1000 acelera mucho
    rows_per_page = 500
    t0 = time.time()
    print(f"[1] Pidiendo pagina 1 (rows={rows_per_page}, dias={dias_atras})...")
    r = h.buscar_entregas(dias=dias_atras, page=1, rows=rows_per_page)
    total_filas = r.get('total_filas')
    total_paginas = r.get('total_paginas')
    print(f"  Total global Hivimar (todas LN): {total_filas:,}")
    print(f"  Total paginas (rows={rows_per_page}): {total_paginas}")
    print(f"  Pagina 1 ent: {len(r.get('entregas',[]))} en {time.time()-t0:.1f}s\n")

    # Si total > 50k, advertir
    if total_filas > 30000:
        print(f"  AVISO: total {total_filas:,} grande. Descarga toda puede tomar varios min.")

    # Descargar TODAS las paginas filtrando solo lineaNegocio=1100 in-memory
    all_ind = []
    pag = 1
    t1 = time.time()
    while True:
        if pag > 1:
            r = h.buscar_entregas(dias=dias_atras, page=pag, rows=rows_per_page)
        ents = r.get('entregas', [])
        if not ents: break
        for e in ents:
            if e.get('lineaNegocio') == '1100':
                all_ind.append(e)
        if pag >= total_paginas: break
        pag += 1
        if pag % 10 == 0:
            print(f"  pag {pag}/{total_paginas}  Industria acum={len(all_ind):,}  "
                  f"({time.time()-t1:.0f}s)")
    print(f"\n  Industria LN=1100 desde {inicio}: {len(all_ind):,} entregas "
          f"(descarga {time.time()-t1:.0f}s)\n")

    if not all_ind:
        print("Nada que analizar.")
        return

    df = pd.DataFrame(all_ind)
    print(f"DataFrame: {len(df)} filas x {len(df.columns)} cols\n")

    # Cruce contra los 2,985 clientes Industria via codigoDestinatario
    # codigoDestinatario viene como '0000462642' (10 digits con leading zeros);
    # dim_clientes.cliente_id_std viene como '465656' sin ceros.
    df['cod_norm'] = df['codigoDestinatario'].astype(str).str.lstrip('0').replace('', '0')
    in_lista = df['cod_norm'].isin(codigos_ind)
    print(f"=== Cruce contra dim_clientes Industria (2,985) ===")
    print(f"  En la lista:  {in_lista.sum():,}")
    print(f"  Fuera:        {(~in_lista).sum():,}")
    print(f"  % cobertura cruce: {100*in_lista.sum()/len(df):.1f}%")
    print()

    # Cobertura de fechas
    df['t_creacion'] = df['fechaHoraCreacionPedido'].apply(parse_dt)
    df['t_liberac']  = df['fechaHoraLiberacionPedido'].apply(parse_dt)
    df['t_entrega']  = df['fechaEntregado'].apply(parse_dt)
    print(f"=== Cobertura de fechas (sobre {len(df):,} Industria) ===")
    print(f"  Con fechaCreacionPedido:   {df['t_creacion'].notna().sum():,}  "
          f"({100*df['t_creacion'].notna().sum()/len(df):.1f}%)")
    print(f"  Con fechaLiberacionPedido: {df['t_liberac'].notna().sum():,}  "
          f"({100*df['t_liberac'].notna().sum()/len(df):.1f}%)")
    print(f"  Con fechaEntregado:        {df['t_entrega'].notna().sum():,}  "
          f"({100*df['t_entrega'].notna().sum()/len(df):.1f}%)")
    estricto = df[df['t_creacion'].notna() & df['t_liberac'].notna() & df['t_entrega'].notna()]
    print(f"  Estricto (3 fechas):       {len(estricto):,}  "
          f"({100*len(estricto)/len(df):.1f}%)  <-- base para KPIs")
    print()

    # Tiempos
    if len(estricto):
        e = estricto.copy()
        e['lead_total_d'] = (e['t_entrega'] - e['t_creacion']).dt.total_seconds() / 86400
        e['credito_d']    = (e['t_liberac']  - e['t_creacion']).dt.total_seconds() / 86400
        e['logistica_d']  = (e['t_entrega'] - e['t_liberac']).dt.total_seconds() / 86400
        # Solo positivos (datos sanos)
        e = e[(e['lead_total_d']>=0) & (e['credito_d']>=0) & (e['logistica_d']>=0)]
        print(f"=== Tiempos (mediana / promedio / max) -- Industria estricto ({len(e):,}) ===")
        for col, lab in [('lead_total_d','Lead Time Total'),
                          ('credito_d',   'Credito'),
                          ('logistica_d', 'Logistica')]:
            print(f"  {lab:18s}: mediana={e[col].median():>5.2f} d   "
                  f"prom={e[col].mean():>5.2f} d   max={e[col].max():>6.1f} d")
        print()

    # Categorizacion Hivitrack vs dim_clientes
    print("=== categorizacionCliente Hivitrack en universo Industria ===")
    print(df['categorizacionCliente'].value_counts(dropna=False).head(10))
    print()

    # Validar coincidencia: para los que estan en la lista Industria,
    # comparar categorizacion Hivitrack vs Clasificacion dim_clientes
    xls = (r'C:\Users\consultorindustrial\OneDrive - Hivimar\Documentos'
           r'\HIVIMAR\Proyectos\SGI 2.0\Indicadores Tablero'
           r'\Base Tablero Industria.xlsx')
    dc = pd.read_excel(xls, sheet_name='dim_clientes',
                       dtype={'cliente_id_std': str})
    dc_lookup = dict(zip(
        dc['cliente_id_std'].astype(str).str.strip(),
        dc['Clasificación']
    ))
    sub = df[in_lista].copy()
    sub['clas_dim'] = sub['cod_norm'].map(dc_lookup)
    match = (sub['categorizacionCliente'] == sub['clas_dim']).sum()
    diff = ((sub['categorizacionCliente'].notna() & sub['clas_dim'].notna()) &
            (sub['categorizacionCliente'] != sub['clas_dim'])).sum()
    print(f"=== Match categorizacion Hivitrack vs dim_clientes ===")
    print(f"  Iguales: {match:,}  Diferentes: {diff:,}")
    print(f"  % match: {100*match/max(1,len(sub)):.1f}%")
    print()

    # Retira agente
    sub['retira_agente'] = (
        sub['rutaSAP'].astype(str).str.contains('RETIROS AGENTE', case=False, na=False) |
        (sub['transporteSAP'].astype(str).str.upper() == 'RETIRA AGENTE')
    )
    print(f"=== Retira Agente (Industria 2026) ===")
    print(f"  Total:        {sub['retira_agente'].sum():,}  "
          f"({100*sub['retira_agente'].sum()/max(1,len(sub)):.1f}%)")
    print(f"  Top 10 agentes con mas retiros:")
    print(sub[sub['retira_agente']]['nombreAgente'].value_counts().head(10).to_string())
    print()

    # Tamano estimado CSV
    print(f"=== Estimacion archivo CSV (Industria estricto) ===")
    print(f"  ~{len(estricto)*250/1024/1024:.1f} MB con 9 columnas relevantes")


if __name__ == '__main__':
    main()
