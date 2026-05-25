"""
generar_tablero_director.py

Genera una version reducida del tablero, llamada "Tablero Director Industria",
que solo contiene la pestania GERENCIA (con sus secciones: Ventas, Presupuesto,
Cartera, CRM, Visitas, Inventario, Proyeccion).

Toma como base 'Tablero_Director_FUENTE.html' y elimina:
  - Botones de pestania: Supervisores, Vendedores, CRM, Perspectiva Producto.
  - Vistas <div class="view" id="vs|vv|vp|vo">.
  - Grupo de filtros 'fgrpProd' (Jefe Producto, Grupo Articulo) que solo
    aplica a la vista Producto.
  - Cambia el title y header a "Tablero Director Industria".

NO toca el JS ni la variable DB, asi se actualiza con los mismos datos del
pipeline diario. Solo las funciones de render de las vistas eliminadas
quedan definidas en el JS pero nunca se llaman (los botones ya no existen).

Uso:
  python automatizacion/generar_tablero_director.py
"""
import os
import re
import json

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
SRC = os.path.join(PROJECT_DIR, 'Tablero_Director_FUENTE.html')
DST = os.path.join(PROJECT_DIR, 'Tablero_Director_Industria.html')


# Campos a CONSERVAR en cada array detalle (drop el resto).
# Identificados leyendo las funciones fCots/fOpps/fActs/fVis/fCart y los
# render rGV, rGC, rGCRM, rGVIS, rGINV, rGP del HTML del Director.
_KEEP_FIELDS = {
    'cotizaciones': {'f', 'co', 'su', 'gr', 'resp', 't', 'e', 'tg'},
    'opps_activas': {'etapa', 'co', 'su', 'gr', 'resp', 'ingreso', 'tg',
                     'seg', 'vencida_act', 'anio_c', 'mes_c'},
    'opps_cerradas': {'etapa', 'co', 'su', 'gr', 'resp', 'ingreso', 'tg',
                      'seg', 'vencida_act', 'anio_c', 'mes_c',
                      'anio', 'mes', 'ganada', 'perdida'},
    'actividades': {'asignado', 'supervisor', 'vencida'},
    'visitas': {'responsable', 'supervisor', 'anio', 'mes', 'estado',
                'hora_entrada'},
    # cartera: conserva todos sus campos (todos se usan)
    # clientes_cartera: conserva todos sus campos
}

# Top N de clientes por vendedor en ventas_vend_cli (Gerencia agrega y
# toma top 10 global; con 30 por vendedor sobra holgadamente).
_VENDCLI_TOPN = 30

# Claves del DB que SOLO se usan en pestañas Supervisores/Vendedores/
# Producto (eliminadas del Director). Drop completo, ahorra ~0.4 MB.
_DROP_KEYS = ('ppto_vend_mes', 'ppto_vend_grupo_mes')

# Campos string a "internar" (deduplicar) por array. Cada string repetido
# (ej: 'CARLOS PEREZ' aparece 800 veces como vendedor) se reemplaza por un
# indice entero hacia un diccionario lookup. Un hidratador JS reconstruye
# los strings al cargar la pagina, antes de que corra cualquier render.
_INTERN_FIELDS = {
    'cotizaciones':     ('co', 'su', 'gr', 'resp', 'e', 'tg', 'f'),
    'opps_activas':     ('etapa', 'co', 'su', 'gr', 'resp', 'tg', 'seg'),
    'opps_cerradas':    ('etapa', 'co', 'su', 'gr', 'resp', 'tg', 'seg'),
    'actividades':      ('asignado', 'supervisor'),
    'visitas':          ('responsable', 'supervisor', 'estado'),
    'cartera':          ('agente', 'supervisor', 'cliente', 'vencimiento',
                         'segmento'),
    'clientes_cartera': ('vend', 'sup', 'cli', 'ultima_visita'),
    'entregas_rows':    ('sup', 'vend', 'clas', 'ag', 'mes'),
}

# Snippet JS que se inyecta INMEDIATAMENTE despues de "var DB={...};"
# (dentro del mismo <script>) para que corra ANTES de cualquier render.
# Reconstruye los strings desde DB._dict.
# IMPORTANTE: NO lleva tags <script> porque va inline en el JS existente.
_HYDRATOR_JS = (
    '(function(){'
    'var d=DB._dict;if(!d)return;'
    'for(var name in d){'
    'var arr=DB[name];if(!Array.isArray(arr))continue;'
    'var fields=d[name];'
    'for(var i=0;i<arr.length;i++){'
    'var row=arr[i];if(!row)continue;'
    'for(var f in fields){'
    'var v=row[f];'
    'if(typeof v==="number")row[f]=fields[f][v];'
    '}}}delete DB._dict;})();'
)


def _slim_db_inline(text):
    """Encuentra la línea 'var DB = {...};', recorta campos no usados por
    Gerencia y reescribe el blob. Reduce el HTML de ~8 MB a ~1-3 MB."""
    m = re.search(r'var DB\s*=\s*(\{.*?\});\s*$', text,
                  flags=re.MULTILINE | re.DOTALL)
    if not m:
        # Intento alternativo: línea sola con var DB
        m = re.search(r'var DB\s*=\s*(\{.*\})\s*;', text, flags=re.DOTALL)
    if not m:
        print('[slim] no se encontro "var DB = {...}"; no se recorta')
        return text

    blob_start = m.start(1)
    blob_end = m.end(1)
    raw = m.group(1)
    size_in = len(raw)

    try:
        db = json.loads(raw)
    except json.JSONDecodeError as e:
        print(f'[slim] error parseando DB JSON: {e}; no se recorta')
        return text

    # Drop claves completas no usadas por Gerencia
    dropped_keys = 0
    for k in _DROP_KEYS:
        if k in db:
            del db[k]
            dropped_keys += 1

    dropped_fields = 0
    # Recortar arrays detalle
    for key, keep in _KEEP_FIELDS.items():
        arr = db.get(key)
        if not isinstance(arr, list):
            continue
        for i, row in enumerate(arr):
            if not isinstance(row, dict):
                continue
            removed = [k for k in list(row.keys()) if k not in keep]
            for k in removed:
                del row[k]
            dropped_fields += len(removed)

    # Recortar ventas_vend_cli: top N por vendedor
    vvc = db.get('ventas_vend_cli')
    if isinstance(vvc, dict):
        for vend, clis in list(vvc.items()):
            if not isinstance(clis, dict):
                continue
            if len(clis) <= _VENDCLI_TOPN:
                continue
            top = sorted(
                clis.items(),
                key=lambda kv: (kv[1].get('venta', 0)
                                if isinstance(kv[1], dict) else 0),
                reverse=True,
            )[:_VENDCLI_TOPN]
            vvc[vend] = dict(top)

    # String interning: deduplicar strings repetidos en arrays detalle
    # reemplazandolos por indices a un diccionario lookup en DB._dict.
    interned_strings = 0
    dict_block = {}
    for arr_name, fields in _INTERN_FIELDS.items():
        arr = db.get(arr_name)
        if not isinstance(arr, list) or not arr:
            continue
        # Construir lookup por campo
        per_field = {}  # field -> {string: index}
        for fname in fields:
            per_field[fname] = {}
        for row in arr:
            if not isinstance(row, dict):
                continue
            for fname in fields:
                v = row.get(fname)
                if not isinstance(v, str) or not v:
                    # Strings vacios se dejan tal cual: preservan el
                    # comportamiento JS de "if(!c.f)return true".
                    continue
                lookup = per_field[fname]
                if v not in lookup:
                    # Indices 1-based: el 0 queda libre para que nunca
                    # colisione con un check JS "!val".
                    lookup[v] = len(lookup) + 1
                row[fname] = lookup[v]
                interned_strings += 1
        # Guardar lookup invertido (idx -> string) en _dict.
        # Posicion 0 = null (placeholder); a partir de 1, los strings.
        dict_block[arr_name] = {
            f: [None] + list(per_field[f].keys())
            for f in fields if per_field[f]
        }
    if dict_block:
        db['_dict'] = dict_block

    # Reserializar compacto (sin espacios)
    new_blob = json.dumps(db, ensure_ascii=False, separators=(',', ':'))
    size_out = len(new_blob)
    print(f'[slim] DB: {size_in/1024/1024:.2f} MB -> '
          f'{size_out/1024/1024:.2f} MB '
          f'(-{(1-size_out/size_in)*100:.0f}%, '
          f'{dropped_keys} claves + {dropped_fields:,} campos + '
          f'{interned_strings:,} strings internados)')

    # Inyectar el hidratador JS INLINE, inmediatamente despues del ";"
    # que cierra "var DB={...};". Asi corre antes de que se ejecute
    # cualquier funcion de render (que vendran mas abajo en el mismo
    # <script>). El blob match incluyo solo los {...} sin el ";", asi
    # que insertamos despues del ";".
    suffix_after_blob = ';' + _HYDRATOR_JS if dict_block else ';'
    # text[blob_end] es el ";" que cierra. Lo reemplazamos por
    # ";<hidratador>".
    new_text = (text[:blob_start] + new_blob + suffix_after_blob
                + text[blob_end + 1:])
    if dict_block:
        print('[slim] hidratador JS inyectado inline tras "var DB={...};"')

    return new_text


def main():
    with open(SRC, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    print(f"[in] {SRC}  ({len(lines)} lineas)")

    # 1) Cambiar <title>
    for i, ln in enumerate(lines):
        if '<title>' in ln:
            lines[i] = '<title>Hivimar — Tablero Director Industria</title>\n'
            break

    # 2) Cambiar header visible
    for i, ln in enumerate(lines):
        if 'class="hdr-t"' in ln:
            lines[i] = ('  <div class="hdr-t">Tablero Director Industria'
                        '<div class="hdr-sub">Vista ejecutiva · Hivimar '
                        'Industrial</div></div>\n')
            break

    # 3) Eliminar lineas de botones de pestania que NO son Gerencia
    #    (data-view="s|v|o|p")
    out = []
    skip_lines = 0
    for ln in lines:
        if skip_lines > 0:
            skip_lines -= 1
            continue

        # Botones de tabs no-gerencia
        if re.search(r'data-view="[svop]"', ln):
            continue

        out.append(ln)
    lines = out
    print(f"[step] removidos botones de tabs s/v/o/p")

    # 4) Eliminar grupo de filtros Producto (div con id=fgrpProd hasta su </div>)
    out = []
    inside_fgrp_prod = False
    depth = 0
    for ln in lines:
        if not inside_fgrp_prod:
            if 'id="fgrpProd"' in ln:
                inside_fgrp_prod = True
                depth = ln.count('<div') - ln.count('</div>')
                continue  # skip this opening line
            out.append(ln)
        else:
            depth += ln.count('<div') - ln.count('</div>')
            if depth <= 0:
                inside_fgrp_prod = False
            # skip every line inside the group (including the closing </div>)
    lines = out
    print(f"[step] removido fgrpProd")

    # 5) Eliminar las views completas: vs, vv, vp, vo (y su comentario header)
    out = []
    skip_view_id = None
    view_depth = 0
    pending_remove_comment = False
    for ln in lines:
        # Detectar comentarios de seccion a eliminar -> marcar para drop
        if (re.match(r'\s*<!--\s*══\s*(SUPERVISORES|VENDEDOR|PRODUCTO|CRM)',
                     ln, re.IGNORECASE)):
            pending_remove_comment = True
            continue

        if skip_view_id is None:
            m = re.search(r'<div class="view[^"]*" id="(v[svpo])"', ln)
            if m:
                skip_view_id = m.group(1)
                view_depth = ln.count('<div') - ln.count('</div>')
                pending_remove_comment = False
                continue
            if pending_remove_comment:
                # comentario huerfano sin view inmediato (no deberia pasar)
                pending_remove_comment = False
            out.append(ln)
        else:
            view_depth += ln.count('<div') - ln.count('</div>')
            if view_depth <= 0:
                skip_view_id = None
    lines = out
    print(f"[step] removidas views vs/vv/vp/vo")

    # 6) Asegurar que el boton de Gerencia tenga clase "on" (debe ya tenerla)
    for i, ln in enumerate(lines):
        if 'data-view="g"' in ln and 'class="nb on"' not in ln:
            lines[i] = ln.replace('class="nb"', 'class="nb on"')

    # 7) Limpiar el comentario de filtros que mencione "Supervisores, Vendedores"
    for i, ln in enumerate(lines):
        if 'Filtros COMERCIAL' in ln:
            lines[i] = '  <!-- Filtros para Director Industria (Gerencia) -->\n'
            break

    # 8) Hacer null-safe las referencias a getElementById('fgrpProd') porque
    #    ese elemento ya no existe (lo eliminamos). Sin esto, el handler
    #    que ajusta visibilidad de filtros tira null.classList.toggle.
    for i, ln in enumerate(lines):
        if "getElementById('fgrpProd')" in ln and '?.' not in ln:
            lines[i] = ln.replace(
                "getElementById('fgrpProd').classList",
                "getElementById('fgrpProd')?.classList")

    # 9) Parchar el listener de cambios para que ignore IDs faltantes (fGR
    #    pertenece al grupo Producto eliminado).
    for i, ln in enumerate(lines):
        if ("['fM','fV','fSeg','fTG','fGR']" in ln
                and 'addEventListener' in ln):
            lines[i] = (
                "['fM','fV','fSeg','fTG','fGR'].forEach(id=>{"
                "var e=document.getElementById(id);"
                "if(e)e.addEventListener('change',render);});\n"
            )

    # 10) Hacer condicional la inicializacion de fJP y fGR (filtros de
    #     Producto eliminados, sus selects no existen y appendChild tira
    #     null).
    text = ''.join(lines)
    text = text.replace(
        "  // Jefes de Producto\n"
        "  const fJP=document.getElementById('fJP');\n"
        "  const jefes=[...new Set(Object.values(DB.inv_grupo).map(d=>d.resp))].filter(j=>j&&j!=='REVISAR'&&j!=='COMERCIO').sort();\n"
        "  jefes.forEach(j=>{const o=document.createElement('option');o.value=j;o.textContent=j;fJP.appendChild(o);});\n",
        "  // Jefes de Producto (omitido: filtro Producto no existe en Director)\n"
        "  const fJP=document.getElementById('fJP');\n"
        "  if(fJP){const jefes=[...new Set(Object.values(DB.inv_grupo).map(d=>d.resp))].filter(j=>j&&j!=='REVISAR'&&j!=='COMERCIO').sort();\n"
        "    jefes.forEach(j=>{const o=document.createElement('option');o.value=j;o.textContent=j;fJP.appendChild(o);});}\n"
    )
    text = text.replace(
        "  // Grupos de artículo\n"
        "  const fGR=document.getElementById('fGR');\n"
        "  Object.keys(DB.inv_grupo).sort().forEach(g=>{const o=document.createElement('option');o.value=g;o.textContent=g;fGR.appendChild(o);});\n",
        "  // Grupos de articulo (omitido: filtro Producto no existe en Director)\n"
        "  const fGR=document.getElementById('fGR');\n"
        "  if(fGR){Object.keys(DB.inv_grupo).sort().forEach(g=>{const o=document.createElement('option');o.value=g;o.textContent=g;fGR.appendChild(o);});}\n"
    )

    # 11) Hacer null-safe el .value en el getter de filtros activos
    text = text.replace(
        "    jp:document.getElementById('fJP').value,\n"
        "    gr:document.getElementById('fGR').value\n",
        "    jp:document.getElementById('fJP')?.value||'',\n"
        "    gr:document.getElementById('fGR')?.value||''\n"
    )

    # 12) Hacer condicional el listener de fJP -> fGR
    text = text.replace(
        "document.getElementById('fJP').addEventListener('change',function(){\n",
        "document.getElementById('fJP')?.addEventListener('change',function(){\n"
    )
    text = text.replace(
        "  const fGR=document.getElementById('fGR');\n"
        "  fGR.innerHTML='<option value=\"\">Todos</option>';\n",
        "  const fGR=document.getElementById('fGR');\n"
        "  if(!fGR){return;}\n"
        "  fGR.innerHTML='<option value=\"\">Todos</option>';\n"
    )

    # 13) Slim del blob "var DB = {...};" inline.
    #     Gerencia solo agrega KPIs y muestra top clientes; no necesita
    #     campos como Oportunidad, Cliente en cotizaciones, etc.
    text = _slim_db_inline(text)

    lines = text.splitlines(keepends=True)

    with open(DST, 'w', encoding='utf-8') as f:
        f.writelines(lines)
    size_mb = os.path.getsize(DST) / (1024*1024)
    print(f"[out] {DST}  ({len(lines)} lineas, {size_mb:.1f} MB)")
    print("\nOK. Abre el archivo en el navegador para verificar.")


if __name__ == '__main__':
    main()
