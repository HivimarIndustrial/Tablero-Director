"""
cifrar_y_publicar.py

Toma `Tablero_Director_Industria.html` (HTML con datos en claro embebidos
como `var DB = {...}`), cifra los datos con AES-256-GCM y key-wrapping
por usuario (PBKDF2-SHA256), inyecta una pantalla de login y escribe el
resultado en `docs/index.html` para servir desde GitHub Pages.

Esquema de cifrado (mismo del Tablero Supervisores):
  master_key = random(32 bytes)
  ciphertext = AES-GCM(master_key, iv, JSON.stringify(DB))
  Por cada usuario u:
    salt_u  = random(16 bytes)
    key_u   = PBKDF2(password_u, salt_u, 200_000, SHA256, 32 bytes)
    wrap_u  = AES-GCM(key_u, iv_u, master_key)

Las contraseñas NUNCA salen de claves.json (gitignored). En el HTML
publicado solo viajan los `wrap` cifrados; sin la clave de cada usuario,
desencriptar es computacionalmente inviable.

Uso: python cifrar_y_publicar.py
"""
import os
import re
import json
import base64
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
HTML_IN    = os.path.join(SCRIPT_DIR, 'Tablero_Director_Industria.html')
DOCS_DIR   = os.path.join(SCRIPT_DIR, 'docs')
HTML_OUT   = os.path.join(DOCS_DIR, 'index.html')
CLAVES     = os.path.join(SCRIPT_DIR, 'claves.json')

PBKDF2_ITER = 200_000


# ---------------------------------------------------------------------------
# Cifrado
# ---------------------------------------------------------------------------
def encrypt_data_for_users(plaintext: str) -> dict:
    """Cifra plaintext y devuelve payload listo para incrustar en HTML."""
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives.hashes import SHA256
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM

    if not os.path.exists(CLAVES):
        sys.exit(f"FATAL: no existe {CLAVES}. Crear con la lista de usuarios y claves.")

    with open(CLAVES, encoding='utf-8') as f:
        cfg = json.load(f)
    usuarios = cfg.get('usuarios', [])
    if not usuarios:
        sys.exit("FATAL: claves.json no tiene usuarios.")

    def b64(b): return base64.b64encode(b).decode('ascii')

    master_key = os.urandom(32)
    iv_data    = os.urandom(12)
    ct = AESGCM(master_key).encrypt(iv_data, plaintext.encode('utf-8'), None)

    wrappers = []
    for u in usuarios:
        salt = os.urandom(16)
        kdf  = PBKDF2HMAC(algorithm=SHA256(), length=32, salt=salt, iterations=PBKDF2_ITER)
        key  = kdf.derive(u['clave'].encode('utf-8'))
        iv_u = os.urandom(12)
        wrap = AESGCM(key).encrypt(iv_u, master_key, None)
        wrappers.append({
            'id':     u['id'],
            'nombre': u['nombre'],
            'salt':   b64(salt),
            'iv':     b64(iv_u),
            'wrap':   b64(wrap),
        })

    return {
        'v':      1,
        'cipher': 'AES-GCM',
        'kdf':    'PBKDF2-SHA256',
        'iter':   PBKDF2_ITER,
        'iv':     b64(iv_data),
        'ct':     b64(ct),
        'users':  wrappers,
    }


# ---------------------------------------------------------------------------
# Login overlay (HTML + JS) — se monta sobre el #loader original
# ---------------------------------------------------------------------------
LOGIN_OVERLAY_CSS = """
<style id="login-overlay-css">
#login-overlay{position:fixed;inset:0;z-index:99999;display:flex;align-items:center;justify-content:center;
  background:linear-gradient(135deg,#1B2F6E 0%,#0e1a47 100%);font-family:system-ui,-apple-system,Segoe UI,Roboto,sans-serif;}
#login-overlay .card{background:#fff;border-radius:14px;box-shadow:0 20px 60px rgba(0,0,0,.4);padding:34px 30px;width:340px;max-width:92vw;}
#login-overlay .logo-wrap{text-align:center;margin-bottom:14px;}
#login-overlay .logo-wrap img{height:54px;max-width:100%;object-fit:contain;}
#login-overlay h1{font-size:18px;margin:0 0 4px 0;color:#1B2F6E;text-align:center;}
#login-overlay .sub{font-size:12px;color:#64748b;text-align:center;margin-bottom:18px;}
#login-overlay label{display:block;font-size:12px;color:#334155;margin:10px 0 4px 0;font-weight:500;}
#login-overlay select,#login-overlay input[type=password]{width:100%;border:1px solid #cbd5e1;border-radius:8px;padding:8px 10px;font-size:13px;outline:none;}
#login-overlay select:focus,#login-overlay input:focus{border-color:#2563eb;box-shadow:0 0 0 3px rgba(37,99,235,.18);}
#login-overlay .row{display:flex;align-items:center;gap:6px;margin-top:10px;font-size:12px;color:#475569;}
#login-overlay button.primary{width:100%;margin-top:14px;background:#1B2F6E;color:#fff;border:0;padding:9px;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;}
#login-overlay button.primary:disabled{opacity:.55;cursor:not-allowed;}
#login-overlay .err{display:none;background:#fef2f2;border:1px solid #fecaca;color:#b91c1c;font-size:12px;padding:7px 9px;border-radius:7px;margin-top:10px;}
#login-overlay .foot{font-size:10.5px;color:#94a3b8;text-align:center;margin-top:14px;}
</style>
"""

LOGIN_OVERLAY_HTML = """
<div id="login-overlay">
  <div class="card">
    <div class="logo-wrap"><img id="login-logo-img" alt="Hivimar Industrial" src="__LOGIN_LOGO_SRC__"></div>
    <h1>Tablero Director Industria</h1>
    <div class="sub">Hivimar · Acceso autorizado</div>
    <form id="login-form" autocomplete="on">
      <label for="login-user">Usuario</label>
      <select id="login-user" name="username" autocomplete="username">
        <option value="">Selecciona tu usuario...</option>
      </select>
      <label for="login-pass">Clave</label>
      <input type="password" id="login-pass" name="password" autocomplete="current-password" placeholder="Tu clave">
      <div class="row"><input type="checkbox" id="login-remember"><label for="login-remember" style="margin:0">Recordarme en este dispositivo</label></div>
      <div class="err" id="login-err"></div>
      <button type="submit" class="primary" id="login-submit">Entrar</button>
    </form>
    <div class="foot">Si olvidaste tu clave, contacta al Consultor Industrial.</div>
  </div>
</div>
"""

LOGIN_JS = r"""
<script id="login-overlay-js">
(function(){
  function _b64(b64){var bin=atob(b64),a=new Uint8Array(bin.length);for(var i=0;i<bin.length;i++)a[i]=bin.charCodeAt(i);return a;}
  async function _deriveKey(pwd, salt, iter){
    var enc=new TextEncoder();
    var bk=await crypto.subtle.importKey('raw', enc.encode(pwd), {name:'PBKDF2'}, false, ['deriveKey']);
    return crypto.subtle.deriveKey({name:'PBKDF2', salt:salt, iterations:iter, hash:'SHA-256'}, bk, {name:'AES-GCM', length:256}, true, ['decrypt','encrypt']);
  }
  async function _decryptForUser(uid, pwd){
    var SD=window.SECURE_DATA;
    var u=SD.users.find(function(x){return x.id===uid;});
    if(!u) throw new Error('Usuario no encontrado');
    var dk = await _deriveKey(pwd, _b64(u.salt), SD.iter);
    var masterRaw;
    try { masterRaw = await crypto.subtle.decrypt({name:'AES-GCM', iv:_b64(u.iv)}, dk, _b64(u.wrap)); }
    catch(e){ throw new Error('Clave incorrecta'); }
    var mk = await crypto.subtle.importKey('raw', masterRaw, {name:'AES-GCM'}, true, ['decrypt']);
    var plain = await crypto.subtle.decrypt({name:'AES-GCM', iv:_b64(SD.iv)}, mk, _b64(SD.ct));
    var rawMK = new Uint8Array(masterRaw);
    var mkB64 = btoa(String.fromCharCode.apply(null, rawMK));
    return { json: new TextDecoder().decode(plain), mkB64: mkB64 };
  }
  function _populateUsers(){
    var sel=document.getElementById('login-user');
    if(!sel||!window.SECURE_DATA) return;
    window.SECURE_DATA.users.forEach(function(u){
      var o=document.createElement('option'); o.value=u.id; o.textContent=u.nombre; sel.appendChild(o);
    });
    var last=localStorage.getItem('td_last_user');
    if(last && window.SECURE_DATA.users.find(function(x){return x.id===last;})) sel.value=last;
  }
  function _showErr(m){ var e=document.getElementById('login-err'); e.textContent=m; e.style.display='block'; }
  function _hideErr(){ document.getElementById('login-err').style.display='none'; }
  function _hideOverlay(){
    var ov=document.getElementById('login-overlay'); if(ov) ov.parentNode.removeChild(ov);
    var css=document.getElementById('login-overlay-css'); if(css) css.parentNode.removeChild(css);
  }
  function _hydrateAndBoot(jsonStr){
    window.DB = JSON.parse(jsonStr);
    if (typeof window.__hydrate === 'function') window.__hydrate();
    if (typeof window.__boot === 'function')    window.__boot();
  }
  async function _doLogin(uid, pwd, remember){
    _hideErr();
    var btn=document.getElementById('login-submit');
    btn.disabled=true; btn.textContent='Verificando...';
    try{
      var r = await _decryptForUser(uid, pwd);
      sessionStorage.setItem('td_session', JSON.stringify({uid:uid, mk:r.mkB64}));
      if(remember) localStorage.setItem('td_remember', JSON.stringify({uid:uid, mk:r.mkB64}));
      else         localStorage.removeItem('td_remember');
      localStorage.setItem('td_last_user', uid);
      _hideOverlay();
      _hydrateAndBoot(r.json);
    } catch(e){
      _showErr(e.message || 'Error al iniciar sesion');
      btn.disabled=false; btn.textContent='Entrar';
    }
  }
  async function _tryAuto(){
    var raw = sessionStorage.getItem('td_session') || localStorage.getItem('td_remember');
    if(!raw) return false;
    try{
      var s = JSON.parse(raw);
      var SD = window.SECURE_DATA;
      var mk = await crypto.subtle.importKey('raw', _b64(s.mk), {name:'AES-GCM'}, false, ['decrypt']);
      var plain = await crypto.subtle.decrypt({name:'AES-GCM', iv:_b64(SD.iv)}, mk, _b64(SD.ct));
      _hideOverlay();
      _hydrateAndBoot(new TextDecoder().decode(plain));
      return true;
    } catch(e){
      sessionStorage.removeItem('td_session');
      localStorage.removeItem('td_remember');
      return false;
    }
  }
  function init(){
    _populateUsers();
    document.getElementById('login-form').addEventListener('submit', function(e){
      e.preventDefault();
      var uid=document.getElementById('login-user').value;
      var pwd=document.getElementById('login-pass').value;
      var rem=document.getElementById('login-remember').checked;
      if(!uid){ _showErr('Selecciona tu usuario'); return; }
      if(!pwd){ _showErr('Ingresa tu clave'); return; }
      _doLogin(uid, pwd, rem);
    });
    _tryAuto();
  }
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
</script>
"""


# ---------------------------------------------------------------------------
# Transformaciones del HTML
# ---------------------------------------------------------------------------
def _split_db_line(line: str):
    """Devuelve (json_db_str, hidratador_body_str) o (None, None) si no matchea."""
    PREFIX = 'var DB = '
    if not line.startswith(PREFIX):
        return None, None
    # Encontrar el JSON balanceando llaves (ignorando llaves dentro de strings).
    body = line[len(PREFIX):]
    if not body.startswith('{'):
        return None, None
    depth = 0
    in_str = False
    esc = False
    end = -1
    for i, ch in enumerate(body):
        if esc:
            esc = False
            continue
        if ch == '\\' and in_str:
            esc = True
            continue
        if ch == '"':
            in_str = not in_str
            continue
        if in_str:
            continue
        if ch == '{':
            depth += 1
        elif ch == '}':
            depth -= 1
            if depth == 0:
                end = i
                break
    if end < 0:
        return None, None
    json_db = body[:end + 1]
    rest = body[end + 1:].rstrip(';')  # tolerar `;` final opcional
    M_OPEN = ';(function(){'
    M_CLOSE = '})()'
    if not rest.startswith(M_OPEN) or not rest.endswith(M_CLOSE):
        return None, None
    hyd_body = rest[len(M_OPEN):-len(M_CLOSE)]
    return json_db, hyd_body


def transformar_html(html: str, secure_payload: dict) -> str:
    # 1) Localizar la línea con `var DB = {...};(function(){...})()` (línea ~264)
    lines = html.split('\n')
    db_line_idx = None
    for i, line in enumerate(lines):
        if line.startswith('var DB = {') and 'delete DB._dict' in line:
            db_line_idx = i
            break
    if db_line_idx is None:
        sys.exit("FATAL: no se encontró la linea 'var DB = {...}; ...delete DB._dict;})()'")

    json_db, hidratador_body = _split_db_line(lines[db_line_idx])
    if json_db is None:
        sys.exit("FATAL: la linea de DB no matchea el patron esperado "
                 "(var DB = {...};(function(){...})()).")

    # Validar que es JSON parseable
    try:
        json.loads(json_db)
    except Exception as e:
        sys.exit(f"FATAL: DB no es JSON valido: {e}")

    # 2) Cifrar el JSON
    secure_payload = encrypt_data_for_users(json_db)
    secure_json = json.dumps(secure_payload, ensure_ascii=False, separators=(',', ':'))

    # 3) Reemplazar la linea con stub + SECURE_DATA + __hydrate
    new_db_line = (
        'var DB = {};'
        f'window.SECURE_DATA = {secure_json};'
        f'window.__hydrate = function(){{{hidratador_body}}};'
    )
    lines[db_line_idx] = new_db_line

    # 4) Convertir el bootstrap IIFE `(function init(){...})();` a `window.__boot=function(){...};`
    new_html = '\n'.join(lines)

    # Localizar el ultimo `})();` antes de `</script>` final (es el cierre del IIFE init)
    # y convertirlo. Tambien convertir el `(function init(){` del inicio.
    new_html = new_html.replace('(function init(){', 'window.__boot = function(){', 1)

    # Reemplazar la ultima ocurrencia de `})();\n</script>` por `};\n</script>`
    # (cierre del IIFE init seguido inmediatamente del cierre de script)
    rx_close = re.compile(r'\}\)\(\);\s*</script>\s*</body>\s*</html>\s*$')
    new_html, n = rx_close.subn('};\n</script>\n</body>\n</html>\n', new_html)
    if n != 1:
        sys.exit("FATAL: no se pudo localizar el cierre `})();</script></body></html>` del IIFE init.")

    # 5) Inyectar CSS del login en <head>
    new_html = new_html.replace('</head>', LOGIN_OVERLAY_CSS + '\n</head>', 1)

    # 6) Inyectar overlay HTML justo despues de <body>, con el logo del HTML fuente
    body_open = re.search(r'<body[^>]*>', new_html)
    if not body_open:
        sys.exit("FATAL: no se encontró <body>.")
    # Extraer el logo data URI del <img id="logo"> del HTML original (no sensible)
    logo_match = re.search(r'<img id="logo" src="(data:image/[^"]+)"', new_html)
    logo_src = logo_match.group(1) if logo_match else ''
    overlay_html = LOGIN_OVERLAY_HTML.replace('__LOGIN_LOGO_SRC__', logo_src)
    insertion_point = body_open.end()
    new_html = new_html[:insertion_point] + '\n' + overlay_html + new_html[insertion_point:]

    # 7) Inyectar JS del login antes de </body>
    new_html = new_html.replace('</body>', LOGIN_JS + '\n</body>', 1)

    return new_html


def main():
    if not os.path.exists(HTML_IN):
        sys.exit(f"FATAL: no existe {HTML_IN}. Corre el pipeline para generarlo.")
    with open(HTML_IN, 'r', encoding='utf-8') as f:
        html = f.read()
    print(f"Leido HTML fuente: {len(html)/1024/1024:.2f} MB")

    out = transformar_html(html, None)

    os.makedirs(DOCS_DIR, exist_ok=True)
    with open(HTML_OUT, 'w', encoding='utf-8') as f:
        f.write(out)
    size_mb = os.path.getsize(HTML_OUT) / 1024 / 1024
    print(f"Escrito HTML cifrado: {HTML_OUT} ({size_mb:.2f} MB)")
    # Pequeno .nojekyll para que GitHub Pages no procese con Jekyll
    with open(os.path.join(DOCS_DIR, '.nojekyll'), 'w') as f:
        f.write('')
    print("Listo. docs/index.html publicable en GitHub Pages.")


if __name__ == '__main__':
    main()
