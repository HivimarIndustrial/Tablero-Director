"""
Detecta si estamos en la red interna HIVICORP. Si no, asegura que la VPN
Sophos Connect (connection "181.39.159.50") este activa.

Logica:
  1) netsh wlan show interfaces -> buscar SSID. Si SSID == 'HIVICORP' -> no
     hacemos nada, estamos en la red interna.
  2) Si no, consultar a sccli si la conexion VPN ya esta levantada
     (campo "Virtual IP" no vacio en `sccli get -n <name>`).
  3) Si no, ejecutar sccli enable con las credenciales del keyring.
  4) Esperar hasta que Virtual IP este presente.

No intenta desconectar — la VPN la deja el usuario cuando quiera.

Credenciales:
  keyring servicio 'hivimar-tablero-vpn' (si existe).
  Fallback: 'hivimar-tablero-qlik' (mismo login NTLM corporativo).

Uso CLI:
  python conectar_vpn.py          # conecta si hace falta
  python conectar_vpn.py --status # solo reporta estado
"""
import keyring
import os
import re
import subprocess
import sys
import time
from typing import Optional

SCCLI = r"C:\Program Files (x86)\Sophos\Connect\sccli.exe"
VPN_NAME = '181.39.159.50'
SSID_CORP = 'HIVICORP'
SERVICE_VPN = 'hivimar-tablero-vpn'
SERVICE_QLIK = 'hivimar-tablero-qlik'  # fallback

WAIT_TIMEOUT_S = 60
WAIT_POLL_S = 2


def _run(cmd: list, timeout: int = 30) -> tuple:
    """Ejecuta un comando y devuelve (rc, stdout, stderr). No levanta excepciones."""
    try:
        p = subprocess.run(cmd, capture_output=True, text=True,
                           timeout=timeout, encoding='utf-8', errors='replace')
        return p.returncode, p.stdout or '', p.stderr or ''
    except subprocess.TimeoutExpired as e:
        return 124, e.stdout or '', e.stderr or f"timeout after {timeout}s"
    except FileNotFoundError:
        return 127, '', f"comando no encontrado: {cmd[0]}"


def get_current_ssid() -> Optional[str]:
    """Devuelve el SSID WiFi actual o None si no hay WiFi conectado."""
    rc, out, _ = _run(['netsh', 'wlan', 'show', 'interfaces'])
    if rc != 0:
        return None
    # Buscar 'SSID :' sin 'BSSID'. Soporta salida en espanol/ingles.
    for line in out.splitlines():
        s = line.strip()
        # "SSID                   : NETLIFE-Espino" -- excluir BSSID y "SSID del perfil"
        m = re.match(r'^SSID\s*:\s*(.+)$', s)
        if m and 'BSSID' not in s.upper() and 'PERFIL' not in s.upper() \
               and 'PROFILE' not in s.upper():
            return m.group(1).strip()
    return None


def vpn_is_connected() -> bool:
    """True si sccli reporta que la VPN tiene Virtual IP asignada."""
    rc, out, _ = _run([SCCLI, 'get', '-n', VPN_NAME])
    if rc != 0:
        return False
    for line in out.splitlines():
        m = re.match(r'^\s*Virtual IP\s*:\s*(\S+)', line)
        if m:
            ip = m.group(1).strip()
            # si hay IP real (no vacio), esta conectado
            return bool(ip) and ip.lower() != 'none'
    return False


def _load_vpn_credentials() -> tuple:
    """Intenta keyring servicio VPN; fallback al servicio Qlik."""
    cred = keyring.get_credential(SERVICE_VPN, None)
    if cred and cred.password:
        return cred.username, cred.password, SERVICE_VPN
    cred = keyring.get_credential(SERVICE_QLIK, None)
    if cred and cred.password:
        return cred.username, cred.password, SERVICE_QLIK
    raise SystemExit(
        f"ERROR: no hay credenciales VPN guardadas (keyring services "
        f"'{SERVICE_VPN}' ni '{SERVICE_QLIK}')."
    )


def connect_vpn(verbose: bool = True) -> bool:
    """Enciende la VPN. Devuelve True si quedo conectada."""
    user, pwd, src = _load_vpn_credentials()
    if verbose:
        print(f"[vpn] levantando conexion '{VPN_NAME}' con credenciales de '{src}' (user={user})")
    rc, out, err = _run([SCCLI, 'enable', '-n', VPN_NAME,
                         '-u', user, '-p', pwd], timeout=30)
    if verbose and out:
        print(out.strip())
    if rc != 0:
        print(f"[vpn] sccli enable rc={rc}: {err.strip()}", file=sys.stderr)
        return False
    # Esperar a que Virtual IP aparezca
    t0 = time.time()
    while time.time() - t0 < WAIT_TIMEOUT_S:
        if vpn_is_connected():
            if verbose:
                print(f"[vpn] conectada en {time.time()-t0:.1f}s")
            return True
        time.sleep(WAIT_POLL_S)
    print(f"[vpn] timeout esperando Virtual IP tras {WAIT_TIMEOUT_S}s", file=sys.stderr)
    return False


def ensure_vpn(verbose: bool = True) -> dict:
    """
    Asegura conectividad a la red interna.
    Devuelve dict con estado: {'ssid': str|None, 'in_corp_lan': bool,
                               'vpn_connected': bool, 'action': str}
    action in {'none_needed_corp_lan', 'already_connected', 'connected_now', 'failed'}
    """
    ssid = get_current_ssid()
    if verbose:
        print(f"[vpn] SSID actual: {ssid!r}")
    if ssid and ssid.upper() == SSID_CORP.upper():
        return {'ssid': ssid, 'in_corp_lan': True,
                'vpn_connected': False, 'action': 'none_needed_corp_lan'}
    if vpn_is_connected():
        if verbose:
            print(f"[vpn] ya estaba conectada")
        return {'ssid': ssid, 'in_corp_lan': False,
                'vpn_connected': True, 'action': 'already_connected'}
    if verbose:
        print(f"[vpn] no estaba conectada; conectando...")
    ok = connect_vpn(verbose=verbose)
    return {'ssid': ssid, 'in_corp_lan': False,
            'vpn_connected': ok, 'action': 'connected_now' if ok else 'failed'}


# ============ CLI ============
if __name__ == '__main__':
    if '--status' in sys.argv:
        ssid = get_current_ssid()
        in_corp = ssid and ssid.upper() == SSID_CORP.upper()
        conn = vpn_is_connected()
        print(f"SSID: {ssid!r}")
        print(f"En LAN corporativa HIVICORP: {in_corp}")
        print(f"VPN Sophos Connect conectada: {conn}")
        sys.exit(0)
    result = ensure_vpn()
    print()
    print(f"Resultado: {result}")
    sys.exit(0 if (result['vpn_connected'] or result['in_corp_lan']) else 2)
