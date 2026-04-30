"""
notificar.py - Notificacion de fallas del pipeline.

Canales en orden de prioridad:
  1) Archivo ULTIMO_ERROR.txt en la raiz del proyecto (siempre).
     Si falla un run, crea el archivo. Si el siguiente run sale OK, lo borra.
  2) Windows Toast notification (via PowerShell, no requiere instalar nada).
  3) Teams webhook (opcional). Si hay una URL guardada en keyring
     (servicio 'hivimar-tablero-teams', user 'webhook'), manda un mensaje.

Uso:
  from notificar import notificar_exito, notificar_error
  notificar_exito(resumen_texto)
  notificar_error(mensaje, log_path=..., traceback_txt=...)
"""
import os
import subprocess
import sys
import traceback as _traceback_module
from datetime import datetime
from urllib import request as _urlreq, error as _urlerr
import json
import keyring

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
ERROR_FILE = os.path.join(PROJECT_DIR, 'ULTIMO_ERROR.txt')
EXITO_FILE = os.path.join(PROJECT_DIR, 'ULTIMO_OK.txt')

TEAMS_SERVICE = 'hivimar-tablero-teams'
TEAMS_USER = 'webhook'


def _teams_webhook_url():
    try:
        return keyring.get_password(TEAMS_SERVICE, TEAMS_USER)
    except Exception:
        return None


def _toast_windows(title: str, message: str, error: bool = False):
    """Muestra un toast de Windows via PowerShell. No bloquea si falla."""
    # Escape de comillas simples para PowerShell
    t = (title or '').replace("'", "''")[:60]
    m = (message or '').replace("'", "''")[:250]
    icon = 'Error' if error else 'Info'
    ps_script = f"""
Add-Type -AssemblyName System.Windows.Forms
$ni = New-Object System.Windows.Forms.NotifyIcon
$ni.Icon = [System.Drawing.SystemIcons]::{icon}
$ni.BalloonTipTitle = '{t}'
$ni.BalloonTipText = '{m}'
$ni.Visible = $true
$ni.ShowBalloonTip(8000)
Start-Sleep -Seconds 8
$ni.Dispose()
"""
    try:
        subprocess.Popen(
            ['powershell.exe', '-NoProfile', '-WindowStyle', 'Hidden',
             '-ExecutionPolicy', 'Bypass', '-Command', ps_script],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
    except Exception:
        pass  # no bloquear si no hay powershell


def _teams_send(webhook_url: str, title: str, message: str, ok: bool = True):
    """Manda un mensaje a un canal Teams via webhook. Silent-fail."""
    if not webhook_url:
        return False
    color = '00b894' if ok else 'e74c3c'
    payload = {
        "@type": "MessageCard",
        "@context": "http://schema.org/extensions",
        "themeColor": color,
        "title": title[:120],
        "text": message[:2000],
    }
    try:
        data = json.dumps(payload).encode('utf-8')
        req = _urlreq.Request(
            webhook_url, data=data,
            headers={'Content-Type': 'application/json'},
            method='POST',
        )
        with _urlreq.urlopen(req, timeout=10) as resp:
            return resp.status in (200, 202)
    except Exception as e:
        print(f"  [notificar] teams webhook fallo: {e}", file=sys.stderr)
        return False


def notificar_exito(resumen: str):
    """Llamar al final si el pipeline termino bien.
    Borra ULTIMO_ERROR.txt si existe, escribe ULTIMO_OK.txt, opcionalmente
    manda OK a Teams.
    """
    ts = datetime.now().isoformat(timespec='seconds')
    # Borrar archivo de error si existia
    if os.path.exists(ERROR_FILE):
        try:
            os.remove(ERROR_FILE)
        except Exception:
            pass
    # Escribir archivo OK
    try:
        with open(EXITO_FILE, 'w', encoding='utf-8') as f:
            f.write(f"Pipeline Tablero Industrial — OK\n")
            f.write(f"Timestamp: {ts}\n\n")
            f.write(resumen + '\n')
    except Exception:
        pass
    # Teams (si hay webhook)
    url = _teams_webhook_url()
    if url:
        _teams_send(url, "Tablero Industrial — OK", f"**{ts}**\n\n{resumen}", ok=True)


def notificar_error(mensaje: str, log_path: str = None, traceback_txt: str = None):
    """Llamar si el pipeline fallo.
    Crea ULTIMO_ERROR.txt con detalle, muestra toast Windows, y si hay
    webhook Teams manda alerta con los detalles.
    """
    ts = datetime.now().isoformat(timespec='seconds')
    # Archivo ULTIMO_ERROR.txt con todos los detalles
    try:
        with open(ERROR_FILE, 'w', encoding='utf-8') as f:
            f.write(f"!! Pipeline Tablero Industrial FALLO !!\n")
            f.write(f"Timestamp: {ts}\n\n")
            f.write(f"Mensaje:\n{mensaje}\n\n")
            if log_path:
                f.write(f"Log completo: {log_path}\n\n")
            if traceback_txt:
                f.write("Traceback:\n")
                f.write(traceback_txt + '\n')
    except Exception as e:
        print(f"  [notificar] no pude escribir ULTIMO_ERROR.txt: {e}", file=sys.stderr)

    # Borrar archivo de OK si existia
    if os.path.exists(EXITO_FILE):
        try:
            os.remove(EXITO_FILE)
        except Exception:
            pass

    # Toast Windows
    corto = mensaje[:200] if mensaje else 'Falla desconocida'
    _toast_windows("Tablero Industrial FALLÓ", corto + " | ver ULTIMO_ERROR.txt", error=True)

    # Teams (si hay webhook)
    url = _teams_webhook_url()
    if url:
        cuerpo = f"**{ts}**\n\n**Error:** {mensaje}\n\n"
        if log_path:
            cuerpo += f"**Log:** `{log_path}`\n\n"
        if traceback_txt:
            cuerpo += f"```\n{traceback_txt[:1500]}\n```"
        _teams_send(url, "Tablero Industrial FALLÓ", cuerpo, ok=False)


def capture_traceback(exc: Exception) -> str:
    """Devuelve el traceback de una excepcion como string."""
    return ''.join(_traceback_module.format_exception(type(exc), exc, exc.__traceback__))


if __name__ == '__main__':
    # Pruebas manuales
    if len(sys.argv) > 1 and sys.argv[1] == 'test-ok':
        notificar_exito("Prueba de mensaje OK desde notificar.py")
        print("Enviado mensaje OK")
    elif len(sys.argv) > 1 and sys.argv[1] == 'test-error':
        notificar_error("Prueba de mensaje de error desde notificar.py",
                        log_path='automatizacion/logs/test.log',
                        traceback_txt='fake traceback')
        print("Enviado mensaje de error. Revisa ULTIMO_ERROR.txt y toast de Windows.")
    else:
        print("Uso: python notificar.py {test-ok|test-error}")
