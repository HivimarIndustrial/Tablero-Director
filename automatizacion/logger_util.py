"""
logger_util.py - Logging para el pipeline del Tablero Industrial.

Funciones:
  setup_logging(nombre_ejecucion) -> (log_path, tee)
    Crea archivo automatizacion/logs/<YYYYMMDD_HHMMSS>_<nombre>.log
    Devuelve un objeto "tee" que redirige print() a stdout Y al archivo.

  cleanup_logs(dias=30)
    Elimina logs mas antiguos que N dias.

Todo queda en texto plano, una linea por evento, con timestamp.
"""
import os
import sys
import datetime as _dt
from io import TextIOBase

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LOGS_DIR = os.path.join(SCRIPT_DIR, 'logs')


class _Tee(TextIOBase):
    """Escribe a consola y a archivo al mismo tiempo, con timestamp al inicio de linea."""
    def __init__(self, file_obj, console=None, with_timestamp=True):
        self.file = file_obj
        self.console = console if console is not None else sys.__stdout__
        self.with_timestamp = with_timestamp
        self._line_start = True

    def write(self, data):
        if not data:
            return 0
        # Escribir en consola SIN modificar
        try:
            self.console.write(data)
            self.console.flush()
        except Exception:
            pass
        # Escribir en archivo con timestamp por linea
        try:
            out = data
            if self.with_timestamp:
                ts = _dt.datetime.now().strftime('%H:%M:%S')
                # Insertar timestamp al inicio de cada nueva linea no-vacia
                parts = []
                i = 0
                while i < len(out):
                    if self._line_start and out[i] != '\n':
                        parts.append(f"[{ts}] ")
                        self._line_start = False
                    ch = out[i]
                    parts.append(ch)
                    if ch == '\n':
                        self._line_start = True
                    i += 1
                out = ''.join(parts)
            self.file.write(out)
            self.file.flush()
        except Exception:
            pass
        return len(data)

    def flush(self):
        try:
            self.file.flush()
        except Exception:
            pass
        try:
            self.console.flush()
        except Exception:
            pass

    def close(self):
        try:
            self.file.close()
        except Exception:
            pass


def setup_logging(nombre_ejecucion='run'):
    """Inicia logging a archivo + consola. Devuelve (log_path, tee_obj)."""
    os.makedirs(LOGS_DIR, exist_ok=True)
    ts = _dt.datetime.now().strftime('%Y%m%d_%H%M%S')
    log_path = os.path.join(LOGS_DIR, f'{ts}_{nombre_ejecucion}.log')

    # Encabezado del archivo
    fh = open(log_path, 'w', encoding='utf-8', errors='replace')
    fh.write(f"=== Tablero Industrial - {nombre_ejecucion} ===\n")
    fh.write(f"Inicio : {_dt.datetime.now().isoformat(timespec='seconds')}\n")
    fh.write(f"PID    : {os.getpid()}\n")
    fh.write(f"CWD    : {os.getcwd()}\n")
    fh.write(f"Python : {sys.version.split()[0]}\n")
    fh.write(f"User   : {os.environ.get('USERNAME', '?')}\n")
    fh.write("=" * 60 + "\n\n")
    fh.flush()

    # Reemplazar stdout y stderr por un Tee que escriba al archivo
    tee_out = _Tee(fh, console=sys.stdout, with_timestamp=True)
    sys.stdout = tee_out
    sys.stderr = tee_out
    return log_path, tee_out


def finalize_logging(tee_obj, exit_code=0, error=None):
    """Cierra el log con resumen final."""
    try:
        sys.stdout.flush()
    except Exception:
        pass
    try:
        footer_ts = _dt.datetime.now().isoformat(timespec='seconds')
        sys.stdout.write("\n" + "=" * 60 + "\n")
        sys.stdout.write(f"Fin    : {footer_ts}\n")
        sys.stdout.write(f"Status : {'OK' if exit_code == 0 else 'ERROR (exit={0})'.format(exit_code)}\n")
        if error:
            sys.stdout.write(f"Error  : {error}\n")
        sys.stdout.write("=" * 60 + "\n")
        sys.stdout.flush()
    except Exception:
        pass
    # Restaurar stdout/stderr originales
    try:
        if isinstance(tee_obj, _Tee):
            tee_obj.close()
    except Exception:
        pass
    sys.stdout = sys.__stdout__
    sys.stderr = sys.__stderr__


def cleanup_logs(dias=30, verbose=False):
    """Elimina archivos de log mas antiguos que `dias`. Devuelve cuantos borro."""
    if not os.path.exists(LOGS_DIR):
        return 0
    cutoff = _dt.datetime.now() - _dt.timedelta(days=dias)
    borrados = 0
    for fn in os.listdir(LOGS_DIR):
        p = os.path.join(LOGS_DIR, fn)
        if not os.path.isfile(p):
            continue
        try:
            mtime = _dt.datetime.fromtimestamp(os.path.getmtime(p))
            if mtime < cutoff:
                os.remove(p)
                borrados += 1
                if verbose:
                    print(f"  log antiguo borrado: {fn}")
        except Exception:
            pass
    return borrados
