"""
Guarda las credenciales de Qlik Sense en Windows Credential Manager.
Uso: corre este script UNA SOLA VEZ.
Se te pedira el usuario y la clave en la terminal; la clave NO se muestra en
pantalla al escribirla. Nada se escribe en archivos.
Despues los otros scripts las leen con:  keyring.get_password('hivimar-tablero-qlik', usuario)
"""
import keyring
import getpass
import sys

SERVICE = 'hivimar-tablero-qlik'

print("=" * 60)
print("  Guardar credenciales de Qlik Sense (Hivimar)")
print("=" * 60)
print()
print("Estas credenciales se guardan cifradas en el llavero de Windows")
print("(Credential Manager). No se escriben en ningun archivo.")
print()

# Aviso si ya hay credenciales guardadas
try:
    existing = keyring.get_credential(SERVICE, None)
    if existing and existing.username:
        print(f"[AVISO] Ya existen credenciales guardadas para usuario: {existing.username}")
        resp = input("Sobrescribir? (s/N): ").strip().lower()
        if resp != 's':
            print("Cancelado.")
            sys.exit(0)
        # Borrar las anteriores
        try:
            keyring.delete_password(SERVICE, existing.username)
        except Exception:
            pass
        print()
except Exception:
    pass

# Pedir usuario
usuario = input("Usuario Qlik (ej: hivimar\\eespino o eespino@hivimar.com): ").strip()
if not usuario:
    print("Usuario vacio, cancelado.")
    sys.exit(1)

# Pedir clave (no se muestra al escribir)
clave1 = getpass.getpass("Clave Qlik (no se mostrara): ")
if not clave1:
    print("Clave vacia, cancelado.")
    sys.exit(1)
clave2 = getpass.getpass("Repite la clave para confirmar: ")
if clave1 != clave2:
    print("ERROR: las claves no coinciden. Cancelado.")
    sys.exit(1)

# Guardar
keyring.set_password(SERVICE, usuario, clave1)

# Verificar
leida = keyring.get_password(SERVICE, usuario)
if leida == clave1:
    print()
    print("OK - credenciales guardadas correctamente.")
    print(f"    Servicio: {SERVICE}")
    print(f"    Usuario : {usuario}")
    print(f"    Backend : {keyring.get_keyring()}")
    print()
    print("Para recuperarlas mas tarde (en otros scripts):")
    print("    keyring.get_password('hivimar-tablero-qlik', '<usuario>')")
else:
    print("ERROR: no se pudieron verificar las credenciales guardadas.")
    sys.exit(1)
