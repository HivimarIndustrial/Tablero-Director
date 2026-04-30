"""
Guarda las credenciales SMTP (Office 365) en Windows Credential Manager
para el envio automatico del correo de Stock Valvulas a Peru.

Servidor por defecto: smtp.office365.com:587 (STARTTLS).
Usuario = correo del remitente (ej: consultorindustrial@hivimar.com).
Clave  = un App Password generado en https://mysignins.microsoft.com/security-info
         (NO usar la contrasena normal si tienes MFA activado).

Uso: doble clic al .bat o `python guardar_credenciales_smtp.py`.
Despues los scripts leen con:
    keyring.get_credential('hivimar-tablero-smtp', None)
"""
import keyring
import getpass
import sys

SERVICE = 'hivimar-tablero-smtp'

print("=" * 60)
print("  Guardar credenciales SMTP M365 (Hivimar)")
print("=" * 60)
print()
print("Se guardan cifradas en el Credential Manager de Windows.")
print("Si tienes MFA, debes generar un APP PASSWORD en:")
print("  https://mysignins.microsoft.com/security-info")
print()

try:
    existing = keyring.get_credential(SERVICE, None)
    if existing and existing.username:
        print(f"[AVISO] Ya existen credenciales para usuario: {existing.username}")
        resp = input("Sobrescribir? (s/N): ").strip().lower()
        if resp != 's':
            print("Cancelado.")
            sys.exit(0)
        try:
            keyring.delete_password(SERVICE, existing.username)
        except Exception:
            pass
        print()
except Exception:
    pass

usuario = input("Correo del remitente (ej: consultorindustrial@hivimar.com): ").strip()
if not usuario or '@' not in usuario:
    print("Correo invalido. Cancelado.")
    sys.exit(1)

clave1 = getpass.getpass("App Password (no se mostrara): ")
if not clave1:
    print("Clave vacia. Cancelado.")
    sys.exit(1)
clave2 = getpass.getpass("Repite la clave para confirmar: ")
if clave1 != clave2:
    print("ERROR: claves no coinciden. Cancelado.")
    sys.exit(1)

keyring.set_password(SERVICE, usuario, clave1)

leida = keyring.get_password(SERVICE, usuario)
if leida == clave1:
    print()
    print("OK - credenciales guardadas.")
    print(f"    Servicio: {SERVICE}")
    print(f"    Usuario : {usuario}")
    print()
    print("Prueba ahora con:")
    print("    python extraer_stock_valvulas_peru.py --display-mail")
else:
    print("ERROR: no se pudieron verificar las credenciales.")
    sys.exit(1)
