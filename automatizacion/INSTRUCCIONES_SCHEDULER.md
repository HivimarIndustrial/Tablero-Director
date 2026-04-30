# Actualización automática — Task Scheduler

Este documento explica cómo funciona la ejecución automática del pipeline
del Tablero Industrial, cómo verificarla, y qué hacer si algo falla.

---

## ✅ Estado: YA CONFIGURADO

La tarea **`Tablero Industrial - Actualizar`** está registrada en Task
Scheduler de Windows. Se disparará automáticamente:

- **Todos los días a las 06:00 AM** (hora local de la PC).
- **Al iniciar sesión** (con 2 min de retraso) — compensa si la PC estaba
  apagada a las 6 AM.

Se ejecuta con tu usuario (`consultorindustrial`) sin necesidad de
permisos admin.

---

## 📁 Qué pasa en cada corrida

```
06:00 AM
  |-- Verifica VPN (si no estás en HIVICORP, levanta Sophos)
  |-- Descarga Qlik: ventas, cartera, inventario, rotación
  |-- Descarga Odoo: cotizaciones, oportunidades, actividades, visitas
  |-- Lee dim_* del Excel (READ-ONLY)
  |-- Enriquece los datos (_std columns, VLOOKUPs en Python)
  |-- Escribe CSVs enriquecidos en salida_raw/
  |-- Genera db_output.js
  |-- Inyecta en Hivimar_Tablero_Industrial 8.0.html
  |-- Inyecta timestamp "Actualizado: DD/MM/YYYY HH:MM"
  `-- Cierra y notifica

Tiempo: ~2 minutos si todo bien.
```

---

## 📜 Dónde ver los logs

Cada corrida crea un archivo en:
```
automatizacion/logs/YYYYMMDD_HHMMSS_update_tablero.log
```

El log incluye:
- Timestamp de cada línea
- Todos los pasos del pipeline con tiempos
- Detalle de filas descargadas por cada fuente
- Error completo con traceback si algo falla

Los logs de más de **30 días** se borran automáticamente en cada corrida.

---

## 🔔 Cómo te enteras si algo falla

### Señal visible 1 — Archivo en la raíz del proyecto

Si la corrida falló, aparece el archivo:
```
C:\Users\consultorindustrial\OneDrive - Hivimar\...\Indicadores Tablero\ULTIMO_ERROR.txt
```

Con el detalle del error, ruta al log y traceback. Si el siguiente run
sale bien, se borra automáticamente.

Si todo salió bien, en su lugar existe:
```
ULTIMO_OK.txt
```

**Rutina recomendada**: al llegar a la oficina, un vistazo a la carpeta.
Si ves `ULTIMO_ERROR.txt`, algo pasó.

### Señal visible 2 — Toast de Windows

Cuando falla, Windows muestra una notificación tipo "globo" en la bandeja
del sistema: "Tablero Industrial FALLÓ".

Dura ~8 segundos. Si estabas frente al PC, lo ves.

### Señal visible 3 (opcional) — Teams webhook

Si configuras un webhook de Teams (ver sección más abajo), cada vez que
falle el pipeline llega un mensaje a ese canal. Igual para éxitos si lo
quieres activar.

---

## 🧪 Cómo probar la tarea manualmente (sin esperar a las 6 AM)

### Desde PowerShell:
```powershell
Start-ScheduledTask -TaskName "Tablero Industrial - Actualizar"
```

### Desde el Task Scheduler (GUI):
1. Tecla Windows → escribir "Task Scheduler" → Abrir.
2. Panel izquierdo → "Biblioteca del Programador de tareas".
3. Buscar "Tablero Industrial - Actualizar".
4. Botón derecho → "Ejecutar".

La corrida empieza de inmediato. Se ve el estado en la misma ventana.
Después revisa el log en `automatizacion/logs/`.

---

## 🛠️ Administración de la tarea

### Ver estado

```powershell
Get-ScheduledTask -TaskName "Tablero Industrial - Actualizar"
Get-ScheduledTaskInfo -TaskName "Tablero Industrial - Actualizar"
```

### Desactivar temporalmente (ej. durante vacaciones)

```powershell
Disable-ScheduledTask -TaskName "Tablero Industrial - Actualizar"
```

Para reactivar:
```powershell
Enable-ScheduledTask -TaskName "Tablero Industrial - Actualizar"
```

### Cambiar la hora

```powershell
cd "C:\Users\consultorindustrial\OneDrive - Hivimar\Documentos\HIVIMAR\Proyectos\SGI 2.0\Indicadores Tablero\automatizacion"
powershell -ExecutionPolicy Bypass -File configurar_scheduler.ps1 -HoraDiaria "07:30"
```

Esto re-registra la tarea con la nueva hora.

### Desregistrar completamente

```powershell
powershell -ExecutionPolicy Bypass -File configurar_scheduler.ps1 -Remove
```

---

## ⚠️ Limitaciones a tener en cuenta

1. **Tu PC debe estar encendido** a las 6 AM para que corra.
   - Si no, corre al siguiente logon (con retraso de 2 min).
   - En vacaciones, la tarea no corre y los usuarios ven datos viejos.
   - Considera dejar el PC en modo "suspensión" en vez de "apagado".

2. **Tu usuario debe estar logueado** (LogonType Interactive).
   - Si necesitas que corra con sesión cerrada, hay que cambiar el
     principal y guardar tu contraseña en Task Scheduler (no recomendado).

3. **La VPN debe estar configurada**.
   - El pipeline intenta levantar Sophos si no estás en HIVICORP.
   - Si Sophos no tiene credenciales guardadas o cambió tu password,
     fallará. El error llegará como notificación.

4. **Credenciales en keyring deben ser válidas**.
   - `hivimar-tablero-qlik` → Industria
   - `hivimar-tablero-odoo` → consultorindustrial@hivimar.com
   - Si cambias contraseñas, actualiza estas credenciales.

---

## 🔧 Configurar notificación por Teams (opcional)

Si quieres recibir alertas de falla en un canal de Teams, necesitas
crear un **Incoming Webhook** en ese canal y guardar la URL en keyring.

### Paso 1 — Crear webhook en Teams

1. En Teams, ve al canal donde quieras las alertas.
2. Clic en "..." junto al nombre del canal → "Conectores".
3. Busca "Incoming Webhook" → "Configurar".
4. Dale un nombre (ej. "Tablero Industrial Alerts") y un icono.
5. Copia la URL que te genera (empieza con `https://hivimar.webhook.office.com/...`).

### Paso 2 — Guardar URL en keyring

```powershell
python -c "import keyring; keyring.set_password('hivimar-tablero-teams', 'webhook', 'PEGA_AQUI_TU_URL_COMPLETA')"
```

### Paso 3 — Probar

```powershell
cd automatizacion
python notificar.py test-error
```

Debería llegar un mensaje de prueba al canal. Si llega, quedará
configurado para todos los errores futuros del pipeline.

Para desactivar: borrar la credencial del keyring.

---

## 🧯 Troubleshooting

### "La tarea muestra 'Last Run Result: 0x1' o similar"

Abre el log más reciente en `automatizacion/logs/` y busca el error.
Probablemente:
- VPN caída
- Credenciales expiradas
- Excel abierto durante la corrida (no debería afectar, pero por si acaso)

### "El tablero no se actualiza pero no veo ULTIMO_ERROR.txt"

Puede que la tarea ni siquiera haya corrido. Verifica:
```powershell
Get-ScheduledTaskInfo -TaskName "Tablero Industrial - Actualizar" | Format-List
```
Mira `LastRunTime` y `LastTaskResult`. Si `LastTaskResult = 0` = OK.
Cualquier otro valor es error.

### "Chrome falla con net::ERR_TIMED_OUT"

Síntoma común cuando hay múltiples procesos de Chrome corriendo. El
pipeline tiene retry automático, pero si persiste:
```powershell
Get-Process chrome | Stop-Process -Force
```

### "El pipeline toma demasiado tiempo (> 5 min)"

Típicamente es internet lento o Qlik/Odoo saturados. Revisa el log.
Si es recurrente, contactar a IT para revisar red.

---

## 🔄 Flujo completo al cambiar algo en el pipeline

Si modificas código Python en `automatizacion/`, no hace falta
re-registrar la tarea — apunta al script por ruta, no toma snapshot.

Si mueves la carpeta del proyecto, sí hay que re-registrar:
```powershell
cd AUTOMATIZACION_NUEVA_RUTA
powershell -ExecutionPolicy Bypass -File configurar_scheduler.ps1
```

---

**Última actualización**: 2026-04-19
**Responsable**: consultorindustrial@hivimar.com
