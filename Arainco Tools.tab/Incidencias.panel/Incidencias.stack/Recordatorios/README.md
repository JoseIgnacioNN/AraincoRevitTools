# Recordatorios de incidencias

Envia recordatorios por correo a las personas asignadas a incidencias pendientes via **SMTP**. No depende de Revit ni de Outlook.

## Dependencias

```bash
pip install mjml-python
```

## Configuracion

1. Edita `recordatorios_config.json`
2. Configura `smtp.user` y `smtp.password` (App Password para Gmail) o usa variable de entorno
3. **Logo**: Coloca `logo.png` en esta carpeta. Si no existe, se usa encabezado con texto ARAINCO.

## Ejecucion

**Doble clic** en `ejecutar_recordatorios.bat` o desde terminal:

```
cd Recordatorios
python recordatorios_smtp.py
```

Detener: **Ctrl+C**

## Criterios y cadencia (configurables)

- **Estados:** Abierto, En revision
- **Cadencia por prioridad** (dias desde creacion):
  - Critica: cada 1 dia
  - Alta: cada 2 dias
  - Media: cada 4 dias
  - Baja: cada 6 dias

- **Solo laborables:** Lunes a Viernes (no envia en fin de semana)

## Tarea programada (9:00 AM Lun-Vie)

Ejecutar como Administrador:
```
powershell -ExecutionPolicy Bypass -File crear_tarea_9am.ps1
```

Crea la tarea "Arainco Recordatorios Incidencias" que ejecuta el script cada dia laborable a las 9:00.

## Seguridad

**No commitear** `recordatorios_config.json` con la contraseña real.

### Opcion 1: Variable de entorno (recomendado)

Define `RECORDATORIOS_SMTP_PASSWORD` antes de ejecutar:

- **Windows (PowerShell):** `$env:RECORDATORIOS_SMTP_PASSWORD = "tu-app-password"`
- **Tarea programada:** En las propiedades de la tarea, Configurar para -> Opciones -> "Ejecutar con los privilegios mas altos" y en Variables de entorno añadir `RECORDATORIOS_SMTP_PASSWORD`

Deja `smtp.password` vacio o como `"REEMPLAZAR_CON_PASSWORD_O_APP_PASSWORD"` en el config.

### Opcion 2: Archivo local (gitignored)

Crea `recordatorios_config.local.json` en la misma carpeta con la contraseña. Se fusiona sobre el config base. Ejemplo:

```json
{"smtp": {"password": "tu-app-password"}}
```

El archivo `.local.json` esta en `.gitignore` y no se sube al repositorio.
