# Reporte Diario Implementación

Migración del reporte de Google Apps Script a Python + GitHub Actions.

## ¿Qué hace?

Todos los días a las 7am Chile:
1. Lee la hoja "Actividades" del Google Sheet.
2. Filtra las implementaciones activas hoy y las agrupa por campaña.
3. Por cada campaña: genera un Excel (Resumen + BBDD) y un PPT con fotos.
4. Envía un correo con ambos archivos adjuntos.

Mismo flujo que tu Apps Script, pero **sin límite de 6 minutos**, sin "auto-continuación" con triggers y mucho más rápido (especialmente el PPT).

## Estructura

```
reporte-implementacion/
├── .github/workflows/reporte.yml   ← cuándo correr (cron diario)
├── src/
│   ├── main.py                     ← orquestador (lo que se ejecuta)
│   ├── config.py                   ← CONFIG e IDX (espejo del GS)
│   ├── google_clients.py           ← autenticación a Google
│   ├── leer_sheets.py              ← lee y filtra la planilla
│   ├── generar_excel.py            ← arma el Excel
│   ├── generar_ppt.py              ← arma el PPT con fotos
│   ├── enviar_correo.py            ← manda el mail
│   └── utils.py                    ← funciones helper (fmt, san, etc.)
└── requirements.txt
```

## Setup paso a paso

### 1. Crear cuenta de servicio en Google Cloud

1. Andá a https://console.cloud.google.com/ con tu cuenta de Google.
2. Creá un proyecto nuevo (botón arriba a la izquierda → "Nuevo proyecto"). Nombre: `reporte-implementacion`.
3. En el buscador de arriba, escribí "API Library" → entrá → habilitá estas dos APIs:
   - **Google Sheets API**
   - **Google Drive API**
4. Buscá "Service Accounts" → "Crear cuenta de servicio".
   - Nombre: `reporte-bot`
   - Saltá los pasos de roles (opcional) → Crear.
5. Clic en la cuenta creada → pestaña "Keys" → "Add Key" → "Create new key" → JSON.
6. Se descarga un archivo `.json`. **Guardalo bien, lo vas a usar después.**
7. **Importante**: copiá el email de la cuenta de servicio (algo como `reporte-bot@reporte-implementacion.iam.gserviceaccount.com`).

### 2. Compartir recursos con la cuenta de servicio

Como si fuera un compañero más, compartí con ese email (con permiso "Editor"):

- La Google Sheet de Actividades (`1c6qE_qtcNkMfL6j3g_CZo8MbuX4f2DmhoXAYZlhrfkM`).
- El template Slides (`1DhZFRZpknziQqdwJdtw5OVjWoV5O3u5yr5rVQjGmU3I`).
- La carpeta raíz donde están las fotos en Drive.

### 3. Contraseña de aplicación de Gmail

1. Tu cuenta de Gmail debe tener verificación en 2 pasos activada.
2. Andá a https://myaccount.google.com/apppasswords
3. Generá una contraseña de aplicación. Te da 16 caracteres tipo `abcd efgh ijkl mnop`.
4. Guardala (la usás solo una vez, después no se puede ver de nuevo).

> Si querés usar un correo corporativo (no Gmail), avisame y te paso la configuración SMTP de Office 365 / Outlook.

### 4. Crear el repositorio en GitHub

1. Creá un repositorio **privado** en https://github.com → "New repository".
2. Subí los archivos de esta carpeta (podés arrastrarlos a la web de GitHub directamente).

### 5. Cargar los secretos en GitHub

En tu repo: **Settings → Secrets and variables → Actions → New repository secret**.

Creá estos tres secretos:

| Nombre | Valor |
|---|---|
| `GOOGLE_CREDENTIALS_JSON` | Pegá el **contenido completo** del JSON descargado en el paso 1 (no la ruta, el JSON entero). |
| `GMAIL_USER` | Tu correo Gmail (ej. `tunombre@gmail.com`). |
| `GMAIL_APP_PASS` | La contraseña de aplicación de 16 caracteres (sin espacios). |

### 6. Probarlo manualmente

1. En tu repo, pestaña **Actions**.
2. En el panel izquierdo, "Reporte Diario Implementación" → botón "Run workflow".
3. Esperá a que termine. Si todo va bien, verás los logs en verde y te llegará el correo.

Si falla, abrí el log y mirá qué dice. Los errores más típicos son:
- "permission denied" → falta compartir un archivo con la cuenta de servicio.
- "no such file" en `buscar_foto` → la foto no existe en Drive con ese nombre.
- "Authentication failed" en SMTP → la contraseña de aplicación está mal o tiene espacios.

### 7. ¡Listo!

A partir de mañana corre solo todos los días a las 7am Chile. Podés ver el historial en la pestaña Actions.

## Probarlo en tu PC (opcional)

```bash
pip install -r requirements.txt

# En Mac/Linux:
export GOOGLE_CREDENTIALS_FILE=/ruta/a/credentials.json
export GMAIL_USER=tunombre@gmail.com
export GMAIL_APP_PASS=abcd efgh ijkl mnop

# En Windows PowerShell:
$env:GOOGLE_CREDENTIALS_FILE = "C:\ruta\credentials.json"
$env:GMAIL_USER = "tunombre@gmail.com"
$env:GMAIL_APP_PASS = "abcdefghijklmnop"

python src/main.py
```

## Diferencias frente al Apps Script

| Apps Script | Esta versión |
|---|---|
| Timeout 6 min, se corta con campañas grandes | Sin límite (GitHub Actions: 6 horas) |
| `SlidesApp` lentísimo con muchas imágenes | `python-pptx` ~10–30x más rápido |
| Excel copiado del template + 1 llamada por celda | `openpyxl` arma el Excel desde cero, en memoria |
| Dos scripts encadenados con triggers + estado en PropertiesService | Un solo proceso lineal |
| Guarda archivos en Drive | Por ahora solo manda el correo. Si querés que también suba copias a Drive, avisame. |

## Modificaciones comunes

- **Cambiar destinatario del correo**: `src/config.py`, variable `EMAIL_DESTINATARIO`.
- **Mandar a varios**: en `enviar_correo.py`, cambiar `msg["To"]` a `"a@x.com, b@y.com"`.
- **Cambiar hora**: `.github/workflows/reporte.yml`, línea del `cron`. Acordate que está en UTC (Chile = UTC-3 o UTC-4 según época del año).
- **Agregar columnas nuevas a la BBDD**: como ahora se vuelcan todas las columnas del Sheet en orden, simplemente agregá la columna en el Sheet y se incluye sola.
