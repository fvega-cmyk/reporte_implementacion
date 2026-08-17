"""
Configuración del reporte SEMANAL DE INGRESOS.

Flujo paralelo al de Implementación. Diferencias clave:

  Implementación                        Ingresos
  ------------------------------------  ------------------------------------
  ACTIVIDAD = "Implementacion"          ACTIVIDAD = "Ingreso"
                                        TIPO ACTIVIDAD = "Activación"
  Excel + PPT por campaña               SOLO PPT por campaña
  Carpeta Cliente/Campaña               Carpeta SEMANA 33 - 2026 (plana)
  1 correo POR campaña                  1 SOLO correo con todos los PPT
  Ventana: hoy (o martes→lunes)         Ventana: semana ISO completa (lu→do)

Se reutiliza de config.py: SHEET_ID, HOJA_DATOS, IDX, SMTP, PROCESOS_CANCELADO.
"""
from config import IDX  # noqa: F401  (re-export por comodidad)

# ============================================================
# FILTRO DE LA PLANILLA
# ============================================================
# Valores esperados en las columnas ACTIVIDAD y TIPO ACTIVIDAD.
# La comparación es normalizada (sin tildes, minúsculas), así que da lo mismo
# si en la planilla dice "Activación", "ACTIVACION" o "activacion".
ACTIVIDAD_INGRESO = "Ingreso"
TIPO_ACTIVIDAD_INGRESO = "Activación"

# Si en algún momento hay más de un TIPO DE ACTIVIDAD que deba entrar al mismo
# reporte, agregalos acá y el filtro los acepta a todos.
TIPOS_ACTIVIDAD_ACEPTADOS = [TIPO_ACTIVIDAD_INGRESO]

# ============================================================
# TEMPLATE
# ============================================================
# El template vive en "Mi unidad" → carpeta "Templates_Reportes".
# IMPORTANTE: compartí ESA CARPETA (o al menos el archivo) con el email de la
# cuenta de servicio, con permiso de Lector. Solo se lee, nunca se escribe ahí.
NOMBRE_TEMPLATE_INGRESOS = "Template_Reporte_Ingresos"
NOMBRE_CARPETA_TEMPLATES = "Templates_Reportes"

# Atajo opcional: si pegás acá el ID del archivo del template, se usa directo
# y se saltea la búsqueda por nombre (más rápido y a prueba de homónimos).
# Se saca de la URL: docs.google.com/presentation/d/<ESTE_ID>/edit
TEMPLATE_PPT_INGRESOS_ID = ""

# ============================================================
# DESTINO EN DRIVE
# ============================================================
# Unidad Compartida donde se crean las carpetas de semana.
# Se puede sobrescribir con la variable de entorno DRIVE_CARPETA_INGRESOS_ID.
# Estructura resultante:
#   [Unidad Compartida]
#     ├── SEMANA 32 - 2026
#     │     ├── Campana_A_SEMANA_32-2026.pptx
#     │     └── Campana_B_SEMANA_32-2026.pptx
#     └── SEMANA 33 - 2026
#           └── Campana_A_SEMANA_33-2026.pptx
DRIVE_CARPETA_INGRESOS_ID_DEFAULT = "0ABVN4w8aqhhqUk9PVA"

# Nombre del archivo. {campana} viene sanitizado (sin tildes ni símbolos).
PATRON_NOMBRE_PPT = "{campana}_SEMANA_{semana}-{anio}.pptx"

# ============================================================
# CORREO (uno solo por semana)
# ============================================================
ASUNTO_INGRESOS = "Reporte Semanal de Ingresos"

CORREOS_INGRESOS = [
    "fvega@sell-out.cl",
    # TODO: agregar acá el resto del equipo que debe recibir el reporte.
]

# Copia visible, si hiciera falta. Dejar vacío si no.
CORREOS_INGRESOS_CC = []

# ============================================================
# UNA SLIDE POR LOCAL (no por material)
# ============================================================
# En Implementación cada fila (sala + material) genera su propia slide.
# En Activación/Ingreso NO: las fotos del local se repiten fila por fila, así
# que se agrupa por LOCAL y cada local es UNA slide, con el listado de sus
# materiales adentro.
#
# La agrupación se hace por ID_LOCAL. Si una fila viene sin ID_LOCAL, se cae
# a "COD|NOMBRE_SALA" para no perderla ni mezclarla con otro local.
AGRUPAR_POR = "ID_LOCAL"

# Formato de cada línea del listado de materiales, cuando el template no trae
# una línea propia con los tokens. Si el template SÍ trae una línea tipo
# "[MATERIAL] ([CANTIDAD]) : [PROCESO]", se respeta esa (con su formato,
# viñetas, tamaño, etc.) y este valor se ignora.
FORMATO_LINEA_MATERIAL = "[MATERIAL] ([CANTIDAD]) : [PROCESO]"

# Si dos filas del mismo local traen el MISMO material, cantidad y proceso,
# se colapsan en una sola línea (duplicado de carga).
#   False (default) → "Exhibidor (1) : Realizado" queda una vez, sin sumar.
#   True            → suma las cantidades: "Exhibidor (2) : Realizado".
# Ojo: sumar es riesgoso si en realidad son dos registros legítimos distintos.
SUMAR_CANTIDADES_REPETIDAS = False

# ============================================================
# FOTOS
# ============================================================
# Fotos por local. Las fotos vienen repetidas en todas las filas del local,
# así que se juntan todas, se quitan duplicados (por nombre de archivo) y se
# toman las primeras N.
MAX_FOTOS_INGRESOS = 2

# --- Dónde viven las fotos de las activaciones ---
# Ruta real: Mi unidad / Control Interno / <valor de la columna FOTO 1 o FOTO 2>
# La columna completa el resto de la ruta (subcarpeta + nombre de archivo).
#
# IMPORTANTE: compartí esta carpeta con la cuenta de servicio (permiso Lector).
NOMBRE_CARPETA_FOTOS_INGRESOS = "Control Interno"

# Recomendado: pegá acá el ID de la carpeta para saltear la búsqueda por
# nombre. Se saca de la URL al abrir la carpeta en Drive:
#   drive.google.com/drive/folders/ESTE_ES_EL_ID
# Es lo más seguro si existe más de una carpeta llamada "Control Interno".
CARPETA_FOTOS_INGRESOS_ID = ""

# Si no encuentra la foto en su ruta, ¿buscarla por nombre en TODO el Drive?
#   False (default) → no. Mejor que falte una foto a que aparezca la de otra
#                     campaña: un nombre repetido en la carpeta de
#                     implementaciones traería la foto equivocada y nadie lo
#                     notaría en el PPT que ve el cliente.
#   True            → sí, como último recurso. Queda avisado en el log.
FALLBACK_BUSQUEDA_GLOBAL = False

# ============================================================
# CANCELADAS
# ============================================================
# Si True, las campañas 100% canceladas no se notifican (igual que en
# Implementación). El actualizador las sigue generando en Drive.
EXCLUIR_CANCELADAS_INGRESOS = True
