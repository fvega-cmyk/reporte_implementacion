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
# FOTOS
# ============================================================
# Cantidad máxima de fotos por slide (el template define cuántos placeholders
# [FOTO n] hay; esto es el techo).
MAX_FOTOS_INGRESOS = 4

# Si para algún PROCESO querés recortar la cantidad de fotos, como se hace en
# Implementación (Realizado→3, Rechazado→1), definilo acá. Vacío = todas.
#   Ej: {"rechazado": 1}
FOTOS_POR_PROCESO_INGRESOS = {}

# ============================================================
# CANCELADAS
# ============================================================
# Si True, las campañas 100% canceladas no se notifican (igual que en
# Implementación). El actualizador las sigue generando en Drive.
EXCLUIR_CANCELADAS_INGRESOS = True
