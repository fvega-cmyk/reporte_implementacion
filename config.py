"""
Configuración global del reporte.
Espejo de las constantes CONFIG e IDX de tu Apps Script.
"""

# ============================================================
# CONFIGURACIÓN PRINCIPAL
# ============================================================
SHEET_ID = "1c6qE_qtcNkMfL6j3g_CZo8MbuX4f2DmhoXAYZlhrfkM"
HOJA_DATOS = "Actividades"

EMAIL_DESTINATARIO = "fvega@sell-out.cl"
ASUNTO_EMAIL = "Reporte Diario Implementacion -"

CARPETA_REPORTES = "Reportes Diarios"
CARPETA_PADRE = "Control Interno"
CARPETA_FOTOS_RAIZ = "Fotos Implementación"

# IDs de templates (Google Slides + Google Sheets en Drive)
TEMPLATE_PPT_ID = "1DhZFRZpknziQqdwJdtw5OVjWoV5O3u5yr5rVQjGmU3I"
TEMPLATE_EXCEL_ID = "1ULHh7FbNyMrue32wFjZs4RFqK9Ab-ZICnt_yxm_-6og"

# ============================================================
# ÍNDICES DE COLUMNAS (espejo de IDX en tu Apps Script)
# ============================================================
IDX = {
    "ID_ACTIVIDAD": 0, "ID_CAMPANA": 1, "N_PRESUPUESTO": 2, "CAMPANA": 3,
    "EJECUTIVO": 4, "CLIENTE": 5, "ID_LOCAL": 6, "CADENA": 7, "COD": 8,
    "NOMBRE_SALA": 9, "DIRECCION": 10, "COMUNA": 11, "REGION": 12,
    "MARCA": 13, "CATEGORIA": 14, "GUIA": 15, "SKU": 16, "MATERIAL": 17,
    "CANTIDAD": 18, "FECHA_INICIO": 19, "FECHA_TERMINO": 20,
    "FECHA_TENTATIVA": 21, "FECHA_ENTREGA": 22, "REAGENDA": 23,
    "N_REAGENDA": 24, "FECHA_REAGENDA": 25, "FECHA_COMPROMISO": 26,
    "HORA_INICIO": 27, "HORA_TERMINO": 28, "PROCESO": 29, "DETALLE": 30,
    "OBSERVACIONES": 31, "LOGISTICO1": 32, "LOGISTICO2": 33,
    "ACTIVIDAD": 34, "TIPO_ACTIVIDAD": 35,
    "FOTO1": 36, "FOTO2": 37, "FOTO3": 38, "FOTO4": 39,
    "VIDEO": 40, "TIPO_ENVIO": 41, "COURIER": 42, "OT": 43,
    "SEGUIMIENTO": 44, "CREACION_PDF": 45, "INFORME_PDF": 46,
}

INDICES_FECHA = [19, 20, 21, 22, 25, 26]
INDICES_FOTO = [36, 37, 38, 39]

# URL base para construir el link público de las fotos (AppSheet)
FOTO_URL_BASE = (
    "https://www.appsheet.com/template/gettablefileurl?"
    "appName=OperacionesSellOut-333263235&tableName=Actividades&fileName="
)

# Estados que se cuentan en la hoja Resumen
ESTADOS = ["Realizado", "Entregado PDV", "Reagenda", "Rechazado", "Cancelado", "En proceso"]

# Configuración SMTP de Gmail (se sobrescribe con variables de entorno)
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587
