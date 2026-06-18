"""
Configuración global del reporte.
Espejo de las constantes CONFIG e IDX de tu Apps Script.
"""

# ============================================================
# CONFIGURACIÓN PRINCIPAL
# ============================================================
SHEET_ID = "1c6qE_qtcNkMfL6j3g_CZo8MbuX4f2DmhoXAYZlhrfkM"
HOJA_DATOS = "Actividades"

EMAIL_DESTINATARIO = "fvega@sell-out.cl"   # fallback / pruebas
ASUNTO_EMAIL = "Reporte Diario Implementacion -"

# ============================================================
# DISTRIBUCIÓN DE CORREOS (correo interno)
# ============================================================
# Correos que reciben TODAS las campañas, sin importar el cliente.
CORREOS_FIJOS = [
    "fvega@sell-out.cl",
    "jparraguez@sell-out.cl",
    "fnunez@sell-out.cl",
    "bbaeza@sell-out.cl"
]

# Correos que se SUMAN según el CLIENTE de la campaña.
# El nombre del cliente debe coincidir con la columna CLIENTE de la planilla.
CORREOS_POR_CLIENTE = {
    "Mondelez":            ["msaavedra@sell-out.cl"],
    "Softys":              ["pmunoz@sell-out.cl", "gcastillo@sell-out.cl"],
    "Unilever":            ["dperera@sell-out.cl", "cleon@sell-out.cl"],
    "Andina":              ["pmunoz@sell-out.cl", "gcastillo@sell-out.cl"],
    "Sofruco":             ["pmunoz@sell-out.cl", "gcastillo@sell-out.cl"],
    "Wild Food":           ["msaavedra@sell-out.cl"],
    "Softys Professional": ["pmunoz@sell-out.cl", "gcastillo@sell-out.cl"],
    "Kraft":               ["pmunoz@sell-out.cl", "gcastillo@sell-out.cl"],
    "Agrosuper":           ["pmunoz@sell-out.cl", "gcastillo@sell-out.cl", "lvaliente@sell-out.cl"],
    "Ariztía":             ["pmunoz@sell-out.cl", "gcastillo@sell-out.cl", "dperera@sell-out.cl", "cleon@sell-out.cl"],
    "NotCo":               ["msaavedra@sell-out.cl"],
    "Cleaner":             ["pmunoz@sell-out.cl", "gcastillo@sell-out.cl"],
    "Kimberly Clark":      ["msaavedra@sell-out.cl"],
    "BIC":                 ["pmunoz@sell-out.cl", "gcastillo@sell-out.cl"],
    "Agricovial":          ["pmunoz@sell-out.cl", "gcastillo@sell-out.cl"],
    "Krispy Kreme":        ["msaavedra@sell-out.cl"],
    "CCU":                 ["pmunoz@sell-out.cl", "mvargas@sell-out.cl", "gcastillo@sell-out.cl", "dperera@sell-out.cl", "cleon@sell-out.cl"],
    "CCU Nestlé":          ["pmunoz@sell-out.cl", "mvargas@sell-out.cl", "gcastillo@sell-out.cl", "dperera@sell-out.cl", "cleon@sell-out.cl"],
    "CCU Pepsico":         ["pmunoz@sell-out.cl", "mvargas@sell-out.cl", "gcastillo@sell-out.cl", "dperera@sell-out.cl", "cleon@sell-out.cl"],
    "CCU Carozzi":         ["pmunoz@sell-out.cl", "mvargas@sell-out.cl", "gcastillo@sell-out.cl", "dperera@sell-out.cl", "cleon@sell-out.cl"],
    "Gourmet":             ["pmunoz@sell-out.cl", "mvargas@sell-out.cl", "gcastillo@sell-out.cl", "dperera@sell-out.cl", "cleon@sell-out.cl"],
    "Eucerin":             ["cchamorro@motionx.cl"],
    # Clientes pendientes de configurar: NO se agregan destinatarios (solo van los fijos)
    "Prosud":              [],
    "Natura":              [],
    "Crocs":               [],
}

# Correos placeholder que deben ignorarse si aparecieran en alguna lista.
CORREOS_IGNORAR = {"pendiente@sell-out.cl"}

# ============================================================
# DISTRIBUCIÓN DE CORREOS (correo EXTERNO - clientes)
# ============================================================
# Se envía solo los lunes. Va en copia visible (CC) a los clientes.
# Solo se envía para clientes que estén en CORREOS_POR_CLIENTE_EXTERNO.
CORREOS_FIJOS_EXTERNO = [
    "fvega@sell-out.cl",
]

CORREOS_POR_CLIENTE_EXTERNO = {
    "Softys": [
        "ctoledo@softys.com",
        "ctoledo@softysla.com",
        "pmunoz@sell-out.cl",
        "gcastillo@sell-out.cl",
    ],
}

CARPETA_REPORTES = "Reportes Diarios"
CARPETA_PADRE = "Control Interno"
CARPETA_FOTOS_RAIZ = "Fotos Implementación"

# IDs de templates (Google Slides + Google Sheets en Drive)
TEMPLATE_PPT_ID = "1DhZFRZpknziQqdwJdtw5OVjWoV5O3u5yr5rVQjGmU3I"  # genérico "Sell Out" (fallback)
TEMPLATE_EXCEL_ID = "1ULHh7FbNyMrue32wFjZs4RFqK9Ab-ZICnt_yxm_-6og"

# Carpeta de Drive donde viven los templates por cliente.
# El código busca "Template_Reporte_{Cliente}" dentro de esta carpeta.
# Si no encuentra uno específico, usa TEMPLATE_PPT_ID (el genérico Sell Out).
CARPETA_TEMPLATES_ID = "1ojipqO8-SC_qSFDT16fwoGkQBRjR87S7"
# Prefijo de los nombres de template. Resultado: "Template_Reporte_Softys"
PREFIJO_TEMPLATE = "Template_Reporte_"

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
