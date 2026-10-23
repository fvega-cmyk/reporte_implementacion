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
    "djorquera@sell-out.cl",
    "bbaeza@sell-out.cl"
]

# Correos que se SUMAN según el CLIENTE de la campaña.
# El nombre del cliente debe coincidir con la columna CLIENTE de la planilla.
CORREOS_POR_CLIENTE = {
    "Mondelez":            ["msaavedra@sell-out.cl", "ralvarez@sell-out.cl", "edgardo.gonzalez@mdlz.com"],
    "Softys":              ["pmunoz@sell-out.cl", "gcastillo@sell-out.cl"],
    "Unilever":            ["dperera@sell-out.cl", "mjcarreno@sell-out.cl", "cleon@sell-out.cl"],
    "Andina":              ["pmunoz@sell-out.cl", "gcastillo@sell-out.cl", "jmeza@sell-out.cl"],
    "Sofruco":             ["pmunoz@sell-out.cl", "gcastillo@sell-out.cl", "jmeza@sell-out.cl"],
    "Wild Food":           ["msaavedra@sell-out.cl", "ralvarez@sell-out.cl"],
    "Softys Professional": ["pmunoz@sell-out.cl", "gcastillo@sell-out.cl", "jmeza@sell-out.cl"],
    "Kraft":               ["pmunoz@sell-out.cl", "gcastillo@sell-out.cl"],
    "Agrosuper":           ["gcastillo@sell-out.cl", "mvargas@sell-out.cl", "lvaliente@sell-out.cl"],
    "Ariztía":             ["gcastillo@sell-out.cl", "dperera@sell-out.cl", "cleon@sell-out.cl"],
    "NotCo":               ["msaavedra@sell-out.cl", "ralvarez@sell-out.cl"],
    "Cleaner":             ["pmunoz@sell-out.cl", "gcastillo@sell-out.cl"],
    "Kimberly Clark":      ["msaavedra@sell-out.cl", "ralvarez@sell-out.cl", "sophia.plaza@kcc.com"],
    "BIC":                 ["pmunoz@sell-out.cl", "gcastillo@sell-out.cl"],
    "Agricovial":          ["pmunoz@sell-out.cl", "gcastillo@sell-out.cl"],
    "Krispy Kreme":        ["msaavedra@sell-out.cl"],
    "CCU":                 ["gcastillo@sell-out.cl", "dperera@sell-out.cl", "cleon@sell-out.cl"],
    "CCU Nestle":          ["gcastillo@sell-out.cl", "dperera@sell-out.cl", "cleon@sell-out.cl"],
    "CCU Pepsico":         ["gcastillo@sell-out.cl", "dperera@sell-out.cl", "cleon@sell-out.cl"],
    "CCU Carozzi":         ["gcastillo@sell-out.cl", "dperera@sell-out.cl", "cleon@sell-out.cl"],
    "Gourmet":             ["gcastillo@sell-out.cl", "dperera@sell-out.cl", "cleon@sell-out.cl"],
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

# ============================================================
# CAMPAÑAS CANCELADAS
# ============================================================
# Valores de la columna PROCESO que se consideran "cancelado".
# Se comparan normalizados (sin tildes, en minúsculas).
# Si el 100% de las filas de una campaña tiene uno de estos valores,
# la campaña NO se notifica (ni interno ni externo).
PROCESOS_CANCELADO = {"cancelado", "cancelada", "anulado", "anulada"}

# Si True: las filas con PROCESO vacío se ignoran al calcular el 100%
#   (una campaña con 8 canceladas + 2 vacías se considera cancelada).
# Si False (recomendado y por defecto): una sola fila con PROCESO vacío
#   alcanza para que la campaña SÍ se notifique.
CANCELADAS_IGNORAR_VACIOS = False

CARPETA_REPORTES = "Reportes Diarios"
CARPETA_PADRE = "Control Interno"
CARPETA_FOTOS_RAIZ = "Fotos Implementación"

# IDs de templates (Google Slides + Google Sheets en Drive)
TEMPLATE_PPT_ID = "14b8cFRMHODnJYq0SX4JgIQssSJzg02g1dPSyZiSqktI"  # genérico "Sell Out" (fallback)
TEMPLATE_EXCEL_ID = "1ULHh7FbNyMrue32wFjZs4RFqK9Ab-ZICnt_yxm_-6og"

# Carpeta de Drive donde viven los templates por cliente.
# El código busca "Template_Reporte_{Cliente}" dentro de esta carpeta.
# Si no encuentra uno específico, usa TEMPLATE_PPT_ID (el genérico Sell Out).
CARPETA_TEMPLATES_ID = "1ojipqO8-SC_qSFDT16fwoGkQBRjR87S7"
# Prefijo de los nombres de template. Resultado: "Template_Reporte_Softys"
PREFIJO_TEMPLATE = "Template_Reporte_"

# ============================================================
# ÍNDICES DE COLUMNAS
# ============================================================
# ESTE DICCIONARIO SE LLENA SOLO, en tiempo de ejecución.
#
# Arranca vacío a propósito. Apenas leer_sheets.leer_actividades() lee la
# planilla, columnas.sincronizar_idx() lo rellena buscando cada campo por el
# NOMBRE de su encabezado (fila 1), no por posición.
#
# Por qué: antes las posiciones estaban escritas a mano acá. Al insertar
# columnas en la planilla, todo lo que estaba a la derecha se corría y el
# código leía la columna de al lado sin dar error: devolvía "0 filas" como si
# ese fuera el resultado correcto. Ahora podés insertar o mover columnas sin
# tocar el código, y si falta una columna obligatoria el proceso corta con un
# mensaje que dice cuál.
#
# Los nombres de encabezado que se buscan están en src/columnas.py (ALIAS).
# Si tu planilla usa otro nombre, agregalo ahí.
#
# NO escribas índices acá a mano: se sobrescriben al leer la planilla.
IDX = {}

# Ancho al que se rellena cada fila. También lo setea sincronizar_idx().
ANCHO_FILA = 0

# ============================================================
# REINTENTOS CONTRA LA API DE DRIVE
# ============================================================
# Drive devuelve errores 500 "Internal Error" y corta descargas cada tanto,
# sin razón ni patrón. Son fallas transitorias de Google, no del código.
# Antes no se reintentaba nada: un timeout bajando el template de 790 KB
# mataba la campaña entera, perdiendo sus 28 slides ya procesadas.
#
# La librería de Google trae reintentos con espera creciente incorporados,
# pero hay que pedirlos explícitamente en cada llamada con num_retries.
# El default es 0, o sea ninguno.
REINTENTOS_DRIVE = 5

# ============================================================
# ESCAPE MANUAL (normalmente vacío)
# ============================================================
# Si un campo no se resuelve por nombre (encabezado raro, dos columnas que se
# llaman parecido), fijá acá su posición 0-based y gana sobre la búsqueda.
# Ejemplo:  COLUMNAS_OVERRIDE = {"FECHA_COMPROMISO": 28}
# Ojo: esto vuelve a atarte a la posición para ESE campo. Es un parche de
# emergencia, no la forma normal de configurar.
COLUMNAS_OVERRIDE = {}

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
