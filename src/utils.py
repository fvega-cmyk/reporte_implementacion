"""
Funciones utilitarias compartidas.
Equivalentes a fmt(), san(), fotoUrl(), buscarFoto() de Apps Script.
"""
import re
import unicodedata
from datetime import datetime, date, timedelta
from urllib.parse import quote
from io import BytesIO

from googleapiclient.http import MediaIoBaseDownload
from config import FOTO_URL_BASE


def ventana_semanal(hoy=None):
    """
    Ventana del reporte EXTERNO: martes de la semana anterior → lunes actual.

    Se ancla siempre en el LUNES (no en 'hoy'), para que el período sea el mismo
    aunque el workflow se dispare manualmente un miércoles.

    - Si hoy es lunes      → (martes semana pasada, hoy)
    - Si hoy es miércoles  → (martes semana pasada, lunes de esta semana)

    weekday(): lunes=0, martes=1, ..., domingo=6.
    """
    if hoy is None:
        hoy = date.today()
    lunes = hoy - timedelta(days=hoy.weekday())   # lunes de la semana en curso
    martes_anterior = lunes - timedelta(days=6)   # martes de la semana anterior
    return martes_anterior, lunes


def fmt(v):
    """Formatea fechas a 'dd/MM/yyyy'. Acepta date, datetime, str o serial."""
    if v is None or v == "":
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%d/%m/%Y")
    s = str(v).strip()
    if not s:
        return ""
    # Formatos comunes
    if re.match(r"^\d{2}/\d{2}/\d{4}$", s):
        return s
    if re.match(r"^\d{2}-\d{2}-\d{4}$", s):
        return s.replace("-", "/")
    # ISO
    if re.match(r"^\d{4}-\d{2}-\d{2}", s):
        try:
            return datetime.fromisoformat(s[:10]).strftime("%d/%m/%Y")
        except ValueError:
            pass
    # Serial de Sheets
    if re.match(r"^\d{4,5}$", s):
        try:
            d = datetime(1899, 12, 30) + (datetime.fromtimestamp(0) - datetime.fromtimestamp(0))
            from datetime import timedelta
            d = datetime(1899, 12, 30) + timedelta(days=int(s))
            return d.strftime("%d/%m/%Y")
        except Exception:
            pass
    return s


def parse_fecha(v):
    """Parsea cualquier formato a date (para comparar con hoy)."""
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    if not s:
        return None
    formatos = ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"]
    for f in formatos:
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            continue
    return None


def san(texto):
    """Sanitiza texto para usar en nombres de archivo (espejo de san() en GS)."""
    if not texto:
        return "sin_nombre"
    s = str(texto)
    s = re.sub(r"[^a-zA-Z0-9_\- ]", "", s)
    s = s.replace(" ", "_")
    return s[:50] or "sin_nombre"


def normalizar(texto):
    """Quita tildes y baja a minúsculas (para comparaciones)."""
    if not texto:
        return ""
    s = str(texto).strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s


def foto_url(ruta):
    """Construye la URL pública AppSheet a partir de la ruta de la foto."""
    if not ruta or not str(ruta).strip():
        return ""
    return FOTO_URL_BASE + quote(str(ruta).strip(), safe="")


def buscar_foto_blob(drive_service, ruta):
    """
    Equivalente a buscarFoto() en Apps Script:
    Busca el archivo por nombre en TODO Drive y descarga su contenido.
    Retorna bytes o None.
    """
    if not ruta:
        return None
    nombre = ruta.split("/")[-1]
    # Escapar comillas simples para la query de Drive
    nombre_esc = nombre.replace("'", "\\'")
    query = f"name = '{nombre_esc}' and trashed = false"

    try:
        resp = drive_service.files().list(
            q=query, fields="files(id, name)", pageSize=1,
            supportsAllDrives=True, includeItemsFromAllDrives=True,
        ).execute()
        archivos = resp.get("files", [])
        if not archivos:
            return None
        file_id = archivos[0]["id"]
        req = drive_service.files().get_media(fileId=file_id)
        buf = BytesIO()
        downloader = MediaIoBaseDownload(buf, req)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        buf.seek(0)
        return buf
    except Exception as e:
        print(f"  [WARN] Error buscando foto '{nombre}': {e}")
        return None
