"""
Inicializa los clientes de Google (Sheets, Drive) usando una cuenta de servicio.
Las credenciales se cargan desde la variable de entorno GOOGLE_CREDENTIALS_JSON
(contenido completo del JSON, NO la ruta).
"""
import os
import json
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def _cargar_credenciales():
    """Carga las credenciales desde variable de entorno o archivo local."""
    json_env = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if json_env:
        info = json.loads(json_env)
        return Credentials.from_service_account_info(info, scopes=SCOPES)
    # Fallback para pruebas locales
    path = os.environ.get("GOOGLE_CREDENTIALS_FILE", "credentials.json")
    if os.path.exists(path):
        return Credentials.from_service_account_file(path, scopes=SCOPES)
    raise RuntimeError(
        "No se encontraron credenciales. Define GOOGLE_CREDENTIALS_JSON "
        "(contenido del JSON) o GOOGLE_CREDENTIALS_FILE (ruta al archivo)."
    )


_creds = None
_gspread_client = None
_drive_service = None


def get_gspread():
    global _creds, _gspread_client
    if _gspread_client is None:
        if _creds is None:
            _creds = _cargar_credenciales()
        _gspread_client = gspread.authorize(_creds)
    return _gspread_client


def get_drive():
    global _creds, _drive_service
    if _drive_service is None:
        if _creds is None:
            _creds = _cargar_credenciales()
        _drive_service = build("drive", "v3", credentials=_creds, cache_discovery=False)
    return _drive_service
