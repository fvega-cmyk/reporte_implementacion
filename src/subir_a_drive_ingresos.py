"""
Sube/actualiza los PPT de Ingresos en una carpeta por SEMANA:

    [Unidad Compartida]
      ├── SEMANA 32 - 2026
      │     ├── Campana_A_SEMANA_32-2026.pptx
      │     └── Campana_B_SEMANA_32-2026.pptx
      └── SEMANA 33 - 2026
            └── Campana_A_SEMANA_33-2026.pptx

Todas las campañas de la semana comparten carpeta (estructura plana, sin
subcarpeta por cliente), que es lo que pide el flujo de Ingresos.

Igual que en Implementación, si el archivo ya existe se ACTUALIZA su contenido
con files().update(): mismo file ID, mismo link. Así el actualizador puede
correr N veces en la semana sin generar duplicados ni romper links ya enviados.

Destino: DRIVE_CARPETA_INGRESOS_ID (env) o el default de config_ingresos.py.
"""
import os
from io import BytesIO

from googleapiclient.http import MediaIoBaseUpload

from utils import san
from google_clients import get_drive
from config_ingresos import (
    DRIVE_CARPETA_INGRESOS_ID_DEFAULT,
    PATRON_NOMBRE_PPT,
)

MIME_PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
MIME_FOLDER = "application/vnd.google-apps.folder"


def carpeta_raiz_id():
    return (os.environ.get("DRIVE_CARPETA_INGRESOS_ID")
            or DRIVE_CARPETA_INGRESOS_ID_DEFAULT).strip()


def _drive_id_de(drive, folder_id):
    """
    driveId de la Unidad Compartida que contiene la carpeta.
    Devuelve None si la carpeta está en "Mi unidad" (no es unidad compartida).
    """
    info = drive.files().get(
        fileId=folder_id, fields="id, name, driveId", supportsAllDrives=True
    ).execute()
    return info.get("driveId")


def _listar(drive, q, drive_id):
    """files().list() que funciona tanto en Unidad Compartida como en Mi unidad."""
    params = {
        "q": q,
        "fields": "files(id, name)",
        "supportsAllDrives": True,
        "includeItemsFromAllDrives": True,
        "pageSize": 25,
    }
    if drive_id:
        params["corpora"] = "drive"
        params["driveId"] = drive_id
    return drive.files().list(**params).execute().get("files", [])


def _buscar_o_crear_carpeta(drive, padre_id, drive_id, nombre):
    """Busca la subcarpeta por nombre exacto; si no existe la crea."""
    nombre_q = nombre.replace("'", "\\'")
    q = (f"name = '{nombre_q}' and mimeType = '{MIME_FOLDER}' "
         f"and '{padre_id}' in parents and trashed = false")
    encontradas = _listar(drive, q, drive_id)
    if encontradas:
        return encontradas[0]["id"]

    nueva = drive.files().create(
        body={"name": nombre, "mimeType": MIME_FOLDER, "parents": [padre_id]},
        fields="id", supportsAllDrives=True,
    ).execute()
    print(f"  [DRIVE] Carpeta creada: {nombre}")
    return nueva["id"]


def _buscar_archivo(drive, carpeta_id, drive_id, nombre):
    nombre_q = nombre.replace("'", "\\'")
    q = f"name = '{nombre_q}' and '{carpeta_id}' in parents and trashed = false"
    encontrados = _listar(drive, q, drive_id)
    return encontrados[0]["id"] if encontrados else None


def nombre_archivo_ppt(campana, semana, anio):
    """Ej: 'Verano_2026_SEMANA_33-2026.pptx'."""
    return PATRON_NOMBRE_PPT.format(
        campana=san(campana), semana=semana, anio=anio
    )


def preparar_carpeta_semana(etiqueta_semana):
    """
    Crea (o encuentra) la carpeta de la semana.
    Devuelve (carpeta_id, drive_id, link_carpeta).
    """
    raiz = carpeta_raiz_id()
    if not raiz:
        raise RuntimeError(
            "Falta el ID de la Unidad Compartida de Ingresos "
            "(DRIVE_CARPETA_INGRESOS_ID o DRIVE_CARPETA_INGRESOS_ID_DEFAULT)."
        )
    drive = get_drive()
    drive_id = _drive_id_de(drive, raiz)
    if not drive_id:
        print("  [AVISO] La carpeta raíz no es una Unidad Compartida. "
              "Si es 'Mi unidad', la cuenta de servicio va a fallar por cuota "
              "al crear archivos.")
    carpeta_id = _buscar_o_crear_carpeta(drive, raiz, drive_id, etiqueta_semana)
    link = f"https://drive.google.com/drive/folders/{carpeta_id}"
    return carpeta_id, drive_id, link


def subir_ppt_semana(carpeta_id, drive_id, campana, semana, anio, ppt_bytes):
    """
    Sube o actualiza el PPT de una campaña dentro de la carpeta de la semana.
    Devuelve (nombre_archivo, link).
    """
    drive = get_drive()
    nombre = nombre_archivo_ppt(campana, semana, anio)
    media = MediaIoBaseUpload(BytesIO(ppt_bytes), mimetype=MIME_PPTX, resumable=False)

    existente = _buscar_archivo(drive, carpeta_id, drive_id, nombre)
    if existente:
        archivo = drive.files().update(
            fileId=existente, media_body=media,
            fields="id, webViewLink", supportsAllDrives=True,
        ).execute()
        print(f"        [DRIVE] Actualizado: {nombre}")
    else:
        archivo = drive.files().create(
            body={"name": nombre, "parents": [carpeta_id]},
            media_body=media, fields="id, webViewLink",
            supportsAllDrives=True,
        ).execute()
        # Permiso público con link, solo la primera vez
        try:
            drive.permissions().create(
                fileId=archivo["id"],
                body={"type": "anyone", "role": "reader"},
                supportsAllDrives=True,
            ).execute()
        except Exception as e:
            print(f"        [WARN] No se pudo dar permiso público: {e}")
        print(f"        [DRIVE] Creado: {nombre}")

    link = archivo.get(
        "webViewLink", f"https://drive.google.com/file/d/{archivo['id']}/view"
    )
    return nombre, link
