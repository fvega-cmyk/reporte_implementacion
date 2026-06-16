"""
Sube/actualiza los archivos (Excel + PPT) en una estructura permanente en Drive:

    [Unidad Compartida]
      └── Cliente
            └── Campaña
                  ├── Reporte_Campaña.xlsx
                  └── Fotos_Campaña.pptx

Estrategia de REEMPLAZO DE CONTENIDO:
- Si el archivo ya existe (mismo nombre en la misma carpeta), se actualiza su
  contenido con files().update() conservando el MISMO file ID y el MISMO link.
- Si no existe, se crea con files().create() y se le pone permiso público.

Esto garantiza que el link de cada archivo sea PERMANENTE: se manda una vez
y siempre apunta a la última versión.

Variable de entorno requerida:
  - DRIVE_CARPETA_REPORTES_ID: ID de la Unidad Compartida raíz.
"""
import os
from io import BytesIO

from googleapiclient.http import MediaIoBaseUpload

from utils import san
from google_clients import get_drive

MIME_PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MIME_FOLDER = "application/vnd.google-apps.folder"


def _get_drive_id(drive, padre_id):
    """Obtiene el driveId raíz de la Unidad Compartida."""
    info = drive.files().get(
        fileId=padre_id,
        fields="id, driveId",
        supportsAllDrives=True,
    ).execute()
    return info.get("driveId", padre_id)


def _buscar_o_crear_carpeta(drive, padre_id, drive_id, nombre):
    """
    Busca una subcarpeta por nombre dentro de padre_id. Si no existe, la crea.
    Retorna el ID de la subcarpeta.
    """
    # Escapar comillas simples en el nombre para la query
    nombre_q = nombre.replace("'", "\\'")
    q = (
        f"name = '{nombre_q}' and mimeType = '{MIME_FOLDER}' "
        f"and '{padre_id}' in parents and trashed = false"
    )
    resp = drive.files().list(
        q=q,
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        corpora="drive",
        driveId=drive_id,
    ).execute()
    archivos = resp.get("files", [])
    if archivos:
        return archivos[0]["id"]
    body = {
        "name": nombre,
        "mimeType": MIME_FOLDER,
        "parents": [padre_id],
    }
    nueva = drive.files().create(
        body=body, fields="id", supportsAllDrives=True,
    ).execute()
    return nueva["id"]


def _buscar_archivo(drive, carpeta_id, drive_id, nombre):
    """Busca un archivo por nombre exacto dentro de una carpeta. Retorna el ID o None."""
    nombre_q = nombre.replace("'", "\\'")
    q = f"name = '{nombre_q}' and '{carpeta_id}' in parents and trashed = false"
    resp = drive.files().list(
        q=q,
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        corpora="drive",
        driveId=drive_id,
    ).execute()
    archivos = resp.get("files", [])
    return archivos[0]["id"] if archivos else None


def _subir_o_actualizar(drive, carpeta_id, drive_id, nombre, contenido_bytes, mime):
    """
    Si el archivo existe → actualiza su contenido (mismo file ID, mismo link).
    Si no existe → lo crea y le pone permiso 'cualquiera con el link puede ver'.
    Retorna (file_id, web_view_link).
    """
    media = MediaIoBaseUpload(BytesIO(contenido_bytes), mimetype=mime, resumable=False)
    existente_id = _buscar_archivo(drive, carpeta_id, drive_id, nombre)

    if existente_id:
        # Actualizar contenido conservando el ID y el link
        archivo = drive.files().update(
            fileId=existente_id,
            media_body=media,
            fields="id, webViewLink",
            supportsAllDrives=True,
        ).execute()
        return archivo["id"], archivo.get("webViewLink", f"https://drive.google.com/file/d/{existente_id}/view")
    else:
        # Crear nuevo
        metadata = {"name": nombre, "parents": [carpeta_id]}
        archivo = drive.files().create(
            body=metadata,
            media_body=media,
            fields="id, webViewLink",
            supportsAllDrives=True,
        ).execute()
        file_id = archivo["id"]
        # Permiso público con link (solo la primera vez)
        drive.permissions().create(
            fileId=file_id,
            body={"type": "anyone", "role": "reader"},
            supportsAllDrives=True,
        ).execute()
        return file_id, archivo.get("webViewLink", f"https://drive.google.com/file/d/{file_id}/view")


def subir_reportes(cliente, campana, excel_bytes, ppt_bytes):
    """
    Sube/actualiza Excel + PPT en [Unidad]/Cliente/Campaña/.
    Retorna dict con los links: {"excel": url, "ppt": url, "carpeta": url}
    """
    padre_id = os.environ.get("DRIVE_CARPETA_REPORTES_ID")
    if not padre_id:
        raise RuntimeError(
            "Falta DRIVE_CARPETA_REPORTES_ID. Cargá el ID de la Unidad Compartida en GitHub."
        )

    drive = get_drive()
    drive_id = _get_drive_id(drive, padre_id)

    # Navegar/crear estructura Cliente → Campaña
    cliente_limpio = (cliente or "SIN CLIENTE").strip()
    campana_limpia = (campana or "SIN CAMPAÑA").strip()

    carpeta_cliente = _buscar_o_crear_carpeta(drive, padre_id, drive_id, cliente_limpio)
    carpeta_campana = _buscar_o_crear_carpeta(drive, carpeta_cliente, drive_id, campana_limpia)

    nombre_excel = f"Reporte_{san(campana)}.xlsx"
    nombre_ppt = f"Fotos_{san(campana)}.pptx"

    _, link_excel = _subir_o_actualizar(
        drive, carpeta_campana, drive_id, nombre_excel, excel_bytes, MIME_XLSX
    )
    _, link_ppt = _subir_o_actualizar(
        drive, carpeta_campana, drive_id, nombre_ppt, ppt_bytes, MIME_PPTX
    )

    # Link a la carpeta de la campaña
    link_carpeta = f"https://drive.google.com/drive/folders/{carpeta_campana}"

    return {"excel": link_excel, "ppt": link_ppt, "carpeta": link_carpeta}
