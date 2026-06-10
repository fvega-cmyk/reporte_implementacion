"""
Envía el correo con Excel + PPT adjuntos (SMTP con Gmail App Password).

Estrategia híbrida según peso del PPT:
- Si peso total estimado < UMBRAL_MB → adjunta ambos archivos al correo.
- Si peso total estimado >= UMBRAL_MB → sube el PPT a Drive (público con link)
  y manda un link en el cuerpo.

Variables de entorno:
  - GMAIL_USER:                  email del remitente (ej. tunombre@sell-out.cl)
  - GMAIL_APP_PASS:              contraseña de aplicación de 16 caracteres
  - DRIVE_CARPETA_REPORTES_ID:   ID de la carpeta padre en TU Drive donde se
                                 subirán los PPT grandes (compartida con la
                                 cuenta de servicio como Editor)
"""
import os
import smtplib
from email.message import EmailMessage
from io import BytesIO

from googleapiclient.http import MediaIoBaseUpload

from config import SMTP_HOST, SMTP_PORT, ASUNTO_EMAIL, EMAIL_DESTINATARIO
from utils import san
from google_clients import get_drive


# Umbral en bytes. Gmail acepta 25 MB. Usamos 22 MB para dejar margen
# (overhead de base64 codifica ~33% más, headers, multipart, etc).
UMBRAL_BYTES = 22 * 1024 * 1024

# Constantes Drive
MIME_PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
MIME_FOLDER = "application/vnd.google-apps.folder"


def _crear_carpeta_dia(drive, padre_id, nombre):
    """
    Crea (o reutiliza) una subcarpeta dentro de la carpeta padre dada por ID.
    Retorna el ID de la subcarpeta.
    """
    # Buscar si ya existe la subcarpeta del día
    q = (
        f"name = '{nombre}' and mimeType = '{MIME_FOLDER}' "
        f"and '{padre_id}' in parents and trashed = false"
    )
    resp = drive.files().list(
        q=q, fields="files(id, name)",
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    archivos = resp.get("files", [])
    if archivos:
        return archivos[0]["id"]
    # Crearla dentro del padre
    body = {
        "name": nombre,
        "mimeType": MIME_FOLDER,
        "parents": [padre_id],
    }
    nueva = drive.files().create(
        body=body, fields="id", supportsAllDrives=True,
    ).execute()
    return nueva["id"]


def _subir_ppt_y_obtener_link(drive, ppt_bytes, nombre_archivo, hoy):
    """
    Sube el PPT a Drive dentro de:
      [carpeta padre cuyo ID viene en DRIVE_CARPETA_REPORTES_ID] / Reporte YYYYMMDD /
    Le pone permiso 'cualquiera con el link puede ver'.
    Retorna la URL pública.
    """
    padre_id = os.environ.get("DRIVE_CARPETA_REPORTES_ID")
    if not padre_id:
        raise RuntimeError(
            "Falta la variable DRIVE_CARPETA_REPORTES_ID. "
            "Cargá el ID de la carpeta 'Reportes Diarios' como secreto en GitHub."
        )

    nombre_dia = f"Reporte {hoy.strftime('%Y%m%d')}"
    carpeta_dia_id = _crear_carpeta_dia(drive, padre_id, nombre_dia)

    # Subir el archivo
    media = MediaIoBaseUpload(
        BytesIO(ppt_bytes), mimetype=MIME_PPTX, resumable=True,
    )
    metadata = {
        "name": nombre_archivo,
        "parents": [carpeta_dia_id],
    }
    archivo = drive.files().create(
        body=metadata, media_body=media,
        fields="id, webViewLink",
        supportsAllDrives=True,
    ).execute()
    file_id = archivo["id"]

    # Darle permiso público con link
    drive.permissions().create(
        fileId=file_id,
        body={"type": "anyone", "role": "reader"},
        supportsAllDrives=True,
    ).execute()

    return archivo.get("webViewLink", f"https://drive.google.com/file/d/{file_id}/view")


def enviar_email(campana, hoy, excel_bytes, ppt_bytes, destinatario=None):
    fecha_str = hoy.strftime("%d/%m/%Y")
    fecha_archivo = hoy.strftime("%Y%m%d")
    destinatario = destinatario or EMAIL_DESTINATARIO

    gmail_user = os.environ.get("GMAIL_USER")
    gmail_pass = os.environ.get("GMAIL_APP_PASS")
    if not gmail_user or not gmail_pass:
        raise RuntimeError("Faltan variables GMAIL_USER / GMAIL_APP_PASS")

    nombre_campana = san(campana)
    nombre_excel = f"Reporte_{nombre_campana}_{fecha_archivo}.xlsx"
    nombre_ppt = f"Fotos_{nombre_campana}_{fecha_archivo}.pptx"

    # Decidir si va como adjunto o como link
    peso_total = len(excel_bytes) + len(ppt_bytes)
    peso_mb = peso_total / 1024 / 1024
    usar_link = peso_total >= UMBRAL_BYTES

    msg = EmailMessage()
    msg["Subject"] = f"{ASUNTO_EMAIL} {campana} | {fecha_str}"
    msg["From"] = f"Reporte Implementación <{gmail_user}>"
    msg["To"] = destinatario

    if usar_link:
        print(f"        Peso total: {peso_mb:.1f} MB → modo LINK (subiendo PPT a Drive)")
        drive = get_drive()
        link_ppt = _subir_ppt_y_obtener_link(drive, ppt_bytes, nombre_ppt, hoy)
        print(f"        PPT subido: {link_ppt}")

        cuerpo = (
            f"Estimado equipo,\n\n"
            f"Adjunto el reporte diario de implementación del {fecha_str}.\n\n"
            f"Campaña: {campana}\n\n"
            f"Se incluyen:\n"
            f"- Excel con resumen y detalle (adjunto)\n"
            f"- PPT con fotos (descargar desde el siguiente link, por tamaño):\n"
            f"  {link_ppt}\n\n"
            f"Saludos,\nSistema de Reportes"
        )
        msg.set_content(cuerpo)
        msg.add_attachment(
            excel_bytes,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=nombre_excel,
        )
    else:
        print(f"        Peso total: {peso_mb:.1f} MB → modo ADJUNTO")
        cuerpo = (
            f"Estimado equipo,\n\n"
            f"Adjunto el reporte diario de implementación del {fecha_str}.\n\n"
            f"Campaña: {campana}\n\n"
            f"Se incluyen:\n"
            f"- Excel con resumen y detalle\n"
            f"- PPT con fotos\n\n"
            f"Saludos,\nSistema de Reportes"
        )
        msg.set_content(cuerpo)
        msg.add_attachment(
            excel_bytes,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=nombre_excel,
        )
        msg.add_attachment(
            ppt_bytes,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.presentationml.presentation",
            filename=nombre_ppt,
        )

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
        s.starttls()
        s.login(gmail_user, gmail_pass)
        s.send_message(msg)
    print(f"  [EMAIL OK] enviado a {destinatario}")
