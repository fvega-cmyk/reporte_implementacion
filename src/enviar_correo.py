"""
Envía el correo con Excel + PPT adjuntos (SMTP con Gmail App Password).

Estrategia híbrida según peso del PPT:
- Si peso total estimado < UMBRAL_MB → adjunta ambos archivos al correo.
- Si peso total estimado >= UMBRAL_MB → sube el PPT a una Unidad Compartida
  (Shared Drive) y manda un link en el cuerpo.

Variables de entorno:
  - GMAIL_USER:                  email del remitente (ej. tunombre@sell-out.cl)
  - GMAIL_APP_PASS:              contraseña de aplicación de 16 caracteres
  - DRIVE_CARPETA_REPORTES_ID:   ID de la UNIDAD COMPARTIDA donde se subirán
                                 los PPT grandes. La cuenta de servicio debe
                                 tener rol "Administrador de contenido" en ella.
"""
import os
import smtplib
from email.message import EmailMessage
from io import BytesIO

from googleapiclient.http import MediaIoBaseUpload

from config import SMTP_HOST, SMTP_PORT, ASUNTO_EMAIL, EMAIL_DESTINATARIO
from utils import san
from google_clients import get_drive


# Umbral en bytes. Gmail acepta 25 MB. Usamos 22 MB para dejar margen.
UMBRAL_BYTES = 22 * 1024 * 1024

MIME_PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
MIME_FOLDER = "application/vnd.google-apps.folder"


def _obtener_drive_id_raiz(drive, padre_id):
    """
    Dado un folder en una Unidad Compartida, obtiene el driveId raíz de esa unidad.
    Es necesario para que las requests funcionen con Shared Drives.
    """
    info = drive.files().get(
        fileId=padre_id,
        fields="id, name, driveId, mimeType",
        supportsAllDrives=True,
    ).execute()
    # Si el padre ES la raíz de la Shared Drive, su 'id' ES el driveId
    # Si es una subcarpeta dentro de la Shared Drive, viene en 'driveId'
    return info.get("driveId", padre_id)


def _crear_carpeta_dia(drive, padre_id, drive_id, nombre):
    """
    Crea (o reutiliza) una subcarpeta dentro de la unidad compartida.
    """
    q = (
        f"name = '{nombre}' and mimeType = '{MIME_FOLDER}' "
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
        body=body,
        fields="id",
        supportsAllDrives=True,
    ).execute()
    return nueva["id"]


def _subir_ppt_y_obtener_link(drive, ppt_bytes, nombre_archivo, hoy):
    """
    Sube el PPT a la Unidad Compartida dentro de:
      [Unidad cuyo ID viene en DRIVE_CARPETA_REPORTES_ID] / Reporte YYYYMMDD /
    Le pone permiso 'cualquiera con el link puede ver'.
    Retorna la URL pública.
    """
    padre_id = os.environ.get("DRIVE_CARPETA_REPORTES_ID")
    if not padre_id:
        raise RuntimeError(
            "Falta la variable DRIVE_CARPETA_REPORTES_ID. "
            "Cargá el ID de la Unidad Compartida 'Reportes Diarios' como secreto en GitHub."
        )

    # Obtener el driveId raíz (necesario para Shared Drives)
    drive_id = _obtener_drive_id_raiz(drive, padre_id)

    nombre_dia = f"Reporte {hoy.strftime('%Y%m%d')}"
    carpeta_dia_id = _crear_carpeta_dia(drive, padre_id, drive_id, nombre_dia)

    # Subir el archivo
    media = MediaIoBaseUpload(
        BytesIO(ppt_bytes), mimetype=MIME_PPTX, resumable=False,
    )
    metadata = {
        "name": nombre_archivo,
        "parents": [carpeta_dia_id],
    }
    archivo = drive.files().create(
        body=metadata,
        media_body=media,
        fields="id, webViewLink",
        supportsAllDrives=True,
    ).execute()
    file_id = archivo["id"]

    # Permiso 'cualquiera con el link puede ver'
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
