"""
Envía el correo informativo diario.

Los archivos YA fueron subidos/actualizados en Drive por subir_a_drive.py.
Este módulo solo arma y manda el correo:
- Si el peso total (Excel + PPT) es < UMBRAL → adjunta ambos archivos.
- Si es >= UMBRAL → adjunta solo el Excel (si entra) y pone el link del PPT.
- Siempre incluye el link a la carpeta de Drive de la campaña.

Variables de entorno:
  - GMAIL_USER:     email del remitente
  - GMAIL_APP_PASS: contraseña de aplicación
"""
import os
import smtplib
from email.message import EmailMessage

from config import SMTP_HOST, SMTP_PORT, ASUNTO_EMAIL, EMAIL_DESTINATARIO
from utils import san

# Gmail acepta 25 MB. Margen a 22 MB.
UMBRAL_BYTES = 22 * 1024 * 1024
# Para el Excel solo, margen individual
UMBRAL_EXCEL = 20 * 1024 * 1024


def enviar_email(campana, hoy, excel_bytes, ppt_bytes, links, destinatario=None):
    """
    links: dict con {"excel": url, "ppt": url, "carpeta": url} de subir_a_drive.
    """
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
    adjuntar_todo = peso_total < UMBRAL_BYTES

    msg = EmailMessage()
    msg["Subject"] = f"{ASUNTO_EMAIL} {campana} | {fecha_str}"
    msg["From"] = f"Reporte Implementación <{gmail_user}>"
    msg["To"] = destinatario

    # Pie común con los links permanentes de Drive
    pie_links = (
        f"\nLos reportes también están siempre disponibles y actualizados en Drive:\n"
        f"- Carpeta de la campaña: {links['carpeta']}\n"
        f"- Excel: {links['excel']}\n"
        f"- PPT de fotos: {links['ppt']}\n"
    )

    if adjuntar_todo:
        print(f"        Peso total: {peso_mb:.1f} MB → ADJUNTA ambos")
        cuerpo = (
            f"Estimado equipo,\n\n"
            f"Adjunto el reporte diario de implementación del {fecha_str}.\n\n"
            f"Campaña: {campana}\n\n"
            f"Se incluyen:\n"
            f"- Excel con resumen y detalle\n"
            f"- PPT con fotos\n"
            f"{pie_links}\n"
            f"Saludos,\nSistema de Reportes"
        )
        msg.set_content(cuerpo)
        msg.add_attachment(
            excel_bytes, maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=nombre_excel,
        )
        msg.add_attachment(
            ppt_bytes, maintype="application",
            subtype="vnd.openxmlformats-officedocument.presentationml.presentation",
            filename=nombre_ppt,
        )
    else:
        print(f"        Peso total: {peso_mb:.1f} MB → ADJUNTA solo Excel + link PPT")
        cuerpo = (
            f"Estimado equipo,\n\n"
            f"Adjunto el reporte diario de implementación del {fecha_str}.\n\n"
            f"Campaña: {campana}\n\n"
            f"Se incluyen:\n"
            f"- Excel con resumen y detalle (adjunto)\n"
            f"- PPT con fotos (por tamaño, descargar desde el link):\n"
            f"  {links['ppt']}\n"
            f"{pie_links}\n"
            f"Saludos,\nSistema de Reportes"
        )
        msg.set_content(cuerpo)
        # Adjuntar el Excel solo si entra
        if len(excel_bytes) < UMBRAL_EXCEL:
            msg.add_attachment(
                excel_bytes, maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename=nombre_excel,
            )

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
        s.starttls()
        s.login(gmail_user, gmail_pass)
        s.send_message(msg)
    print(f"  [EMAIL OK] enviado a {destinatario}")
