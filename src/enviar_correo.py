"""
Envía el correo con Excel + PPT adjuntos.
Usa SMTP de Gmail con contraseña de aplicación (16 caracteres).
Variables de entorno requeridas:
  - GMAIL_USER:     tu_correo@gmail.com (o el corporativo)
  - GMAIL_APP_PASS: la contraseña de aplicación
"""
import os
import smtplib
from email.message import EmailMessage

from config import SMTP_HOST, SMTP_PORT, ASUNTO_EMAIL, EMAIL_DESTINATARIO
from utils import san


def enviar_email(campana, hoy, excel_bytes, ppt_bytes, destinatario=None):
    fecha_str = hoy.strftime("%d/%m/%Y")
    fecha_archivo = hoy.strftime("%Y%m%d")
    destinatario = destinatario or EMAIL_DESTINATARIO

    gmail_user = os.environ.get("GMAIL_USER")
    gmail_pass = os.environ.get("GMAIL_APP_PASS")
    if not gmail_user or not gmail_pass:
        raise RuntimeError("Faltan variables GMAIL_USER / GMAIL_APP_PASS")

    msg = EmailMessage()
    msg["Subject"] = f"{ASUNTO_EMAIL} {campana} | {fecha_str}"
    msg["From"] = f"Reporte Implementación <{gmail_user}>"
    msg["To"] = destinatario
    msg.set_content(
        f"Estimado equipo,\n\n"
        f"Adjunto el reporte diario de implementación del {fecha_str}.\n\n"
        f"Campaña: {campana}\n\n"
        f"Se incluyen:\n"
        f"- Excel con resumen y detalle\n"
        f"- PPT con fotos\n\n"
        f"Saludos,\nSistema de Reportes"
    )

    nombre_campana = san(campana)
    msg.add_attachment(
        excel_bytes,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"Reporte_{nombre_campana}_{fecha_archivo}.xlsx",
    )
    msg.add_attachment(
        ppt_bytes,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=f"Fotos_{nombre_campana}_{fecha_archivo}.pptx",
    )

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
        s.starttls()
        s.login(gmail_user, gmail_pass)
        s.send_message(msg)
    print(f"  [EMAIL OK] enviado a {destinatario}")
