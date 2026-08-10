"""
Envía el correo informativo en formato HTML. Dos tipos:

  - "interno": diario. Va a correos fijos + correos del cliente (todos en To).
               Lleva la firma de correo automático interno.
  - "externo": semanal (lunes). Va al fijo externo (To) + clientes en copia
               visible (CC). Sin firma interna. Solo para clientes con config
               externa. Informa el período martes anterior → lunes actual.

Los archivos YA fueron subidos/actualizados en Drive por subir_a_drive.py.
- Si el peso total (Excel + PPT) es < UMBRAL → adjunta ambos.
- Si es >= UMBRAL → adjunta solo el Excel (si entra) y deja el PPT por link.
- Siempre incluye los links a Drive (carpeta, Excel, PPT).

Variables de entorno:
  - GMAIL_USER:     email del remitente
  - GMAIL_APP_PASS: contraseña de aplicación
"""
import os
import smtplib
from email.message import EmailMessage

from config import (
    SMTP_HOST, SMTP_PORT, ASUNTO_EMAIL, EMAIL_DESTINATARIO,
    CORREOS_FIJOS, CORREOS_POR_CLIENTE, CORREOS_IGNORAR,
    CORREOS_FIJOS_EXTERNO, CORREOS_POR_CLIENTE_EXTERNO,
)
from utils import san

# Gmail acepta 25 MB. Margen a 22 MB.
UMBRAL_BYTES = 22 * 1024 * 1024
UMBRAL_EXCEL = 20 * 1024 * 1024

# Paleta Sell Out
AZUL = "#1F3864"
GRIS_TEXTO = "#333333"
GRIS_SUAVE = "#666666"


def _dedupe(correos):
    """Quita duplicados (case-insensitive), ignorados y vacíos. Conserva orden."""
    vistos = set()
    final = []
    for correo in correos:
        c = (correo or "").strip().lower()
        if not c or c in CORREOS_IGNORAR or c in vistos:
            continue
        vistos.add(c)
        final.append(correo.strip())
    return final


def _parsear_destinatarios(valor):
    """
    Acepta None, lista, o string "a@x.cl, b@y.cl; c@z.cl".
    Devuelve lista limpia y deduplicada (vacía si no hay nada).
    """
    if not valor:
        return []
    if isinstance(valor, str):
        crudos = valor.replace(";", ",").split(",")
    else:
        crudos = list(valor)
    return _dedupe([c for c in (x.strip() for x in crudos) if c])


def _construir_html(campana, fecha_str, links, mostrar_firma=True, periodo_str=None):
    """Arma el cuerpo HTML del correo."""
    firma = ""
    if mostrar_firma:
        firma = f"""
      <hr style="border:none;border-top:1px solid #e3e8ef;margin:24px 0 12px;">
      <p style="margin:0;font-size:12px;color:{GRIS_SUAVE};">
        Este es un correo automático del Sistema de Reportes de Sell Out.
      </p>"""

    if periodo_str:
        titulo_bloque = "Reporte Semanal de Implementación"
        parrafo_intro = (
            f"Adjunto reportes de implementación de la campaña "
            f"<strong>{campana}</strong> correspondientes al período "
            f"<strong>{periodo_str}</strong>."
        )
    else:
        titulo_bloque = "Reporte Diario de Implementación"
        parrafo_intro = (
            f"Adjunto reportes de implementación de la campaña "
            f"<strong>{campana}</strong> actualizado al <strong>{fecha_str}</strong>."
        )

    return f"""\
<!DOCTYPE html>
<html lang="es">
<head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;background-color:#f4f6f9;">
  <div style="max-width:640px;margin:0 auto;padding:24px;
              font-family:'Segoe UI',Roboto,Helvetica,Arial,sans-serif;
              color:{GRIS_TEXTO};line-height:1.55;">

    <div style="background-color:{AZUL};border-radius:10px 10px 0 0;padding:22px 28px;">
      <div style="color:#ffffff;font-size:13px;letter-spacing:1px;
                  text-transform:uppercase;opacity:0.85;">
        {titulo_bloque}
      </div>
      <div style="color:#ffffff;font-size:22px;font-weight:700;margin-top:4px;">
        {campana}
      </div>
    </div>

    <div style="background-color:#ffffff;border:1px solid #e3e8ef;border-top:none;
                border-radius:0 0 10px 10px;padding:28px;">

      <p style="margin:0 0 14px;">Estimado equipo,</p>

      <p style="margin:0 0 18px;">
        {parrafo_intro}
      </p>

      <p style="margin:0 0 8px;font-weight:600;color:{AZUL};">Se incluyen:</p>
      <ol style="margin:0 0 20px;padding-left:22px;">
        <li style="margin-bottom:4px;">Excel con resumen y detalle.</li>
        <li>PPT con fotos.</li>
      </ol>

      <p style="margin:0 0 12px;">
        Los reportes están disponibles y actualizados en los siguientes links:
      </p>

      <table cellpadding="0" cellspacing="0" style="width:100%;margin:0 0 20px;">
        <tr><td style="padding:8px 0;">
          <a href="{links['carpeta']}"
             style="display:inline-block;background-color:{AZUL};color:#ffffff;
                    text-decoration:none;padding:10px 18px;border-radius:6px;
                    font-weight:600;font-size:14px;">📁 CARPETA {campana}</a>
        </td></tr>
        <tr><td style="padding:8px 0;">
          <a href="{links['excel']}"
             style="display:inline-block;background-color:#107C41;color:#ffffff;
                    text-decoration:none;padding:10px 18px;border-radius:6px;
                    font-weight:600;font-size:14px;">📊 EXCEL {campana}</a>
        </td></tr>
        <tr><td style="padding:8px 0;">
          <a href="{links['ppt']}"
             style="display:inline-block;background-color:#C43E1C;color:#ffffff;
                    text-decoration:none;padding:10px 18px;border-radius:6px;
                    font-weight:600;font-size:14px;">🖼️ FOTOS {campana}</a>
        </td></tr>
      </table>

      <p style="margin:0;">Saludos.</p>
{firma}
    </div>
  </div>
</body>
</html>"""


def _texto_plano(campana, fecha_str, links, periodo_str=None):
    if periodo_str:
        intro = (f"Adjunto reportes de implementación de la campaña {campana} "
                 f"correspondientes al período {periodo_str}.")
    else:
        intro = (f"Adjunto reportes de implementación de la campaña {campana} "
                 f"actualizado al {fecha_str}.")
    return (
        f"Estimado equipo,\n\n"
        f"{intro}\n\n"
        f"Se incluyen:\n"
        f"1. Excel con resumen y detalle.\n"
        f"2. PPT con fotos.\n\n"
        f"Los reportes están disponibles y actualizados en los siguientes links:\n"
        f"- CARPETA {campana}: {links['carpeta']}\n"
        f"- EXCEL {campana}: {links['excel']}\n"
        f"- FOTOS {campana}: {links['ppt']}\n\n"
        f"Saludos."
    )


def enviar_email(campana, hoy, excel_bytes, ppt_bytes, links, cliente=None,
                 tipo="interno", periodo=None,
                 solo_cliente=False, destinatarios=None):
    """
    tipo:    "interno" (diario) o "externo" (semanal, a clientes en CC).
    periodo: None o tupla (date_desde, date_hasta). Si viene, el correo habla
             de un PERÍODO (reporte semanal) en vez de una fecha puntual.

    Overrides para REENVÍOS manuales:
      solo_cliente:  True  → omite CORREOS_FIJOS y manda solo a los correos
                             configurados para ese cliente. Sirve para no
                             volver a spamear al equipo interno en un reenvío.
      destinatarios: lista o string "a@x.cl, b@y.cl" → manda EXACTAMENTE a esos
                     correos, ignorando toda la configuración. Tiene prioridad
                     sobre solo_cliente.

    Devuelve True si se envió, False si se omitió.
    """
    fecha_str = hoy.strftime("%d/%m/%Y")
    periodo_str = None
    if periodo:
        d, h = periodo
        periodo_str = f"{d.strftime('%d/%m/%Y')} al {h.strftime('%d/%m/%Y')}"
    fecha_archivo = hoy.strftime("%Y%m%d")
    cliente_limpio = (cliente or "").strip()

    gmail_user = os.environ.get("GMAIL_USER")
    gmail_pass = os.environ.get("GMAIL_APP_PASS")
    if not gmail_user or not gmail_pass:
        raise RuntimeError("Faltan variables GMAIL_USER / GMAIL_APP_PASS")

    # --- Resolver destinatarios ---
    cc = []
    override = _parsear_destinatarios(destinatarios)

    if override:
        # Override total: manda exactamente a estos correos, sin CC.
        to = override
        mostrar_firma = (tipo != "externo")
        print(f"        [OVERRIDE] destinatarios manuales: {', '.join(to)}")

    elif tipo == "externo":
        clientes_ext = _dedupe(CORREOS_POR_CLIENTE_EXTERNO.get(cliente_limpio, []))
        if not clientes_ext:
            print(f"        [EXTERNO] '{cliente_limpio}' sin destinatarios externos → se omite")
            return False
        if solo_cliente:
            # Solo el cliente, sin el fijo interno de respaldo.
            to = clientes_ext
        else:
            to = _dedupe(CORREOS_FIJOS_EXTERNO)
            cc = clientes_ext
        mostrar_firma = False

    else:
        del_cliente = _dedupe(CORREOS_POR_CLIENTE.get(cliente_limpio, []))
        if solo_cliente:
            # Omite los CORREOS_FIJOS: útil en reenvíos, para no volver a
            # notificar al equipo interno que ya recibió el correo.
            to = del_cliente
            if not to:
                print(f"        [INTERNO] '{cliente_limpio}' no tiene correos propios "
                      f"configurados y solo_cliente está activo → se omite")
                return False
        else:
            to = _dedupe(list(CORREOS_FIJOS) + del_cliente)
            if not to:
                to = [EMAIL_DESTINATARIO]
        mostrar_firma = True

    print(f"        [{tipo.upper()}] To: {', '.join(to)}" + (f" | CC: {', '.join(cc)}" if cc else ""))

    nombre_campana = san(campana)
    nombre_excel = f"Reporte_{nombre_campana}_{fecha_archivo}.xlsx"
    nombre_ppt = f"Fotos_{nombre_campana}_{fecha_archivo}.pptx"

    peso_total = len(excel_bytes) + len(ppt_bytes)
    peso_mb = peso_total / 1024 / 1024
    adjuntar_todo = peso_total < UMBRAL_BYTES

    msg = EmailMessage()
    if periodo_str:
        msg["Subject"] = f"Reporte Semanal Implementacion - {campana} | {periodo_str}"
    else:
        msg["Subject"] = f"{ASUNTO_EMAIL} {campana} | {fecha_str}"
    msg["From"] = f"Reporte Implementación <{gmail_user}>"
    msg["To"] = ", ".join(to)
    if cc:
        msg["Cc"] = ", ".join(cc)

    msg.set_content(_texto_plano(campana, fecha_str, links, periodo_str))
    msg.add_alternative(
        _construir_html(campana, fecha_str, links, mostrar_firma, periodo_str),
        subtype="html",
    )

    if adjuntar_todo:
        print(f"        Peso total: {peso_mb:.1f} MB → ADJUNTA ambos")
        msg.add_attachment(excel_bytes, maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=nombre_excel)
        msg.add_attachment(ppt_bytes, maintype="application",
            subtype="vnd.openxmlformats-officedocument.presentationml.presentation", filename=nombre_ppt)
    else:
        print(f"        Peso total: {peso_mb:.1f} MB → ADJUNTA solo Excel (PPT por link)")
        if len(excel_bytes) < UMBRAL_EXCEL:
            msg.add_attachment(excel_bytes, maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=nombre_excel)

    # Todos los destinatarios reales (To + CC) para el envío
    todos = to + cc
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
        s.starttls()
        s.login(gmail_user, gmail_pass)
        s.send_message(msg, to_addrs=todos)
    print(f"  [EMAIL OK] ({tipo}) enviado a {len(todos)} destinatario(s)")
    return True
