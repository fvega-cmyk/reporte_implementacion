"""
Envía UN SOLO correo semanal con todos los PPT de Ingresos de la semana.

A diferencia del flujo de Implementación (un correo por campaña), acá se
consolida todo: un correo, N adjuntos, uno por campaña.

Manejo de peso: Gmail corta en 25 MB. Se adjunta en orden hasta llegar al
umbral; los PPT que no entran quedan solo con su link a Drive y el correo lo
dice explícitamente, para que nadie asuma que falta información.

Variables de entorno: GMAIL_USER, GMAIL_APP_PASS
"""
import os
import smtplib
from email.message import EmailMessage

from config import SMTP_HOST, SMTP_PORT, CORREOS_IGNORAR
from config_ingresos import (
    ASUNTO_INGRESOS, CORREOS_INGRESOS, CORREOS_INGRESOS_CC,
)

UMBRAL_BYTES = 22 * 1024 * 1024

AZUL = "#1F3864"
GRIS_TEXTO = "#333333"
GRIS_SUAVE = "#666666"
NARANJA = "#C43E1C"


def _dedupe(correos):
    vistos, final = set(), []
    for correo in correos or []:
        c = (correo or "").strip().lower()
        if not c or c in CORREOS_IGNORAR or c in vistos:
            continue
        vistos.add(c)
        final.append(correo.strip())
    return final


def _parsear_destinatarios(valor):
    if not valor:
        return []
    crudos = valor.replace(";", ",").split(",") if isinstance(valor, str) else list(valor)
    return _dedupe([c.strip() for c in crudos if c.strip()])


def _fila_boton(texto, link, color):
    return f"""
        <tr><td style="padding:6px 0;">
          <a href="{link}"
             style="display:inline-block;background-color:{color};color:#ffffff;
                    text-decoration:none;padding:10px 18px;border-radius:6px;
                    font-weight:600;font-size:14px;">{texto}</a>
        </td></tr>"""


def _construir_html(etiqueta_semana, periodo_str, reportes, link_carpeta,
                    solo_link):
    filas_botones = _fila_boton(f"📁 CARPETA {etiqueta_semana}", link_carpeta, AZUL)
    for r in reportes:
        filas_botones += _fila_boton(f"🖼️ {r['campana']}", r["link"], NARANJA)

    aviso = ""
    if solo_link:
        nombres = ", ".join(f"<strong>{c}</strong>" for c in solo_link)
        aviso = f"""
      <p style="margin:0 0 18px;padding:12px 14px;background-color:#fff6e5;
                border-left:4px solid #d98a00;font-size:13px;">
        Por límite de tamaño de correo, {nombres} no viene como adjunto.
        Está completo en su link más abajo.
      </p>"""

    lista = "".join(
        f'<li style="margin-bottom:4px;">{r["campana"]} '
        f'<span style="color:{GRIS_SUAVE};">({r["registros"]} registros)</span></li>'
        for r in reportes
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
        Reporte Semanal de Ingresos
      </div>
      <div style="color:#ffffff;font-size:22px;font-weight:700;margin-top:4px;">
        {etiqueta_semana}
      </div>
    </div>

    <div style="background-color:#ffffff;border:1px solid #e3e8ef;border-top:none;
                border-radius:0 0 10px 10px;padding:28px;">

      <p style="margin:0 0 14px;">Estimado equipo,</p>

      <p style="margin:0 0 18px;">
        Adjunto los reportes de ingresos correspondientes a
        <strong>{etiqueta_semana}</strong> ({periodo_str}).
      </p>

      <p style="margin:0 0 8px;font-weight:600;color:{AZUL};">
        Campañas incluidas ({len(reportes)}):
      </p>
      <ol style="margin:0 0 20px;padding-left:22px;">{lista}</ol>
{aviso}
      <p style="margin:0 0 12px;">
        Todos los archivos quedan guardados en la carpeta de la semana:
      </p>

      <table cellpadding="0" cellspacing="0" style="width:100%;margin:0 0 20px;">{filas_botones}
      </table>

      <p style="margin:0;">Saludos.</p>

      <hr style="border:none;border-top:1px solid #e3e8ef;margin:24px 0 12px;">
      <p style="margin:0;font-size:12px;color:{GRIS_SUAVE};">
        Este es un correo automático del Sistema de Reportes de Sell Out.
      </p>
    </div>
  </div>
</body>
</html>"""


def _texto_plano(etiqueta_semana, periodo_str, reportes, link_carpeta, solo_link):
    lineas = [
        "Estimado equipo,",
        "",
        f"Adjunto los reportes de ingresos correspondientes a {etiqueta_semana} "
        f"({periodo_str}).",
        "",
        f"Campañas incluidas ({len(reportes)}):",
    ]
    for i, r in enumerate(reportes, 1):
        lineas.append(f"{i}. {r['campana']} ({r['registros']} registros)")
    if solo_link:
        lineas += ["", "Por límite de tamaño, estos no van adjuntos (ver link): "
                   + ", ".join(solo_link)]
    lineas += ["", f"CARPETA {etiqueta_semana}: {link_carpeta}", ""]
    for r in reportes:
        lineas.append(f"- {r['campana']}: {r['link']}")
    lineas += ["", "Saludos."]
    return "\n".join(lineas)


def enviar_email_ingresos(etiqueta_semana, lunes, domingo, reportes,
                          link_carpeta, destinatarios=None):
    """
    Manda el correo único de la semana.

    reportes: lista de dicts, uno por campaña:
        {"campana": str, "nombre": str, "link": str,
         "bytes": bytes, "registros": int}

    destinatarios: opcional. Lista o "a@x.cl, b@y.cl" → sobrescribe la config.

    Devuelve True si se envió.
    """
    gmail_user = os.environ.get("GMAIL_USER")
    gmail_pass = os.environ.get("GMAIL_APP_PASS")
    if not gmail_user or not gmail_pass:
        raise RuntimeError("Faltan variables GMAIL_USER / GMAIL_APP_PASS")

    if not reportes:
        print("  [EMAIL] Sin reportes en la semana → no se envía correo.")
        return False

    override = _parsear_destinatarios(destinatarios)
    if override:
        to, cc = override, []
        print(f"  [OVERRIDE] destinatarios manuales: {', '.join(to)}")
    else:
        to = _dedupe(CORREOS_INGRESOS)
        cc = _dedupe(CORREOS_INGRESOS_CC)
        if not to:
            raise RuntimeError("CORREOS_INGRESOS está vacío en config_ingresos.py")

    periodo_str = f"{lunes.strftime('%d/%m/%Y')} al {domingo.strftime('%d/%m/%Y')}"

    # --- Decidir qué adjuntar según el peso acumulado ---
    reportes = sorted(reportes, key=lambda r: r["campana"].lower())
    adjuntar, solo_link, acumulado = [], [], 0
    for r in reportes:
        peso = len(r["bytes"])
        if acumulado + peso < UMBRAL_BYTES:
            adjuntar.append(r)
            acumulado += peso
        else:
            solo_link.append(r["campana"])

    print(f"  [EMAIL] {len(adjuntar)} adjunto(s), {acumulado/1024/1024:.1f} MB total"
          + (f" | solo link: {', '.join(solo_link)}" if solo_link else ""))

    msg = EmailMessage()
    msg["Subject"] = f"{ASUNTO_INGRESOS} - {etiqueta_semana}"
    msg["From"] = f"Reporte Ingresos <{gmail_user}>"
    msg["To"] = ", ".join(to)
    if cc:
        msg["Cc"] = ", ".join(cc)

    msg.set_content(_texto_plano(etiqueta_semana, periodo_str, reportes,
                                 link_carpeta, solo_link))
    msg.add_alternative(_construir_html(etiqueta_semana, periodo_str, reportes,
                                        link_carpeta, solo_link), subtype="html")

    for r in adjuntar:
        msg.add_attachment(
            r["bytes"], maintype="application",
            subtype="vnd.openxmlformats-officedocument.presentationml.presentation",
            filename=r["nombre"],
        )

    todos = to + cc
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
        s.starttls()
        s.login(gmail_user, gmail_pass)
        s.send_message(msg, to_addrs=todos)
    print(f"  [EMAIL OK] enviado a {len(todos)} destinatario(s)")
    return True
