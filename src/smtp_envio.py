"""
Envío SMTP compartido por los dos flujos de correo.

POR QUÉ EXISTE
--------------
Gmail cortaba la conexión durante el AUTH, cuatro intentos seguidos, con más
de un minuto y medio de espera entre el primero y el último. Eso descarta el
corte transitorio: algo de la combinación puerto + credencial + runner no
estaba funcionando.

Qué hace distinto:

1. DOS PUERTOS. Si 587 con STARTTLS falla, prueba 465 con SSL directo antes
   de darse por vencido. Son dos caminos distintos hacia el mismo servidor y
   no siempre fallan juntos: 587 negocia TLS sobre una conexión en claro, y
   hay redes que lo interfieren.

2. CREDENCIALES LIMPIAS. Google muestra las contraseñas de aplicación como
   "abcd efgh ijkl mnop", con espacios. Si esos espacios quedan guardados en
   el secret, el AUTH se arma mal. Lo mismo con el salto de línea que queda
   al copiar y pegar. Se limpian antes de usarlas.

3. DIAGNÓSTICO ÚTIL. Cuando falla, dice qué pasó en cada puerto, en vez de un
   único "Connection unexpectedly closed" sin contexto.

Variables de entorno: GMAIL_USER, GMAIL_APP_PASS
"""
import os
import ssl
import time
import smtplib

from config import SMTP_HOST

# Caminos a probar, en orden. 587 primero por compatibilidad.
PUERTOS = [(587, "starttls"), (465, "ssl")]

INTENTOS = 3                 # vueltas completas sobre la lista de puertos
ESPERA_INICIAL = 10          # segundos; se duplica en cada vuelta
TIMEOUT = 90

# Reintentar con credenciales malas puede hacer que Google bloquee la cuenta.
_NO_REINTENTAR = (
    smtplib.SMTPAuthenticationError,
    smtplib.SMTPRecipientsRefused,
)

_REINTENTABLES = (
    smtplib.SMTPServerDisconnected,
    smtplib.SMTPConnectError,
    smtplib.SMTPHeloError,
    ConnectionError,
    TimeoutError,
    OSError,
)


def credenciales():
    """
    Lee GMAIL_USER / GMAIL_APP_PASS y las limpia.

    La contraseña de aplicación de Google son 16 caracteres sin espacios,
    pero la interfaz la muestra en grupos de 4. Si quedó guardada con los
    espacios, o con un salto de línea del copiar y pegar, el AUTH falla.
    """
    usuario = (os.environ.get("GMAIL_USER") or "").strip()
    clave = (os.environ.get("GMAIL_APP_PASS") or "").strip()
    clave = clave.replace(" ", "").replace("\t", "").replace("\u00a0", "")
    if not usuario or not clave:
        raise RuntimeError("Faltan variables GMAIL_USER / GMAIL_APP_PASS")
    return usuario, clave


def _conectar(puerto, modo):
    """Abre la conexión según el modo y devuelve el objeto SMTP ya cifrado."""
    if modo == "ssl":
        contexto = ssl.create_default_context()
        return smtplib.SMTP_SSL(SMTP_HOST, puerto,
                                timeout=TIMEOUT, context=contexto)
    s = smtplib.SMTP(SMTP_HOST, puerto, timeout=TIMEOUT)
    s.ehlo()
    s.starttls(context=ssl.create_default_context())
    s.ehlo()          # el EHLO se renegocia después de TLS
    return s


def enviar_mensaje(msg, destinatarios, prefijo="  "):
    """
    Manda el mensaje probando cada puerto y reintentando.

    Devuelve (puerto, modo) del camino que funcionó.
    Lanza RuntimeError con el detalle de cada falla si se agotan los intentos.
    """
    usuario, clave = credenciales()
    fallas = []

    for vuelta in range(1, INTENTOS + 1):
        for puerto, modo in PUERTOS:
            try:
                with _conectar(puerto, modo) as s:
                    s.login(usuario, clave)
                    s.send_message(msg, to_addrs=destinatarios)
                if vuelta > 1 or (puerto, modo) != PUERTOS[0]:
                    print(f"{prefijo}[EMAIL] enviado por puerto {puerto} ({modo}), "
                          f"vuelta {vuelta}")
                return puerto, modo

            except _NO_REINTENTAR as e:
                # Credencial o destinatario mal: no insistir en ningún puerto.
                raise RuntimeError(
                    f"Gmail rechazó las credenciales o el destinatario: "
                    f"{type(e).__name__}: {e}"
                ) from e

            except _REINTENTABLES as e:
                detalle = f"puerto {puerto} ({modo}): {type(e).__name__}: {e}"
                fallas.append(f"vuelta {vuelta}, {detalle}")
                print(f"{prefijo}[EMAIL] falló {detalle}")

        if vuelta < INTENTOS:
            espera = ESPERA_INICIAL * (2 ** (vuelta - 1))
            print(f"{prefijo}[EMAIL] ningún puerto respondió. "
                  f"Reintento en {espera}s...")
            time.sleep(espera)

    raise RuntimeError(
        "No se pudo enviar el correo por ningún puerto.\n  "
        + "\n  ".join(fallas)
        + "\n\n  Si todas las fallas dicen SMTPServerDisconnected durante el "
          "login, el problema no es la red: revisá que GMAIL_APP_PASS sea una "
          "contraseña de aplicación vigente de la cuenta GMAIL_USER. "
          "Corré el workflow 'Diagnóstico de correo' para aislarlo."
    )
