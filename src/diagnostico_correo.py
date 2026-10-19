"""
Diagnóstico del envío de correo. Corre en segundos, no genera ningún reporte.

Sirve para aislar el problema sin tener que esperar los 7 minutos que tarda
el flujo de ingresos en llegar al paso del correo.

Qué prueba, en este orden:
  1. Que los secrets existan y no tengan basura (espacios, saltos de línea).
  2. Que se pueda ABRIR la conexión a Gmail, por 587 y por 465.
  3. Que el LOGIN funcione en cada puerto. Acá es donde viene fallando.
  4. Opcional: manda un correo de prueba de 1 KB, sin adjuntos.

NUNCA imprime la contraseña. Solo su largo y si tiene caracteres raros, que
es lo que hace falta para diagnosticar.

    python src/diagnostico_correo.py                      # solo conecta
    python src/diagnostico_correo.py --enviar a@b.cl      # además manda uno
"""
import os
import ssl
import sys
import socket
import smtplib
from email.message import EmailMessage

from config import SMTP_HOST
from smtp_envio import PUERTOS, TIMEOUT


def _arg(flag):
    if flag not in sys.argv:
        return None
    i = sys.argv.index(flag)
    return sys.argv[i + 1] if i + 1 < len(sys.argv) else None


def revisar_secrets():
    """Revisa el contenido de los secrets sin revelarlos."""
    print("=" * 70)
    print("1. SECRETS")
    print("=" * 70)

    crudo_user = os.environ.get("GMAIL_USER")
    crudo_pass = os.environ.get("GMAIL_APP_PASS")

    if not crudo_user:
        print("  [FALLA] GMAIL_USER no está definido o está vacío.")
        return None, None
    if not crudo_pass:
        print("  [FALLA] GMAIL_APP_PASS no está definido o está vacío.")
        return None, None

    usuario = crudo_user.strip()
    print(f"  GMAIL_USER: {usuario}")
    if crudo_user != usuario:
        print("  [AVISO] GMAIL_USER tenía espacios o saltos de línea al borde.")
    if "@" not in usuario:
        print("  [AVISO] GMAIL_USER no parece un correo.")

    clave = crudo_pass.strip().replace(" ", "").replace("\t", "").replace("\u00a0", "")
    print(f"  GMAIL_APP_PASS: {len(crudo_pass)} caracteres guardados, "
          f"{len(clave)} después de limpiar")

    if crudo_pass.strip() != crudo_pass:
        print("  [AVISO] la contraseña tiene espacios o saltos de línea al borde.")
    if " " in crudo_pass.strip():
        print("  [AVISO] la contraseña tiene espacios en el medio. Google la")
        print("          muestra como 'abcd efgh ijkl mnop' pero hay que")
        print("          guardarla SIN espacios. Se limpian automáticamente.")
    if len(clave) != 16:
        print(f"  [AVISO] las contraseñas de aplicación de Google tienen 16")
        print(f"          caracteres; esta tiene {len(clave)}. Si es la")
        print(f"          contraseña normal de la cuenta, no va a funcionar:")
        print(f"          hay que generar una de aplicación.")
    if not clave.isalnum():
        print("  [AVISO] la contraseña tiene caracteres no alfanuméricos.")

    return usuario, clave


def probar_puertos(usuario, clave):
    """Prueba conexión y login en cada puerto por separado."""
    print()
    print("=" * 70)
    print("2. CONEXIÓN Y LOGIN POR PUERTO")
    print("=" * 70)

    resultados = {}
    for puerto, modo in PUERTOS:
        print(f"\n  --- {SMTP_HOST}:{puerto} ({modo}) ---")
        s = None
        try:
            print("      abriendo conexión...", end=" ", flush=True)
            if modo == "ssl":
                s = smtplib.SMTP_SSL(SMTP_HOST, puerto, timeout=TIMEOUT,
                                     context=ssl.create_default_context())
            else:
                s = smtplib.SMTP(SMTP_HOST, puerto, timeout=TIMEOUT)
                s.ehlo()
                s.starttls(context=ssl.create_default_context())
                s.ehlo()
            print("OK")

            print("      autenticando...", end=" ", flush=True)
            s.login(usuario, clave)
            print("OK")
            resultados[puerto] = "OK"

        except smtplib.SMTPAuthenticationError as e:
            print("FALLA")
            print(f"      Gmail rechazó las credenciales: {e.smtp_code} "
                  f"{e.smtp_error!r}")
            print("      → la contraseña de aplicación no es válida o fue revocada.")
            resultados[puerto] = "credenciales"

        except smtplib.SMTPServerDisconnected as e:
            print("FALLA")
            print(f"      Gmail cortó la conexión: {e}")
            print("      → si pasa en los DOS puertos, casi siempre es la")
            print("        contraseña de aplicación (revocada, de otra cuenta,")
            print("        o 2FA desactivado, que las invalida todas).")
            resultados[puerto] = "cortada"

        except (socket.timeout, TimeoutError):
            print("FALLA")
            print("      timeout. El puerto puede estar bloqueado en la red.")
            resultados[puerto] = "timeout"

        except Exception as e:
            print("FALLA")
            print(f"      {type(e).__name__}: {e}")
            resultados[puerto] = type(e).__name__

        finally:
            if s is not None:
                try:
                    s.quit()
                except Exception:
                    pass

    return resultados


def enviar_prueba(usuario, clave, destino, resultados):
    """Manda un correo chico por el primer puerto que haya funcionado."""
    print()
    print("=" * 70)
    print("3. CORREO DE PRUEBA")
    print("=" * 70)

    puerto_ok = next((p for p, r in resultados.items() if r == "OK"), None)
    if puerto_ok is None:
        print("  Ningún puerto autenticó. No se intenta enviar.")
        return False

    modo = dict(PUERTOS)[puerto_ok]
    msg = EmailMessage()
    msg["Subject"] = "Prueba de envío - reporte_implementacion"
    msg["From"] = f"Reporte Sell-Out <{usuario}>"
    msg["To"] = destino
    msg.set_content(
        "Correo de prueba del workflow de diagnóstico.\n"
        "Si te llegó, el envío funciona y el problema está en otro lado.\n"
    )

    try:
        print(f"  Enviando a {destino} por el puerto {puerto_ok}...", end=" ", flush=True)
        if modo == "ssl":
            s = smtplib.SMTP_SSL(SMTP_HOST, puerto_ok, timeout=TIMEOUT,
                                 context=ssl.create_default_context())
        else:
            s = smtplib.SMTP(SMTP_HOST, puerto_ok, timeout=TIMEOUT)
            s.ehlo()
            s.starttls(context=ssl.create_default_context())
            s.ehlo()
        with s:
            s.login(usuario, clave)
            s.send_message(msg)
        print("OK")
        print(f"  Revisá la bandeja de {destino} (y el spam).")
        return True
    except Exception as e:
        print("FALLA")
        print(f"  {type(e).__name__}: {e}")
        return False


def main():
    usuario, clave = revisar_secrets()
    if not usuario:
        return 1

    resultados = probar_puertos(usuario, clave)

    destino = _arg("--enviar")
    if destino:
        enviar_prueba(usuario, clave, destino, resultados)

    print()
    print("=" * 70)
    print("CONCLUSIÓN")
    print("=" * 70)
    ok = [p for p, r in resultados.items() if r == "OK"]
    if ok:
        print(f"  El login funciona en: {', '.join(str(p) for p in ok)}")
        print("  Las credenciales están bien. Si igual falla el envío real,")
        print("  el problema es el mensaje (tamaño o destinatarios), no la cuenta.")
        return 0

    if all(r in ("cortada", "credenciales") for r in resultados.values()):
        print("  Ningún puerto autenticó, y los dos fallan en el login.")
        print("  La red no es el problema: la conexión se abre bien.")
        print()
        print("  Revisá, en este orden:")
        print("   1. Que la cuenta de GMAIL_USER tenga la verificación en dos")
        print("      pasos ACTIVA. Si se desactivó, Google invalida todas las")
        print("      contraseñas de aplicación de esa cuenta en el momento.")
        print("   2. Generá una contraseña de aplicación NUEVA en")
        print("      myaccount.google.com/apppasswords y actualizá el secret")
        print("      GMAIL_APP_PASS, pegándola SIN espacios.")
        print("   3. Que GMAIL_USER sea exactamente la cuenta donde generaste")
        print("      esa contraseña.")
    else:
        print("  Fallas mezcladas. Detalle por puerto:")
        for p, r in resultados.items():
            print(f"    {p}: {r}")
    return 1


if __name__ == "__main__":
    sys.exit(main())
