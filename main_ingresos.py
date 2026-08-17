"""
Orquestador del REPORTE SEMANAL DE INGRESOS (Activación).

Modos:
  python src/main_ingresos.py             → ACTUALIZADOR: genera los PPT de la
                                            semana y los sube a la carpeta
                                            SEMANA XX - AAAA. Sin correo.
  python src/main_ingresos.py --enviar    → NOTIFICADOR: genera, sube y manda
                                            UN SOLO correo con todos los PPT.

Semana informada:
  Por defecto, la SEMANA EN CURSO (el flujo está pensado para correr viernes o
  sábado, con la semana ya cargada).

Flags:
  --semana-anterior            → informa la semana que ya cerró.
  --semana 33 [--anio 2026]    → fuerza una semana puntual (para regenerar).
  --cliente "Softys"           → solo las campañas de ese cliente.
  --campana "Verano 2026"      → solo esa campaña.
  --destinatarios "a@x,b@y"    → manda exactamente a esos correos.
  --incluir-canceladas         → no omite las campañas 100% canceladas.
"""
import sys
import time
import traceback
from datetime import date

from semana import resolver_semana
from leer_sheets_ingresos import cargar_ingresos
from generar_ppt_ingresos import generar_ppt_ingresos
from subir_a_drive_ingresos import preparar_carpeta_semana, subir_ppt_semana
from enviar_correo_ingresos import enviar_email_ingresos
from config import IDX
from config_ingresos import EXCLUIR_CANCELADAS_INGRESOS


def _arg_texto(flag):
    if flag not in sys.argv:
        return None
    i = sys.argv.index(flag)
    if i + 1 >= len(sys.argv):
        raise SystemExit(f"Falta el valor de {flag}")
    return sys.argv[i + 1].strip() or None


def _arg_entero(flag):
    valor = _arg_texto(flag)
    if valor is None:
        return None
    try:
        return int(valor)
    except ValueError:
        raise SystemExit(f"{flag} debe ser un número entero")


def main():
    enviar = "--enviar" in sys.argv
    excluir_canceladas = (
        EXCLUIR_CANCELADAS_INGRESOS and "--incluir-canceladas" not in sys.argv
    )
    filtro_cliente = _arg_texto("--cliente")
    filtro_campana = _arg_texto("--campana")
    destinatarios = _arg_texto("--destinatarios")

    hoy = date.today()
    anio, semana, lunes, domingo, etiqueta = resolver_semana(
        hoy,
        anterior="--semana-anterior" in sys.argv,
        semana=_arg_entero("--semana"),
        anio=_arg_entero("--anio"),
    )

    modo = "NOTIFICADOR (genera, sube y manda 1 correo)" if enviar \
        else "ACTUALIZADOR (genera y sube a Drive, sin correo)"

    inicio = time.time()
    print(f"=== Reporte Ingresos | {hoy.isoformat()} | modo: {modo} ===")
    print(f"=== {etiqueta} | {lunes.strftime('%d/%m/%Y')} → "
          f"{domingo.strftime('%d/%m/%Y')} ===")
    if filtro_cliente:
        print(f"=== Filtro de cliente: {filtro_cliente} ===")
    if filtro_campana:
        print(f"=== Filtro de campaña: {filtro_campana} ===")
    if destinatarios:
        print(f"=== Destinatarios forzados: {destinatarios} ===")

    _, grupos = cargar_ingresos(
        lunes, domingo,
        excluir_canceladas=excluir_canceladas,
        cliente=filtro_cliente,
        campana=filtro_campana,
    )
    if not grupos:
        print("Sin ingresos para procesar esta semana. Fin.")
        return 0

    # La carpeta de la semana se crea UNA vez y la comparten todas las campañas
    carpeta_id, drive_id, link_carpeta = preparar_carpeta_semana(etiqueta)
    print(f"Carpeta de la semana: {link_carpeta}")

    reportes, errores = [], []
    for campana, filas in grupos.items():
        print(f"\n--- Procesando campaña: {campana} ({len(filas)} filas) ---")
        t0 = time.time()
        try:
            cliente = str(filas[0][IDX["CLIENTE"]] or "SIN CLIENTE").strip()
            print(f"  Cliente: {cliente}")

            ppt_bytes = generar_ppt_ingresos(
                campana, filas, etiqueta, semana, anio, lunes, domingo, hoy
            )
            print(f"  PPT generado ({len(ppt_bytes)/1024:.1f} KB)")

            nombre, link = subir_ppt_semana(
                carpeta_id, drive_id, campana, semana, anio, ppt_bytes
            )
            reportes.append({
                "campana": campana,
                "cliente": cliente,
                "nombre": nombre,
                "link": link,
                "bytes": ppt_bytes,
                "registros": len(filas),
            })
            print(f"  [OK] {campana} en {time.time()-t0:.1f}s")

        except Exception as e:
            print(f"  [ERROR] {campana}: {e}")
            traceback.print_exc()
            errores.append((campana, str(e)))

    # --- Un solo correo con todo ---
    if enviar and reportes:
        print(f"\n--- Enviando correo único de {etiqueta} ---")
        try:
            enviar_email_ingresos(
                etiqueta, lunes, domingo, reportes, link_carpeta,
                destinatarios=destinatarios,
            )
        except Exception as e:
            print(f"  [ERROR] Envío de correo: {e}")
            traceback.print_exc()
            errores.append(("CORREO", str(e)))
    elif enviar:
        print("\nNo se generó ningún PPT: no se envía correo.")

    print(f"\n=== Fin. {len(reportes)} PPT en {etiqueta}. "
          f"Tiempo total: {time.time()-inicio:.1f}s ===")
    if errores:
        print(f"Errores ({len(errores)}):")
        for c, e in errores:
            print(f"  - {c}: {e}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
