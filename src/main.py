"""
Orquestador del reporte.

Modos:
  python src/main.py                   → ACTUALIZADOR: genera Excel+PPT y sube a Drive.
  python src/main.py --enviar          → NOTIFICADOR INTERNO: regenera fresco, sube y manda correo interno.
  python src/main.py --enviar-externo  → NOTIFICADOR EXTERNO: reporte SEMANAL (martes→lunes) a clientes.

Diferencias de período:
  - ACTUALIZADOR e INTERNO: implementaciones activas HOY.
  - EXTERNO: implementaciones activas en cualquier día de la ventana
    martes de la semana anterior → lunes actual (7 días). Así una campaña que
    corrió de miércoles a viernes SÍ se informa, y una que arranca justo hoy
    lunes también se informa hoy (no la próxima semana).

Ambos notificadores (interno y externo) OMITEN las campañas cuyo PROCESO está
100% en "Cancelado". El actualizador las sigue generando en Drive.

Los tres modos escriben sobre el MISMO archivo de Drive por campaña
(Reporte_Campana.xlsx / Fotos_Campana.pptx), conservando file ID y link
permanente. El link siempre refleja la última corrida.

Flags extra:
  --desde YYYY-MM-DD  --hasta YYYY-MM-DD   → fuerza la ventana (pruebas).
  --incluir-canceladas                     → no omite las 100% canceladas.
"""
import sys
import time
from datetime import date, datetime
import traceback

from leer_sheets import cargar_datos
from generar_excel import generar_excel
from generar_ppt import generar_ppt
from subir_a_drive import subir_reportes
from enviar_correo import enviar_email
from utils import ventana_semanal
from config import IDX


def _arg_fecha(flag):
    """Lee un flag tipo --desde 2026-08-04 desde sys.argv. Devuelve date o None."""
    if flag not in sys.argv:
        return None
    i = sys.argv.index(flag)
    if i + 1 >= len(sys.argv):
        raise SystemExit(f"Falta el valor de {flag} (formato YYYY-MM-DD)")
    try:
        return datetime.strptime(sys.argv[i + 1], "%Y-%m-%d").date()
    except ValueError:
        raise SystemExit(f"Valor inválido en {flag}: use YYYY-MM-DD")


def _generar_y_subir(campana, filas, headers, fecha_ref):
    """Genera Excel + PPT y los sube a Drive. Devuelve (excel_bytes, ppt_bytes, links)."""
    print("  [1/3] Generando Excel...")
    excel_bytes = generar_excel(campana, filas, headers, fecha_ref)
    print(f"        OK ({len(excel_bytes)/1024:.1f} KB)")

    print("  [2/3] Generando PPT con fotos...")
    ppt_bytes = generar_ppt(campana, filas, fecha_ref)
    print(f"        OK ({len(ppt_bytes)/1024:.1f} KB)")

    cliente = str(filas[0][IDX["CLIENTE"]] or "SIN CLIENTE").strip()
    print("  [3/3] Subiendo/actualizando en Drive...")
    links = subir_reportes(cliente, campana, excel_bytes, ppt_bytes)
    print(f"        OK carpeta: {links['carpeta']}")
    return excel_bytes, ppt_bytes, links


def main():
    enviar = "--enviar" in sys.argv
    enviar_externo = "--enviar-externo" in sys.argv
    es_notificador = enviar or enviar_externo
    excluir_canceladas = es_notificador and "--incluir-canceladas" not in sys.argv

    hoy = date.today()

    # --- Determinar la ventana de fechas según el modo ---
    if enviar_externo:
        desde, hasta = ventana_semanal(hoy)
        modo = "NOTIFICADOR EXTERNO (semanal martes→lunes, correo a clientes)"
        tipo_correo = "externo"
    elif enviar:
        desde = hasta = hoy
        modo = "NOTIFICADOR INTERNO (diario, correo interno)"
        tipo_correo = "interno"
    else:
        desde = hasta = hoy
        modo = "ACTUALIZADOR (genera y sube a Drive)"
        tipo_correo = None

    # Override manual para pruebas
    desde = _arg_fecha("--desde") or desde
    hasta = _arg_fecha("--hasta") or hasta

    # Fecha de referencia que se imprime en los archivos y el correo
    fecha_ref = hasta

    inicio = time.time()
    print(f"=== Reporte {hoy.isoformat()} | modo: {modo} ===")
    if desde != hasta:
        print(f"=== Período informado: {desde.strftime('%d/%m/%Y')} → "
              f"{hasta.strftime('%d/%m/%Y')} ===")
    if es_notificador and not excluir_canceladas:
        print("=== ATENCIÓN: --incluir-canceladas activo, NO se omiten las 100% canceladas ===")

    headers, grupos = cargar_datos(desde, hasta, excluir_canceladas=excluir_canceladas)
    if not grupos:
        print("Sin campañas para procesar. Fin.")
        return 0

    periodo = (desde, hasta) if desde != hasta else None

    errores = []
    for campana, filas in grupos.items():
        print(f"\n--- Procesando campaña: {campana} ({len(filas)} filas) ---")
        t0 = time.time()
        try:
            cliente = str(filas[0][IDX["CLIENTE"]] or "SIN CLIENTE").strip()
            print(f"  Cliente: {cliente}")

            if not es_notificador:
                # MODO ACTUALIZADOR: generar y subir
                _generar_y_subir(campana, filas, headers, fecha_ref)
                print(f"  [OK] {campana} actualizado en {time.time()-t0:.1f}s")
                continue

            # MODO NOTIFICADOR: SIEMPRE regenera fresco antes de enviar.
            # Así el correo lleva datos del momento, sin depender de que el
            # actualizador horario haya corrido (el cron de GitHub es errático).
            print("  [1/2] Generando reportes frescos y subiendo a Drive...")
            excel_bytes, ppt_bytes, links = _generar_y_subir(
                campana, filas, headers, fecha_ref
            )

            print("  [2/2] Enviando correo...")
            enviar_email(campana, fecha_ref, excel_bytes, ppt_bytes, links,
                         cliente=cliente, tipo=tipo_correo, periodo=periodo)
            print(f"  [OK] {campana} notificado en {time.time()-t0:.1f}s")

        except Exception as e:
            print(f"  [ERROR] {campana}: {e}")
            traceback.print_exc()
            errores.append((campana, str(e)))

    total = time.time() - inicio
    print(f"\n=== Fin. Tiempo total: {total:.1f}s ===")
    if errores:
        print(f"Campañas con error ({len(errores)}):")
        for c, e in errores:
            print(f"  - {c}: {e}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
