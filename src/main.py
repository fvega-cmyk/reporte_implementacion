"""
Orquestador del reporte.

Modos:
  python src/main.py                   → ACTUALIZADOR: genera Excel+PPT y sube a Drive.
  python src/main.py --enviar          → NOTIFICADOR INTERNO: descarga de Drive y manda correo interno.
  python src/main.py --enviar-externo  → NOTIFICADOR EXTERNO: descarga de Drive y manda correo externo (lunes).

Los notificadores NO regeneran los archivos: descargan lo último que dejó el
actualizador en Drive (rápido). Si por alguna razón no existen en Drive, los
regeneran como fallback (red de seguridad).
"""
import sys
import time
from datetime import date
import traceback

from leer_sheets import cargar_datos_del_dia
from generar_excel import generar_excel
from generar_ppt import generar_ppt
from subir_a_drive import subir_reportes, obtener_reportes_existentes
from enviar_correo import enviar_email
from config import IDX


def _generar_y_subir(campana, filas, headers, hoy):
    """Genera Excel + PPT y los sube a Drive. Devuelve (excel_bytes, ppt_bytes, links)."""
    print("  [1/3] Generando Excel...")
    excel_bytes = generar_excel(campana, filas, headers, hoy)
    print(f"        OK ({len(excel_bytes)/1024:.1f} KB)")

    print("  [2/3] Generando PPT con fotos...")
    ppt_bytes = generar_ppt(campana, filas, hoy)
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

    if enviar_externo:
        modo = "NOTIFICADOR EXTERNO (descarga de Drive + correo a clientes)"
        tipo_correo = "externo"
    elif enviar:
        modo = "NOTIFICADOR INTERNO (descarga de Drive + correo interno)"
        tipo_correo = "interno"
    else:
        modo = "ACTUALIZADOR (genera y sube a Drive)"
        tipo_correo = None

    inicio = time.time()
    hoy = date.today()
    print(f"=== Reporte {hoy.isoformat()} | modo: {modo} ===")

    headers, grupos = cargar_datos_del_dia(hoy)
    if not grupos:
        print("Sin implementaciones activas hoy. Fin.")
        return 0

    errores = []
    for campana, filas in grupos.items():
        print(f"\n--- Procesando campaña: {campana} ({len(filas)} filas) ---")
        t0 = time.time()
        try:
            cliente = str(filas[0][IDX["CLIENTE"]] or "SIN CLIENTE").strip()
            print(f"  Cliente: {cliente}")

            if not es_notificador:
                # MODO ACTUALIZADOR: generar y subir
                _generar_y_subir(campana, filas, headers, hoy)
                print(f"  [OK] {campana} actualizado en {time.time()-t0:.1f}s")
                continue

            # MODO NOTIFICADOR: intentar descargar de Drive (rápido)
            print("  [1/2] Descargando reportes desde Drive...")
            datos = obtener_reportes_existentes(cliente, campana)

            if datos is None:
                # Fallback: no existen en Drive todavía → generarlos
                print("        No estaban en Drive. Generando (fallback)...")
                excel_bytes, ppt_bytes, links = _generar_y_subir(campana, filas, headers, hoy)
            else:
                excel_bytes = datos["excel_bytes"]
                ppt_bytes = datos["ppt_bytes"]
                links = datos["links"]
                print(f"        OK descargado (Excel {len(excel_bytes)/1024:.1f} KB, "
                      f"PPT {len(ppt_bytes)/1024:.1f} KB)")

            print("  [2/2] Enviando correo...")
            enviar_email(campana, hoy, excel_bytes, ppt_bytes, links,
                         cliente=cliente, tipo=tipo_correo)
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
