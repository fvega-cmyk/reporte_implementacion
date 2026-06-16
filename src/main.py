"""
Orquestador del reporte.

Dos modos:
  python src/main.py            → modo ACTUALIZADOR (solo sube/actualiza Drive)
  python src/main.py --enviar   → modo NOTIFICADOR (sube/actualiza Drive + manda correo)

El modo actualizador corre cada hora (11-20h). El notificador corre 1 vez al día.
"""
import sys
import time
from datetime import date
import traceback

from leer_sheets import cargar_datos_del_dia
from generar_excel import generar_excel
from generar_ppt import generar_ppt
from subir_a_drive import subir_reportes
from enviar_correo import enviar_email
from config import IDX


def main():
    enviar = "--enviar" in sys.argv
    modo = "NOTIFICADOR (Drive + correo)" if enviar else "ACTUALIZADOR (solo Drive)"

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
            # Cliente: lo tomamos de la primera fila (siempre está poblado)
            cliente = str(filas[0][IDX["CLIENTE"]] or "SIN CLIENTE").strip()
            print(f"  Cliente: {cliente}")

            print("  [1/4] Generando Excel...")
            excel_bytes = generar_excel(campana, filas, headers, hoy)
            print(f"        OK ({len(excel_bytes)/1024:.1f} KB)")

            print("  [2/4] Generando PPT con fotos...")
            ppt_bytes = generar_ppt(campana, filas, hoy)
            print(f"        OK ({len(ppt_bytes)/1024:.1f} KB)")

            print("  [3/4] Subiendo/actualizando en Drive...")
            links = subir_reportes(cliente, campana, excel_bytes, ppt_bytes)
            print(f"        OK carpeta: {links['carpeta']}")

            if enviar:
                print("  [4/4] Enviando correo...")
                enviar_email(campana, hoy, excel_bytes, ppt_bytes, links)
            else:
                print("  [4/4] (modo actualizador, no se envía correo)")

            print(f"  [OK] {campana} terminado en {time.time()-t0:.1f}s")
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
