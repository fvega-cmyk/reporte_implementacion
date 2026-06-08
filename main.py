"""
Orquestador del reporte diario.
Equivalente a generarReporteDiario() + procesarFotosYEmail() de Apps Script,
pero en un solo proceso sin límites de tiempo.
"""
import sys
import time
from datetime import date
import traceback

from leer_sheets import cargar_datos_del_dia
from generar_excel import generar_excel
from generar_ppt import generar_ppt
from enviar_correo import enviar_email


def main():
    inicio = time.time()
    hoy = date.today()
    print(f"=== Reporte diario {hoy.isoformat()} ===")

    headers, grupos = cargar_datos_del_dia(hoy)
    if not grupos:
        print("Sin implementaciones activas hoy. Fin.")
        return 0

    errores = []
    for campana, filas in grupos.items():
        print(f"\n--- Procesando campaña: {campana} ({len(filas)} filas) ---")
        t0 = time.time()
        try:
            print("  [1/3] Generando Excel...")
            excel_bytes = generar_excel(campana, filas, headers, hoy)
            print(f"        OK ({len(excel_bytes)/1024:.1f} KB)")

            print("  [2/3] Generando PPT con fotos...")
            ppt_bytes = generar_ppt(campana, filas, hoy)
            print(f"        OK ({len(ppt_bytes)/1024:.1f} KB)")

            print("  [3/3] Enviando correo...")
            enviar_email(campana, hoy, excel_bytes, ppt_bytes)

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
