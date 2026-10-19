"""
Diagnóstico de columnas de la hoja Actividades.

NO modifica nada. Lee la planilla y te muestra:
  1. Cada encabezado real con su índice y su letra de columna.
  2. A qué columna quedó apuntando cada campo del IDX, y cómo se resolvió
     (exacto / prefijo / contiene / override).
  3. Los campos que quedaron sin columna.

Corrélo cuando cambies la estructura de la planilla, o cuando un reporte
devuelva menos filas de las que esperás.

    python src/diagnostico_columnas.py
"""
import sys

from leer_sheets import leer_actividades
from columnas import ALIAS, OBLIGATORIAS, limpiar, letra_columna, construir_idx
from utils import parse_fecha


def main():
    # leer_actividades() ya sincroniza el IDX. Si falta una columna obligatoria
    # corta acá con el detalle, que es justamente lo que queremos ver.
    headers, filas = leer_actividades()
    if not headers:
        print("La hoja vino vacía. Revisá SHEET_ID y HOJA_DATOS en config.py.")
        return 1

    idx, informe = construir_idx(headers, verbose=False)
    por_indice = {}
    for campo, (i, modo) in informe["resueltos"].items():
        por_indice.setdefault(i, []).append((campo, modo))

    print("=" * 86)
    print(f"ENCABEZADOS REALES: {len(headers)} columnas (A hasta {letra_columna(len(headers)-1)})")
    print(f"Filas de datos: {len(filas)}")
    print("=" * 86)
    for i, h in enumerate(headers):
        muestra = ""
        for fila in filas[:60]:
            if i < len(fila) and str(fila[i]).strip():
                muestra = str(fila[i]).strip()[:20]
                break
        campos = por_indice.get(i, [])
        etiqueta = ", ".join(
            c if m == "exacto" else f"{c}[{m}]" for c, m in campos
        ) or "-"
        print(f"  [{i:>3}] {letra_columna(i):>3}  {str(h)[:30]:<30} | {muestra:<20} | {etiqueta}")

    print()
    print("=" * 86)
    print("CAMPOS DEL CÓDIGO")
    print("=" * 86)
    for campo in ALIAS:
        marca = "*" if campo in OBLIGATORIAS else " "
        if campo in informe["resueltos"]:
            i, modo = informe["resueltos"][campo]
            aviso = "" if modo == "exacto" else f"   <-- resuelto por '{modo}', verificar"
            print(f" {marca} {campo:<18} -> [{i:>3}] {letra_columna(i):>3}  "
                  f"{str(headers[i])[:30]:<30}{aviso}")
        else:
            print(f" {marca} {campo:<18} -> SIN COLUMNA (se lee vacío)")
    print()
    print("  (*) obligatoria: si falta, el proceso corta en vez de seguir con datos malos")

    if informe["aproximados"]:
        print()
        print("REVISAR: estos se resolvieron por parecido, no por nombre exacto.")
        print("Si alguno quedó mal, agregá el nombre real a ALIAS en src/columnas.py:")
        for campo in informe["aproximados"]:
            i, modo = informe["resueltos"][campo]
            print(f"  - {campo}: '{headers[i]}' ({modo}). Alias probados: {ALIAS[campo]}")

    if informe["ambiguos"]:
        print()
        print("AMBIGUOS: varios encabezados coinciden. Fijalos en config.COLUMNAS_OVERRIDE:")
        for campo in informe["ambiguos"]:
            print(f"  - {campo}  (alias: {ALIAS[campo]})")

    # Chequeo de salud: ¿las columnas de fecha realmente tienen fechas?
    print()
    print("=" * 86)
    print("CHEQUEO DE CONTENIDO (primeras 500 filas)")
    print("=" * 86)
    for campo in ("FECHA_INICIO", "FECHA_TERMINO", "FECHA_ENTREGA",
                  "FECHA_COMPROMISO", "FECHA_REAGENDA", "FECHA_TENTATIVA"):
        i = idx.get(campo)
        if i is None or i >= len(headers):
            continue
        valores = [f[i] for f in filas[:500] if i < len(f) and str(f[i]).strip()]
        ok = sum(1 for v in valores if parse_fecha(v))
        estado = "OK" if ok == len(valores) else "<<< HAY VALORES QUE NO SON FECHA"
        print(f"  {campo:<18} col {letra_columna(i):>3}: {ok}/{len(valores)} parsean como fecha  {estado}")
        if ok != len(valores):
            malos = [v for v in valores if not parse_fecha(v)][:3]
            print(f"       ejemplos no-fecha: {malos}")

    for campo in ("ACTIVIDAD", "TIPO_ACTIVIDAD", "PROCESO"):
        i = idx.get(campo)
        if i is None or i >= len(headers):
            continue
        vistos = {}
        for f in filas[:500]:
            if i < len(f):
                v = str(f[i]).strip()
                if v:
                    vistos[v] = vistos.get(v, 0) + 1
        top = sorted(vistos.items(), key=lambda x: -x[1])[:6]
        print(f"  {campo:<18} col {letra_columna(i):>3}: {dict(top)}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
