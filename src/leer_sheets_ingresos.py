"""
Lee la hoja Actividades y deja SOLO las filas de Ingresos (Activación) que
caen dentro de una semana ISO, agrupadas por campaña.

Filtro triple:
    ACTIVIDAD       = "Ingreso"
    TIPO ACTIVIDAD  = "Activación"
    FECHA DEL INGRESO dentro de la semana (lunes a domingo)

LA FECHA QUE MANDA
------------------
NO se usa FECHA_INICIO / FECHA_TERMINO. En activaciones esas dos columnas son
el rango de la CAMPAÑA completa ("DEGUSTACIÓN OREO ... DE JUNIO A JULIO"), no
la fecha del ingreso. Como ese rango se solapa con cualquier semana, filtrar
por ahí metía toda la historia de la campaña en todas las semanas: una misma
sala aparecía 5 veces, una por cada ingreso viejo.

Se usa la fecha del evento: FECHA_ENTREGA, y si está vacía, la primera con
valor de COLUMNAS_FECHA_INGRESO (reagenda, compromiso, tentativa). Eso cubre
tanto los ingresos que sucedieron como los que debían suceder esa semana.

Se reutiliza toda la maquinaria de leer_sheets.py (lectura, agrupación,
canceladas, filtros manuales); acá solo cambia el criterio de filas.
"""
from leer_sheets import (
    leer_actividades,
    _rellenar,
    agrupar_por_campana,
    excluir_campanas_canceladas,
    filtrar_por_cliente,
    filtrar_por_campana,
    clientes_disponibles,
)
from config import IDX
from config_ingresos import (
    ACTIVIDAD_INGRESO,
    TIPOS_ACTIVIDAD_ACEPTADOS,
    COLUMNAS_FECHA_INGRESO,
)
from utils import normalizar, parse_fecha

_ACTIVIDAD_OBJETIVO = normalizar(ACTIVIDAD_INGRESO)
_TIPOS_OBJETIVO = {normalizar(t) for t in TIPOS_ACTIVIDAD_ACEPTADOS}


def fecha_de_ingreso(row):
    """
    Fecha del ingreso de una fila. Devuelve (fecha, nombre_de_columna).

    Recorre COLUMNAS_FECHA_INGRESO en orden y se queda con la primera que
    tenga un valor parseable. Si ninguna tiene fecha, devuelve (None, None):
    la fila no se puede ubicar en ninguna semana y se descarta.
    """
    for columna in COLUMNAS_FECHA_INGRESO:
        idx = IDX.get(columna)
        if idx is None:
            continue
        fecha = parse_fecha(row[idx])
        if fecha:
            return fecha, columna
    return None, None


def es_fila_ingreso(row):
    """True si la fila es ACTIVIDAD=Ingreso y TIPO ACTIVIDAD=Activación."""
    if normalizar(row[IDX["ACTIVIDAD"]]) != _ACTIVIDAD_OBJETIVO:
        return False
    return normalizar(row[IDX["TIPO_ACTIVIDAD"]]) in _TIPOS_OBJETIVO


def filtrar_ingresos(filas, desde, hasta):
    """
    Filas de Ingreso cuya FECHA DEL INGRESO cae dentro de [desde, hasta].

    Es una fecha puntual, no un rango: el ingreso ocurrió (o debía ocurrir) un
    día concreto y pertenece a esa semana y a ninguna otra.
    """
    activas = []
    sin_fecha = 0
    fuera_semana = 0
    por_columna = {}

    for row in filas:
        row = _rellenar(row)
        if not es_fila_ingreso(row):
            continue

        fecha, columna = fecha_de_ingreso(row)
        if not fecha:
            sin_fecha += 1
            continue

        if desde <= fecha <= hasta:
            activas.append(row)
            por_columna[columna] = por_columna.get(columna, 0) + 1
        else:
            fuera_semana += 1

    if activas:
        detalle = ", ".join(f"{col}: {n}" for col, n in por_columna.items())
        print(f"  Fecha usada para ubicar el ingreso → {detalle}")
    if fuera_semana:
        print(f"  {fuera_semana} fila(s) de Ingreso de otras semanas. Se omiten.")
    if sin_fecha:
        print(f"  [AVISO] {sin_fecha} fila(s) de Ingreso sin ninguna fecha de "
              f"{', '.join(COLUMNAS_FECHA_INGRESO)}. Se omiten.")

    return activas


def cargar_ingresos(desde, hasta, excluir_canceladas=False,
                    cliente=None, campana=None):
    """
    Lee, filtra y agrupa por campaña. Retorna (headers, grupos).
    `desde`/`hasta` son el lunes y el domingo de la semana ISO a informar.
    """
    print(f"Filtro: {ACTIVIDAD_INGRESO} / {', '.join(TIPOS_ACTIVIDAD_ACEPTADOS)} "
          f"entre {desde.strftime('%d/%m/%Y')} y {hasta.strftime('%d/%m/%Y')}")

    headers, filas = leer_actividades()
    activas = filtrar_ingresos(filas, desde, hasta)
    print(f"Filas de ingreso en la semana: {len(activas)}")

    grupos = agrupar_por_campana(activas)
    print(f"Campañas encontradas ({len(grupos)}): "
          f"{', '.join(grupos.keys()) if grupos else '(ninguna)'}")

    if cliente and grupos:
        disponibles = clientes_disponibles(grupos)
        grupos = filtrar_por_cliente(grupos, cliente)
        if not grupos:
            print(f"[AVISO] Ninguna campaña del cliente '{cliente}' en esta semana.")
            print(f"        Clientes con ingresos: {', '.join(disponibles)}")
            return headers, {}
        print(f"Filtro cliente '{cliente}' → {len(grupos)} campaña(s)")

    if campana and grupos:
        posibles = sorted(grupos.keys())
        grupos = filtrar_por_campana(grupos, campana)
        if not grupos:
            print(f"[AVISO] La campaña '{campana}' no tiene ingresos esta semana.")
            print(f"        Disponibles: {', '.join(posibles)}")
            return headers, {}

    if excluir_canceladas and grupos:
        grupos, excluidas = excluir_campanas_canceladas(grupos)
        if excluidas:
            print(f"Campañas excluidas por estar 100% canceladas ({len(excluidas)}): "
                  f"{', '.join(excluidas)}")
        print(f"Campañas a reportar ({len(grupos)}): "
              f"{', '.join(grupos.keys()) if grupos else '(ninguna)'}")

    return headers, grupos
