"""
Lee la hoja Actividades y deja SOLO las filas de Ingresos (Activación) que
caen dentro de una semana ISO, agrupadas por campaña.

Filtro doble:
    ACTIVIDAD       = "Ingreso"
    TIPO ACTIVIDAD  = "Activación"

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
)
from utils import normalizar, parse_fecha

_ACTIVIDAD_OBJETIVO = normalizar(ACTIVIDAD_INGRESO)
_TIPOS_OBJETIVO = {normalizar(t) for t in TIPOS_ACTIVIDAD_ACEPTADOS}


def _fechas_de_fila(row):
    """
    Determina el período [ini, ter] de una fila de Ingreso.

    En la práctica muchas filas de activación traen solo una de las dos fechas,
    así que se completa la que falte con la otra. Si no hay ninguna, la fila
    no se puede ubicar en una semana y se descarta.
    """
    ini = parse_fecha(row[IDX["FECHA_INICIO"]])
    ter = parse_fecha(row[IDX["FECHA_TERMINO"]])

    if not ini and not ter:
        # Último recurso: fecha tentativa (algunas filas se cargan así)
        ini = ter = parse_fecha(row[IDX["FECHA_TENTATIVA"]])

    if ini and not ter:
        ter = ini
    if ter and not ini:
        ini = ter

    if ini and ter and ter < ini:
        ini, ter = ter, ini

    return ini, ter


def es_fila_ingreso(row):
    """True si la fila es ACTIVIDAD=Ingreso y TIPO ACTIVIDAD=Activación."""
    if normalizar(row[IDX["ACTIVIDAD"]]) != _ACTIVIDAD_OBJETIVO:
        return False
    return normalizar(row[IDX["TIPO_ACTIVIDAD"]]) in _TIPOS_OBJETIVO


def filtrar_ingresos(filas, desde, hasta):
    """
    Filas de Ingreso cuyo período se SOLAPA con [desde, hasta].

    Solape (no "empieza dentro"): una activación de jueves a martes entra
    tanto en la semana del jueves como en la del martes. Es a propósito: si
    la actividad estuvo viva en la semana, corresponde informarla.
    """
    activas = []
    descartadas_sin_fecha = 0

    for row in filas:
        row = _rellenar(row)
        if not es_fila_ingreso(row):
            continue

        ini, ter = _fechas_de_fila(row)
        if not ini or not ter:
            descartadas_sin_fecha += 1
            continue

        if ini <= hasta and ter >= desde:
            activas.append(row)

    if descartadas_sin_fecha:
        print(f"  [AVISO] {descartadas_sin_fecha} fila(s) de Ingreso sin fecha usable. Se omiten.")

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
