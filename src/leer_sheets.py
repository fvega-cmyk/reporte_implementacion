"""
Lee la hoja Actividades, filtra implementaciones y agrupa por campaña.

Dos modos de filtro:

  1) UN SOLO DÍA (actualizador y notificador interno):
         cargar_datos(hoy)
     Incluye filas cuyo período [FECHA_INICIO, FECHA_TERMINO] contiene a hoy.

  2) RANGO DE FECHAS (notificador externo, martes → lunes):
         cargar_datos(desde=martes, hasta=lunes)
     Incluye toda fila cuyo período se SOLAPE con la ventana, no solo las que
     empiezan dentro de ella. Así una campaña que corrió de miércoles a viernes
     entra igual, y una que empieza justo hoy lunes también.

Además: se pueden excluir las campañas cuyo PROCESO está 100% en "Cancelado"
(se usa en ambos notificadores; el actualizador las sigue generando en Drive).
"""
from datetime import date
from collections import defaultdict

from google_clients import get_gspread
from config import (
    SHEET_ID, HOJA_DATOS, IDX,
    PROCESOS_CANCELADO, CANCELADAS_IGNORAR_VACIOS,
)
from utils import normalizar, parse_fecha

ANCHO_MINIMO = max(IDX.values()) + 1


def leer_actividades():
    """Devuelve (headers, filas) leyendo toda la hoja."""
    gc = get_gspread()
    sh = gc.open_by_key(SHEET_ID)
    ws = sh.worksheet(HOJA_DATOS)
    datos = ws.get_all_values()  # lista de listas (strings)
    if not datos:
        return [], []
    headers = datos[0]
    filas = datos[1:]
    return headers, filas


def _rellenar(row):
    """Padding por si la fila viene corta."""
    if len(row) < ANCHO_MINIMO:
        return row + [""] * (ANCHO_MINIMO - len(row))
    return row


# ============================================================
# FILTRO POR FECHAS
# ============================================================
def filtrar_activas(filas, desde, hasta=None):
    """
    Filtra filas con ACTIVIDAD='implementacion' cuyo período de fechas se
    SOLAPA con la ventana [desde, hasta].

    Si `hasta` es None, se asume hasta = desde (comportamiento de un solo día,
    idéntico al original: ini <= hoy <= ter).
    """
    if hasta is None:
        hasta = desde

    activas = []
    for row in filas:
        row = _rellenar(row)

        if normalizar(row[IDX["ACTIVIDAD"]]) != "implementacion":
            continue

        ini = parse_fecha(row[IDX["FECHA_INICIO"]])
        ter = parse_fecha(row[IDX["FECHA_TERMINO"]])
        if not ini or not ter:
            continue

        # Solape de intervalos: [ini, ter] ∩ [desde, hasta] != vacío
        if ini <= hasta and ter >= desde:
            activas.append(row)

    return activas


# ============================================================
# AGRUPACIÓN
# ============================================================
def agrupar_por_campana(activas):
    """Agrupa filas por nombre de campaña."""
    grupos = defaultdict(list)
    for row in activas:
        campana = (row[IDX["CAMPANA"]] or "SIN CAMPAÑA").strip()
        grupos[campana].append(row)
    return dict(grupos)


# ============================================================
# CAMPAÑAS 100% CANCELADAS
# ============================================================
def estado_cancelacion(filas):
    """
    Analiza la columna PROCESO de una campaña.
    Devuelve (esta_100_cancelada, cancelados, considerados, vacios).
    """
    cancelados = 0
    considerados = 0
    vacios = 0

    for row in filas:
        proc = normalizar(row[IDX["PROCESO"]])
        if not proc:
            vacios += 1
            if CANCELADAS_IGNORAR_VACIOS:
                continue          # no cuenta ni a favor ni en contra
            considerados += 1     # cuenta como "no cancelada"
            continue
        considerados += 1
        if proc in PROCESOS_CANCELADO:
            cancelados += 1

    cancelada = considerados > 0 and cancelados == considerados
    return cancelada, cancelados, considerados, vacios


def excluir_campanas_canceladas(grupos):
    """
    Quita del dict de grupos las campañas cuyo PROCESO está 100% en Cancelado.
    Devuelve (grupos_filtrados, lista_de_campanas_excluidas).
    """
    vivas = {}
    excluidas = []

    for campana, filas in grupos.items():
        cancelada, canc, total, vacios = estado_cancelacion(filas)
        if cancelada:
            excluidas.append(campana)
            print(f"  [CANCELADA] '{campana}' → {canc}/{total} filas en Cancelado. Se omite.")
        else:
            vivas[campana] = filas
            if canc:
                detalle = f"  [PARCIAL] '{campana}' → {canc}/{total} en Cancelado"
                if vacios:
                    detalle += f" ({vacios} sin PROCESO)"
                print(detalle + ". Se notifica.")

    return vivas, excluidas


# ============================================================
# ENTRADA PRINCIPAL
# ============================================================
def cargar_datos(desde=None, hasta=None, excluir_canceladas=False):
    """
    Lee, filtra y agrupa. Retorna (headers, grupos).

    - Un día:  cargar_datos(hoy)
    - Rango:   cargar_datos(martes, lunes)
    """
    if desde is None:
        desde = date.today()
    if hasta is None:
        hasta = desde

    if desde == hasta:
        print(f"Filtro: implementaciones activas el {desde.strftime('%d/%m/%Y')}")
    else:
        print(f"Filtro: implementaciones activas entre "
              f"{desde.strftime('%d/%m/%Y')} y {hasta.strftime('%d/%m/%Y')} "
              f"(solape de fechas)")

    headers, filas = leer_actividades()
    activas = filtrar_activas(filas, desde, hasta)
    print(f"Filas activas: {len(activas)}")

    grupos = agrupar_por_campana(activas)
    print(f"Campañas encontradas ({len(grupos)}): "
          f"{', '.join(grupos.keys()) if grupos else '(ninguna)'}")

    if excluir_canceladas and grupos:
        grupos, excluidas = excluir_campanas_canceladas(grupos)
        if excluidas:
            print(f"Campañas excluidas por estar 100% canceladas ({len(excluidas)}): "
                  f"{', '.join(excluidas)}")
        print(f"Campañas a notificar ({len(grupos)}): "
              f"{', '.join(grupos.keys()) if grupos else '(ninguna)'}")

    return headers, grupos


# Alias retrocompatible con el nombre viejo.
def cargar_datos_del_dia(hoy=None, excluir_canceladas=False):
    return cargar_datos(desde=hoy, excluir_canceladas=excluir_canceladas)
