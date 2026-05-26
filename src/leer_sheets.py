"""
Lee la hoja Actividades, filtra implementaciones activas hoy y agrupa por campaña.
Equivalente al primer tramo de generarReporteDiario() en Apps Script.
"""
from datetime import date
from collections import defaultdict

from google_clients import get_gspread
from config import SHEET_ID, HOJA_DATOS, IDX
from utils import normalizar, parse_fecha


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


def filtrar_activas(filas, hoy=None):
    """Filtra filas con actividad='implementacion' y rango activo hoy."""
    if hoy is None:
        hoy = date.today()
    activas = []
    for row in filas:
        # Padding por si la fila viene corta
        if len(row) < max(IDX.values()) + 1:
            row = row + [""] * (max(IDX.values()) + 1 - len(row))
        act = normalizar(row[IDX["ACTIVIDAD"]])
        if act != "implementacion":
            continue
        ini = parse_fecha(row[IDX["FECHA_INICIO"]])
        ter = parse_fecha(row[IDX["FECHA_TERMINO"]])
        if not ini or not ter:
            continue
        if ini <= hoy <= ter:
            activas.append(row)
    return activas


def agrupar_por_campana(activas):
    """Agrupa filas por nombre de campaña."""
    grupos = defaultdict(list)
    for row in activas:
        campana = (row[IDX["CAMPANA"]] or "SIN CAMPAÑA").strip()
        grupos[campana].append(row)
    return dict(grupos)


def cargar_datos_del_dia(hoy=None):
    """One-shot: lee, filtra y agrupa. Retorna (headers, grupos)."""
    headers, filas = leer_actividades()
    activas = filtrar_activas(filas, hoy)
    print(f"Filas activas: {len(activas)}")
    grupos = agrupar_por_campana(activas)
    print(f"Campañas: {', '.join(grupos.keys()) if grupos else '(ninguna)'}")
    return headers, grupos
