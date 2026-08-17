"""
Helpers de SEMANA ISO para el reporte de Ingresos.

Se usa semana ISO (lunes a domingo) porque coincide con la numeración que ya
usan en la operación: el 12/08/2026 es la SEMANA 33 de 2026.

OJO con el borde de año: se devuelve el AÑO ISO, no el año calendario.
  - 03/01/2027 (domingo) → SEMANA 53 - 2026
  - 04/01/2027 (lunes)   → SEMANA 1 - 2027
Si usáramos el año calendario, el domingo 03/01/2027 caería en una carpeta
"SEMANA 53 - 2027" que no existe, partiendo la semana en dos carpetas.
"""
from datetime import date, timedelta


def semana_iso(dia=None):
    """Devuelve (anio_iso, numero_semana) del día indicado."""
    if dia is None:
        dia = date.today()
    iso = dia.isocalendar()
    # isocalendar() devuelve (year, week, weekday) en py3.8 y un objeto en 3.9+
    return iso[0], iso[1]


def rango_semana(anio, semana):
    """Devuelve (lunes, domingo) de la semana ISO indicada."""
    lunes = date.fromisocalendar(anio, semana, 1)
    return lunes, lunes + timedelta(days=6)


def etiqueta_semana(anio, semana):
    """Nombre de la carpeta de Drive. Ej: 'SEMANA 33 - 2026'."""
    return f"SEMANA {semana} - {anio}"


def resolver_semana(hoy=None, anterior=False, semana=None, anio=None):
    """
    Resuelve qué semana se debe procesar.

    Prioridad:
      1. --semana / --anio explícitos.
      2. --semana-anterior  → la semana que ya cerró.
      3. Default            → la semana EN CURSO (el flujo corre viernes/sábado).

    Devuelve (anio, semana, lunes, domingo, etiqueta).
    """
    if hoy is None:
        hoy = date.today()

    if semana:
        anio_final = anio or hoy.isocalendar()[0]
        semana_final = int(semana)
    else:
        ref = hoy - timedelta(days=7) if anterior else hoy
        anio_final, semana_final = semana_iso(ref)

    lunes, domingo = rango_semana(anio_final, semana_final)
    return anio_final, semana_final, lunes, domingo, etiqueta_semana(anio_final, semana_final)
