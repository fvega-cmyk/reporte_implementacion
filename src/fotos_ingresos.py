"""
Resuelve las fotos de ACTIVACIONES por RUTA, no por nombre.

EL PROBLEMA QUE RESUELVE
------------------------
`utils.buscar_foto_blob` (el del flujo de Implementación) busca por NOMBRE de
archivo en todo el Drive y se queda con el primer resultado:

    q = "name = 'foto.jpg' and trashed = false"   →   pageSize=1

Si ese nombre existe en más de una carpeta, agarra cualquiera. En la práctica
eso significa que un reporte de activaciones puede terminar mostrando la foto
de una implementación, que es peor que no mostrar nada: nadie se da cuenta.

Acá se navega la ruta real:

    Mi unidad / Control Interno / <valor de la columna FOTO 1 o FOTO 2>

Si la columna trae "Activaciones_Images/abc.jpg", se busca la carpeta
"Activaciones_Images" DENTRO de "Control Interno" y el archivo "abc.jpg"
dentro de ella. Sin ambigüedad posible.

Las carpetas resueltas se cachean: 40 fotos en la misma subcarpeta hacen una
sola llamada a la API para ubicarla.
"""
import os
from io import BytesIO

from googleapiclient.http import MediaIoBaseDownload

from utils import normalizar
from config_ingresos import (
    NOMBRE_CARPETA_FOTOS_INGRESOS,
    CARPETA_FOTOS_INGRESOS_ID,
    FALLBACK_BUSQUEDA_GLOBAL,
)

MIME_FOLDER = "application/vnd.google-apps.folder"

# Cache de la corrida: {(padre_id, nombre, solo_carpetas): id | None}
_cache_hijos = {}
_cache_raiz = None

# Contadores para el resumen final
stats = {"ok": 0, "no_encontrada": 0, "global": 0}


def _q(texto):
    """Escapa comillas simples para las queries de Drive."""
    return str(texto).replace("\\", "\\\\").replace("'", "\\'")


def _listar(drive, q):
    return drive.files().list(
        q=q,
        fields="files(id, name, mimeType)",
        pageSize=25,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute().get("files", [])


def carpeta_raiz_fotos(drive):
    """
    Carpeta base de las fotos. Devuelve (id, nombre_real).

    El NOMBRE real importa: si se fija el ID de "Fotos Activación" pero la
    columna de la planilla ya trae "Fotos Activación/..." adelante, hay que
    detectar esa repetición y saltearla. Por eso no alcanza con el nombre del
    config: se usa el que devuelve Drive.
    """
    global _cache_raiz
    if _cache_raiz:
        return _cache_raiz

    id_fijo = (CARPETA_FOTOS_INGRESOS_ID
               or os.environ.get("DRIVE_CARPETA_FOTOS_INGRESOS_ID") or "").strip()
    if id_fijo:
        info = drive.files().get(
            fileId=id_fijo, fields="id, name", supportsAllDrives=True
        ).execute()
        print(f"  [FOTOS] Carpeta base (ID fijo): '{info['name']}' ({info['id']})")
        _cache_raiz = (info["id"], info["name"])
        return _cache_raiz

    carpetas = _listar(
        drive,
        f"name = '{_q(NOMBRE_CARPETA_FOTOS_INGRESOS)}' "
        f"and mimeType = '{MIME_FOLDER}' and trashed = false",
    )
    if not carpetas:
        raise RuntimeError(
            f"No se encontró la carpeta '{NOMBRE_CARPETA_FOTOS_INGRESOS}'. "
            f"Compartila con la cuenta de servicio (permiso Lector) o pegá su "
            f"ID en CARPETA_FOTOS_INGRESOS_ID (config_ingresos.py)."
        )
    if len(carpetas) > 1:
        ids = ", ".join(c["id"] for c in carpetas)
        print(f"  [FOTOS] ATENCIÓN: hay {len(carpetas)} carpetas llamadas "
              f"'{NOMBRE_CARPETA_FOTOS_INGRESOS}'. Uso la primera. "
              f"Para fijar la correcta, pegá el ID en CARPETA_FOTOS_INGRESOS_ID. "
              f"Candidatas: {ids}")

    _cache_raiz = (carpetas[0]["id"], carpetas[0]["name"])
    print(f"  [FOTOS] Carpeta base: '{_cache_raiz[1]}' ({_cache_raiz[0]})")
    return _cache_raiz


def _buscar_hijo(drive, padre_id, nombre, solo_carpetas=False):
    """Busca un hijo directo por nombre exacto. Devuelve el ID o None. Cachea."""
    clave = (padre_id, nombre.lower(), solo_carpetas)
    if clave in _cache_hijos:
        return _cache_hijos[clave]

    q = f"name = '{_q(nombre)}' and '{padre_id}' in parents and trashed = false"
    if solo_carpetas:
        q += f" and mimeType = '{MIME_FOLDER}'"

    encontrados = _listar(drive, q)
    resultado = encontrados[0]["id"] if encontrados else None
    _cache_hijos[clave] = resultado
    return resultado


def resolver_id_foto(drive, ruta):
    """
    Navega la ruta y devuelve (file_id, detalle).
    Si no la encuentra, devuelve (None, motivo) para que quede en el log.
    """
    ruta = str(ruta or "").strip().replace("\\", "/")
    if not ruta:
        return None, "ruta vacía"

    segmentos = [s.strip() for s in ruta.split("/") if s.strip()]
    if not segmentos:
        return None, "ruta vacía"

    raiz, raiz_nombre = carpeta_raiz_fotos(drive)

    # Si la columna repite el nombre de la carpeta base al principio, se saltea.
    # Pasa cuando se fija el ID de "Fotos Activación" y la planilla guarda
    # "Fotos Activación/CAMPAÑA/LOCAL/Ingreso/foto.jpg": el prefijo ya está
    # cubierto por la carpeta base y buscarlo de nuevo adentro falla siempre.
    if len(segmentos) > 1 and normalizar(segmentos[0]) == normalizar(raiz_nombre):
        segmentos = segmentos[1:]

    actual = raiz
    recorrido = [raiz_nombre]
    motivo = ""

    # Todos los segmentos menos el último son carpetas.
    # Si alguna no existe, NO se corta acá: se anota el motivo y se sigue con
    # los intentos de abajo (el archivo puede estar guardado un nivel arriba).
    for seg in segmentos[:-1]:
        siguiente = _buscar_hijo(drive, actual, seg, solo_carpetas=True)
        if not siguiente:
            motivo = (f"no existe la carpeta '{seg}' dentro de "
                      f"'{'/'.join(recorrido)}'")
            actual = None
            break
        actual = siguiente
        recorrido.append(seg)

    archivo = segmentos[-1]

    if actual:
        file_id = _buscar_hijo(drive, actual, archivo)
        if file_id:
            return file_id, "/".join(recorrido + [archivo])
        motivo = f"no está '{archivo}' en '{'/'.join(recorrido)}'"

    # Segundo intento: el archivo suelto en la carpeta base (a veces la columna
    # trae la subcarpeta pero el archivo quedó guardado un nivel arriba)
    if actual != raiz:
        file_id = _buscar_hijo(drive, raiz, archivo)
        if file_id:
            return file_id, f"{raiz_nombre}/{archivo} (sin subcarpeta)"

    # Último recurso, desactivado por default: búsqueda global por nombre.
    # Es la que provoca traer la foto equivocada, así que solo se usa si se
    # activa a propósito y queda avisado en el log.
    if FALLBACK_BUSQUEDA_GLOBAL:
        globales = _listar(drive, f"name = '{_q(archivo)}' and trashed = false")
        if globales:
            stats["global"] += 1
            return globales[0]["id"], (
                f"BÚSQUEDA GLOBAL: '{archivo}' apareció en otra carpeta "
                f"({len(globales)} coincidencia/s). Puede NO ser la foto correcta."
            )

    return None, motivo or f"no está '{archivo}' en '{raiz_nombre}'"


def buscar_foto_ingreso(drive, ruta):
    """
    Descarga la foto de una activación. Devuelve BytesIO o None.
    Firma compatible con utils.buscar_foto_blob, así que es reemplazo directo.
    """
    try:
        file_id, detalle = resolver_id_foto(drive, ruta)
    except Exception as e:
        print(f"      [ERROR-FOTO] {ruta}: {e}")
        stats["no_encontrada"] += 1
        return None

    if not file_id:
        print(f"      [WARN] Foto no encontrada ({detalle}) → {ruta}")
        stats["no_encontrada"] += 1
        return None

    if detalle.startswith("BÚSQUEDA GLOBAL"):
        print(f"      [OJO] {detalle}")

    try:
        req = drive.files().get_media(fileId=file_id, supportsAllDrives=True)
        buf = BytesIO()
        downloader = MediaIoBaseDownload(buf, req)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        buf.seek(0)
        stats["ok"] += 1
        return buf
    except Exception as e:
        print(f"      [ERROR-FOTO] No se pudo descargar '{ruta}': {e}")
        stats["no_encontrada"] += 1
        return None


def resumen_fotos():
    """Línea de resumen para el final de la corrida."""
    partes = [f"{stats['ok']} descargada(s)"]
    if stats["no_encontrada"]:
        partes.append(f"{stats['no_encontrada']} NO encontrada(s)")
    if stats["global"]:
        partes.append(f"{stats['global']} por búsqueda global (revisar)")
    return "Fotos: " + ", ".join(partes)
