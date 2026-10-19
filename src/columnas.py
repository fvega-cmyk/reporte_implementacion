"""
Resolución de columnas por NOMBRE de encabezado, no por posición.

EL PROBLEMA QUE RESUELVE
------------------------
Antes, config.IDX tenía las posiciones escritas a mano ("PROCESO": 29). Cuando
alguien insertaba una columna en la planilla, todas las de la derecha se
corrían y el código empezaba a leer la columna de al lado. Lo peor: no fallaba.
Leía basura, no matcheaba ningún filtro, y el log decía "0 filas" como si ese
fuera el resultado correcto.

Ahora el IDX se arma leyendo la fila 1 de la planilla y buscando cada campo por
su nombre. Podés insertar, mover o reordenar columnas: el código se acomoda.

QUÉ PASA SI FALTA UNA COLUMNA
-----------------------------
  - Si es OBLIGATORIA (las que usan los filtros): corta con error y te dice
    cuál falta y qué encabezados sí encontró. Preferimos que explote a que
    devuelva "0 filas" en silencio.
  - Si es OPCIONAL (fotos, OT, PDFs): queda apuntando a una columna virtual
    vacía, se avisa en el log, y el reporte sale sin ese dato.

SI EL NOMBRE NO MATCHEA
-----------------------
Agregá el nombre real a ALIAS, o fijá la posición a mano en
config.COLUMNAS_OVERRIDE, que siempre gana.
"""
from utils import normalizar


class ColumnaFaltante(RuntimeError):
    """Falta una columna obligatoria en la planilla."""


# ============================================================
# NOMBRES POSIBLES DE CADA CAMPO
# ============================================================
# Se comparan normalizados: sin tildes, sin signos, en minúscula y con los
# espacios colapsados. "FECHA DE TÉRMINO" y "fecha de termino" son lo mismo.
# El primero que matchee gana, así que poné primero el nombre más probable.
ALIAS = {
    "ID_ACTIVIDAD":     ["id actividad", "id de actividad"],
    "ID_CAMPANA":       ["id campana", "id de campana"],
    "N_PRESUPUESTO":    ["n presupuesto", "no presupuesto", "numero presupuesto"],
    "CAMPANA":          ["campana", "nombre campana", "nombre de campana"],
    "EJECUTIVO":        ["ejecutivo", "ejecutiva", "ejecutivo a cargo"],
    "CLIENTE":          ["cliente"],
    "ID_LOCAL":         ["id local", "id de local", "id sala"],
    "CADENA":           ["cadena"],
    "COD":              ["cod", "codigo", "cod sala", "codigo sala", "cod local"],
    "NOMBRE_SALA":      ["nombre sala", "nombre local", "nombre de sala", "nombre"],
    "DIRECCION":        ["direccion"],
    "COMUNA":           ["comuna"],
    "REGION":           ["region"],
    "MARCA":            ["marca"],
    "CATEGORIA":        ["categoria"],
    "GUIA":             ["n guia", "no guia", "numero guia", "guia"],
    "SKU":              ["sku"],
    "MATERIAL":         ["material"],
    "CANTIDAD":         ["cantidad"],
    "FECHA_INICIO":     ["fecha inicio", "fecha de inicio"],
    "FECHA_TERMINO":    ["fecha termino", "fecha de termino", "fecha fin"],
    "FECHA_TENTATIVA":  ["fecha tentativa", "fecha de tentativa"],
    "FECHA_ENTREGA":    ["fecha entrega", "fecha de entrega"],
    "REAGENDA":         ["reagenda", "se reagenda", "reagendamiento"],
    "N_REAGENDA":       ["n reagenda", "no reagenda", "numero reagenda"],
    "FECHA_REAGENDA":   ["fecha reagenda", "fecha de reagenda"],
    "FECHA_COMPROMISO": ["fecha compromiso", "fecha de compromiso"],
    "HORA_INICIO":      ["hora inicio", "hora de inicio"],
    "HORA_TERMINO":     ["hora termino", "hora de termino"],
    "PROCESO":          ["proceso", "estado proceso", "estado del proceso"],
    "DETALLE":          ["detalle"],
    "OBSERVACIONES":    ["observaciones", "observacion"],
    "LOGISTICO1":       ["logistico 1", "logistico1", "logistico"],
    "LOGISTICO2":       ["logistico 2", "logistico2"],
    "ACTIVIDAD":        ["actividad"],
    "TIPO_ACTIVIDAD":   ["tipo actividad", "tipo de actividad"],
    "FOTO1":            ["foto 1", "foto1"],
    "FOTO2":            ["foto 2", "foto2"],
    "FOTO3":            ["foto 3", "foto3"],
    "FOTO4":            ["foto 4", "foto4"],
    "VIDEO":            ["video"],
    "TIPO_ENVIO":       ["tipo envio", "tipo de envio"],
    "COURIER":          ["courier", "transporte"],
    "OT":               ["ot", "n ot", "numero ot"],
    "SEGUIMIENTO":      ["seguimiento", "n seguimiento", "numero seguimiento"],
    "CREACION_PDF":     ["creacion pdf", "creacion de pdf"],
    "INFORME_PDF":      ["informe pdf", "informe de pdf"],
}

# ============================================================
# OBLIGATORIAS: si falta una de estas, el proceso corta
# ============================================================
# Son las que usan los FILTROS. Si una se corre, el resultado no es un reporte
# incompleto: es un reporte vacío o, peor, uno con datos de otras filas.
OBLIGATORIAS = {
    "CAMPANA", "CLIENTE", "ACTIVIDAD", "TIPO_ACTIVIDAD", "PROCESO",
    "FECHA_INICIO", "FECHA_TERMINO", "FECHA_ENTREGA", "FECHA_COMPROMISO",
    "FECHA_TENTATIVA", "FECHA_REAGENDA", "MATERIAL", "CANTIDAD",
    "COD", "NOMBRE_SALA", "ID_LOCAL",
}


# Signos que se descartan al comparar encabezados. OJO: solo puntuación y
# símbolos. Nunca letras: un "n" suelto acá borraría la N de CAMPAÑA, CLIENTE
# y CANTIDAD, y todos esos campos dejarían de resolverse.
_SIGNOS = "°ºª¿?¡!().,:;\"'*#/\\|-_[]{}+"


def limpiar(texto):
    """Normaliza un encabezado para poder compararlo."""
    t = normalizar(str(texto or ""))
    for ch in _SIGNOS:
        t = t.replace(ch, " ")
    return " ".join(t.split())


def letra_columna(i):
    """0 -> A, 25 -> Z, 26 -> AA, 60 -> BI. Para mensajes legibles."""
    letras = ""
    i += 1
    while i > 0:
        i, resto = divmod(i - 1, 26)
        letras = chr(65 + resto) + letras
    return letras


def _buscar(limpios, nombres):
    """
    Busca un campo entre los encabezados. Devuelve (indice, modo).

    Tres pasadas, de más estricta a menos:
      1. exacto      → el encabezado es igual a un alias.
      2. prefijo     → el encabezado EMPIEZA con un alias ("cantidad bultos").
      3. contiene    → el alias aparece dentro del encabezado.

    Las pasadas 2 y 3 exigen que haya UNA sola coincidencia. Si hay dos, se
    devuelve ambiguo en vez de elegir al azar: con columnas como "DIRECCION" y
    "DIRECCION ENVIO" elegir mal es peor que no elegir.
    """
    for nombre in nombres:
        for i, h in enumerate(limpios):
            if h == nombre:
                return i, "exacto"

    for nombre in nombres:
        candidatos = [i for i, h in enumerate(limpios) if h and h.startswith(nombre)]
        if len(candidatos) == 1:
            return candidatos[0], "prefijo"
        if len(candidatos) > 1:
            return None, "ambiguo"

    for nombre in nombres:
        candidatos = [i for i, h in enumerate(limpios) if h and nombre in h]
        if len(candidatos) == 1:
            return candidatos[0], "contiene"
        if len(candidatos) > 1:
            return None, "ambiguo"

    return None, None


def construir_idx(headers, overrides=None, verbose=True):
    """
    Arma el diccionario IDX desde la fila de encabezados.

    Devuelve (idx, informe) donde informe trae:
        resueltos  {campo: (indice, modo)}
        aproximados [campo, ...]   resueltos por prefijo/contiene
        faltantes   [campo, ...]   opcionales no encontrados
        ambiguos    [campo, ...]
        ancho       int, columnas a las que hay que rellenar cada fila

    Lanza ColumnaFaltante si falta alguna columna OBLIGATORIA.
    """
    overrides = overrides or {}
    limpios = [limpiar(h) for h in headers]
    # Los alias se limpian con la MISMA función que los encabezados, para que
    # las dos puntas lleguen normalizadas igual.
    alias_limpios = {campo: [limpiar(a) for a in nombres]
                     for campo, nombres in ALIAS.items()}
    n = len(headers)

    # Columna virtual: índice n, siempre vacía. Ahí apuntan las opcionales
    # que no existen en la planilla, para que row[IDX[...]] no reviente.
    COLUMNA_VACIA = n

    idx = {}
    resueltos, aproximados, faltantes, ambiguos = {}, [], [], []

    for campo in ALIAS:
        if campo in overrides:
            idx[campo] = overrides[campo]
            resueltos[campo] = (overrides[campo], "override")
            continue

        i, modo = _buscar(limpios, alias_limpios[campo])

        if i is not None:
            idx[campo] = i
            resueltos[campo] = (i, modo)
            if modo != "exacto":
                aproximados.append(campo)
        else:
            if modo == "ambiguo":
                ambiguos.append(campo)
            idx[campo] = COLUMNA_VACIA
            faltantes.append(campo)

    criticas = [c for c in faltantes if c in OBLIGATORIAS]
    if criticas:
        detalle = "\n".join(
            f"    [{i:>3}] {letra_columna(i):>3}  {h}" for i, h in enumerate(headers)
        )
        raise ColumnaFaltante(
            "No se encontraron estas columnas OBLIGATORIAS en la hoja:\n"
            + "".join(f"    - {c}  (nombres probados: {ALIAS[c]})\n" for c in criticas)
            + "\n  Encabezados que SÍ tiene la planilla:\n" + detalle
            + "\n\n  Arreglo: agregá el nombre real a ALIAS en src/columnas.py, "
              "o fijá la posición a mano en config.COLUMNAS_OVERRIDE."
        )

    informe = {
        "resueltos": resueltos,
        "aproximados": aproximados,
        "faltantes": faltantes,
        "ambiguos": ambiguos,
        "ancho": COLUMNA_VACIA + 1,
    }

    if verbose:
        _log(headers, informe)

    return idx, informe


def _log(headers, informe):
    """Resumen corto en el log. Solo grita cuando hay algo raro."""
    n_ok = sum(1 for _, (_, m) in informe["resueltos"].items() if m == "exacto")
    print(f"Columnas: {n_ok}/{len(ALIAS)} resueltas por nombre exacto "
          f"({len(headers)} columnas en la hoja)")

    for campo in informe["aproximados"]:
        i, modo = informe["resueltos"][campo]
        print(f"  [APROX] {campo} -> [{i}] {letra_columna(i)} '{headers[i]}' ({modo})")

    for campo in informe["ambiguos"]:
        print(f"  [AMBIGUO] {campo}: varios encabezados coinciden. Fijalo en "
              f"config.COLUMNAS_OVERRIDE.")

    opcionales = [c for c in informe["faltantes"] if c not in informe["ambiguos"]]
    if opcionales:
        print(f"  [SIN COLUMNA] {', '.join(opcionales)} -> se leen vacías")


def sincronizar_idx(headers, verbose=True):
    """
    Arma el IDX desde los encabezados y lo escribe DENTRO de config.IDX.

    Se modifica el diccionario en su lugar (clear + update) a propósito: los
    módulos hacen `from config import IDX` al importar, así que todos tienen
    una referencia al MISMO objeto. Reasignar config.IDX = {...} no los
    actualizaría; mutarlo sí.

    Devuelve el informe de construir_idx.
    """
    import config

    idx, informe = construir_idx(
        headers,
        overrides=getattr(config, "COLUMNAS_OVERRIDE", None),
        verbose=verbose,
    )
    config.IDX.clear()
    config.IDX.update(idx)
    config.ANCHO_FILA = informe["ancho"]
    return informe
