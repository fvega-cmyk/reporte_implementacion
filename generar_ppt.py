"""
Genera el PPT del reporte de fotos.
Equivalente a generarSlides() en Apps Script, pero con python-pptx (mucho más rápido).

ESTRATEGIA:
1. Descargar el template Google Slides como .pptx (vía Drive API export).
2. Abrirlo con python-pptx.
3. La primera slide es la portada, la segunda es la plantilla de sala (se duplica N veces).
4. Reemplazar placeholders [CAMPAÑA], [FOTO 1], etc.
5. Insertar las imágenes en las posiciones donde estaban los placeholders.
"""
import copy
from io import BytesIO

from pptx import Presentation
from pptx.util import Emu

from googleapiclient.http import MediaIoBaseDownload

from config import IDX, TEMPLATE_PPT_ID
from utils import fmt, san, buscar_foto_blob
from google_clients import get_drive


def _descargar_template_como_pptx():
    """Exporta el Google Slides template a .pptx en memoria."""
    drive = get_drive()
    req = drive.files().export_media(
        fileId=TEMPLATE_PPT_ID,
        mimeType="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
    buf = BytesIO()
    downloader = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    return buf


def _iter_text_frames(slide):
    """Genera todos los text_frames de la slide (incluyendo dentro de tablas)."""
    for shape in slide.shapes:
        if shape.has_text_frame:
            yield shape, shape.text_frame
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    yield shape, cell.text_frame


def _reemplazar_texto(slide, buscar, reemplazar):
    """Reemplaza texto en toda la slide, preservando formato del primer run."""
    reemplazar = str(reemplazar or "")
    for _, tf in _iter_text_frames(slide):
        for para in tf.paragraphs:
            full = "".join(run.text for run in para.runs)
            if buscar in full:
                nuevo = full.replace(buscar, reemplazar)
                # Limpiar todos los runs menos el primero, conservando su formato
                if para.runs:
                    para.runs[0].text = nuevo
                    for run in para.runs[1:]:
                        run.text = ""


def _encontrar_placeholders_fotos(slide):
    """
    Busca las shapes que contienen [FOTO 1]..[FOTO 4] y devuelve sus posiciones.
    Retorna dict {1: shape, 2: shape, ...}
    """
    encontrados = {}
    for shape in list(slide.shapes):
        if not shape.has_text_frame:
            continue
        txt = shape.text_frame.text
        for i in range(1, 5):
            if f"[FOTO {i}]" in txt:
                encontrados[i] = shape
                break
    return encontrados


def _calcular_posiciones(n_fotos, placeholders_info):
    """
    Calcula posiciones (left, top, width, height) según cantidad real de fotos.
    Espejo de calcularPosicionesFotos() en Apps Script.
    placeholders_info: lista de (left, top, width, height) en EMU.
    """
    if n_fotos == 0 or not placeholders_info:
        return []

    lefts = [p[0] for p in placeholders_info]
    tops = [p[1] for p in placeholders_info]
    rights = [p[0] + p[2] for p in placeholders_info]
    bottoms = [p[1] + p[3] for p in placeholders_info]

    min_x, min_y = min(lefts), min(tops)
    max_x, max_y = max(rights), max(bottoms)
    total_w, total_h = max_x - min_x, max_y - min_y
    gap = Emu(80000)  # ~8pt de separación

    if n_fotos == 1:
        return [(min_x, min_y, total_w, total_h)]
    if n_fotos == 2:
        w = (total_w - gap) // 2
        return [
            (min_x, min_y, w, total_h),
            (min_x + w + gap, min_y, w, total_h),
        ]
    if n_fotos == 3:
        if len(placeholders_info) == 3:
            return placeholders_info[:3]
        w = (total_w - gap * 2) // 3
        return [(min_x + i * (w + gap), min_y, w, total_h) for i in range(3)]
    # 4 fotos: grilla 2x2
    w = (total_w - gap) // 2
    h = (total_h - gap) // 2
    return [
        (min_x, min_y, w, h),
        (min_x + w + gap, min_y, w, h),
        (min_x, min_y + h + gap, w, h),
        (min_x + w + gap, min_y + h + gap, w, h),
    ]


def _duplicar_slide(pres, slide_origen):
    """
    Duplica una slide. python-pptx no tiene .duplicate() nativo, así que
    copiamos el XML manualmente.
    """
    blank_layout = slide_origen.slide_layout
    nueva = pres.slides.add_slide(blank_layout)
    # Borrar los placeholders del layout en la nueva slide
    for shape in list(nueva.shapes):
        sp = shape._element
        sp.getparent().remove(sp)
    # Copiar shapes del slide origen
    for shape in slide_origen.shapes:
        el = shape._element
        nueva.shapes._spTree.insert_element_before(copy.deepcopy(el), "p:extLst")
    return nueva


def _mover_slide_a_posicion(pres, slide, posicion):
    """Mueve una slide a la posición dada (0-indexed)."""
    xml_slides = pres.slides._sldIdLst
    slides_list = list(xml_slides)
    # Encontrar el elemento de esta slide
    for el in slides_list:
        if el.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        ) == slide.part.partname:
            # Difícil identificar 1:1; usamos el rId
            pass
    # Más simple: trabajamos con el último (recién agregado) y lo movemos
    new_el = slides_list[-1]
    xml_slides.remove(new_el)
    xml_slides.insert(posicion, new_el)


def _limpiar_placeholder_foto(shape):
    """Borra el shape del placeholder de foto."""
    sp = shape._element
    sp.getparent().remove(sp)


def generar_ppt(campana, filas, hoy):
    """
    Genera el PPT y devuelve los bytes.
    Equivalente a generarSlides() en Apps Script.
    """
    drive = get_drive()
    template_buf = _descargar_template_como_pptx()
    pres = Presentation(template_buf)

    fecha_str = hoy.strftime("%d/%m/%Y")
    cliente = str(filas[0][IDX["CLIENTE"]] or "")

    # ---- PORTADA (slide 0) ----
    portada = pres.slides[0]
    _reemplazar_texto(portada, "[FECHA ENVIO DEL REPORTE]", fecha_str)
    _reemplazar_texto(portada, "[CAMPAÑA]", campana)
    _reemplazar_texto(portada, "[CLIENTE]", cliente)

    # ---- AGRUPAR POR SALA ----
    salas_map = {}
    for row in filas:
        key = f"{row[IDX['COD']] or ''}__{row[IDX['NOMBRE_SALA']] or ''}"
        salas_map.setdefault(key, []).append(row)

    sala_keys = list(salas_map.keys())
    print(f"  Salas: {len(sala_keys)}")

    # Slide template de sala (la segunda en el template)
    template_sala = pres.slides[1]

    # Lista de slides ya creadas para sala
    slides_de_sala = [template_sala]

    # Duplicar para las salas adicionales (la primera reutiliza template_sala)
    for _ in range(len(sala_keys) - 1):
        nueva = _duplicar_slide(pres, template_sala)
        slides_de_sala.append(nueva)

    # ---- LLENAR CADA SLIDE DE SALA ----
    for sala_idx, key in enumerate(sala_keys):
        filas_sala = salas_map[key]
        r0 = filas_sala[0]
        slide = slides_de_sala[sala_idx]
        print(f"    Sala {sala_idx + 1}/{len(sala_keys)}: {key}")

        # Reemplazar textos
        _reemplazar_texto(
            slide, "[CÓD.] - [NOMBRE SALA]",
            f"{r0[IDX['COD']] or ''} - {r0[IDX['NOMBRE_SALA']] or ''}",
        )
        _reemplazar_texto(
            slide, "[DIRECCIÓN], [COMUNA] - [REGION]",
            f"{r0[IDX['DIRECCION']] or ''}, {r0[IDX['COMUNA']] or ''} - {r0[IDX['REGION']] or ''}",
        )
        _reemplazar_texto(slide, "[MATERIAL]", r0[IDX["MATERIAL"]] or "")
        _reemplazar_texto(slide, "[CANTIDAD]", r0[IDX["CANTIDAD"]] or "")
        _reemplazar_texto(slide, "[PROCESO]", r0[IDX["PROCESO"]] or "")
        _reemplazar_texto(slide, "[FECHA ENTREGA]", fmt(r0[IDX["FECHA_ENTREGA"]]))

        # Recolectar rutas únicas (max 4)
        rutas = []
        for row in filas_sala:
            for fi in [IDX["FOTO1"], IDX["FOTO2"], IDX["FOTO3"], IDX["FOTO4"]]:
                ruta = (row[fi] or "").strip()
                if ruta and ruta not in rutas:
                    rutas.append(ruta)
        rutas = rutas[:4]

        # Localizar placeholders de foto y sus posiciones
        placeholders = _encontrar_placeholders_fotos(slide)
        placeholders_info = []
        for i in sorted(placeholders.keys()):
            sh = placeholders[i]
            placeholders_info.append((sh.left, sh.top, sh.width, sh.height))

        # Borrar todos los placeholders de foto (los reemplazaremos por imágenes)
        for sh in placeholders.values():
            _limpiar_placeholder_foto(sh)

        # Calcular posiciones según cantidad real de fotos
        posiciones = _calcular_posiciones(len(rutas), placeholders_info)

        # Insertar fotos
        for i, ruta in enumerate(rutas):
            if i >= len(posiciones):
                break
            left, top, width, height = posiciones[i]
            blob = buscar_foto_blob(drive, ruta)
            if blob is None:
                print(f"      [WARN] Foto {i+1} no encontrada: {ruta}")
                continue
            try:
                slide.shapes.add_picture(blob, left, top, width=width, height=height)
                print(f"      OK foto {i+1}")
            except Exception as e:
                print(f"      [ERROR] foto {i+1}: {e}")

    # Guardar a bytes
    buf = BytesIO()
    pres.save(buf)
    buf.seek(0)
    return buf.getvalue()
