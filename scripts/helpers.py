"""Helpers para la generación del PPT mensual del PMO TI Banreservas.

Estas funciones encapsulan los patrones más usados al generar el
"Reporte de Avance de Proyectos Banreservas {Mes} {Año}.pptx" desde el
workbook Excel correspondiente. Diseñadas para usarse con `python-pptx`
y `openpyxl`.

Importar con:

    from scripts.helpers import (
        COLORS, FONTS, EMU, FOOTER_TOP,
        duplicate_slide, replace_run,
        remove_extra_table_rows, remove_shape,
        clean_section_title, posicionar_riesgos,
    )
"""

from __future__ import annotations

import re
from copy import deepcopy
from typing import Iterable, Sequence

# ---------------------------------------------------------------------------
# Paleta oficial Banreservas (tokens)
# ---------------------------------------------------------------------------

COLORS = {
    # Azules
    "azul_br": "#00558C",          # Banreservas primario
    "azul_claro": "#00AEF0",       # acentos portada
    "azul_gantt_medio": "#77A5C5", # barra % avance Gantt
    "azul_gantt_claro": "#C9D1D9", # barra duración Gantt
    "azul_muy_claro": "#F3F6FB",   # fondo notas, headers
    # Naranjas / rojos
    "naranja_br": "#FF9E1B",
    "naranja_marker": "#F2A340",   # marcador "Hoy" en Gantt
    "naranja_oscuro": "#EF6C00",   # impacto ALTO badge, EN RIESGO
    "rojo_alerta": "#BC2D2C",      # OFF TRACK, BLOQUEADO
    "rojo_impacto": "#E74C3C",     # fondo impacto ALTO
    "amarillo_impacto": "#F1C40F", # fondo impacto MEDIO
    # Verdes
    "verde_primario": "#00B98B",   # check de notas
    "verde_on_track_txt": "#1B5E20",
    "verde_on_track_fondo": "#E8F5E9",
    "verde_impacto": "#27AE60",
    # Grises
    "gris_texto": "#232529",
    "gris_secundario": "#63666A",
    "gris_bordes": "#E1E5E9",
    "gris_linea": "#9AA3AB",
    "blanco": "#FFFFFF",
}

# Atajos por uso semántico (estado de entregables / impacto / seguimiento)
ESTADO_ENTREGABLE = {
    "COMPLETADO": {"color": "#63666A", "bold": False, "italic": False},
    "EN CURSO": {"color": "#00558C", "bold": True, "italic": False},
    "PENDIENTE": {"color": "#63666A", "bold": False, "italic": True},
    "BLOQUEADO": {"color": "#BC2D2C", "bold": True, "italic": False},
    "PLANIFICADO": {"color": "#EF6C00", "bold": False, "italic": False},
}

IMPACTO_RIESGO = {
    "ALTO": {"fondo": "#E74C3C", "texto": "#FFFFFF", "bold": True},
    "MEDIO": {"fondo": "#F1C40F", "texto": "#232529", "bold": False},
    "BAJO": {"fondo": "#27AE60", "texto": "#FFFFFF", "bold": False},
}

BADGE_SEGUIMIENTO = {
    "ON TRACK":     {"texto": "✓ ON TRACK",     "fondo": "#E8F5E9", "borde": "#1B5E20"},
    "EN RIESGO":    {"texto": "⚠ EN RIESGO",    "fondo": "#FFF3E0", "borde": "#EF6C00"},
    "OFF TRACK":    {"texto": "✗ OFF TRACK",    "fondo": "#FFEBEE", "borde": "#BC2D2C"},
    "EN REVISIÓN":  {"texto": "⟳ EN REVISIÓN",  "fondo": "#F3F6FB", "borde": "#63666A"},
}

# ---------------------------------------------------------------------------
# Constantes geométricas (unidades EMU absolutos; 1 unidad python-pptx = 9525)
# ---------------------------------------------------------------------------

EMU = 9525                  # multiplicador a EMU absolutos
SLIDE_W = 12192000          # 13.333"
SLIDE_H = 6858000           # 7.5"
FOOTER_TOP = 690            # en unidades EMU/9525

# Caja superior (5 columnas), unidades EMU/9525:
CAJA_SUPERIOR = {
    "AvCard":       (141,  63, 1148, 65),
    "AvHdr_alc":    (148,  72,  333, None),
    "Alcance_txt":  (161,  61,  422, None),
    "AvHdr_est":    (148, 496,  333, None),
    "AvNum":        (170, 496,  104, None),
    "AvTrk":        (179, 605,  224, None),  # bar bg
    "SegBtn":       (171, 854,  130, None),
    "FIHdr":        (148, 990,  100, None),
    "FIValue":      (173, 990,  100, None),
    "FFHdr":        (148, 1100, 110, None),
    "FFValue":      (173, 1100, 110, None),
}

# Tipografía (puntos)
FONTS = {
    "portada_titulo": 54,
    "portada_subtitulo": 24,
    "portada_fecha": 16,
    "gantt_titulo": 40,
    "gantt_header": 14,
    "gantt_proyectos": 12,
    "detalle_projname": 32,
    "detalle_eyebrow": 13,
    "detalle_estatus_pct": 32,
    "detalle_titulo_seccion": 14,
    "detalle_tabla": 10,
    "detalle_tabla_densa": 8,
    "detalle_tabla_compresion": 7.5,
    "detalle_notas": 10.5,
    "detalle_footer": 8,
}

# Límites de chars para resúmenes coherentes (sin elipsis)
LIMITES_RESUMEN = {
    "nota_relevante": 140,
    "riesgo_descripcion": 109,
    "riesgo_mitigacion": 109,
}

NS_DRAWINGML = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

# ---------------------------------------------------------------------------
# Manipulación de slides y shapes
# ---------------------------------------------------------------------------

def duplicate_slide(prs, slide_idx: int):
    """Clona un slide con todos sus shapes, sin reusar placeholders.

    Útil cuando aparece un proyecto nuevo en el Índice que no existía en
    el PPT del mes anterior: se clona un slide-modelo (por convención
    el último de detalle, p. ej. Paperless) y se puebla con los datos
    del proyecto nuevo.
    """
    source = prs.slides[slide_idx]
    new_slide = prs.slides.add_slide(source.slide_layout)
    # Limpiar shapes heredados del layout
    for shp in list(new_slide.shapes):
        shp.element.getparent().remove(shp.element)
    # Clonar shapes uno por uno
    for shp in source.shapes:
        new_el = deepcopy(shp.element)
        new_slide.shapes._spTree.insert_element_before(new_el, "p:extLst")
    return new_slide


def replace_run(shape, txt) -> None:
    """Reemplaza el texto del primer run de un shape preservando el formato.

    Limpia los runs y párrafos extra para evitar mezclas de estilos. Si
    el texto es None o vacío, deja el shape en blanco.
    """
    tf = shape.text_frame
    if not tf.paragraphs:
        return
    para = tf.paragraphs[0]
    new = str(txt) if txt is not None else ""
    if para.runs:
        para.runs[0].text = new
        for r in para.runs[1:]:
            r.text = ""
    else:
        para.text = new
    for extra in tf.paragraphs[1:]:
        extra.text = ""


def remove_shape(slide, shape) -> None:
    """Elimina un shape del slide por manipulación directa de XML."""
    shape._element.getparent().remove(shape._element)


def remove_extra_table_rows(tbl, keep_rows: int) -> None:
    """Elimina filas sobrantes de una tabla python-pptx vía XML.

    `keep_rows` incluye el header. Las cabeceras se preservan siempre,
    incluso cuando solo hay una fila de datos.
    """
    tbl_el = tbl._tbl
    rows = tbl_el.findall(f".//{NS_DRAWINGML}tr")
    while len(rows) > keep_rows:
        last = rows.pop()
        last.getparent().remove(last)


# ---------------------------------------------------------------------------
# Excel → títulos limpios
# ---------------------------------------------------------------------------

_PREFIX_RE = re.compile(r"^\d+\.\s*")


def clean_section_title(cell_value: str | None) -> str:
    """Limpia el prefijo numérico (`2. `, `3. `, `4. `) de un título de sección.

    Si el valor es None, retorna cadena vacía.
    """
    if cell_value is None:
        return ""
    return _PREFIX_RE.sub("", str(cell_value)).strip()


_FORMULA_SHEET_RE = re.compile(r"=\s*'?([^'!]+)'?!\s*B3", re.IGNORECASE)


def extract_sheet_name_from_formula(formula: str | None) -> str | None:
    """Extrae el sheet name desde una fórmula tipo `='SheetName'!B3` del Índice."""
    if not formula:
        return None
    m = _FORMULA_SHEET_RE.search(str(formula))
    return m.group(1).strip() if m else None


# ---------------------------------------------------------------------------
# Posicionamiento dinámico del bloque Riesgos
# ---------------------------------------------------------------------------

def _bottom(shape) -> int:
    return shape.top + shape.height


def posicionar_riesgos(
    hitos_tbl,
    riesgos_tbl,
    notas_cards: Sequence,
    *,
    gap: int = 24,
    footer_top: int = FOOTER_TOP,
    title_block_h: int = 27 + 3 + 10,
) -> dict:
    """Calcula la posición vertical del bloque Riesgos y reporta si comprime.

    Devuelve un dict con:
      - `rk_top`: posición vertical sugerida del título Riesgos (`RkTitle`).
      - `comprimir`: True si el contenido no entra y debe aplicarse compresión.
      - `min_rk_top`, `max_rk_top`: valores intermedios para debugging.

    No modifica el slide — solo calcula. El caller decide si setea
    `riesgos_tbl.top` y aplica la compresión.
    """
    hitos_bottom = _bottom(hitos_tbl)

    notas_bottoms = [_bottom(c) for c in notas_cards if c is not None]
    last_card_bottom = max(notas_bottoms) if notas_bottoms else 0

    riesgos_h = sum(row.height for row in riesgos_tbl.rows)

    min_rk_top = max(hitos_bottom, last_card_bottom) + gap
    max_rk_top = footer_top - 5 - riesgos_h - title_block_h

    if max_rk_top >= min_rk_top:
        rk_top = min_rk_top
        comprimir = False
    else:
        rk_top = min_rk_top
        comprimir = True

    return {
        "rk_top": rk_top,
        "comprimir": comprimir,
        "min_rk_top": min_rk_top,
        "max_rk_top": max_rk_top,
        "riesgos_h": riesgos_h,
    }


# ---------------------------------------------------------------------------
# Validación final
# ---------------------------------------------------------------------------

def validar_bordes(prs, tolerancia: int = EMU) -> list[str]:
    """Inspecciona todos los shapes del PPT y devuelve violaciones de borde.

    Útil como sanity check antes de entregar el archivo. Retorna lista de
    strings descriptivos; vacía si todo está bien.
    """
    issues: list[str] = []
    for slide_idx, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            right = (shape.left or 0) + (shape.width or 0)
            bottom = (shape.top or 0) + (shape.height or 0)
            if right > SLIDE_W + tolerancia:
                issues.append(f"Slide {slide_idx} · {shape.name}: se sale a la derecha ({right} > {SLIDE_W})")
            if bottom > SLIDE_H + tolerancia:
                issues.append(f"Slide {slide_idx} · {shape.name}: se sale abajo ({bottom} > {SLIDE_H})")
            if (shape.left or 0) < -tolerancia:
                issues.append(f"Slide {slide_idx} · {shape.name}: sale a la izquierda ({shape.left})")
            if (shape.top or 0) < -tolerancia:
                issues.append(f"Slide {slide_idx} · {shape.name}: sale arriba ({shape.top})")
    return issues


def validar_sin_elipsis(prs) -> list[str]:
    """Reporta apariciones de `...`, `…` o elipsis en el PPT.

    Estas no deben existir nunca en el output (regla dura del skill).
    """
    issues: list[str] = []
    for slide_idx, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for p in shape.text_frame.paragraphs:
                txt = "".join(r.text for r in p.runs)
                if "..." in txt or "…" in txt:
                    issues.append(f"Slide {slide_idx} · {shape.name}: contiene elipsis — '{txt[:80]}…'")
    return issues
