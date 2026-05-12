# Troubleshooting y casos especiales

Problemas conocidos del template y del flujo, con su solución validada.

---

## 1. El template Abril tiene FIHdr/FIValue off-screen en (-525, -525)

Algunos shapes del template original quedaron ocultos. Para hacer visible FECHA INICIO:

- Mover `FIHdr` a `(T=148, L=990, W=100)`.
- Mover `FIValue` a `(T=173, L=990, W=100)`.
- Comprimir `SegBtn` / `SegBtnTxt` a `W=130` (era 178).
- Mover `FFHdr` / `FFValue` a `L=1100, W=110`.
- Extender `AvCard` a `W=1148` para acomodar.

---

## 2. Los hipervínculos del Índice apuntan a pestañas viejas tras renombrar

`openpyxl` NO actualiza hyperlinks al renombrar hojas. Algoritmo de re-vinculación:

1. Borrar todos los hyperlinks del Índice.
2. Por cada fila de proyecto, parsear la fórmula de col C para extraer el sheet name actual.
3. Recrear: `cell.hyperlink = Hyperlink(ref=coord, location=f"'{sheet}'!A1", display=sheet)`.

---

## 3. Tablas crecen verticalmente más de lo declarado en row.height

PowerPoint hace auto-fit de filas con texto wrap, pero `python-pptx` no expone el resultado. Soluciones:

- Estimar altura real basada en `len(text) / chars_per_line` (ver `posicionamiento.md` §"Altura estimada de fila").
- Aplicar font reducido (8 pt o 7.5 pt) en slides densos.
- Recalcular posición del bloque Riesgos con `helpers.posicionar_riesgos`.

---

## 4. Filas vacías quedan como espacios en blanco

`python-pptx` no tiene API para eliminar filas. Manipular XML directamente:

```python
def remove_extra_table_rows(tbl, keep_rows):
    """keep_rows incluye el header."""
    tbl_el = tbl._tbl
    ns = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    rows = tbl_el.findall(f'.//{ns}tr')
    while len(rows) > keep_rows:
        last = rows.pop()
        last.getparent().remove(last)
```

Ya disponible en `scripts/helpers.py`.

---

## 5. Proyectos nuevos no existen como slide en el PPT del mes anterior

Clonar el slide de un proyecto existente (por convención, el último — Paperless en Mayo 2026) y poblar con los datos del proyecto nuevo:

```python
new_slide = duplicate_slide(prs, paperless_idx)
populate_detail(new_slide, data_by_sheet[sheet])
```

`duplicate_slide` está en `scripts/helpers.py`.

---

## 6. Sheet name excede 31 chars al renombrar

Excel impone un límite duro de 31 caracteres. Usar abreviaciones manteniendo identificadores únicos (H1, H2, H3, F1, V4.0, etc.):

```python
rename_map = {
    'Adopción de Nube - SAS (H1) - Oracle ADB…': 'SAS H1 - Adopción Nube Oracle',   # 29
    'SAS - Cumplimiento: Anti-Money Laundering…': 'SAS H3 - Cumplimiento AML',     # 25
}
```

Caracteres prohibidos en sheet name: `\ / ? * [ ] :`.

---

## 7. Texto desborda y el script lo trunca con `...`

**Esto no debe ocurrir.** Si ves un `...` o `…` en el output, el script falló la regla principal. Soluciones, en orden:

1. Reescribir el texto como frase coherente de ≤ 109 chars (riesgos/mitigación) o ≤ 140 chars (notas) preservando datos cuantitativos. Ver `SKILL.md` §"Resúmenes coherentes".
2. Si aún no cabe: reducir font 1 pt (mínimo 9 pt).
3. Si aún no cabe: aplicar compresión del slide (ver `posicionamiento.md`).

Nunca dejes elipsis en el output final.
