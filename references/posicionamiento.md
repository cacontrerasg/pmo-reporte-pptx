# Algoritmo de posicionamiento dinámico

## Problema

Las tablas Entregables y Riesgos varían en altura según cantidad de filas y wrap de texto. Las posiciones fijas del template colapsan en proyectos densos.

---

## Altura estimada de fila

Empíricamente validado con LibreOffice render a 80 dpi. Unidades EMU/9525.

| Tabla | Chars por línea en col principal | 1 línea | 2 líneas | 3 líneas |
|---|---|---|---|---|
| Entregables (col 1) | ~48 | 30 | 42 | 56 |
| Riesgos (col 1 y 3) | ~35 | 30 | 42 | 56 |
| Header | n/a | 26-28 | — | — |

**Modo densidad alta** — aplica cuando `n_hitos + n_riesgos ≥ 9` o `n_hitos ≥ 7`:

- Font: 8 pt.
- Alturas reducidas: 24 / 32 / 42 EMU.

---

## Posicionamiento del bloque Riesgos

```python
hitos_h = sum(row.height for row in hitos_tbl.rows)
hitos_bottom = hitos_tbl.top + hitos_h

riesgos_h = sum(row.height for row in riesgos_tbl.rows)
title_block_h = 27 + 3 + 10  # RkTitle + RkAcc + gap

last_card_bottom = max(LgCard[i].bottom for i in range(4))  # tarjetas Notas

GAP = 24  # EMU/9525

# Riesgos debe ir BAJO la tabla Entregables Y BAJO las tarjetas de Notas:
min_rk_top = max(hitos_bottom, last_card_bottom) + GAP

# Y NO puede pasar el footer:
FOOTER_TOP = 690
max_rk_top = FOOTER_TOP - 5 - riesgos_h - title_block_h

# Posición final:
rk_top = min(min_rk_top, max_rk_top) if max_rk_top >= min_rk_top else min_rk_top
```

`scripts/helpers.py` expone `posicionar_riesgos(hitos_tbl, riesgos_tbl, notes_cards)` que encapsula este cálculo.

---

## Pase de compresión (cuando no cabe)

Si `min_rk_top > max_rk_top` (el contenido no entra en el slide), aplicar en orden:

1. Reducir fuente de tablas a 7.5 pt.
2. Multiplicar altura de filas de Riesgos por 0.7.
3. Multiplicar altura de filas de Entregables por 0.85 (solo si tabla ≥ 6 filas).
4. Recalcular posición.

**Tolerancia aceptable** (referencia Mayo 2026): 17/20 slides limpios sin compresión, 3 con compresión densa (SAS, Conciliación Core, Falcon). Si necesitas comprimir más de ~20 % de los slides, revisa la calidad del Excel — probablemente hay textos sin resumir que deberían pasar por §"Resúmenes coherentes".

---

## Por qué este enfoque y no auto-fit nativo

PowerPoint hace auto-fit de filas con texto wrap, pero no expone el resultado a `python-pptx`. La consecuencia es que las tablas crecen verticalmente sin que el script lo sepa, y empujan al footer. Estimar nosotros la altura real con `len(text) / chars_per_line` da control determinista y nos permite tomar la decisión de comprimir _antes_ de renderizar.
