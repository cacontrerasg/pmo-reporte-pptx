# Layout del PPT — Slide a slide

Especificación visual del archivo `.pptx` 16:9 (13.333" × 7.5" = `12192000 × 6858000 EMU`). Las coordenadas concretas están en unidades `EMU/9525` (lo que `python-pptx` usa internamente como `Emu`).

---

## Slide 1 — Portada

- **Fondo:** rectángulo azul Banreservas `#00558C` cubriendo todo el slide.
- **Forma decorativa blanca:** cuarto de círculo redondeado en esquina inferior derecha.
- **Título** `{{titulo}}` — centrado, blanco, bold, 54 pt, Aptos Display o Calibri Bold.
- **Línea decorativa** — azul claro `#00AEF0`, 3 pt, centrada bajo título.
- **Subtítulo** `{{subtitulo}}` — blanco bold, 24 pt.
- **Fecha** `{{fecha_reporte_texto}}` — blanco bold, 16 pt.
- **Logo Banreservas** — abajo a la derecha, sobre la zona blanca (~2.4" × 1.0").

---

## Slide 2 — Roadmap consolidado (Gantt)

- **Título** `Portafolio de Proyectos` — azul `#00558C` bold 40 pt, arriba izquierda.
- **Subtítulo** — gris `#63666A` 11 pt: `Año detallado: {{anio}} · Comprimidos: {{anios_comprimidos}} · Corte: {{fecha_corte}}`.
- **Logo** — arriba derecha, ~1.5" × 0.7".
- **Leyenda** — `Duración:` (círculo `#C9D1D9`) y `% Avance:` (círculo `#77A5C5`).
- **Header columnas** — banda azul `#00558C` con texto blanco bold 14 pt.
  - Columnas estándar: `2024 · 2025 · 26Q1 · 26Q2 · 26Q3 · 26Q4 · 2027`
  - Años comprimidos (no detallados): ~0.4 × ancho de un trimestre.
- **Líneas verticales** — gris claro `#E1E5E9` entre columnas.
- **Eje Y** — nombre de cada proyecto (desde `B3`) en gris oscuro `#232529` 12 pt, alineado izquierda.
- **Marcador "Hoy":**
  - Línea naranja `#F2A340` de 2 pt vertical.
  - Badge naranja con texto blanco bold: `DD mes` (mes en español, minúsculas).
- **Por proyecto, 2 barras superpuestas:**
  - Duración: gris azulado `#C9D1D9`, bordes redondeados, de `fecha_inicio` a `fecha_fin`.
  - Avance: azul `#77A5C5`, de `fecha_inicio` a `fecha_inicio + (% * (fin-inicio))`.
  - Etiqueta `XX%` a la derecha del extremo de la barra de avance.

### Cálculo de posición X

```python
ancho_columna = ancho_total / (anios_comprimidos * 0.4 + trimestres_detallados)

# Fecha en año comprimido:
x = x_inicio_columna + (dia_del_año / 365) * ancho_columna

# Fecha en año detallado:
trimestre = (mes - 1) // 3
x = x_inicio_columna_trimestre + (dia_del_trimestre / 90) * ancho_columna_trimestre
```

---

## Slides 3..N — Detalle de Proyecto

Layout fijo replicado por cada proyecto en el Índice.

### Estructura vertical

```
[Header]        Eyebrow (subtitulo) + ProjName (B3) + Logo                          ~T=38-110
[Caja superior] ALCANCE · ESTATUS · BADGE · FECHA INICIO · FECHA FIN                ~T=141-206
[Sección 2]     ENTREGABLES (HtTitle) + tabla 5 columnas                            ~T=231-X
[Sección 3]     NOTAS RELEVANTES (LgTitle) + tarjetas (col derecha, paralela)       ~T=231-X
[Sección 4]     RIESGOS Y PLAN DE MITIGACIÓN (RkTitle) + tabla 6 columnas           ~T=Y-Z
[Footer]        Versión · Autor · Confidencial · Fecha reporte                      ~T=690
[Decorativo]    ⌂ esquina inferior derecha
```

### Coordenadas de la caja superior (5 columnas)

Validadas con el template Abril/Mayo 2026. Unidades EMU/9525.

| Shape | Top | Left | Width | Height |
|---|---|---|---|---|
| AvCard | 141 | 63 | 1148 | 65 |
| AvHdr ALCANCE | 148 | 72 | 333 | — |
| CuadroTexto 2 (alcance) | 161 | 61 | 422 | — |
| AvHdr ESTATUS | 148 | 496 | 333 | — |
| AvNum (%) | 170 | 496 | 104 | — |
| AvTrk (bar bg) | 179 | 605 | 224 | — |
| AvFill (bar fill) | 179 | 605 | `% * 224 / 100` | — |
| SegBtn (badge) | 171 | 854 | 130 | — |
| FIHdr | 148 | 990 | 100 | — |
| FIValue | 173 | 990 | 100 | — |
| FFHdr | 148 | 1100 | 110 | — |
| FFValue | 173 | 1100 | 110 | — |

### Tablas

| Tabla | Columnas | Tipo |
|---|---|---|
| Entregables | Ref · Entregable · Inicio · Fin · Estado | 5 cols, anclada izquierda, ~47 % ancho |
| Riesgos | Ref · Descripción · Impacto · Plan · Responsable · Deadline | 6 cols, ancho completo |

### Estado en Entregables (color del texto)

| Estado | Color |
|---|---|
| `COMPLETADO` | gris `#63666A` |
| `EN CURSO` | azul `#00558C` bold |
| `PENDIENTE` | gris `#63666A` itálica |
| `BLOQUEADO` | rojo `#BC2D2C` bold |
| `PLANIFICADO` | naranja `#EF6C00` |

### Impacto en Riesgos (semáforo de fondo)

| Impacto | Fondo | Texto |
|---|---|---|
| `ALTO` | `#E74C3C` | blanco bold |
| `MEDIO` | `#F1C40F` | oscuro |
| `BAJO` | `#27AE60` | blanco |

### Badge de Seguimiento (F10)

| Estado | Texto badge | Fondo | Borde/texto |
|---|---|---|---|
| ON TRACK | `✓ ON TRACK` | `#E8F5E9` | `#1B5E20` |
| EN RIESGO | `⚠ EN RIESGO` | `#FFF3E0` | `#EF6C00` |
| OFF TRACK | `✗ OFF TRACK` | `#FFEBEE` | `#BC2D2C` |
| EN REVISIÓN | `⟳ EN REVISIÓN` | `#F3F6FB` | `#63666A` |

### Notas Relevantes (columna derecha)

Hasta 4 tarjetas apiladas, cada una con: icono (`✓` / `⚠️` / `📓`) + texto.

- Left fijo: `766`
- Width: `416`
- Top arranca a la altura del título `LgTitle` y crece hacia abajo.
- Si una tarjeta queda vacía: eliminar `LgTxt{i}`, `LgChkTxt{i}` y `LgCard{i}` del XML.

### Iconos de Notas

- `✓` Logro del mes — verde primario `#00B98B`
- `⚠️` Punto de atención — naranja `#F2A340`
- `📓` Nota relevante — azul `#00558C`

---

## Footer común a todos los slides de detalle

- **Top:** `690 EMU/9525` (≈ `6571500 EMU` absolutos).
- **Texto:** generado en Excel con fórmula `="Versión: V1.0 · Preparado por: PMO TI — "&C7&" · Confidencial — Uso Interno Banreservas"`
- **Tamaño:** 8 pt, gris secundario `#63666A`.
- **Decorativo:** ⌂ en esquina inferior derecha.

Ningún contenido del slide debe pasar este límite. Si una tabla de Riesgos empuja contra él, aplicar compresión (ver `posicionamiento.md`).
