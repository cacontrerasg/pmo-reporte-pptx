# Paleta, tipografía y constantes

## Paleta oficial Banreservas

| Token | Hex | Uso |
|---|---|---|
| Azul Banreservas | `#00558C` | Títulos, tablas header, progreso |
| Azul claro | `#00AEF0` | Acentos portada |
| Azul medio Gantt | `#77A5C5` | Barra % avance Gantt |
| Azul gantt claro | `#C9D1D9` | Barra duración Gantt |
| Azul muy claro | `#F3F6FB` | Fondo notas, headers tabla |
| Naranja BR | `#FF9E1B` / `#F2A340` | Marcador "Hoy", Impacto MEDIO |
| Naranja oscuro | `#EF6C00` | Impacto ALTO, EN RIESGO |
| Rojo alerta | `#BC2D2C` | OFF TRACK, BLOQUEADO |
| Rojo impacto | `#E74C3C` | Fondo Impacto ALTO |
| Amarillo impacto | `#F1C40F` | Fondo Impacto MEDIO |
| Verde primario | `#00B98B` | Check notas |
| Verde ON TRACK txt | `#1B5E20` | Badge ON TRACK texto |
| Verde ON TRACK fondo | `#E8F5E9` | Badge ON TRACK fondo |
| Verde impacto | `#27AE60` | Fondo Impacto BAJO |
| Gris texto | `#232529` | Texto general |
| Gris secundario | `#63666A` | Etiquetas, mensajes vacíos |
| Gris bordes | `#E1E5E9` | Tablas, tarjetas |
| Gris línea | `#9AA3AB` | Underline corto bajo títulos |
| Blanco | `#FFFFFF` | Fondo slides 2-N |

Estos tokens están disponibles como constantes en `scripts/helpers.py` (módulo `COLORS`).

---

## Tipografía

- **Preferida:** Aptos Display (template Banreservas).
- **Fallbacks:** Calibri / Arial.
- **En cuerpo del Excel:** Open Sans 10.

| Elemento | Tamaño |
|---|---|
| Portada — Título | 54 pt |
| Portada — Subtítulo | 24 pt |
| Portada — Fecha | 16 pt |
| Gantt — Título | 40 pt |
| Gantt — Header cols | 14 pt |
| Gantt — Proyectos | 12 pt |
| Detalle — ProjName | 32 pt |
| Detalle — Eyebrow | 13 pt |
| Detalle — % ESTATUS | 32 pt |
| Detalle — Títulos sección | 14 pt |
| Detalle — Tablas | 10 pt (8 pt en slides densos, 7.5 pt en compresión extrema) |
| Detalle — Notas | 10.5 pt |
| Detalle — Footer | 8 pt |

---

## Constantes geométricas

- **Slide:** `12192000 × 6858000 EMU` (1 EMU para python-pptx = `9525 absolutos`).
- **Footer:** `T=690 EMU/9525` (≈ `6571500 EMU absolutos`).
- **Tabla Entregables:** ancho ~682 unidades, `left=63`.
- **Tabla Riesgos:** ancho ~1118 unidades, `left=63`.
- **Notas (columna derecha):** `left=766`, `width=416`.

---

## Convenciones de formato (Excel fuente)

- Fechas: `dd-mm-aaaa`.
- Tipografía cuerpo: Open Sans 10.
- Dropdown Notas Relevantes: `✓` Logro del mes · `⚠️` Punto de atención · `📓` Nota relevante.
- Semáforo Impacto: ALTO rojo · MEDIO amarillo · BAJO verde.
- Estado Entregables: COMPLETADO verde · EN CURSO azul · PENDIENTE gris · BLOQUEADO rojo · PLANIFICADO amarillo.
