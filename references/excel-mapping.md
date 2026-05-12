# Estructura del Excel de entrada — Mapeo celda → campo PPT

## Convención general

Un libro Excel con **N+4 pestañas**, donde N es el número de proyectos activos.

### Pestañas no-proyecto

| Pestaña | Contenido | Uso en PPT |
|---|---|---|
| `Instrucciones` | Texto explicativo para usuarios humanos | ❌ no se procesa |
| `Índice` | Resumen ejecutivo de toda la cartera (1 fila por proyecto) | Lista canónica de proyectos a renderizar |
| `Plantilla` | Template vacío para crear nuevas hojas de proyecto | ❌ no se procesa |
| `Ejemplo - Plataforma Digital` | Ejemplo ilustrativo | ❌ no se procesa |

### Columnas del Índice

`#`, `Proyecto`, `Dirección`, `PM`, `Estatus`, `% Avance`, `% Desv`, `Fecha Inicio`, `Fecha Fin`, `Interesados TI`.

La columna `Proyecto` usa una fórmula `='SheetName'!B3` para enlazar al nombre real. Para obtener el sheet name, parsea esa fórmula con regex: `r"='?([^'!]+)'?!B3"`.

---

## Hoja de proyecto — Layout estándar

Cada proyecto vive en una pestaña dedicada. El nombre de la pestaña puede estar truncado a 31 caracteres (límite Excel). El nombre real vive en `B3`.

| Fila(s) | Sección | Contenido |
|---|---|---|
| 1 | Header institucional | `B1`: `DIRECCIÓN PMO TI · Vicepresidencia Oficina de Transformación` |
| 2 | Subtítulo | `B2`: `Reporte de Avance de Proyecto` |
| 3 | **Nombre del proyecto** | `B3`: nombre completo (fuente de verdad para PPT) |
| 4 | Fecha de actualización | `E4`: label, `F4`: fecha |
| 5 | `1. INFORMACIÓN GENERAL DEL PROYECTO` | merged B:G |
| 6 | Nombre del Proyecto / Dirección | `B6`/`C6` (=B3) · `E6`/`F6` |
| 7 | Project Manager TI / Patrocinador | `B7`/`C7` · `E7`/`F7` |
| 8 | Project Manager PMO / Fecha Inicio | `B8`/`C8` · `E8`/`F8` |
| 9 | Fecha del Reporte / Fecha Fin | `B9`/`C9` · `E9`/`F9` |
| 10 | Líder Técnico / Seguimiento | `B10` · `E10`/`F10` (ON TRACK / EN RIESGO / OFF TRACK) |
| 11 | % de Avance / % de Desviación | `B11`/`C11` · `E11`/`F11` |
| 12 | Alcance del Proyecto | `B12` (label) · `C12:G12` merged (texto) |
| 13 | `2. ENTREGABLES` | merged B:G (título sección) |
| 14 | Header tabla Entregables | Ref · Entregable · Inicio · Fin · Estado · Comentario |
| 15..N | Filas H-01..H-NN | datos de hitos |
| (sig) | `3. NOTAS RELEVANTES` | merged B:G |
| +1..+4 | Notas (hasta 4) | B: icono (`✓` ⚠️ 📓) · C:G merged: texto |
| (sig) | Leyenda visual | "Leyenda: ✓ Logro del mes · ⚠️ Punto de atención · 📓 Nota relevante" |
| (sig) | `4. RIESGOS Y PLAN DE MITIGACIÓN` | merged B:G |
| +1 | Header tabla Riesgos | Ref · Descripción · Impacto · Plan · Responsable · Deadline |
| +2..+N | Filas R-01..R-NN | datos de riesgos |
| (última) | Footer | fórmula `="Versión: V1.0 · Preparado por: PMO TI — "&C7&" · Confidencial — Uso Interno Banreservas"` |

**Notas importantes:**
- Las posiciones exactas de secciones 3 y 4 dependen de la cantidad de hitos del proyecto — las filas se insertan/eliminan dinámicamente. NO confíes en filas absolutas para 3 y 4: búscalas por prefijo `3. ` / `4. ` en columna B.
- Los títulos de sección llevan prefijo numérico (`2. `, `3. `, `4. `) — el PPT debe **parsearlo y eliminarlo** con `re.sub(r'^\d+\.\s*', '', cell_value)`.
- Fechas en formato `dd-mm-aaaa` o `dd/mm/aaaa` o datetime nativo.

---

## Caja superior del slide detalle (5 columnas)

**Orden:** `ALCANCE · ESTATUS · BADGE · FECHA INICIO · FECHA FIN`

| Columna | % ancho | Origen Excel | Notas render |
|---|---|---|---|
| ALCANCE | 40 % | `C12` (texto merged C12:G12) | Wrap a 2-3 líneas, 10.5 pt |
| ESTATUS | 22 % | `C11` (% Avance) | Número grande 32 pt + barra de progreso |
| BADGE | 13 % | `F10` (Seguimiento) | Cápsula con color según estado |
| FECHA INICIO | 12 % | `F8` | dd/mm/yyyy 12 pt |
| FECHA FIN | 12 % | `F9` | dd/mm/yyyy 12 pt |

---

## Slide de portada — Mapeo

| Campo PPT | Origen Excel |
|---|---|
| `{{titulo}}` | Si existe `Metadata.titulo`, úsala. Si no: `Portafolio de Proyectos TI` (default fijo) |
| `{{subtitulo}}` | Si existe `Metadata.subtitulo`, úsala. Default: `Vicepresidencia Ejecutiva de Tecnología y Operaciones` |
| `{{fecha_reporte_texto}}` | Derivar del nombre del libro `... {Mes} {Año}.xlsx` |

## Slide Gantt — Mapeo

| Campo PPT | Origen |
|---|---|
| Título | Fijo: `Portafolio de Proyectos` |
| Eje Y | Nombre del proyecto (**`B3` de cada hoja**, no el nombre de la pestaña) |
| Fechas barras | `F8` / `F9` (inicio / fin) de cada hoja |
| % de avance | `C11` de cada hoja |
| Fecha "Hoy" | `today()` en formato `DD mes` (mes español, minúsculas) |

## Slides de detalle — Mapeo del header del slide

| Shape PPT | Origen Excel |
|---|---|
| Eyebrow | `{{subtitulo}}` (mismo que portada) |
| `ProjName` | **`B3` de la hoja del proyecto** |
| `HtTitle` | Celda con prefijo `2. ` en columna B (parsear) |
| `LgTitle` | Celda con prefijo `3. ` en columna B (parsear) |
| `RkTitle` | Celda con prefijo `4. ` en columna B (parsear) |
