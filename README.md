---
name: pmo-reporte-pptx
description: Genera el PPT mensual `Reporte de Avance de Proyectos Banreservas {Mes} {Año}.pptx` desde el Excel del PMO TI (VP Ejecutiva de Tecnología y Operaciones), replicando el modelo `assets/modelo-referencia.pptx` (Abril 2026 actualizado). Actívalo cuando el usuario diga "reporte de avance", "PPT del PMO", "presentación mensual de proyectos TI", "regenerar el PPT", "informe de cartera", "roadmap consolidado", "slide de detalle por proyecto", o adjunte/referencie `Reporte de Avance de Proyectos PMO TI Banreservas - {Mes} {Año}.xlsx`. También para actualizar fechas, regenerar el Gantt, ajustar un slide específico, o aplicar resúmenes coherentes (sin elipsis) en notas/riesgos. Trabaja con `banreservas-marca` (tono, paleta) y `banreservas-pptx` (template); el LAYOUT VINCULANTE sale de `references/modelo-referencia.md` y el archivo en `assets/`.
---

# PMO TI · Generación del Reporte Mensual de Avance de Proyectos

Skill para producir el PPT mensual del portafolio TI de Banreservas a partir del workbook Excel del PMO. La salida es un `.pptx` 16:9 con: 1 portada + 1 roadmap consolidado (Gantt) + N slides de detalle (uno por proyecto del Índice).

> **Versión 3.0** — incorpora `assets/modelo-referencia.pptx` (Reporte de Abril 2026 actualizado por Carlos el 8-may-2026) como **MODELO VINCULANTE** del que toda generación debe replicar layout, tipografía, secciones, colores, footer y decorativos.

---

## ⚠️ MODELO DE REFERENCIA VINCULANTE

**Antes de generar nada, lee `references/modelo-referencia.md` y abre `assets/modelo-referencia.pptx`.**

Este modelo es la fuente de verdad absoluta para el diseño del reporte mensual. Reglas duras:

1. **Replica TODO del modelo**: layout, tipografía, paleta, posiciones, redacción de títulos de sección (`HITOS Y CRONOGRAMA`, `NOTAS GENERALES`, `RIESGOS Y PLAN DE MITIGACIÓN`), footer, casita ⌂ decorativa.
2. **La caja superior tiene 4 columnas**: `ALCANCE · ESTATUS · SEGUIMIENTO · FECHA FIN`. NO incluyas `FECHA INICIO`.
3. **Notas Generales** se renderizan con círculo verde sólido + texto, sin tarjeta de fondo.
4. **Tablas sin alternancia de color** (todas las filas blancas).
5. **Impacto en Riesgos** solo colorea el TEXTO, no la celda completa.
6. **Footer**: `Versión: V1.0 · Dirección Senior de Proyectos TI · Confidencial — Uso Interno Banreservas · Fecha de reporte: dd/mm/yyyy` en itálica.
7. **Portada sin eyebrow** institucional. Solo: título centrado + subtítulo + fecha + logo abajo-derecha.
8. **No agregar slide de cierre azul** (Master5). El reporte termina en el último proyecto.

Si una decisión de diseño no aparece en el modelo, consulta a Carlos antes de inventarla. Cualquier desviación del modelo debe estar explícitamente justificada por una indicación del usuario en la sesión actual.

---

## Skills complementarios (cargar SIEMPRE antes de construir el PPT)

Este skill **NO replica** la guía de marca ni el template institucional. Antes de empezar a generar slides, Claude DEBE invocar (`Skill`) los dos siguientes skills y aplicar lo que indiquen:

1. **`banreservas-marca`** — guía completa de la marca Banreservas: paleta institucional, tipografía oficial, tono de voz, lineamientos de comunicación interna/externa, eslogan y reglas del logo. Úsalo para:
   - Validar los colores aplicados (deben coincidir con la paleta oficial documentada por el skill, no con valores improvisados).
   - Redactar el `eyebrow`, subtítulo y cualquier texto institucional en el tono Banreservas (cercano, profesional, claro).
   - Verificar que el uso del nombre "Banreservas" y la mención a la Vicepresidencia Ejecutiva de Tecnología y Operaciones sea consistente.

2. **`banreservas-pptx`** — skill de presentaciones PowerPoint con identidad Banreservas y acceso al template oficial. Úsalo para:
   - **Punto de partida**: clonar/abrir el template oficial del banco como base del `.pptx` mensual (slide masters institucionales, theme de colores, layouts de portada/contenido/cierre). NUNCA generar el deck desde un `Presentation()` vacío si el template está disponible.
   - Obtener el logo Banreservas como imagen (ubicado en los assets del skill) en vez de simular el logo con texto.
   - Heredar los slide masters de portada y cierre institucional.
   - Reutilizar layouts ya formateados con la marca cuando aplique.

**Orden recomendado al iniciar:**

```text
1. Skill: banreservas-marca       (cargar guías de marca y tono)
2. Skill: banreservas-pptx        (clonar template oficial)
3. Skill: pmo-reporte-pptx        (este skill — lógica del reporte)
```

Si por alguna razón uno de los dos skills no está disponible, Claude debe avisarle al usuario antes de proseguir con valores fallback (paleta documentada en `references/branding.md`, logo simulado con texto). El output debería SIEMPRE tender al template oficial; las constantes locales son solo respaldo.

---

## Cuándo activar este skill

Activa este skill cuando Carlos (u otro PM del PMO TI) pida regenerar, actualizar o ajustar el reporte mensual. Señales típicas:

- "Genera el PPT de mayo a partir del Excel."
- "Actualiza el Gantt consolidado con la fecha de corte de hoy."
- "Regenera el slide de SAS H1 con los nuevos hitos."
- "Aplica los resúmenes coherentes (sin elipsis) a las notas y riesgos."
- "Convierte el workbook del mes en presentación."

El workbook fuente vive en `…/Gestión Proyectos Ti - Informe Mensual/{Mes} {Año}/Reporte de Avance de Proyectos PMO TI Banreservas - {Mes} {Año}.xlsx`. La salida va a la misma carpeta como `Reporte de Avance de Proyectos Banreservas {Mes} {Año}.pptx`.

---

## Filosofía y reglas duras

**El Excel es la única fuente de verdad.** Nada de nombres hardcodeados, nada que dependa del nombre de la pestaña, nada inventado. Si `B3` está vacío, marcar `[SIN TÍTULO — REVISAR B3]` y loggear; no inventar el nombre.

**Discrepancia pestaña vs B3 → prevalece B3.** El nombre de la pestaña puede estar truncado a 31 caracteres (límite Excel), pero `B3` tiene el nombre completo.

**Títulos de secciones internas se parsean del Excel.** Las celdas con prefijo `2. `, `3. `, `4. ` en columna B alimentan `HtTitle`, `LgTitle`, `RkTitle`. Algoritmo:

```python
import re
titulo_limpio = re.sub(r'^\d+\.\s*', '', valor_celda)
```

**Prohibido el truncamiento.** Nunca uses `...`, `…`, ni elipsis. Si un texto no cabe, reescríbelo como **frase completa coherente** (ver §"Resúmenes coherentes"). El usuario ha sido enfático: jamás aparece `...` en el output.

**Filas vacías se eliminan del XML, no se renderizan como espacios.** Aplica a hitos sin entregable, riesgos sin descripción, tarjetas de notas vacías. Las cabeceras de tabla siempre se preservan, incluso con una sola fila de datos.

**Preserva cambios manuales del usuario.** Carlos a veces edita celdas entre sesiones (ver `CLAUDE.md` del proyecto). Si detectas contenido manual cerca del área que vas a modificar y la corrección pedida es para OTRA cosa: NO lo sobrescribas. Si hay duda, pregunta.

---

## Workflow de alto nivel

0. **Cargar skills de marca y abrir el modelo.** Antes de tocar el Excel:
   - Invoca `banreservas-marca` (paleta, tono) y `banreservas-pptx` (template oficial).
   - Lee `references/modelo-referencia.md` para tener presentes todas las decisiones de diseño vinculantes.
   - Si vas a comparar visualmente, abre `assets/modelo-referencia.pptx` (es el Reporte Abril 2026 actualizado).
1. **Localizar el workbook del mes.** Carpeta `…/Gestión Proyectos Ti - Informe Mensual/{Mes} {Año}/`. El nombre del libro tiene el mes y año.
2. **Leer la lista canónica de proyectos** desde la pestaña `Índice` (no del listado de pestañas). Extraer el nombre de la hoja real parseando la fórmula `='SheetName'!B3` en la columna correspondiente.
3. **Por cada proyecto, extraer las celdas relevantes** (ver `references/excel-mapping.md` para el mapeo exacto).
4. **Validar integridad** (B3 no vacío, fechas coherentes, % en [0,100]).
5. **Aplicar resúmenes coherentes** a textos que excedan el límite de su columna (ver §"Resúmenes coherentes"). El tono de los resúmenes debe respetar la guía de `banreservas-marca` (cercano, profesional, claro; sin tecnicismos innecesarios).
6. **Construir el PPT replicando el modelo.** Opción preferida: clonar `assets/modelo-referencia.pptx` y reemplazar contenidos slide por slide. Opción alternativa: partir del template oficial de `banreservas-pptx` y construir los slides siguiendo exactamente las especificaciones de `references/modelo-referencia.md`. **NUNCA** partir de `Presentation()` vacío. **NUNCA** introducir variaciones de layout no documentadas en el modelo.
7. **Eliminar filas/tarjetas vacías** del XML (usar `scripts/helpers.py`).
8. **Reposicionar el bloque de Riesgos** dinámicamente con el algoritmo de `references/posicionamiento.md`.
9. **Aplicar compresión** si el contenido no entra (font 7.5 pt, alturas reducidas).
10. **Guardar** como `Reporte de Avance de Proyectos Banreservas {Mes} {Año}.pptx`.
11. **QA visual:** convertir a PDF con LibreOffice, exportar a JPG con `pdftoppm`, inspeccionar.
12. **Reportar al usuario** los proyectos con resúmenes aplicados, los con compresión densa, y cualquier validación fallida.

---

## Stack técnico

- Python 3 + `python-pptx` + `openpyxl`.
- LibreOffice (`soffice --headless --convert-to pdf`) para QA.
- Poppler (`pdftoppm -jpeg -r 80`) para slides → JPG.

Los helpers Python ya están en `scripts/helpers.py`. Úsalos en lugar de reinventarlos: `duplicate_slide`, `replace_run`, `remove_extra_table_rows`, `remove_shape`, `resumir_coherente_check`, `posicionar_riesgos`.

---

## Resúmenes coherentes (anti-desbordamiento)

Algunos campos vienen con párrafos largos. Reescribirlos **solo si no caben** como frase completa de máximo 2 líneas, preservando datos clave (fechas, %, sistemas, contrapartes).

| Campo | Límite chars | Estilo |
|---|---|---|
| Nota relevante (`LgTxt`) | ~140 | Preserva datos clave |
| Descripción Riesgo | ~109 | Sujeto + verbo + objeto |
| Plan de Mitigación | ~109 | Acción + contrapartes |

**Reglas duras del resumen:**

- **Prohibido `...`, `…`, elipsis.** Siempre frase gramatical completa.
- No resumir si el texto cabe sin tocar.
- Preserva datos cuantitativos: fechas, %, sistemas, releases, contrapartes.
- Elimina intros redundantes (`Existe el riesgo de que…` → arranca con el sujeto real).
- Si tras resumir aún no cabe: reduce fuente 1 pt (mínimo 9 pt) y reintenta.
- Loggea: `[Proyecto X] R-02 desc. reescrita de N→M chars`.

**Por qué importa:** el usuario revisa cada mes los slides uno a uno. La presencia de `...` se interpreta como "Claude se rindió". Una frase reescrita coherente, en cambio, comunica que el resumen es una decisión editorial deliberada.

**Ejemplos validados (Mayo 2026):**

| Original (chars) | Resumen (chars) |
|---|---|
| "Existe el riesgo de retrasos y desviaciones en el cronograma del proyecto debido a la dependencia de desarrollos asignados a proveedores externos…" (276) | "Retrasos en cronograma por dependencia de proveedores externos; pueden desfasar entregables para QA." (100) |
| "Validar con DBA Oracle la estrategia de migración, tiempos estimados, rendimiento, volumen de datos, pruebas previas y plan de contingencia." (140) | "Validar con DBA Oracle estrategia, tiempos, rendimiento, pruebas previas y plan de contingencia." (96) |

---

## Secciones vacías completas (fallback)

Cuando una sección entera está vacía, sustituir con un mensaje único centrado, gris `#63666A`, itálica:

| Sección | Texto sustituto |
|---|---|
| Riesgos | `No hay riesgo para este mes` |
| Notas Relevantes | `No hay notas en este mes` |
| Entregables | `Sin hitos cargados para este corte` |

---

## Referencias detalladas

Cuando necesites el detalle de una parte específica, lee la referencia correspondiente. **No leas todas al inicio**: lee solo la que necesitas para la subtarea actual.

- **`references/modelo-referencia.md`** — **(VINCULANTE)** Especificación exacta del PPT modelo (Abril 2026 actualizado). Lee esto **PRIMERO**, antes que cualquier otro reference. Toda decisión de diseño debe alinearse a este documento.
- **`assets/modelo-referencia.pptx`** — Archivo binario del modelo. Úsalo como base para clonar y reemplazar, o ábrelo en LibreOffice para verificar coordenadas exactas.
- **`references/excel-mapping.md`** — Estructura del workbook, mapeo celda→campo PPT, convenciones de hojas y del Índice. Léelo apenas necesites tocar el Excel.
- **`references/layout.md`** — Layout slide-a-slide histórico. **Donde haya conflicto con `modelo-referencia.md`, gana el modelo.**
- **`references/posicionamiento.md`** — Algoritmo de posicionamiento dinámico del bloque Riesgos, alturas estimadas de fila, criterios de compresión. Léelo cuando ajustes layout denso.
- **`references/branding.md`** — Paleta de colores oficial Banreservas, tipografía, constantes geométricas (EMU). **Fallback local**: el skill `banreservas-marca` es la fuente primaria; usa este archivo solo si el skill de marca no está cargado.
- **`references/qa.md`** — Checklist final, validación programática de bordes, QA visual con subagente. Úsalo antes de entregar.
- **`references/troubleshooting.md`** — Casos especiales: shapes off-screen del template viejo, hyperlinks rotos al renombrar pestañas, auto-fit de tablas, sheet name >31 chars, normalización Unicode NFC/NFD al guardar en OneDrive/macOS. Consulta antes de improvisar.

### Skills externos relacionados (cargarlos como Skill, no como Read)

- **`banreservas-marca`** — guía de marca, paleta, tono de voz. **Cárgalo siempre al inicio.**
- **`banreservas-pptx`** — template `.pptx` institucional con slide masters de portada, contenido y cierre, theme oficial y logo. **Cárgalo siempre al inicio.**
- **`pptx`** (Anthropic, opcional) — utilidades genéricas de manipulación de PowerPoint. Útil para operaciones avanzadas (clonar slides, manipular XML) cuando los helpers locales de `scripts/helpers.py` no alcancen.

---

## Convenciones de entrega

- **Excel entrada:** `Reporte de Avance de Proyectos PMO TI Banreservas - {Mes} {Año}.xlsx`
- **PPT salida:** `Reporte de Avance de Proyectos Banreservas {Mes} {Año}.pptx`
- Ambos en `…/Gestión Proyectos Ti - Informe Mensual/{Mes} {Año}/`

**Cierre del flujo** — al entregar, reporta a Carlos:

1. Total de slides generados (`1 portada + 1 roadmap + N detalle`).
2. Lista de proyectos con resúmenes aplicados (con conteo de chars original→final).
3. Lista de proyectos con compresión densa (font 7.5 pt).
4. Cualquier validación que haya fallado (B3 vacío, fechas inconsistentes, % fuera de rango).
5. Link `computer://…` directo al `.pptx` en la carpeta del mes.

---

*Especificación basada en el template Abril 2026 y refinada con la generación Mayo 2026.*
