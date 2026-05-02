# Log de avance por día — Reportería (mesas de trabajo / POR 2026)

Este archivo resume **qué existe en el proyecto**, **en qué orden se construyó** (según Git) y **cómo mantenerlo**. Sirve para que cualquier asistente o persona nueva lea el estado del trabajo **sin depender solo del historial de chat**.

**Proyecto:** `la-internacional-reporteria`  
**Enfoque:** minutas y trazabilidad de reuniones actuariales bajo el Plan de Optimización y Rentabilidad (POR); vínculo documentado con `../Info`, `../Reportes` y promoción futura a `../La-Internacional-BI`.

---

## Cómo actualizar este log

1. Al cerrar el día (o al hacer un hito), añade una subsección **con la fecha** en formato `AAAA-MM-DD`.
2. Lista **hechos concretos** (archivos nuevos, scripts, decisiones de codificación, cambios en minutas).
3. Si hubo commits, puedes copiar el resumen desde:  
   `git log --reverse --date=short --format="### %ad %h%n%s%n" --stat`
4. Si hay trabajo **solo en disco** (sin `git commit`), indícalo explícitamente para no confundir “estado oficial” con “borrador local”.

---

## 2026-05-02 (workspace padre + reportería — equipo operativo Mercadeo)

- **Nuevo repo hermano** en la raíz del workspace LIS: **`por-equipo-mercadeo-operativo/`** (`git init` + primer commit local): `README.md`, `CONTEXTO-EQUIPO-MERCADEO-POR.md`, `REGISTRO-ACTIVIDADES.md`, `.gitignore`.
- **Minuta mesa Mercadeo:** `reuniones/2026/MIN-POR-MKT-20260430-01.md` (Sesión Mercadeo 2026-04-30, equipo operativo POR).
- **Padre documental:** `DOCUMENTO-MAESTRO-CONTEXTO-PROYECTO-la-internacional.md` (§1, §2, §3, §4.10, §5, §7, historial), `AGENTS.md`, `GUIA-SUBPROYECTOS-LA-INTERNACIONAL.md`, `ESTADO-ULTIMA-ACTIVIDAD.md`; regla Cursor **`../.cursor/rules/por-equipo-mercadeo-operativo-lis.mdc`**.
- **Pendiente:** `git commit` en `la-internacional-reporteria` si el equipo consolida la minuta; crear **remoto GitHub** para `por-equipo-mercadeo-operativo` y empuje según bitácora de sincronía (`instrucciones-cursor/ESTADO-SINCRONIA-ACTUAL.md`).

---

## 2026-04-13 (histórico Git — día de arranque del repo)

Orden cronológico de commits (de más antiguo a más reciente). Todo ocurrió el mismo día calendario.

### Commit `0a6ae15` — Estructura inicial

- Repositorio de **reportería diaria** para reuniones actuariales: notas estructuradas, requerimientos y puente hacia el BI.
- Añadidos: `.gitignore`, `README.md`, `INSUMOS.md` (rutas a `../Info` y `../Reportes`), `SEGUIMIENTO.md` (plantilla de seguimiento breve).
- Plantillas: `plantillas/acta-diaria.md`, `plantillas/requerimiento.md`.
- Carpeta `promocion-bi/README.md` para registrar CSV/datos promovidos a `La-Internacional-BI`.
- Carpetas vacías preparadas: `requerimientos/`, `reuniones/2026/` (`.gitkeep`).

### Commit `96f9872` — Primera minuta en Markdown

- Minuta de presentación gerencia actuarial (Yaritza Oberto), abril 2026.
- Archivo (nombre inicial del flujo): `reuniones/2026/...-2026-04-14-presentacion-gerencia-actuarial.md`.

### Commit `9f2eac1` — Export a PDF con identidad visual corporativa

- Dependencia: `reportlab` en `requirements.txt`.
- Script `scripts/generar_minuta_pdf.py`: PDF con **Arial**, título **#1B3A5C**, cabecera/pie alineados al criterio del documento *Solicitud La Internacional Definitivo v2*.
- PDF generado junto a la minuta.
- `README.md` ampliado con instrucciones de uso.

### Commit `5f0d1c5` — Ajustes de contenido y PDF (mesa actuarial / SUDEASEG)

- Minuta: mención a **SUDEASEG en instalaciones**, **correos sin rutas locales**, ajustes de título.
- Ajustes menores en `generar_minuta_pdf.py` y regeneración del PDF.

### Commit `a4653b4` — Codificación unificada POR + renombre de archivos

- Nuevo `CODIGOS.md`: esquema `{TIPO}-POR-{MESA}-{AAAAMMDD}-{NN}` (ej. `MIN-POR-ACT-20260414-01`); año dentro de la fecha, sin prefijo “LI”.
- Renombrado de la minuta y su PDF a:  `reuniones/2026/MIN-POR-ACT-20260414-01.md` y `.pdf`.
- `generar_minuta_pdf.py`: lectura de la **referencia** desde la tabla del markdown para la cabecera del PDF.
- `README.md` enlaza a `CODIGOS.md`.

### Commit `1da1905` — Minuta: correspondencia y próxima reunión

- Actualización de la minuta con **correos del 13-04-2026** (según PDFs en `Reportes`), **próxima reunión 14-04-2026 14:00**, nota sobre denominación.
- PDF regenerado.

### Commit `1fa72d6` — Minuta y CODIGOS: redacción y alineación

- Minuta: **correos sin nombres de archivo** en tabla; notas POR; ajuste de redacción (sin sufijo “plazo” donde aplicaba).
- `CODIGOS.md` acotado/alineado con el uso real.
- PDF regenerado.

### Commit `d9cdc06` — Participantes, notas POR, export DOCX

- Minuta: sección **1. Participantes**; nota POR sin frase redundante de empresa.
- Dependencia: `python-docx` en `requirements.txt`.
- Nuevo `scripts/generar_minuta_docx.py` (Arial, tablas y **negritas** básicas desde el MD).
- Generado `reuniones/2026/MIN-POR-ACT-20260414-01.docx`; PDF actualizado.
- `README.md` documenta ambos comandos (`generar_minuta_pdf.py` y `generar_minuta_docx.py`) y rutas opcionales por CLI.

**Resumen del día:** quedó definido el **marco del repo**, **plantillas**, **codificación POR**, **primera minuta mesa actuarial** contenido sustantivo, y **pipeline MD → PDF y MD → DOCX** con criterios visuales corporativos.

---

## 2026-04-14

**En Git:** no hay commits con fecha 2026-04-14 en el historial consultado. La **fecha en el código** `MIN-POR-ACT-20260414-01` ancla la minuta a la **reunión de seguimiento** acordada para el **14-04-2026 a las 14:00** (contenido de la minuta elaborado en el flujo del 13-04).

---

## 2026-04-15 (estado del árbol de trabajo — revisar antes de considerar “oficial”)

Al momento de generar este log, `git status` mostró **cambios locales sin commit** y un archivo **nuevo sin seguimiento**:

- Modificados (entre otros): `.gitignore`, `CODIGOS.md`, `INSUMOS.md`, `README.md`, `SEGUIMIENTO.md`, `plantillas/acta-diaria.md`, `plantillas/requerimiento.md`, `promocion-bi/README.md`, `requirements.txt`, `reuniones/2026/MIN-POR-ACT-20260414-01.{md,pdf,docx}`, `scripts/generar_minuta_docx.py`, `scripts/generar_minuta_pdf.py`.
- Sin seguimiento: `plantillas/minuta-mesa-plantilla.md` (plantilla dedicada de minuta por mesa).

**Acción recomendada:** revisar `git diff`, decidir qué incluir en un commit y **tras el commit** volver a copiar aquí un resumen en una entrada “2026-04-15” para que el log refleje el trabajo aceptado en la rama principal.

---

## 2026-04-23 (solo en disco — promoción Mesa Actuarial)

- **Promoción:** copia recursiva completa desde `../Reportes/Carpeta compartida Actuarial` hacia `../Reportes/Mesa Actuarial/Materiales/Promocion-2026-04-23-Carpeta-compartida-Actuarial` (**176 archivos**); el origen en la bandeja de gerencia actuarial **no** se eliminó.
- **Traza:** `../Reportes/Mesa Actuarial/INDICE-PROMOCION-CARPETA-ACTUARIAL-A-MESA.md`; actualización del **documento maestro** del workspace (historial); regeneración de `INDICE-Y-AVANCE-Carpeta-Mesa-Actuarial.docx` (nueva subsección 3.7 en el script `scripts/generar_indice_carpeta_mesa_actuarial_docx.py`).
- **Minuta:** `../Reportes/Mesa Actuarial/Minutas Reuniones/MIN-POR-ACT-20260423-Mesa-Actuarial-vs-Solicitud-v3.docx` generada con el mismo formato que la v2 (`generar_minuta_relevamiento_actuarial_vs_solicitud_docx.py`).
- **Minuta (nombre con fecha de actualización):** `../Reportes/Mesa Actuarial/Minutas Reuniones/MIN-POR-ACT-20260423-Mesa-Actuarial-vs-Solicitud_actualizado-2026-04-23.docx` (salida por defecto del script; sustituye a v3 como referencia operativa).
- **Semana 2 — control:** `../Reportes/Mesa Actuarial/Semanal/La_Internacional_Semana2_2026.pptx`, `REP-POR-ACT-Semana02-2026-BORRADOR.pptx` y `REP-POR-ACT-Semana02-2026-BORRADOR.md` (`scripts/crear_presentacion_semana2_mesa_actuarial.py`).
- **Índice ejecutivo:** si `INDICE-Y-AVANCE-Carpeta-Mesa-Actuarial.docx` está abierto en Word, el script guarda copia `INDICE-Y-AVANCE-Carpeta-Mesa-Actuarial-REGENERADO-2026-04-23.docx` hasta poder sobrescribir el principal.

---

## 2026-04-23 (solo en disco — vaciado Gerencia Legal a Mesa Actuarial)

- **Espejo completo:** `../Reportes/Mesa Actuarial/Gerencia-Legal/vaciado-completo-2026-04-23/` (copia recursiva de `../Reportes/GERENCIA LEGAL`, 176 archivos, sin `Thumbs.db`).
- **Clasificación por tema (mesa):** `../Reportes/Mesa Actuarial/Materiales/De-Gerencia-Legal/` (`01-Reaseguro-Contratos` … `13-Propiedad-intelectual`). Índice: `../Reportes/Mesa Actuarial/Gerencia-Legal/INDICE-VACIADO-Y-MAPA-A-MESA-ACTUARIAL.md`.
- **Presentación Semana 2:** regenerados `../Reportes/Mesa Actuarial/Semanal/La_Internacional_Semana2_2026.pptx` y `REP-POR-ACT-Semana02-2026-BORRADOR.pptx` — **7** diapositivas; síntesis visible en cuerpo de la **diapositiva 5**; **diapositivas 6–7** con detalle Legal y Actuarial (corrección: el textbox superpuesto no se veía en PowerPoint). `REP-POR-ACT-Semana02-2026-BORRADOR.md` alineado.

---

## 2026-04-24 (solo en disco — hallazgos minuta, Legal y matriz reaseguro)

- **Minuta:** `../Reportes/Mesa Actuarial/Minutas Reuniones/MIN-POR-ACT-20260424-Mesa-Actuarial-vs-Solicitud_actualizado-2026-04-24.docx` (matriz apartado 4 revisada: reportería manual **Bajo**, automatización BI **Alto** sin desarrollo, oferta en espera SUDEASEG, certificación reservas **cubierta** en carpeta base, carga analítica distinta de recibos, reaseguro con cruce Legal–SUDEASEG sin nota explícita de accesos).
- **Matriz:** `../Reportes/Mesa Actuarial/Materiales/MATRIZ-REASEGURO-Legal-vs-SUDEASEG-2026-04-24.md` (contratos en `CONTRATOS DE REASEGUROS-…` / `GERENCIA LEGAL` vs `f- Relación de Contratos de Reaseguro Año 2025 SUDEASEG.xlsx`).
- **Carpetas `Reportes`:** `CONDICIONADOS DE POLIZAS - PRODUCTOS`, `CONTRATOS DE REASEGUROS-20260423T200206Z-3-001`, `GERENCIA LEGAL` referenciadas en índice ejecutivo (script `generar_indice_carpeta_mesa_actuarial_docx.py`, revisión 2026-04-24). Si Word bloquea el DOCX principal, revisar `INDICE-Y-AVANCE-Carpeta-Mesa-Actuarial-REGENERADO-2026-04-24.docx`.
- **Compilación multi-mesa (Word):** `../Reportes/Mesa Actuarial/Minutas Reuniones/COMPILACION-Mesas-Diagnostico-POR-Sem1-2-2026-04-24.docx` — síntesis de las **siete mesas** del diagnóstico (Semanas 1–2); script `../.cursor/scripts/generar_compilacion_mesas_docx.py` desde plantilla `CONF-Accionista-Reaseguro-TPA-Salud-2026-04-24.docx`.
- **Documento maestro (raíz workspace):** actualizado `../DOCUMENTO-MAESTRO-CONTEXTO-PROYECTO-la-internacional.md` con **§6.1**, lista de mesas en **§6**, guía del agente **§7 ítem 12** y **Historial**; regla Cursor `../.cursor/rules/documento-maestro-la-internacional.mdc` (comportamiento **§6** multi-mesa).
- **Demo reportería interna POR:** carpeta `../reporteria-interna-por-demo/` (luego renombrada a **`../La-Internacional-BI-Interno/`**); `README.md`, `docs/ARQUEO-REPORTERIA-INTERNA.md`, `app.py` Streamlit leyendo CSV de `Reportes/Mesa Actuarial/BI financiero/` y agregados `bi-la-internacional`; maestro **§4.7** y tabla Git **§5**.
- **Centralización insumos BI interno:** carpeta `../Data Raw/` como **canónica** para Excel de cierres y dominios afines (inventario en `../Data Raw/README.md`); `../bi-la-internacional/lib/report_paths.py` (`DATA_RAW`, `GESTION_GENERAL`, `PAQUETES_CIERRE` → `data_raw_central`); `../DOCUMENTO-MAESTRO-CONTEXTO-PROYECTO-la-internacional.md` **§1.4** y diagrama **§3**; `workspace_paths.py` del BI interno expone `DATA_RAW`; arqueo en `docs/ARQUEO-REPORTERIA-INTERNA.md` actualizado.
- **Rename BI interno (portal):** `../reporteria-interna-por-demo/` → **`../La-Internacional-BI-Interno/`** (visión multi-área, mapa analítico desde `Data Raw`, pestañas en `app.py`); documentación maestro **§4.7**, tabla **§1**, Git **§5**; regla Cursor `documento-maestro-la-internacional.mdc`; `../Data Raw/README.md` alineado al nuevo nombre de carpeta.

---

## Inventario rápido de piezas actuales (referencia)

| Área | Contenido |
|------|-----------|
| Documentación raíz | `README.md`, `INSUMOS.md`, `SEGUIMIENTO.md`, `CODIGOS.md`, este `LOG-AVANCE-POR-DIA.md` |
| Reuniones | `reuniones/2026/MIN-POR-ACT-20260414-01.md` (+ `.pdf` / `.docx` generados) |
| Plantillas | `plantillas/acta-diaria.md`, `plantillas/requerimiento.md`, `plantillas/minuta-mesa-plantilla.md` (última: ver estado Git) |
| Scripts | `scripts/generar_minuta_pdf.py`, `scripts/generar_minuta_docx.py` |
| BI | `promocion-bi/README.md` (registro de promociones a `data/public/` del proyecto BI) |
| Requerimientos | `requerimientos/` (preparado; sin REQ numerados aún en el listado de archivos revisado) |

---

*Generado para lectura por asistentes y equipo; mantener actualizado con el mismo criterio que las minutas: una línea de tiempo clara y verificable.*
