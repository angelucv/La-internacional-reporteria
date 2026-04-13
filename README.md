# Reportería diaria — reuniones actuariales (La Internacional)

Seguimiento de reuniones con el equipo actuarial: minutas, requerimientos y trazabilidad hacia entregables. Responsable del seguimiento: **Angel Colmenares**.

## Insumos fuera de este repo (mismas carpeta padre)

| Ubicación | Contenido |
|-----------|-----------|
| [`../Info`](../Info) | Documentación interna de referencia |
| [`../Reportes`](../Reportes) | Correos y datos adjuntos de reportería |

Este repositorio concentra **notas estructuradas** y **decisiones**; los archivos binarios o correos pueden permanecer en esas carpetas y referenciarse desde aquí por ruta y nombre.

## Estructura

```
reuniones/YYYY/           # Una nota por día (o sesión)
requerimientos/           # REQ-### con estado e historial
plantillas/               # Copiar para nuevas actas o requerimientos
promocion-bi/             # CSV o notas sobre qué subir al BI (La-Internacional-BI)
```

## Flujo sugerido

1. **Diario:** crear o actualizar `reuniones/AAAA/MM-DD-actuarial.md` (o `YYYY-MM-DD-...`).
2. **Nuevo pedido:** archivo en `requerimientos/REQ-###-titulo-corto.md`.
3. **Promoción al BI:** cuando un dato esté listo para el tablero, copiar el CSV a `La-Internacional-BI/data/public/` y registrar en `promocion-bi/README.md` la fecha y el origen.

## Relación con el BI

Proyecto web: `../La-Internacional-BI`. Los datos que consuma la app viven en ese repo (`data/public/`); aquí solo se **documenta** qué se promovió y cuándo.

## Exportar minuta a PDF (arte alineado al documento corporativo)

El script usa **Arial**, color de título **#1B3A5C** y cabecera/pie tipo *Solicitud La Internacional Definitivo v2* (mismo criterio visual que el PDF en `../Info`).

Requisitos: **Python 3** + fuentes Arial en `C:\Windows\Fonts` (Windows).

```bash
cd la-internacional-reporteria
pip install -r requirements.txt
python scripts/generar_minuta_pdf.py
python scripts/generar_minuta_docx.py
```

Codificación de referencias (`MIN-POR-ACT-20260414-01`, etc.): ver **`CODIGOS.md`**.

Por defecto lee `reuniones/2026/MIN-POR-ACT-20260414-01.md` y escribe el **`.pdf`** y el **`.docx`** junto al markdown. Rutas opcionales:

`python scripts/generar_minuta_pdf.py ruta\entrada.md ruta\salida.pdf`  
`python scripts/generar_minuta_docx.py ruta\entrada.md ruta\salida.docx`
