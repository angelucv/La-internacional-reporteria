# Codificación de documentos — Plan de Optimización y Rentabilidad 2026

Todos los entregables bajo este plan comparten el vínculo **POR** (*Plan de Optimización y Rentabilidad*). El **año** no va al inicio del código: va **dentro de la fecha** en formato `AAAAMMDD`.

## Formato

```
{TIPO}-POR-{MESA}-{AAAAMMDD}-{NN}
```

| Segmento | Significado | Ejemplos |
|----------|-------------|----------|
| **TIPO** | Tipo de documento | `MIN` minuta · `REQ` requerimiento · `ENT` entrega |
| **POR** | Siglas de **Plan de Optimización y Rentabilidad** (segmento fijo) | `POR` |
| **MESA** | Mesa o área | `ACT` actuarial · `LEG` legal · `FIN` financiera · `TEC` tecnología · `MKT` mercadeo · `CUM` cumplimiento |
| **AAAAMMDD** | Fecha ancla (reunión, corte o envío acordado) | `20260414` |
| **NN** | Correlativo del mismo día, mesa y tipo (`01`, `02`…) | `01` |

## Ejemplos (mesa actuarial)

`MIN-POR-ACT-20260413-01` → Minuta del **13-04-2026** (presentación inicial).

`MIN-POR-ACT-20260414-01` → Minuta del **14-04-2026** (seguimiento: revisión de insumos, VPN y carpetas compartidas).

## Archivos

Mismo código en nombre de archivo (extensión según formato):

`MIN-POR-ACT-20260413-01.md` · `MIN-POR-ACT-20260414-01.md` (y `.pdf` / `.docx` si se exportan).
