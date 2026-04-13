# Codificación de documentos — Plan de Optimización y Rentabilidad 2026

Todos los entregables bajo este plan comparten el vínculo **POR** (*Plan de Optimización y Rentabilidad*). El **año** no va al inicio del código: va **dentro de la fecha** en formato `AAAAMMDD`.

## Formato

```
{TIPO}-POR-{MESA}-{AAAAMMDD}-{NN}
```

| Segmento | Significado | Ejemplos |
|----------|-------------|----------|
| **TIPO** | Tipo de documento | `MIN` minuta · `REQ` requerimiento · `ENT` entrega |
| **POR** | Plan de Optimización y Rentabilidad (fijo) | `POR` |
| **MESA** | Mesa o área | `ACT` actuarial · `LEG` legal · `FIN` financiera · `TEC` tecnología · `MKT` mercadeo · `CUM` cumplimiento |
| **AAAAMMDD** | Fecha ancla (reunión, corte o envío acordado) | `20260414` |
| **NN** | Correlativo del mismo día, mesa y tipo (`01`, `02`…) | `01` |

### Opcional — fecha de entrega comprometida

Si el documento asume un plazo explícito:

```
{TIPO}-POR-{MESA}-{AAAAMMDD}-{NN}-E{AAAAMMDD}
```

Ejemplo: `MIN-POR-ACT-20260414-01-E20260416`

## Ejemplo

`MIN-POR-ACT-20260414-01` → Primera minuta del plan POR, mesa actuarial, fecha 14-04-2026.

## Archivos

Mismo código en nombre de archivo:

`MIN-POR-ACT-20260414-01.md` · `MIN-POR-ACT-20260414-01.pdf`
