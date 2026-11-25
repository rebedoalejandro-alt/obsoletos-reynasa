# Obsoletos Reynasa · seguimiento

Panel del obsoleto de Reynasa: qué se ha vendido y qué queda del perímetro de obsoletos congelado a 31.12.2025 (16.697 referencias), con unidades e importe por categoría, familia, referencia y planta.

Panel: rebedoalejandro-alt.github.io/obsoletos-reynasa

## Cómo se actualiza

Cada noche se sacan de Sage X3 dos peticionarios y se recalcula el corte del día con `actualizar.py`:

- ZCON13 «Stock sin estado Q ni cascos», todas las plantas.
- ZDL25 «Ventas por referencia», del 01.08.2026 al día del corte.

```
python actualizar.py --stock ZCON13.xlsx --ventas ZDL25.xlsx --fecha AAAA-MM-DD --datos datos
```

El script no publica si el ZCON13 sale incompleto o si el obsoleto cambia más de un 15 % frente al corte anterior.

## Datos

- `datos/detalle.json`: una fila por referencia del perímetro (perímetro, cortes en unidades, tramos de venta, stock y planta).
- `datos/serie.json`: totales de cada corte (unidades y euros a PMP de diciembre 2025).
- `datos/seguimiento.json`: peticiones abiertas y su estado.

## Criterio

- Obsoleto en una fecha = MIN( MAX(stock − umbral; 0) ; unidades obsoletas de diciembre ).
- Tramo 1 (dic → 31.07.2026) fijo; tramo 2 (31.07.2026 → corte) recalculado cada día.
- Vendido = MIN( salida de obsoleto del tramo ; ventas netas del tramo ). El resto de la salida son otras salidas (devoluciones a proveedor, bajas, regularizaciones). Las subidas son reentradas.
- Todo a PMP unitario del obsoleto de diciembre 2025, congelado.

Generado automáticamente: los datos se sobrescriben en cada actualización.
