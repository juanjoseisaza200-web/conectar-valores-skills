# Formato visual del modelo — estándar Conectar

## Colores de fuente (la convención del modelo — leyenda en su Portada)
| Color | Significado | Cuándo |
|---|---|---|
| **Azul** RGB(0,0,255) | Input, o celda IMPORTADA de otra hoja | F-imports `=Inputs_C!$F$x`; filas `=OtraHoja!Jx`; valores de input en Inputs_C/Inputs_Years |
| **Negro** | Cálculo local (misma hoja) | SUMs, rolls, cadenas que solo usan la propia hoja |
| **Rojo** | Celda EXPORTADA a otra hoja | Solo si otra hoja la referencia; nunca en etiquetas/unidades |
| Pestaña fuente blanca | Hoja de inputs | Inputs_C, Inputs_Years |
| Pestaña fuente negra | Hoja de cálculo | resto |

Regla práctica al crear filas: si la fórmula contiene `!` (cruza hoja) → azul; si no → negro. Etiquetas y unidades SIEMPRE negras.

## Layout de fila
| Col | Contenido |
|---|---|
| B–D | Encabezados de sección / sub-sección |
| E | Etiqueta de la fila (nombre EXACTO de la cuenta del cliente) |
| F | Constante de fila — SOLO import `=Inputs_C!$F$x` (azul) |
| G | Unidad: `COP 000` · `%` · `días` · `#años` · `binario` · `x` · `ton` |
| J:AD | Datos 2019–2039, fórmula única (FAST) |

- Una fila en blanco entre sub-bloques; encabezado con estilo de sección (copiar formato de un encabezado existente).
- Bloques nuevos con sufijo de trazabilidad: `— ajuste <tema> <mmm-aa>`.
- NumberFormat: copiar la fila plantilla del mismo ROL; porcentajes `0.0%`; no dejar unidades/márgenes duplicados heredados del copy de formato.

## Crear una fila nueva (procedimiento)
1. Elegir fila plantilla del MISMO ROL (import / cálculo local / input / total / header).
2. `cv.copy_row_format(sh, plantilla, destino)` y limpiar valores arrastrados (cols 2–8).
3. Escribir etiqueta (E), unidad (G), import F si aplica.
4. Escribir fórmula y propagar con `cv.fast_fill`.
5. Ajustar color de fuente según la regla práctica.
6. Al terminar el bloque: `cv.scan_rojos(sh, r1, r2)` debe devolver [].

## Documentos Word para el cliente
Usar la skill `conectar-docx-creator` (formato VF3: Arial, Navy #17375E, Dorado #C9A449, Mint #E7F0EF, Letter 1.2"/0.79") y auditar 10/10 con su `critic/audit_against_vf3.py` antes de entregar.
