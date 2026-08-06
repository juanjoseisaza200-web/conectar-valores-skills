# LECCIONES — registro vivo de errores y reglas

Mantenido por la skill `aprender-de-errores`. Una entrada por lección; reglas imperativas y verificables.

---

## Excel / Modelaje financiero

### [2026-06-11] [arquitectura] Proyecté subcuentas directamente en el balance
- Qué pasó: implementé drivers de CxP, anticipados y run-off como fórmulas dentro de EEFF; el usuario lo rechazó ("se proyecta en WK y se lleva al balance").
- Causa raíz: imité parcialmente el patrón original (que tenía %×ancla en EEFF) en vez de la arquitectura correcta del estándar.
- REGLA: el balance/EEFF SOLO trae valores: `=IF(flag; HojaMotor!fila; Hist!fila)`. Toda proyección vive en la hoja motor (WK/TX/CAPEX/Deuda/Ingresos).

### [2026-06-11] [arquitectura] Dejé inputs sueltos en columnas F de las hojas
- Qué pasó: constantes nuevas (días, bases, switches, tasas) quedaron en F de EEFF/WK/TX/CAPEX/Deuda; el usuario exigió centralizarlas.
- REGLA: escalares en Inputs_C con el patrón del modelo (valor por escenario en K:X, activa F=INDEX(K:X,$I$5)); series anuales en Inputs_Years (J:AD); la hoja consume vía import azul `=Inputs_C!$F$x`. Cero constantes mágicas en hojas de cálculo.

### [2026-06-11] [FAST] Filas con fórmula solo en el rango de proyección (Q:AD)
- Qué pasó: bloques nuevos arrancaban en Q; el usuario exigió el principio FAST (fórmula idéntica desde J hasta el final).
- REGLA: toda fila de proyección lleva UNA fórmula uniforme J:AD. Propagar con PasteSpecial xlPasteFormulas desde una celda correcta y AUDITAR con set(FormulaR1C1) por fila == 1. Los años observados deben dar 0 vía flags, no vía celdas vacías.

### [2026-06-11] [FAST] Rangos que se invierten al extender a la izquierda
- Qué pasó: `SUM($Q$x:Qx)` y `COLUMNS($Q$x:Jx)` se normalizan invertidos en J:P → basura, y un INDEX resolvió celdas FUTURAS creando referencia circular (INDEX registra como precedente la celda resuelta).
- REGLA: prohibido anclar acumulados/ventanas en $Q$. Acumulados: anclar en $J$ (los flujos observados son 0). Rezagos: `n = COLUMN()-COLUMN($J$src)+offset` con `IF(n<1;0;INDEX(rango;n))` para nunca resolver columnas futuras.

### [2026-06-11] [COM] Leí checks con el recálculo asíncrono en curso
- Qué pasó: tras CalculateFullRebuild leí Check_Macros=0 (valor rancio de la copia limpia), el bucle de cierre se saltó y guardé un archivo a medio calcular.
- REGLA: tras todo Calculate/CalculateFullRebuild, esperar `xl.CalculationState == 0` ANTES de leer cualquier celda. Añadir probes de consistencia (celda == f(sus insumos)) antes de guardar.

### [2026-06-11] [COM] Referencias circulares silenciosas
- Qué pasó: con DisplayAlerts=False Excel no avisa de circulares; valores quedaron inconsistentes sin error visible (EEFF mostraba un valor distinto al de su única precedente).
- REGLA: tras cambios de fórmulas, recorrer `ws.CircularReference` en todas las hojas (ignorar los $A$1 con CELL("filename"), son benignos). Valores inconsistentes entre dependientes = sospecha de circular.

### [2026-06-11] [flags] Asumí que el flag de proyección era la misma fila en todas las hojas
- Qué pasó: en WK la fila 7 es "periodo proyección" y la 8 "primera proyección"; en EEFF/TX la 8 es proyección. Una fila de control quedó en 0 por usar el flag equivocado.
- REGLA: verificar la fila de flags POR HOJA leyendo sus etiquetas antes de usar `×Q$7/×Q$8`.

### [2026-06-11] [formato] Heredé fuentes rojas al copiar formatos de fila
- Qué pasó: copy_fmt desde filas "exportadas" (rojas) dejó 66 etiquetas/unidades rojas en bloques nuevos.
- REGLA: tras crear filas con formato copiado, escanear Font.Color==255 en el bloque y normalizar: azul=input/importado, negro=local, rojo SOLO si la celda es exportada a otra hoja.

### [2026-06-11] [supuestos] Asumí composición de datos sin verificar contra la fuente
- Qué pasó: asigné el stock castigado a las 4 intercompañías "porque las cifras calzaban"; el deterioro real por tercero (cta. 1480) tenía otra composición (Soledad 41, municipios ~128, Air-e revertido).
- REGLA: si existe el dato por tercero/detalle en la fuente primaria, verificar SIEMPRE antes de asumir composición. "Las cifras calzan" no es evidencia.

### [2026-06-11] [versiones] Cifras de documentos de revisión pueden ser de versiones viejas
- Qué pasó: las observaciones de Juan citaban tarifa 7% y diferido plano; la versión vigente ya traía 4,5% y DTA móvil.
- REGLA: antes de implementar observaciones de un revisor, verificar CADA cifra citada contra la versión vigente; reportar qué ya estaba resuelto.

### [2026-06-11] [inserts] Inserción de filas y posiciones
- REGLA: insertar de abajo hacia arriba por hoja; las fórmulas existentes se auto-ajustan, pero las fórmulas NUEVAS deben escribirse con posiciones FINALES (mapear offset +n para filas ≥ punto de inserción); asserts de etiquetas en anclas antes y después; los nombres definidos sobreviven.

### [2026-06-11] [escenarios] Parámetros de Inputs_C solo en columna K y fill relativo
- Qué pasó: (1) parámetros con fórmula quedaron solo en K → los escenarios L/M/N leían 0 vía INDEX; (2) al rellenar K:X con paste de fórmulas, las referencias relativas a otras hojas se corrieron de columna (P128→Q128 = año proyectado) creando 44 errores y circularidad.
- REGLA: en Inputs_C, todo parámetro se llena K:X completo; las fórmulas de calibración usan referencias con COLUMNA ABSOLUTA ($P$128) antes de propagar. Probe final: valor idéntico en K, L y N para parámetros que no varían por escenario.

### [2026-06-11] [escenarios] Cada escenario guarda su propio cierre
- REGLA: el selector es Inputs_C!F6 (nombre) → I5=MATCH; la hoja Macros guarda un bloque pegado POR escenario (offsets 2×esc+Variables×(esc−1)). Cambiar de escenario o sus inputs exige re-cerrar ESE escenario; cerrar uno no cierra los demás.

### [2026-06-11] [auditor] "Observados = 0" y "sin rojos" no son criterios universales
- Qué pasó: la primera corrida del auditor automático dio 4 falsos positivos: filas ORIGINALES del modelo legítimamente rojas (exportadas) cayeron en el barrido min-max del rango, y filas ANCLADAS (base fiscal, neteo) llevan valores observados por diseño.
- REGLA: auditar rojos SOLO sobre las filas declaradas (no el span); separar filas de FLUJO (observados deben ser 0 → `--filas-cero`) de filas ANCLADAS (valores históricos intencionales → solo FAST+rojos). Toda herramienta de verificación nueva se calibra contra un caso real conocido-bueno antes de confiar en ella.

## Documentos / Word

### [2026-06-11] [stdout] Wrapper de stdout rompe librerías que imprimen
- Qué pasó: `sys.stdout = io.TextIOWrapper(sys.stdout.buffer)` causó "I/O operation on closed file" dentro de build_from_outline.
- REGLA: en scripts que importan librerías con prints, usar `python -X utf8` sin re-envolver stdout.

## Proceso

### [2026-06-11] [alcance] Preguntar antes de cambiar el modelo
- REGLA: en este proyecto (y por defecto con modelos de clientes), los cambios al modelo se proponen con impacto cuantificado y se ejecutan SOLO tras aprobación explícita; el análisis sí es autónomo.
