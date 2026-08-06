# Plantillas de fórmulas — estándar Conectar (FAST)

Copiar LITERALMENTE estas plantillas (cambiando solo filas/hojas). No inventar variantes.
Convención: J=2019 … P=2025 (observado), Q=2026 … AD=2039 (proyección). Escribir la fórmula en una celda y propagar J:AD con `cv.fast_fill` (PasteSpecial xlPasteFormulas).

## 1. Traer del motor al balance (única fórmula permitida en EEFF)
```
✔ =IF(J$8 =1, WK!J111, 'EEFF Hist.'!J158)
✘ =J$160*Inputs_C!$F$433          ← proyección dentro de EEFF: PROHIBIDO
✘ cualquier constante en columna F del EEFF
```

## 2. Cadena de crecimiento con semilla histórica (en hoja motor)
`flag_primera` = primera proyección (WK: fila 8); `flag_proy` = periodo proyección (WK: fila 7).
```
✔ =IF(J$7=0,0, IF(J$8=1,'EEFF Hist.'!I159,I95)*(1+Inputs_Years!J$20))
✔ =IF(J$7=0,0, IF(J$8=1,'EEFF Hist.'!I162,I98)*EEFF!J326/EEFF!I326)
```
- El wrap exterior `IF(flag=0;0;…)` es OBLIGATORIO si hay división (la columna previa al primer año está vacía → #DIV/0!).
- La semilla usa la columna ANTERIOR del histórico (I dentro de la fórmula escrita en J) porque al propagarse, en Q leerá P (último real). 

## 3. Nivel × driver con base en Inputs_C
```
✔ =$F$104*(EEFF!J317+EEFF!J334)/(EEFF!$P$317+EEFF!$P$334)*J$7      ← base anclada a 2025 ($P$)
✔ =J64*$F$110/365*J$7                                               ← rotación días sobre flujo
```
`$F$x` de la hoja debe ser import: `=Inputs_C!$F$<fila>` (azul).

## 4. Acumulados (cumulative)
```
✔ =SUM($J$418:J418)        ← ancla SIEMPRE en $J$ (los flujos observados son 0)
✘ =SUM($Q$418:Q418)        ← al propagar a J:P el rango se INVIERTE y suma el futuro
```

## 5. Rezagos / vintages (tomar el valor de hace N años)
```
✔ =IF(COLUMN()-COLUMN($Q$368)+1-$F$400<1, 0,
      INDEX($Q$368:$AD$368, COLUMN()-COLUMN($Q$368)+1-$F$400)) * J$7
✘ =IF(COLUMNS($Q$368:J368)>$F$400, INDEX(...), 0)   ← COLUMNS se invierte en J:P y el INDEX
   resuelve celdas FUTURAS → referencia circular (INDEX registra la celda resuelta como precedente)
```

## 6. Roll de saldo (saldo inicial + flujos − salidas)
```
✔ Saldo final:   =IF(cond, SUM(J260:J263), SUM(J261:J263)) - J372      ← appendear castigos/salidas al final
✔ Saldo inicial: =I264                                                  ← prior final
✔ Roll con semilla: =(IF(J$8=1,$F$739,I744)*(1+$F$734)-J737)*J$7
```

## 7. Inputs en Inputs_C
```
Etiqueta col C · unidad col G · valores en K:X (TODAS las columnas; mismo valor si no varía por escenario)
Celda activa:  F = =INDEX(K{r}:X{r},$I$5)
Fórmulas de calibración en K: SIEMPRE columnas absolutas antes de propagar K:X:
✔ ='EEFF Hist.'!$P$160/2          ✘ ='EEFF Hist.'!P160/2  (al llenar K:X se corre a Q160 = año proyectado)
Serie anual → Inputs_Years, valores J:AD; la hoja consume =Inputs_Years!J$<fila>*J$7
```

## 8. MIN/MAX guards
```
✔ Castigo ≤ saldo:  =MIN(programado, MAX(0, J150+J151))*J$7
✔ Run-off a cero:   =MAX(0, IF(J$8=1,'EEFF Hist.'!I160,I96)-$F$96)*J$7
```

## 9. Devolución / release de saldos fiscales (patrón TX)
```
✔ =(-MAX(0,J129+J132)*$F$138 - IF(J$2=$F$121, MAX(0,J129+J132)*(1-$F$138)+J130, 0))*J$8
   (switch en Inputs_C; término del último año conserva el comportamiento de wind-up)
```

## 10. Checks de control en hoja
```
✔ =ROUND(SUM(J426:J431)-J400,0)     ← fila de check de un bloque: debe dar 0 en TODAS las columnas
```
Todo bloque nuevo con redistribuciones DEBE incluir su fila de check.
