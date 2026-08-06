# Bitácora de observaciones — mb-generate

Este archivo registra observaciones reales del uso del skill en producción
(no del stress test sintético) — tanto problemas sospechosos como casos
límite que funcionaron bien. Cualquier instancia de Claude usando este
skill debe agregar una entrada aquí cuando corresponda (ver criterios
abajo), pero **nunca debe modificar el motor (`engine/*.js`) basándose
solo en esto** — eso requiere revisión humana periódica.

**Nota: este archivo es local a esta copia del plugin.** Si varias
personas del equipo tienen su propia instalación, cada una acumula su
propia bitácora de forma independiente — no hay sincronización
automática entre máquinas. Para consolidar observaciones de varias
personas, hay que compartir manualmente el archivo (por el canal que el
equipo ya use para esto) y juntarlas a mano.

## Cuándo agregar una entrada

**Registrar un PROBLEMA cuando:**
- El validador (`mb_validator.js`) dio ✓ PASS pero la revisión visual
  mostró algo desbordado, apretado, o mal alineado
- Un caso de contenido real (no sintético) produjo un resultado distinto
  al esperado — texto cortado, layout roto, overflow no manejado
- Tuviste que hacer un workaround manual para que algo se viera bien
  (eso es una señal de que el motor debería cubrir ese caso solo)

**Registrar un ÉXITO cuando:**
- Un caso de contenido real e inusual (título muy largo/corto, tabla con
  pocas o muchas filas, cuerpo de tarjeta extenso, etc.) se vio bien sin
  intervención manual — esto es evidencia de que esa parte del motor ya
  está bien probada en uso real, no solo en el stress test sintético

**No registrar:** generaciones rutinarias sin nada inusual en el
contenido (título normal, 3-4 bullets típicos, etc.) — eso no aporta
señal nueva, ya está cubierto por los casos base del stress test.

## Cómo agregar una entrada

Usar el formato exacto de abajo (copiar el bloque y rellenar). Mantener
cada entrada breve — 5-8 líneas máximo. Insertar entradas nuevas al
FINAL del archivo, en orden cronológico, nunca reescribir entradas viejas.

```
### [FECHA ISO] — [PROBLEMA | ÉXITO] — [layout afectado: header/card/flowchart/cover/tabla]

**Contexto:** quién lo generó (si se sabe) y para qué documento/cliente,
en una frase.

**Qué pasó:** descripción concreta — qué contenido se usó (cantidad de
texto, número de items, etc.) y qué resultado se observó.

**Por qué importa:** si fue problema, por qué el validador no lo agarró
(o si sí lo agarró, anotar eso también). Si fue éxito, qué parte del
motor queda más confirmada.

**Acción tomada:** si hubo un workaround manual en el momento, cuál. Si
no se hizo nada (se dejó para revisión posterior), decir eso explícitamente.
```

## Entradas

### 2026-06-17 — ÉXITO — card

**Contexto:** caso de demostración del formato de bitácora, simulando un
deck real de estructura de capital (tranches senior/mezzanine/equity)
para un proyecto de financiamiento estructurado.

**Qué pasó:** tres tarjetas con cuerpos que combinan cifras (USD 45M),
porcentajes (250bps, 11.5%, 22%), abreviaturas financieras (SOFR, DSCR,
PIK) y guiones largos como separador visual ("·"). Todo se renderizó
sin overlaps ni desbordes, con buen espaciado entre tarjetas de distinto
largo de contenido.

**Por qué importa:** confirma que `addCard()` no solo funciona con el
texto sintético en español plano del stress test, sino también con la
densidad típica de contenido financiero real de Conectar Valores
(números, símbolos, abreviaturas mezcladas en una sola línea).

**Acción tomada:** ninguna necesaria — se registra como evidencia
adicional de robustez del layout de tarjetas en uso real.

