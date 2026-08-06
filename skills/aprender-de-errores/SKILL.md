---
name: aprender-de-errores
description: Skill global de mejora continua. Usar SIEMPRE que (a) el usuario corrija un error mío, (b) una verificación/auditoría falle, (c) descubra un comportamiento no obvio de una herramienta o convención, o (d) al cerrar una fase de trabajo significativa. También consultarla ANTES de empezar tareas de un tipo ya realizado (modelaje Excel, documentos Word, análisis de datos) para no repetir errores registrados.
---

# Aprender de errores — protocolo de mejora continua

Esta skill mantiene un registro vivo de lecciones (`LECCIONES.md`, en esta misma carpeta) y obliga a consultarlo y alimentarlo. El objetivo: **ningún error se comete dos veces**.

## Cuándo y cómo registrar (obligatorio, inmediato)

Registrar una lección apenas ocurra cualquiera de estos eventos — no al final de la sesión:
1. El usuario corrige algo que hice (formato, arquitectura, criterio, alcance).
2. Una verificación falla (check de balance, auditoría de formato, test, recálculo).
3. Descubro un comportamiento no obvio (API, Excel/COM, convención de un cliente).
4. Tuve que rehacer trabajo por un supuesto equivocado.

Formato de entrada en `LECCIONES.md` (una entrada = 3-5 líneas, concreta y accionable):

```
### [fecha] [área] Título corto
- Qué pasó: ...
- Causa raíz: ...
- REGLA: <instrucción imperativa verificable para la próxima vez>
```

Reglas de calidad: la REGLA debe ser ejecutable sin contexto de la sesión original (nada de "tener más cuidado"); si una lección nueva generaliza otra vieja, fusionarlas; si una regla se gradúa a convención estable de un dominio, moverla a la skill del dominio (p. ej. `modelaje-conectar`) y dejar aquí solo la referencia.

## Cuándo consultar (obligatorio)

- **Antes** de iniciar una tarea de un tipo ya hecho: leer las secciones de LECCIONES.md del área correspondiente (Excel/COM, Word/formatos, datos, proceso).
- **Antes** de entregar: repasar las reglas del área como checklist.
- Si el usuario dice "otra vez lo mismo" o similar: buscar la lección existente, entender por qué no se aplicó, y reforzar la REGLA (hacerla más temprana/automática en el flujo).

## Autoevaluación al cerrar fases

Al terminar una fase significativa, responder por escrito (en la actualización de memoria del proyecto):
1. ¿Qué me corrigió el usuario esta fase? → ¿ya está en LECCIONES.md?
2. ¿Qué verificación fallida me costó más tiempo? → ¿la REGLA la habría evitado?
3. ¿Qué haría distinto si repitiera la fase desde cero?
