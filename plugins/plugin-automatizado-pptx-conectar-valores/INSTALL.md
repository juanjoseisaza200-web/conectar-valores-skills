# Instalación — Plugin Automatizado PPTX Conectar Valores

Este plugin contiene skills de Claude Code para generar (y eventualmente
auditar) presentaciones PowerPoint con el grid Müller-Brockmann y la
identidad de marca de Conectar Valores S.A.S.

## Requisitos previos

- Node.js instalado (cualquier versión reciente — se probó con la
  disponible en un entorno Linux estándar)
- Claude Code instalado y funcionando
- Una API key gratuita de Pexels (ver paso 3)
- Python 3 instalado y accesible como `python3` en tu PATH (requerido
  por mb-format para leer archivos .pptx existentes)
- python-pptx instalado: `pip install python-pptx` (mismo requisito que el punto anterior)

## Paso 1 — Copiar la carpeta del plugin

Copia toda la carpeta `plugin-automatizado-pptx-conectar-valores/` a la
ubicación de skills de
tu proyecto o de tu usuario en Claude Code. Típicamente:

```
~/.claude/skills/plugin-automatizado-pptx-conectar-valores/
```

o, si es un skill específico de un proyecto:

```
<tu-proyecto>/.claude/skills/plugin-automatizado-pptx-conectar-valores/
```

Verifica que la estructura se vea así después de copiar:

```
plugin-automatizado-pptx-conectar-valores/
├── INSTALL.md          (este archivo)
├── CLAUDE.md            (contexto del proyecto para Claude Code)
└── skills/
    ├── mb-generate/
    │   ├── SKILL.md
    │   ├── engine/      (el motor — no editar a mano salvo que sepas qué haces)
    │   ├── assets/       (logo_cv.png)
    │   ├── config/       (api_keys.json — AQUÍ va tu key, ver paso 3)
    │   ├── reference/    (documentación de la API del motor)
    │   └── tests/        (batería de stress test)
    └── mb-format/
        └── SKILL.md      (placeholder — todavía no implementado)
```

## Paso 2 — Instalar dependencias de Node

Desde la carpeta `skills/mb-generate/engine/`, instala las dependencias
necesarias:

```bash
cd plugin-automatizado-pptx-conectar-valores/skills/mb-generate/engine
npm install pptxgenjs sharp adm-zip @xmldom/xmldom
```

Esto instala: `pptxgenjs` (generación de PPTX), `sharp` (procesamiento
de imágenes), `adm-zip` + `@xmldom/xmldom` (lectura del XML interno del
pptx para el validador).

## Paso 3 — Configurar tu API key de Pexels

El motor necesita buscar imágenes reales para las portadas. Esto usa la
API de Pexels, que es gratuita.

1. Ve a https://www.pexels.com/api/ y crea una cuenta (sin tarjeta de
   crédito, toma 2 minutos).
2. Copia tu API key.
3. Abre `skills/mb-generate/config/api_keys.json` y pégala:

```json
{
  "pexels": "TU_API_KEY_AQUÍ",
  "unsplash": null
}
```

**Nota de seguridad:** esta key vive en un archivo de configuración
separado del código del motor, precisamente para que sea fácil
encontrarla, rotarla, o reemplazarla sin tocar la lógica. Si en algún
momento este plugin se sube a un repositorio compartido (incluso
privado) o se distribuye más allá del equipo original, considera mover
este archivo fuera del control de versiones (`.gitignore`) y
distribuirlo por separado — el plan original al construir esto fue
compartirlo solo de máquina a máquina directamente, sin pasar por
GitHub, así que esta keys vive en texto plano dentro del plugin.

Límite del plan gratuito de Pexels: 200 solicitudes/hora, 20,000/mes.
Si tu equipo genera muchas presentaciones simultáneamente, podría
agotarse — en ese caso Pexels devuelve un error 429 y `fetchHeroImage()`
lo propaga; pídele a Claude que use un placeholder o que esperes al
siguiente reseteo de cuota.

## Paso 4 — Probar que funciona

Desde la carpeta `skills/mb-generate/`, corre la batería de tests:

```bash
node tests/stress_test_full.js
```

Deberías ver 15/15 casos con "✓ PASS". Si alguno falla, no continúes —
algo se rompió en la copia o en las dependencias instaladas.

## Paso 5 — Usar el skill en una conversación

Una vez instalado, simplemente pídele a Claude Code algo como:

> "Genérame una presentación de 4 slides sobre [tema] con el formato de
> Conectar Valores"

Claude debería reconocer automáticamente que aplica `/mb-generate` según
la descripción del `SKILL.md`. Si no lo hace, puedes invocarlo
explícitamente mencionando "usa el skill mb-generate" o "con el grid de
Conectar Valores".

## Actualizar el plugin más adelante

Si en el futuro corriges un bug en el motor o agregas un layout nuevo,
los archivos a tocar son los de `skills/mb-generate/engine/`. Después de
cualquier cambio, vuelve a correr `node tests/stress_test_full.js` antes
de redistribuir la carpeta actualizada al resto del equipo — eso evita
que un cambio rompa algo que antes funcionaba sin que nadie lo note
hasta producción.
