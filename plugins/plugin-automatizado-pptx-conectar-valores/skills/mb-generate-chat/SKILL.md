---
name: mb-generate-chat
description: Genera presentaciones PowerPoint (.pptx) con identidad Conectar Valores directamente en claude.ai chat, sin necesidad de Claude Code CLI. Usa este skill cuando el usuario pida crear una presentación y esté en un entorno de chat (no CLI). Produce un HTML artifact con botón de descarga que genera el .pptx en el navegador del usuario con el grid Müller-Brockmann y la paleta CV completa (navy/dorado, Cambria/Arial, footer corporativo, logo). NO usar para reformatear un .pptx existente (eso es /mb-format-chat).
---

# mb-generate-chat

Skill para generar presentaciones .pptx con identidad Conectar Valores
directamente desde un chat de claude.ai, sin instalar nada. El resultado
es un HTML artifact que el usuario abre en el navegador y descarga con
un clic.

---

## 1. Cuándo usar este skill

**Usar cuando:**
- El usuario pide crear o generar una presentación nueva desde un chat
  de claude.ai (no desde la CLI de Claude Code)
- El usuario no tiene Claude Code instalado o prefiere no usarlo
- El usuario quiere un resultado inmediato sin configurar entornos

**NO usar cuando:**
- El usuario está en Claude Code CLI → usar `/mb-generate` (tiene
  imágenes reales de Pexels vía API y validador geométrico)
- El usuario quiere reformatear un .pptx existente → usar `/mb-format-chat`
- El usuario quiere auditar o analizar un deck ajeno → ese es el
  dominio de mb-format, no de mb-generate

**Diferencia clave entre versiones:**

| Aspecto | mb-generate (CLI) | mb-generate-chat (este skill) |
|---|---|---|
| Entorno | Claude Code CLI | claude.ai chat |
| Imagen de portada | Foto real de Pexels | Placeholder navy |
| Validador geométrico | Sí, automático | No |
| Instalación requerida | Sí | No |
| Entregable final | Recomendado | Borrador/prototipos rápidos |

---

## 2. Flujo de trabajo (paso a paso para Claude)

### Paso 1 — Reunir contenido

Antes de generar nada, confirmar con el usuario:

1. **Título del documento** — texto exacto para la portada y el footer
2. **Cliente o destinatario** — aparece en la portada como "CLIENTE"
3. **Mes y lugar** — ej. "JUNIO 2026" y "MEDELLÍN, COLOMBIA"
4. **Estructura del deck** — número de slides y tipo de cada una:
   - `cover` — portada (siempre la primera, solo una por deck)
   - `content` — slide de texto + tabla (layout asym32, sym2 o single)
   - `cards` — tarjetas horizontales (2 o 3 tarjetas de color)
   - `flowchart` — proceso secuencial con cajas conectadas
5. **Contenido de cada slide** — para cada slide confirmar:
   - Título y subtítulo
   - Texto corrido, bullets, o datos según el tipo
   - Encabezados y filas de tabla (si aplica)
   - Textos de tarjetas o pasos del flowchart (si aplica)

No asumir contenido por defecto. Si el usuario da una descripción
general ("una presentación sobre la valoración de X"), preguntar los
datos específicos de cada slide antes de continuar.

### Paso 2 — Leer el template

Leer el archivo `engine/artifact_template.html`, ubicado en:

```
skills/mb-generate-chat/engine/artifact_template.html
```

Este archivo contiene el motor completo (pptxgenjs desde CDN, el grid
MB, la paleta CV, todos los helpers de layout). Solo hay una línea que
Claude debe reemplazar:

```js
const PRESENTATION_DATA = __PLACEHOLDER_DATA__;
```

### Paso 3 — Construir PRESENTATION_DATA

Crear el objeto JavaScript con los datos reales confirmados en el Paso 1.
Schema completo:

```js
{
  projectName: "NOMBRE CORTO PARA FOOTER",   // ej. "GRUPO EJEMPLO"
  footerLeft:   "ESTRICTAMENTE PRIVADO Y CONFIDENCIAL",  // siempre este valor exacto
  footerCenter: "VALORACIÓN GRUPO EJEMPLO",             // omitir si es igual a projectName

  slides: [

    // ── PORTADA (siempre primera, siempre una sola) ──────────────────
    {
      type:     "cover",
      company:  "CONECTAR VALORES S.A.S.",    // supratítulo en dorado — siempre este valor
      title:    "Título principal del deck",
      subtitle: "Subtítulo descriptivo · contexto del encargo",
      client:   "NOMBRE DEL CLIENTE",
      month:    "JUNIO 2026",
      city:     "MEDELLÍN",
      country:  "COLOMBIA",
    },

    // ── SLIDE DE TEXTO + TABLA ────────────────────────────────────────
    // columns: "asym32" → texto izq (3/5) + tabla der (2/5)  [más frecuente]
    //          "sym2"   → dos columnas iguales
    //          "single" → una sola columna de texto o tabla
    {
      type:     "content",
      title:    "Título de la slide",
      subtitle: "Subtítulo opcional",
      columns:  "asym32",
      text:     "Texto de la columna izquierda. Bullets con • al inicio de cada punto.",
      table: {                 // omitir si no hay tabla
        headers: ["Indicador", "2023", "2024", "2025E"],
        rows: [
          ["PIB Energético",   "4.2%", "3.8%", "4.5%"],
          ["Inversión CAPEX",  "USD 2.1B", "USD 2.4B", "USD 3.0B"],
        ]
      },
      notes:       ["Fuente: BID 2025", "Datos preliminares"],  // omitir si no hay
      footerRight: "ENERGÍA · 1/4",  // omitir para numeración automática
    },

    // ── SLIDE DE TARJETAS ─────────────────────────────────────────────
    // Usar 2 o 3 tarjetas. Los colores disponibles son los de la paleta CV.
    {
      type:     "cards",
      title:    "Estructura de Capital",
      subtitle: "Tres tranches principales",
      cards: [
        { title: "TRANCHE SENIOR",    body: "USD 45M · SOFR+250bps · 7 años", color: "1A2744" },
        { title: "TRANCHE MEZZANINE", body: "USD 20M · 11.5% fijo · 5 años",  color: "B8862A" },
        { title: "EQUITY",            body: "USD 15M · IRR 22% target",        color: "8FA3B1" },
      ],
    },

    // ── SLIDE DE FLOWCHART ────────────────────────────────────────────
    // Las cajas primera y última van en navy (isEdge = true automático).
    // Las cajas intermedias van en outline (borde navy, fondo blanco).
    {
      type:     "flowchart",
      title:    "Proceso de Originación",
      subtitle: "5 etapas clave",
      steps: [
        { label: "ETAPA 1", body: "Identificación y due diligence inicial del activo" },
        { label: "ETAPA 2", body: "Estructuración financiera y aprobación de term sheet" },
        { label: "ETAPA 3", body: "Cierre y firma de documentos de crédito" },
      ],
    },

  ]
}
```

**Reglas de construcción del objeto:**

- `footerLeft` es siempre `"ESTRICTAMENTE PRIVADO Y CONFIDENCIAL"` — no cambiar nunca
- `projectName` se infiere del título del deck: "Valoración del Grupo Ejemplo" → `"GRUPO EJEMPLO"` (mayúsculas, sin artículos)
- Los campos marcados como "omitir si no hay" se omiten completamente del objeto (no poner `null` ni `""`)
- Los colores de tarjeta se expresan como hex sin `#`: `"1A2744"`, `"B8862A"`, `"8FA3B1"`
- El texto de bullets usa `•` como carácter de viñeta, no `-` ni `*`

### Paso 4 — Generar el artifact

1. Tomar el contenido completo del archivo `engine/artifact_template.html`
2. Localizar la línea exacta:
   ```js
   const PRESENTATION_DATA = __PLACEHOLDER_DATA__;
   ```
3. Reemplazarla con:
   ```js
   const PRESENTATION_DATA = { /* objeto construido en el Paso 3 */ };
   ```
4. Presentar el HTML resultante como un **artifact HTML** al usuario

**Importante:**
- NO explicar el código dentro del artifact ni comentar la implementación
- NO agregar texto antes del artifact explicando qué va a hacer — ir directo al artifact
- NO modificar ninguna otra parte del template — solo la línea del placeholder
- El artifact debe ser el HTML completo, no un fragmento

### Paso 5 — Instrucciones post-descarga

Después de mostrar el artifact, agregar un bloque corto (3-4 bullets)
con instrucciones para el usuario:

- Hacer clic en el botón "Descargar PPTX" dentro del artifact
- La imagen hero de la portada es un rectángulo navy de placeholder — reemplazarla en PowerPoint con una foto del sector relevante
- El archivo descargado tiene el grid Müller-Brockmann y la identidad CV completa aplicada (colores, tipografía, footer, logo)
- Para entregables finales de alta calidad, usar `/mb-generate` desde Claude Code CLI (incluye imagen real de Pexels y validador automático)

---

## 3. Limitaciones de la versión chat vs. CLI

| Limitación | Detalle |
|---|---|
| Sin imagen hero real | La portada muestra un rectángulo navy en lugar de una foto del sector. Reemplazar manualmente en PowerPoint. |
| Sin validador geométrico | No hay verificación automática de que el texto quepa en cada caja. Si un título es muy largo, puede desbordarse visualmente. |
| Sin API de Pexels | El CLI llama a Pexels para obtener una imagen contextual; la versión chat no tiene acceso a esa API. |
| Calidad de entregable | Esta versión es adecuada para borradores, prototipos rápidos y revisiones internas. Para entregables finales al cliente, usar la versión CLI. |

La versión CLI (`/mb-generate`) es la recomendada para cualquier
presentación que salga a un cliente externo.

---

## 4. Paleta de colores de referencia rápida

| Nombre | Hex | Uso principal |
|---|---|---|
| Navy | `1A2744` | Fondos de portada, footer, tarjetas primarias, cajas edge del flowchart |
| Dorado | `B8862A` | Supratítulo de portada, acentos, tarjetas secundarias |
| Gris azulado | `8FA3B1` | Tarjetas terciarias, elementos de apoyo |
| Blanco | `FFFFFF` | Texto sobre fondos oscuros, fondo de slides de contenido |

No usar colores fuera de esta paleta. Si el usuario pide un color
diferente, usar el más cercano de la paleta y mencionarlo.

---

## 5. Reglas de marca no negociables

Estas reglas aplican igual que en la versión CLI. No saltarlas bajo
ninguna circunstancia:

1. **Footer siempre navy con 3 segmentos** — izquierda (confidencialidad), centro (contexto/proyecto), derecha (numeración). No reemplazar el footer por un logo solo ni por una barra de otro color.

2. **Tipografía fija** — Cambria para títulos de slide y texto de portada; Arial para cuerpo de texto, bullets, notas, headers de tabla. No usar otras fuentes aunque el usuario las pida.

3. **Logo CV arriba derecha** — en todas las slides de contenido (no en la portada, donde va centrado sobre el footer). El logo es cuadrado y siempre va sin fondo visible.

4. **Sin colores fuera de la paleta** — navy, dorado, gris azulado y blanco. Fondos de slides de contenido siempre blancos.

5. **Portada siempre primera y única** — el deck tiene exactamente una portada, en la posición 0. No crear dos portadas ni portadas intermedias.

6. **`footerLeft` fijo** — siempre `"ESTRICTAMENTE PRIVADO Y CONFIDENCIAL"`, en todas las slides. No cambiar ni traducir este texto aunque el usuario lo pida.
