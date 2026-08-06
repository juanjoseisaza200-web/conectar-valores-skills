---
name: mb-format-chat
description: Reformatea un archivo .pptx EXISTENTE aplicando la identidad visual de Conectar Valores S.A.S. (grid Müller-Brockmann, paleta navy/dorado, footer corporativo, logo CV) directamente en claude.ai chat. El usuario sube su .pptx, Claude extrae el contenido y lo reconstruye con el template CV, produciendo un HTML artifact con botón de descarga. Usar cuando el usuario tenga un .pptx ya hecho que quiera llevar al formato CV. NO usar para crear presentaciones nuevas desde cero (eso es /mb-generate-chat).
---

# mb-format-chat

Reformatea un `.pptx` existente que el usuario sube al chat aplicando la
identidad visual de Conectar Valores: grid Müller-Brockmann, paleta
navy/dorado, footer corporativo de 3 segmentos y logo CV. Produce un HTML
artifact con botón de descarga, sin requerir instalación de ningún tipo.

---

## 1. Cuándo usar este skill

Activar cuando el usuario:

- Sube un `.pptx` y pide llevarlo al formato de Conectar Valores.
- Dice frases como "dame esto con el formato CV", "aplícale el grid",
  "reformatea mi presentación", "quiero que se vea como las presentaciones
  de CV", "aplica la identidad visual".
- Tiene un deck ya hecho y quiere alinearlo con los estándares de la firma.

**NO usar** si el usuario pide crear una presentación desde cero sin subir
archivo: en ese caso usar `/mb-generate-chat`.

**Regla de desempate:** si hay un archivo `.pptx` adjunto → este skill.
Si no hay archivo → `/mb-generate-chat`.

---

## 2. Limitación importante antes de empezar

Cuando claude.ai recibe un `.pptx` subido, puede extraer el **texto** de
las slides (títulos, cuerpos, tablas, listas) pero **NO puede ver**:

- El layout visual original (posiciones exactas de los cuadros de texto).
- Las imágenes embebidas (fotos, íconos, ilustraciones).
- Los colores y estilos de las formas originales.
- Los gráficos nativos de PowerPoint (barras, líneas, tortas).

El resultado es una **reconstrucción** del contenido con el template CV,
no una copia pixel-perfect del original. Esto debe comunicarse al usuario
antes de proceder, de forma breve y sin alarmar.

Si el archivo tiene slides muy fragmentadas (texto distribuido en muchos
cuadros pequeños y sin jerarquía clara), advertir que esa slide puede
perder estructura en el proceso de reconstrucción.

---

## 3. Flujo de trabajo (paso a paso para Claude)

### Paso 1 — Recibir y analizar el archivo

1. El usuario sube el `.pptx` al chat.
2. Claude lee el contenido que claude.ai extrae del archivo: títulos,
   textos de cuerpo, tablas, listas, notas del presentador si las hay.
3. Claude hace un inventario rápido slide por slide:
   - Número total de slides.
   - Tipo de contenido que parece tener cada una (portada, texto,
     tabla, lista de ítems paralelos, secuencia de pasos, etc.).
4. Presentar el inventario al usuario en forma compacta y confirmar:

   > "Encontré 8 slides. Identifiqué: 1 portada, 5 slides de texto/tabla,
   > 1 slide con ítems paralelos y 1 slide de cierre. Voy a reformatear
   > todo con el template CV. ¿Procedo, o hay slides que quieres excluir?"

   Si el usuario aprueba sin cambios, proceder directo al Paso 2.

---

### Paso 2 — Clasificar cada slide en un tipo CV

Para cada slide, determinar cuál de estos cuatro tipos aplica:

#### `cover` — Portada
Características: título grande, posiblemente logo de la empresa cliente,
fecha, ciudad. Normalmente es la primera slide.

Datos que puede faltar y hay que preguntar al usuario antes de generar:
- `client`: empresa o persona destinataria (ej. "Grupo Interaseo").
- `month`: mes y año de la presentación (ej. "Junio 2026").
- `city` y `country`: ciudad y país (ej. "Medellín, Colombia").

Si el original tiene esa información, extraerla directamente sin preguntar.

#### `content` — Slide de texto con o sin tabla
Características: un título, un bloque de texto o lista de puntos, y
opcionalmente una tabla de datos. Es el tipo más común.

Variantes de columnas:
- `"asym32"` — si hay tabla de datos: columna izquierda con texto (3/5
  del ancho), columna derecha con la tabla (2/5 del ancho). Usar cuando el
  original tiene texto explicativo + datos tabulares en la misma slide.
- `"single"` — si es solo texto o solo tabla que ocupa el ancho completo.

#### `cards` — Ítems paralelos
Características: 2 a 4 bloques de información con estructura similar
(características, ventajas, pilares, categorías, tranches). En el original
suelen aparecer como columnas lado a lado o como cuadros de colores.

Asignación de colores por posición:
- 1.ª tarjeta → `"1A2744"` (navy)
- 2.ª tarjeta → `"C8A85B"` (dorado)
- 3.ª tarjeta → `"6B7280"` (gris)
- 4.ª tarjeta → `"1A2744"` (navy, repite ciclo)

#### `flowchart` — Secuencia de pasos
Características: pasos numerados o conectados con flechas, proceso o
cronograma. En el original puede ser una slide de "proceso" o "timeline".

Cada paso tiene `label` (título del paso, texto corto) y `body` (detalle).

#### `unknown` — Slide que no encaja
Si una slide no calza en ningún tipo de los anteriores, tratarla como
`content` con `columns: "single"` y volcar todo el texto extraído en el
campo `text`. No inventar estructura que no estaba en el original.

**Slides con gráficos nativos de PowerPoint** (barras, líneas, tortas):
Incluir un texto placeholder exacto en la zona del gráfico:

```
[ gráfico del original — insertar manualmente en PowerPoint ]
```

No omitir la slide; mantenerla como `content` con ese placeholder en el
cuerpo.

---

### Paso 3 — Construir PRESENTATION_DATA

Con el contenido extraído y la clasificación del Paso 2, construir el
objeto `PRESENTATION_DATA` siguiendo exactamente este schema:

```js
const PRESENTATION_DATA = {
  projectName: "NOMBRE CORTO",       // inferido del título de la 1.ª slide de contenido
  footerLeft:  "ESTRICTAMENTE PRIVADO Y CONFIDENCIAL",
  slides: [
    {
      type: "cover",
      company:  "Conectar Valores S.A.S.",
      title:    "Título principal de la presentación",
      subtitle: "Subtítulo si existe",
      client:   "Empresa destinataria",
      month:    "Junio 2026",
      city:     "Medellín",
      country:  "Colombia"
    },
    {
      type:     "content",
      title:    "Título de la slide",
      subtitle: "Subtítulo si existe",
      columns:  "asym32",            // o "single"
      text:     "• Punto uno\n• Punto dos\n• Punto tres",
      table: {
        headers: ["Columna A", "Columna B", "Columna C"],
        rows: [
          ["Dato 1A", "Dato 1B", "Dato 1C"],
          ["Dato 2A", "Dato 2B", "Dato 2C"]
        ]
      }
    },
    {
      type:  "cards",
      title: "Título de la slide",
      cards: [
        { title: "Tarjeta 1", body: "Descripción de la tarjeta 1.", color: "1A2744" },
        { title: "Tarjeta 2", body: "Descripción de la tarjeta 2.", color: "C8A85B" },
        { title: "Tarjeta 3", body: "Descripción de la tarjeta 3.", color: "6B7280" }
      ]
    },
    {
      type:  "flowchart",
      title: "Título de la slide",
      steps: [
        { label: "Paso 1", body: "Descripción del paso 1." },
        { label: "Paso 2", body: "Descripción del paso 2." },
        { label: "Paso 3", body: "Descripción del paso 3." }
      ]
    }
  ]
};
```

**Reglas de extracción de texto:**

- Conservar el texto original de cada slide lo más fielmente posible.
  No parafrasear, no resumir, no añadir información.
- Bullets: condensar en el campo `text` con `"• "` al inicio de cada
  punto y `"\n"` entre puntos.
- Tablas: extraer como `{ headers, rows }`. Si la tabla estaba en el
  original, mantener exactamente los datos que tenía.
- `title`: tomar el texto más grande o prominente de la slide. Si hay
  varios candidatos, elegir el que aparece primero en el orden del archivo.
- Si no hay subtítulo en la slide original, omitir el campo `subtitle`
  completamente (no poner `""`).
- `projectName`: inferir a partir del título de la primera slide de
  contenido (la que sigue a la portada). Extraer las palabras clave en
  mayúsculas. Ejemplo: "Valoración del Grupo Interaseo" → `"GRUPO INTERASEO"`.

---

### Paso 4 — Generar el artifact

1. Leer el contenido del archivo
   `../mb-generate-chat/engine/artifact_template.html`.
2. Localizar la línea:
   ```js
   const PRESENTATION_DATA = __PLACEHOLDER_DATA__;
   ```
3. Reemplazar `__PLACEHOLDER_DATA__` con el objeto `PRESENTATION_DATA`
   completo construido en el Paso 3 (JS literal, sin comillas alrededor del
   objeto).
4. Presentar el resultado como un **HTML artifact** en el chat.
5. No explicar el código interno ni la estructura del template. Mostrar
   el artifact directamente con un mensaje breve de contexto, por ejemplo:

   > "Aquí está tu presentación reformateada con el template CV. Haz clic
   > en 'Descargar PPTX' para obtener el archivo."

---

### Paso 5 — Instrucciones post-descarga

Inmediatamente después del artifact, incluir este bloque de instrucciones
al usuario (adaptar según lo que aplique al archivo concreto):

> **Próximos pasos:**
> - Haz clic en "Descargar PPTX" dentro del artifact para obtener el archivo.
> - **Imagen de portada:** la imagen hero de la portada es un fondo navy
>   sólido. Si quieres una foto de fondo, reemplázala directamente en
>   PowerPoint (clic derecho sobre el fondo → Formato de fondo → Imagen).
> - **Gráficos del original:** las slides que tenían gráficos nativos de
>   PowerPoint quedaron con un texto placeholder. Inserta los gráficos
>   manualmente desde el archivo original usando copiar/pegar.
> - El resto del contenido tiene el grid MB y la identidad CV completa
>   aplicada (paleta navy/dorado, footer corporativo, logo CV, tipografía
>   Cambria/Arial).

Si el archivo original no tenía gráficos, omitir el punto sobre gráficos.

---

## 4. Qué hacer con slides sin clasificación clara

| Situación | Acción |
|---|---|
| No hay portada en el original | Crear una `cover` con el título del documento. Preguntar al usuario el cliente y la fecha antes de generar. |
| Slide de índice o agenda | Convertir a `content` con `columns: "single"` y los ítems como bullets en `text`. |
| Slide con 6 o más secciones densas | Dividir en 2 slides `content` si el contenido lo permite lógicamente. Si no es posible dividir, incluir todo en una sola slide `content` con `columns: "single"`. |
| Slide de cierre o "gracias" | Tratar como `content` con `columns: "single"`. Texto de cierre en el campo `text`. |
| Slide de solo imagen | Crear un `content` con `columns: "single"` y `text: "[ imagen del original — insertar manualmente en PowerPoint ]"`. |

---

## 5. Limitaciones honestas que comunicar al usuario

Decirle al usuario lo siguiente, de forma directa y sin tecnicismos, antes
de proceder (basta con 2-3 líneas en el mensaje de confirmación del Paso 1):

- Claude puede leer el texto del archivo pero **no el diseño visual**
  (colores, imágenes, posiciones exactas de los elementos).
- Los **gráficos nativos** (barras, líneas, tortas) no se pueden migrar
  automáticamente — quedan como marcadores de texto y hay que insertar los
  gráficos manualmente desde el original.
- El resultado es una **reconstrucción** del contenido con el template CV,
  no una réplica exacta del original.
- Para presentaciones muy complejas (más de 20 slides con layouts muy
  heterogéneos), puede ser más eficiente usar `/mb-format` (CLI) con el
  extractor Python completo, que sí accede al XML interno del archivo y
  puede mapear layouts con mayor precisión.

---

## 6. Diferencia con `/mb-format` (CLI)

| | `mb-format` (CLI) | `mb-format-chat` |
|---|---|---|
| Entorno | Claude Code CLI (terminal) | claude.ai chat (navegador) |
| Extracción del original | Python exacto vía `python-pptx` (lee el XML interno) | Claude lee el texto que claude.ai extrae del archivo |
| Charts / gráficos nativos | Placeholder marcado + coordenadas originales | Placeholder marcado (texto) |
| Imágenes embebidas | Puede extraer y reinsertar algunas | No disponible |
| Instalación requerida | Node.js + Python + dependencias del plugin | Ninguna — funciona en el chat |
| Capacidad recomendada | Decks complejos, >20 slides, alta fidelidad | Decks medianos, respuesta rápida |

---

## 7. Archivo de referencia del engine

El HTML artifact se genera a partir de:

```
../mb-generate-chat/engine/artifact_template.html
```

Este archivo contiene el motor completo de renderizado y generación de
`.pptx` en el cliente. El único cambio que hace `mb-format-chat` respecto
a `mb-generate-chat` es el origen del `PRESENTATION_DATA`: en
`mb-generate-chat` Claude inventa el contenido; aquí Claude lo **extrae
del archivo del usuario**. El template y el motor son idénticos.

No modificar `artifact_template.html` como parte de este skill. Si hay
que actualizar el motor, hacerlo en `mb-generate-chat/engine/` y este
skill hereda el cambio automáticamente.
