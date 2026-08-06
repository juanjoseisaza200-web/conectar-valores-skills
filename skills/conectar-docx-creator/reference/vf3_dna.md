# vf3_dna.md — ADN visual literal del Informe Peers VF3

Extraído por inspección directa del XML de `20260513_Informe_Peers_Aseo_VF3.docx` (ver `assets/vf3_dna/*.json` para datos crudos).

⚠️ **Esta es la fuente única de verdad.** Si algo aquí contradice memoria o conversaciones previas, **gana este documento.**

## 1. Página

- Tamaño: Letter (w=12240, h=15840 twips = 8.5 × 11 in)
- Márgenes: **top=1728, right=1138, bottom=1138, left=1138 twips**
  - = **1.20" / 0.79" / 0.79" / 0.79"**
- Header offset: 720 twips (0.5")
- Footer offset: 720 twips (0.5")

## 2. Tipografía global

- **Fuente única:** Arial (declarada en 51 runs explícitos + theme.minorFont)
- **Tamaños** (conteo real de runs en VF3):
  - 9.0 pt → 1,252 runs (DOMINANTE — cuerpo de tabla, párrafos)
  - 9.5 pt → 156 runs (headers de tabla, sub-énfasis)
  - 8.5 pt → 55 runs (footnotes, captions)
  - 12.0 pt → 30 runs (títulos H1)
  - 22.0, 16.0, 13.0, 11.0 → portada / títulos especiales
- **Colores texto** (en runs con color explícito):
  - Navy `17375E` → 399 runs (títulos + headers tabla + fila destacada)
  - Teal `5D9E9D` → 1 run (decoración mínima)
  - Resto: heredado (negro) — no se setea color explícito en cuerpo
- **Alineaciones** (conteo real):
  - `center` → 1,456 párrafos (DOMINANTE — todo lo de tablas)
  - `both` (justify) → 191 párrafos (párrafos de cuerpo)

## 3. Paleta hex auditada (única autorizada)

| Color | Hex | Uso real medido en VF3 | Conteo |
|---|---|---|---|
| Navy | `#17375E` | Texto títulos, headers tabla, bordes ocasionales | 399 runs + 25 bordes |
| Dorado CV | `#C9A449` | Borde inferior header de tabla (ÚNICO uso) | 154 bordes |
| Hairline | `#BFBFBF` | Borde inferior filas cuerpo (ÚNICO uso) | 1,055 bordes |
| Mint | `#E7F0EF` | Fill fila destacada (cliente) — ÚNICO fill | 76 fills |
| Teal | `#5D9E9D` | Texto decorativo aislado | 1 run |

⚠️ **NO existen en VF3:**
- Banded rows en gris claro (`#EDEDED` o similar)
- Filas totales con fondo Navy + texto blanco
- Header de tabla pintado en Navy
- Bordes laterales en tablas
- Cualquier color fuera de esta lista

## 4. ESTILO TABLA — ADN exacto

### 4.1 Propiedades de tabla (`tblPr`)

```xml
<w:tblPr>
  <w:tblW w:w="0" w:type="auto"/>          <!-- ancho automático -->
  <w:jc w:val="center"/>                    <!-- TABLA CENTRADA en página -->
  <w:tblLayout w:type="fixed"/>             <!-- ancho fijo de columnas -->
  <w:tblLook w:val="04A0" w:firstRow="1" 
             w:lastRow="0" w:firstColumn="1" 
             w:lastColumn="0" w:noHBand="0" 
             w:noVBand="1"/>                <!-- sin bandas verticales -->
</w:tblPr>
```

Notar: **no hay `tblBorders`** a nivel tabla (los bordes se ponen por celda).

### 4.2 Header (primera fila — `tblHeader`)

```xml
<w:tr>
  <w:trPr>
    <w:tblHeader/>                <!-- marca como header (se repite si tabla se parte) -->
    <w:jc w:val="center"/>
  </w:trPr>
  <w:tc>
    <w:tcPr>
      <w:tcW w:w="..." w:type="dxa"/>
      <w:tcBorders>
        <w:top val="nil"/>
        <w:left val="nil"/>
        <w:bottom val="single" sz="12" color="C9A449"/>  <!-- DORADO 1.5pt -->
        <w:right val="nil"/>
      </w:tcBorders>
      <w:vAlign val="center"/>
      <!-- SIN <w:shd> — fondo blanco -->
    </w:tcPr>
    <w:p>
      <w:pPr>
        <w:spacing beforeLines="20" before="48" afterLines="20" after="48"/>
        <w:jc val="center"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>                     <!-- BOLD -->
          <w:color val="17375E"/>    <!-- NAVY -->
          <w:sz val="19"/>           <!-- 9.5pt (sz en half-points) -->
        </w:rPr>
        <w:t>Texto header</w:t>
      </w:r>
    </w:p>
  </w:tc>
  ...
</w:tr>
```

**Resumen header:**
- Fondo: **BLANCO** (sin `<w:shd>`)
- Texto: **Navy `#17375E` bold 9.5pt centrado**
- Bordes: SOLO `bottom` dorado `#C9A449` sz=12 (1.5pt); resto `nil`
- Spacing párrafo: `beforeLines=20 before=48 afterLines=20 after=48`
- vAlign: center

### 4.3 Filas cuerpo (regulares)

```xml
<w:tr>
  <w:trPr><w:jc val="center"/></w:trPr>
  <w:tc>
    <w:tcPr>
      <w:tcW w:w="..." w:type="dxa"/>
      <w:tcBorders>
        <w:top val="nil"/>
        <w:left val="nil"/>
        <w:bottom val="single" sz="2" color="BFBFBF"/>  <!-- HAIRLINE 0.25pt -->
        <w:right val="nil"/>
      </w:tcBorders>
      <w:vAlign val="center"/>
      <!-- SIN <w:shd> — fondo blanco -->
    </w:tcPr>
    <w:p>
      <w:pPr>
        <w:spacing beforeLines="20" before="48" afterLines="20" after="48"/>
        <w:jc val="center"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:sz val="18"/>           <!-- 9pt -->
          <!-- SIN bold, SIN color (hereda Normal = negro) -->
        </w:rPr>
        <w:t>Contenido celda</w:t>
      </w:r>
    </w:p>
  </w:tc>
  ...
</w:tr>
```

**Resumen fila cuerpo:**
- Fondo: **BLANCO** (sin `<w:shd>`)
- Texto: **negro regular 9pt centrado**
- Bordes: SOLO `bottom` hairline `#BFBFBF` sz=2 (0.25pt); resto `nil`
- Spacing: idéntico al header
- vAlign: center

### 4.4 Fila destacada (cliente / Interaseo)

```xml
<w:tr>
  <w:trPr><w:jc val="center"/></w:trPr>
  <w:tc>
    <w:tcPr>
      <w:tcW w:w="..."/>
      <w:tcBorders>
        <w:top val="nil"/>
        <w:left val="nil"/>
        <w:bottom val="single" sz="2" color="BFBFBF"/>
        <w:right val="nil"/>
      </w:tcBorders>
      <w:shd val="clear" color="auto" fill="E7F0EF"/>    <!-- MINT FILL -->
      <w:vAlign val="center"/>
    </w:tcPr>
    <w:p>
      <w:pPr>
        <w:spacing before="20" after="20"/>             <!-- spacing simplificado -->
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>                                         <!-- BOLD -->
          <w:color val="17375E"/>                        <!-- NAVY -->
          <w:sz val="18"/>                               <!-- 9pt -->
        </w:rPr>
        <w:t>Interaseo</w:t>
      </w:r>
    </w:p>
  </w:tc>
  ...
</w:tr>
```

**Resumen fila destacada:**
- Fondo: **Mint `#E7F0EF`**
- Texto: **Navy `#17375E` bold 9pt centrado**
- Bordes: SOLO `bottom` hairline `#BFBFBF` sz=2
- Spacing: **`before=20 after=20`** (más simple — sin `beforeLines`)
- vAlign: center

### 4.5 Última fila (puede ser blanca sin borde — separador visual)

Encontrada en última fila de tabla idx=6:
- Bordes TODOS `nil` (incluye bottom)
- Celda vacía o casi vacía
- Marca el cierre visual

### 4.6 Conversión de unidades

- **sz en `<w:sz>`** está en **half-points** → sz=18 → 9pt, sz=19 → 9.5pt, sz=24 → 12pt
- **sz en bordes `<w:bottom sz=...>`** está en **eighths of a point** → sz=2 → 0.25pt, sz=4 → 0.5pt, sz=12 → 1.5pt
- **Twips** (medida XML estándar): 1440 twips = 1 pulgada; 20 twips = 1 point

## 5. ESTILOS PÁRRAFO

(Inventario completo en `assets/vf3_dna/styles.json` — 168 estilos)

Estilos clave usados:
- `Normal` (heredado por todo)
- `Heading 1`, `Heading 2`, `Heading 3` (títulos secciones)
- `Caption` (encima de tablas/figuras)
- Custom CV (varios `pXXX` y `aXXX` con basedOn=Normal)

Spec heading (verificada en `styles.json`):
- H1: 12pt bold Navy `17375E`, centrado
- H2: 10pt bold Navy
- H3: 9.5pt bold Navy

## 6. HEADER y FOOTER

- **Header1:** vacío (no logo, no texto)
- **Footer1:** `"Informe de Peers • Sector Aseo Colombia | Confidencial • Uso interno Conectar Valores S.A.S."`
  - Centrado, Slate, 8.5pt
  - (Ver `assets/vf3_dna/headers_footers.json` para XML completo)

## 7. Diferencias críticas vs memoria/skill anterior

| Asunción previa (ERRADA) | Realidad VF3 |
|---|---|
| Header tabla pintado Navy + texto blanco | Header **blanco con texto Navy bold** + sólo línea inferior dorada |
| Filas alternadas en gris `#EDEDED` | **NO existen bandas** — todas blancas excepto destacada Mint |
| Filas Total con fondo Navy claro | **NO existe estilo Total** distinto en VF3 |
| Hairline sz=4 (0.5pt) entre filas | **sz=2 (0.25pt)** |
| Header tamaño 9pt | Header **9.5pt** |
| Texto cuerpo justificado | **Centrado** (en celdas de tabla) |
| Padding 72/115 twips | `beforeLines=20 before=48 afterLines=20 after=48` (más como margen vertical de párrafo) |

## 8. Reglas de oro para reproducir el formato

1. **Antes de generar cualquier .docx**, leer este archivo y los JSON en `assets/vf3_dna/`.
2. **Para clonar la apariencia exacta**, usar `assets/vf3_blank_template.docx` (template clonado del VF3 con cuerpo vaciado) como punto de partida.
3. **Sólo agregar nuevo contenido**, NO crear estilos paralelos.
4. **Validar después** comparando el XML del nuevo doc contra los JSON de referencia (mismos bordes, mismos colores, mismos sizes).
