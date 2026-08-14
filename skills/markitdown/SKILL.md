---
name: markitdown
description: "Convierte casi cualquier archivo a Markdown con MarkItDown de Microsoft (PDF, Word .docx, PowerPoint .pptx, Excel .xlsx/.xls, imágenes, audio, HTML, CSV, JSON, XML, ZIP, EPub y URLs de YouTube). Usar siempre que haya que leer, extraer o convertir el contenido de un documento de oficina a texto o Markdown para analizarlo, resumirlo, compararlo, auditarlo o citarlo — incluso si el usuario no dice 'markdown', p. ej. 'lee este PDF', 'qué dice esta presentación', 'saca el texto del docx', 'revisa los documentos de este ZIP'."
---

# MarkItDown — archivos a Markdown

Wrapper de [microsoft/markitdown](https://github.com/microsoft/markitdown) (paquete Python, ya instalado con `pip install "markitdown[all]"`). Convierte documentos a Markdown limpio preservando estructura (títulos, tablas, listas, links), que es el formato que mejor entiende un LLM.

Úsala como primera opción para leer el contenido de un documento que la herramienta Read no maneja bien de forma nativa (.docx, .pptx, .xlsx, HTML complejo, EPub, ZIP con varios documentos). Para PDFs, Read ya los muestra; usa markitdown cuando necesites el texto como archivo .md (para diff, grep, citas o alimentar otro proceso).

## Uso CLI (lo normal)

```powershell
markitdown "ruta\al\archivo.pptx" -o "salida.md"
```

- **Siempre `-o archivo.md`, nunca redirigir con `>`** — en PowerShell `>` escribe UTF-16/BOM y corrompe el encoding.
- La salida es **UTF-8**. Leerla con la herramienta Read o con `Get-Content -Encoding utf8`; sin el flag, PowerShell la muestra con mojibake (Ã­, Ã³) aunque el archivo esté bien.
- Acepta URLs de YouTube directamente (baja la transcripción): `markitdown "https://youtube.com/watch?v=..." -o video.md`.
- Un ZIP se convierte entero: itera los archivos internos y concatena sus conversiones en un solo .md.

## Uso Python (batch o texto en proceso)

Para convertir muchos archivos o cuando el texto se necesita en memoria sin archivo intermedio:

```python
from markitdown import MarkItDown

md = MarkItDown()
result = md.convert(r"C:\ruta\informe.docx")
texto = result.text_content  # str en Unicode

# Al guardar, siempre UTF-8:
with open("salida.md", "w", encoding="utf-8") as f:
    f.write(texto)
```

## Cobertura y límites por formato

| Formato | Qué sale | Límite a tener en cuenta |
|---|---|---|
| PDF | Texto por capa de texto (pdfminer) | **Sin OCR**: un PDF escaneado sale vacío — usar Tesseract u otro OCR aparte |
| .docx | Títulos, tablas, listas | Comentarios y track changes no salen |
| .pptx | Texto por slide, tablas, notas | La geometría/diseño se pierde (es solo contenido) |
| .xlsx / .xls | Una tabla Markdown por hoja | Solo valores visibles — no fórmulas, no formato; para auditar fórmulas usar openpyxl/COM |
| Imágenes | Metadatos EXIF | La descripción del contenido requiere pasar un `llm_client` en la API Python; sin eso no hay caption |
| Audio | Transcripción | Requiere ffmpeg instalado (hoy **no** está en esta máquina) y conexión a internet |
| HTML, CSV, JSON, XML, EPub | Conversión directa | — |

## Errores frecuentes

- `RuntimeWarning: Couldn't find ffmpeg` al arrancar: inofensivo salvo que se esté convirtiendo audio; ignorarlo.
- Salida vacía en un PDF: casi seguro es un escaneado sin capa de texto → OCR.
- Caracteres rotos (Ã±, Ã³) al leer la salida: el archivo está bien; se leyó sin `-Encoding utf8`.
