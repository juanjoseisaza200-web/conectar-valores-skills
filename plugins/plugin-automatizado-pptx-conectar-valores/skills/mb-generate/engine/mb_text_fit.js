/**
 * mb_text_fit.js — Estimación compartida de desborde de texto
 * =================================================================
 * Función única usada por deriveHeaderBlock (slides de contenido) y
 * buildCover (portada) para estimar cuántas líneas ocupará un texto
 * y, si hace falta, reducir el tamaño de fuente para que quepa en un
 * máximo de líneas. Evita mantener dos heurísticas distintas que
 * puedan desincronizarse (como pasó: el motor general tenía una
 * heurística, mb_cover.js tenía otra más débil con un umbral de
 * caracteres fijo que no consideraba ancho real ni tamaño de fuente).
 */

/**
 * @param {string} text - el texto real a medir
 * @param {number} widthIn - ancho disponible en inches
 * @param {number} fontSize - tamaño de fuente deseado en pt
 * @param {number} maxLines - máximo de líneas aceptable (default 2)
 * @returns {{fontSize: number, lines: number}}
 */
function fitTextLines(text, widthIn, fontSize, maxLines = 2) {
  if (text == null) text = "";
  // ~0.58 es el ancho promedio de un carácter en pt relativo al tamaño de
  // fuente, calibrado para texto BOLD (más ancho que regular) en fuentes
  // sans/serif estándar. Conservador a propósito: prefiere reducir la
  // fuente de más a dejar que el texto se desborde de su caja.
  const CHAR_WIDTH_RATIO = 0.58;
  // piso absoluto de legibilidad — nunca reducimos más allá de esto,
  // pero TAMPOCO debe superar nunca el fontSize solicitado por el caller
  const ABSOLUTE_FLOOR = 6;
  const floor = Math.min(ABSOLUTE_FLOOR, fontSize);

  let size = fontSize;
  let charsPerLine = (widthIn * 72) / (size * CHAR_WIDTH_RATIO);
  let lines = Math.max(1, Math.ceil(text.length / charsPerLine));

  // reduce en pasos de 1pt hasta que quepa, hasta el piso. Suficientes
  // iteraciones para cubrir el rango completo (ej. 24pt → 6pt = 18 pasos).
  let attempts = 0;
  const maxAttempts = Math.ceil(fontSize - floor) + 2;
  while (lines > maxLines && size > floor && attempts < maxAttempts) {
    size = Math.max(floor, size - 1);
    charsPerLine = (widthIn * 72) / (size * CHAR_WIDTH_RATIO);
    lines = Math.max(1, Math.ceil(text.length / charsPerLine));
    attempts++;
  }

  // reporte HONESTO: si tras reducir al piso el texto sigue ocupando más
  // líneas que maxLines, lo decimos — el caller debe reservar esa altura
  // real, no fingir que cabe en maxLines cuando no es así.
  return { fontSize: size, lines };
}

module.exports = { fitTextLines };
