/**
 * mb_extractor.py — Extrae contenido estructurado de slides .pptx existentes
 *
 * Usa python-pptx porque tiene mejor soporte para leer XML de pptx que las
 * librerías de Node (adm-zip + xmldom solo llegan hasta el nivel de shapes
 * sin expandir grupos ni acceder a propiedades de tablas).
 *
 * Devuelve un objeto JSON con:
 * {
 *   slideW, slideH,           // tamaño real de la slide en inches
 *   elements: [               // elementos reconocidos
 *     { role, text, x, y, w, h, meta }
 *   ]
 * }
 *
 * Los roles posibles: "title", "subtitle", "company", "description",
 * "fin_table", "kpi_group", "chart", "closing_text", "footer_ignore", "unknown"
 *
 * IMPORTANTE: Este módulo solo EXTRAE y CLASIFICA — no toca el archivo
 * original, no escribe nada. El reconstructor (mb_reconstructor.js) es
 * quien genera el .pptx nuevo con la identidad CV aplicada.
 */

const { execSync } = require("child_process");
const path = require("path");
const fs   = require("fs");

const PYTHON_SCRIPT = path.join(__dirname, "extract_slide.py");

/**
 * Extrae el contenido estructurado de UNA slide específica de un .pptx.
 *
 * @param {string} pptxPath - ruta al archivo .pptx
 * @param {number} slideIdx - índice 0-based de la slide a extraer
 * @returns {object} - { slideW, slideH, elements[] }
 */
function extractSlide(pptxPath, slideIdx) {
  const result = execSync(
    `python3 "${PYTHON_SCRIPT}" "${pptxPath}" ${slideIdx}`,
    { encoding: "utf-8", timeout: 15000 }
  );
  return JSON.parse(result);
}

module.exports = { extractSlide };
