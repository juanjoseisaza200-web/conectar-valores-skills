#!/usr/bin/env node
/**
 * mb_format.js — CLI principal de mb-format
 *
 * Uso:
 *   node mb_format.js <input.pptx> <slide_range> <output.pptx>
 *
 * Ejemplos:
 *   node mb_format.js ./deck.pptx 95        ./deck_cv_s95.pptx
 *   node mb_format.js ./deck.pptx 95-119    ./deck_cv.pptx
 *   node mb_format.js ./deck.pptx all       ./deck_cv_full.pptx
 *
 * El script lee el rango de slides indicado, aplica la detección de template
 * y la reconstrucción con identidad CV a cada una, y guarda el resultado.
 * Las slides que no tienen template reconocido se guardan con un placeholder
 * de diagnóstico (no se pierden, pero tampoco se reformatean a ciegas).
 */

const pptxgen = require("pptxgenjs");
const path    = require("path");
const { execSync } = require("child_process");
const { reconstructSlide } = require("./mb_reconstructor.js");

const PYTHON_EXTRACTOR = path.join(__dirname, "extract_slide.py");

// ─── PARSEO DE ARGUMENTOS ────────────────────────────────────────────────────

const [,, inputPath, rangeArg, outputPath] = process.argv;

if (!inputPath || !rangeArg || !outputPath) {
  console.error("Uso: node mb_format.js <input.pptx> <slide_num|range|all> <output.pptx>");
  process.exit(1);
}

function parseRange(arg, totalSlides) {
  if (arg === "all") return Array.from({ length: totalSlides }, (_, i) => i);
  if (arg.includes("-")) {
    const [start, end] = arg.split("-").map(Number);
    return Array.from({ length: end - start + 1 }, (_, i) => start - 1 + i);
  }
  return [Number(arg) - 1];
}

// ─── OBTENER TOTAL DE SLIDES ─────────────────────────────────────────────────

function getTotalSlides(pptxPath) {
  const result = execSync(
    `python3 -c "from pptx import Presentation; p=Presentation('${pptxPath}'); print(len(p.slides))"`,
    { encoding: "utf-8", timeout: 10000 }
  );
  return parseInt(result.trim(), 10);
}

// ─── EXTRAER UNA SLIDE ────────────────────────────────────────────────────────

function extractSlide(pptxPath, slideIdx) {
  const result = execSync(
    `python3 "${PYTHON_EXTRACTOR}" "${pptxPath}" ${slideIdx}`,
    { encoding: "utf-8", timeout: 15000 }
  );
  return JSON.parse(result);
}

// ─── MAIN ────────────────────────────────────────────────────────────────────

async function main() {
  console.log(`\nmb-format — procesando: ${path.basename(inputPath)}`);

  const total = getTotalSlides(inputPath);
  const indices = parseRange(rangeArg, total);
  console.log(`Slides a procesar: ${indices.map(i => i+1).join(", ")} (${indices.length} total)\n`);

  // Detectar tamaño de slide del archivo original para configurar la presentación
  const firstExtracted = extractSlide(inputPath, indices[0]);
  const { slideW, slideH } = firstExtracted;

  const pres = new pptxgen();
  pres.defineLayout({ name: "ORIGINAL", width: slideW, height: slideH });
  pres.layout = "ORIGINAL";

  const report = [];

  for (const idx of indices) {
    process.stdout.write(`  Slide ${idx + 1}... `);
    try {
      const extracted = idx === indices[0] ? firstExtracted : extractSlide(inputPath, idx);
      const slide = pres.addSlide();
      const result = reconstructSlide(slide, pres, extracted);
      const status = result.ok ? "✓" : `⚠ ${result.message}`;
      process.stdout.write(`${status} [${result.template}]\n`);
      report.push({ slide: idx + 1, ...result });
    } catch (e) {
      process.stdout.write(`✗ ERROR: ${e.message}\n`);
      report.push({ slide: idx + 1, ok: false, template: "error", message: e.message });
    }
  }

  await pres.writeFile({ fileName: outputPath });

  // Resumen
  const ok  = report.filter(r => r.ok).length;
  const err = report.filter(r => !r.ok).length;
  console.log(`\n${"─".repeat(60)}`);
  console.log(`  ${ok} slides reformateadas  |  ${err} sin template reconocido`);
  console.log(`  Output: ${outputPath}`);
  console.log(`${"─".repeat(60)}\n`);

  // Registrar en observations.md si hay algo notable
  if (err > 0) {
    const note = `\n### ${new Date().toISOString().split("T")[0]} — PROBLEMA — mb-format clasificación\n\n**Contexto:** ejecución de mb_format.js sobre ${path.basename(inputPath)}, rango ${rangeArg}.\n\n**Qué pasó:** ${err} de ${indices.length} slides no tuvieron template reconocido. Templates detectados: ${[...new Set(report.map(r => r.template))].join(", ")}.\n\n**Por qué importa:** el clasificador de mb_reconstructor.js::detectTemplate() no cubrió estos casos. Ver el output de la consola para los roles detectados en cada slide sin template.\n\n**Acción tomada:** slides con placeholder de diagnóstico incluidas en el output — no se perdió contenido, pero tampoco se reformatearon.\n`;
    const obsPath = path.join(__dirname, "..", "..", "mb-generate", "logs", "observations.md");
    try {
      const fs = require("fs");
      fs.appendFileSync(obsPath, note, "utf-8");
      console.log("  (Observación anotada en logs/observations.md)");
    } catch (e) {
      // no fatal si no puede escribir
    }
  }
}

main().catch(e => { console.error(e); process.exit(1); });
