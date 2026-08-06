/**
 * mb_validator.js — Validador programático de slides MB
 * ============================================================
 * Lee el XML interno de un .pptx y verifica, para cada shape/texto:
 *   1. Ningún elemento sale de los límites del slide
 *   2. Ningún elemento de contenido invade la zona de footer
 *   3. Ningún elemento de contenido invade la zona de notas
 *   4. No hay overlaps entre cajas de texto/shapes que no deberían solaparse
 *      (excluye overlaps intencionales como fondos decorativos bajo texto)
 *
 * Uso: node mb_validator.js <archivo.pptx> [--footer-y=5.345] [--notes-y=4.8]
 *
 * Esto NO reemplaza la revisión visual — detecta violaciones geométricas
 * duras. Problemas de legibilidad, contraste o estética requieren ojo humano.
 */

const fs       = require("fs");
const path     = require("path");
const AdmZip   = require("adm-zip");
const { DOMParser } = require("@xmldom/xmldom");

const EMU_PER_INCH = 914400;

function emuToIn(emu) {
  return emu / EMU_PER_INCH;
}

/** Extrae todos los shapes (posición + tamaño) de un slide XML. */
function extractShapes(slideXml) {
  const doc = new DOMParser().parseFromString(slideXml, "text/xml");
  const shapes = [];

  // namespaces: a (drawingml), p (presentationml)
  const spTree = doc.getElementsByTagName("p:spTree")[0];
  if (!spTree) return shapes;

  function walk(node, depth = 0) {
    for (let i = 0; i < node.childNodes.length; i++) {
      const child = node.childNodes[i];
      if (child.nodeType !== 1) continue; // element nodes only

      const tag = child.tagName;
      if (tag === "p:sp" || tag === "p:pic" || tag === "p:graphicFrame" || tag === "p:cxnSp") {
        const xfrm = findXfrm(child);
        if (xfrm) {
          const name = getShapeName(child);
          shapes.push({ tag, name, ...xfrm });
        }
      }
      walk(child, depth + 1);
    }
  }

  walk(spTree);
  return shapes;
}

function getShapeName(node) {
  // busca p:nvSpPr/p:cNvPr@name o equivalentes
  for (let i = 0; i < node.childNodes.length; i++) {
    const child = node.childNodes[i];
    if (child.nodeType !== 1) continue;
    if (child.tagName && child.tagName.startsWith("p:nv")) {
      const cNvPr = child.getElementsByTagName("p:cNvPr")[0];
      if (cNvPr) return cNvPr.getAttribute("name") || "unnamed";
    }
  }
  return "unnamed";
}

function findXfrm(node) {
  // busca a:xfrm en cualquier nivel del spPr (o xfrm directo en graphicFrame)
  const xfrmNodes = node.getElementsByTagName("a:xfrm");
  if (xfrmNodes.length === 0) {
    // graphicFrame usa p:xfrm
    const pXfrm = node.getElementsByTagName("p:xfrm");
    if (pXfrm.length === 0) return null;
    return parseXfrm(pXfrm[0]);
  }
  return parseXfrm(xfrmNodes[0]);
}

function parseXfrm(xfrmNode) {
  const off = xfrmNode.getElementsByTagName("a:off")[0];
  const ext = xfrmNode.getElementsByTagName("a:ext")[0];
  if (!off || !ext) return null;
  const x = parseInt(off.getAttribute("x"), 10);
  const y = parseInt(off.getAttribute("y"), 10);
  const w = parseInt(ext.getAttribute("cx"), 10);
  const h = parseInt(ext.getAttribute("cy"), 10);
  return {
    xIn: emuToIn(x), yIn: emuToIn(y),
    wIn: emuToIn(w), hIn: emuToIn(h),
    x2In: emuToIn(x + w), y2In: emuToIn(y + h),
  };
}

/** Verifica si dos rectángulos se solapan (con tolerancia mínima para bordes). */
function rectsOverlap(a, b, tolerance = 0.02) {
  return !(
    a.x2In <= b.xIn + tolerance ||
    b.x2In <= a.xIn + tolerance ||
    a.y2In <= b.yIn + tolerance ||
    b.y2In <= a.yIn + tolerance
  );
}

/**
 * Valida un slide contra las reglas MB.
 * @param {Array} shapes - resultado de extractShapes
 * @param {object} bounds - { slideW, slideH, footerY, notesY (opcional) }
 */
function validateSlide(shapes, bounds, slideIndex) {
  const issues = [];
  const { slideW, slideH, footerY, notesY } = bounds;

  // shapes que pertenecen a las zonas de footer/fondo se excluyen de la
  // regla "no debe estar en zona de footer" (se identifican por nombre o posición)
  const footerShapes = shapes.filter(s => s.yIn >= footerY - 0.01);
  const contentShapes = shapes.filter(s => s.yIn < footerY - 0.01);

  // 1. ningún shape de CONTENIDO debe extenderse dentro de la zona de footer
  for (const s of contentShapes) {
    if (s.y2In > footerY + 0.02) {
      issues.push({
        type: "FOOTER_OVERLAP",
        slide: slideIndex,
        shape: s.name,
        detail: `Shape termina en y=${s.y2In.toFixed(3)}in, footer empieza en y=${footerY.toFixed(3)}in`,
      });
    }
  }

  // 2. si hay zona de notas, ningún shape de bloques de contenido debe invadirla
  //    (esto requiere distinguir notas vs contenido — aproximación: shapes
  //     que empiezan antes de notesY pero terminan después, y no son las notas mismas)
  if (notesY != null) {
    for (const s of contentShapes) {
      if (s.yIn < notesY - 0.02 && s.y2In > notesY + 0.02 && !s.name.toLowerCase().includes("nota")) {
        issues.push({
          type: "NOTES_ZONE_OVERLAP",
          slide: slideIndex,
          shape: s.name,
          detail: `Shape cruza la zona de notas en y=${notesY.toFixed(3)}in (${s.yIn.toFixed(3)} → ${s.y2In.toFixed(3)})`,
        });
      }
    }
  }

  // 3. ningún shape debe salir de los límites del slide
  for (const s of shapes) {
    if (s.xIn < -0.02 || s.yIn < -0.02 || s.x2In > slideW + 0.02 || s.y2In > slideH + 0.02) {
      issues.push({
        type: "OUT_OF_BOUNDS",
        slide: slideIndex,
        shape: s.name,
        detail: `Posición (${s.xIn.toFixed(3)}, ${s.yIn.toFixed(3)}) a (${s.x2In.toFixed(3)}, ${s.y2In.toFixed(3)}) excede slide ${slideW}x${slideH}`,
      });
    }
  }

  // 4. overlaps entre shapes de TEXTO que probablemente no deberían solaparse.
  //    Se excluyen los casos de "fondo decorativo + texto encima" — el patrón
  //    típico es: un rectángulo sin texto propio (o con muy poco) que sirve de
  //    fondo, y un textbox que ocupa un área igual o menor totalmente contenida.
  //    Estos son intencionales en tarjetas, cajas de flowchart, headers, etc.
  const textLike = shapes.filter(s => s.tag === "p:sp");
  for (let i = 0; i < textLike.length; i++) {
    for (let j = i + 1; j < textLike.length; j++) {
      const a = textLike[i], b = textLike[j];
      if (!rectsOverlap(a, b)) continue;

      const overlapArea = overlapAreaOf(a, b);
      const areaA = a.wIn * a.hIn;
      const areaB = b.wIn * b.hIn;
      const minArea = Math.min(areaA, areaB);
      const maxArea = Math.max(areaA, areaB);
      if (minArea === 0) continue;

      const overlapRatioOfMin = overlapArea / minArea;
      if (overlapRatioOfMin <= 0.15) continue;

      // patrón fondo+texto: el shape MENOR está geométricamente contenido
      // (con tolerancia) dentro del shape MAYOR. No importa el ratio de áreas,
      // lo que importa es si uno cabe dentro del otro — eso es "texto sobre
      // tarjeta/caja de fondo", un patrón intencional y común.
      const smaller = areaA <= areaB ? a : b;
      const larger  = areaA <= areaB ? b : a;
      const tol = 0.05; // in
      const isContained =
        smaller.xIn  >= larger.xIn  - tol &&
        smaller.yIn  >= larger.yIn  - tol &&
        smaller.x2In <= larger.x2In + tol &&
        smaller.y2In <= larger.y2In + tol;

      if (isContained) continue; // overlap intencional — no reportar

      issues.push({
        type: "TEXT_OVERLAP",
        slide: slideIndex,
        shape: `${a.name} ↔ ${b.name}`,
        detail: `Overlap ${(100 * overlapRatioOfMin).toFixed(0)}% del área menor (no es contención fondo/texto)`,
      });
    }
  }

  return issues;
}

function overlapAreaOf(a, b) {
  const xOverlap = Math.max(0, Math.min(a.x2In, b.x2In) - Math.max(a.xIn, b.xIn));
  const yOverlap = Math.max(0, Math.min(a.y2In, b.y2In) - Math.max(a.yIn, b.yIn));
  return xOverlap * yOverlap;
}

/** Valida un archivo .pptx completo. */
function validatePptx(pptxPath, boundsPerSlide) {
  const zip = new AdmZip(pptxPath);
  const slideEntries = zip.getEntries()
    .filter(e => /^ppt\/slides\/slide\d+\.xml$/.test(e.entryName))
    .sort((a, b) => {
      const na = parseInt(a.entryName.match(/slide(\d+)/)[1], 10);
      const nb = parseInt(b.entryName.match(/slide(\d+)/)[1], 10);
      return na - nb;
    });

  const allIssues = [];
  slideEntries.forEach((entry, idx) => {
    const xml = entry.getData().toString("utf-8");
    const shapes = extractShapes(xml);
    const bounds = boundsPerSlide[idx] || boundsPerSlide.default;
    const issues = validateSlide(shapes, bounds, idx + 1);
    allIssues.push(...issues);
  });

  return allIssues;
}

module.exports = { validatePptx, extractShapes, validateSlide };

// ─── CLI ──────────────────────────────────────────────────────────────────
if (require.main === module) {
  const pptxPath = process.argv[2];
  if (!pptxPath) {
    console.error("Uso: node mb_validator.js <archivo.pptx>");
    process.exit(1);
  }

  // bounds default — slide 16:9 estándar pptxgenjs
  const defaultBounds = {
    slideW: 10, slideH: 5.625,
    footerY: 5.345,
    notesY: null,
  };

  const issues = validatePptx(pptxPath, { default: defaultBounds });

  if (issues.length === 0) {
    console.log(`✓ ${pptxPath} — sin violaciones geométricas detectadas`);
  } else {
    console.log(`✗ ${pptxPath} — ${issues.length} problema(s) detectado(s):\n`);
    issues.forEach((iss) => {
      console.log(`  [Slide ${iss.slide}] ${iss.type}: ${iss.shape}`);
      console.log(`    ${iss.detail}`);
    });
    process.exit(1);
  }
}
