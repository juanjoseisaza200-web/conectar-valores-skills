/**
 * mb_engine.js — Motor de grid Müller-Brockmann para PowerPoint
 * ================================================================
 * Librería reutilizable. Calcula zonas, deriva el grid y expone
 * helpers de construcción. No genera contenido — eso lo decide
 * quien llama (el skill / Claude).
 *
 * Regla de oro: las zonas se reservan de abajo hacia arriba con
 * respiros explícitos. Nada cruza un límite de zona.
 */

const fs    = require("fs");
const path  = require("path");
const { fitTextLines } = require("./mb_text_fit.js");

// ─── DIMENSIONES BASE PPTXGENJS (LAYOUT_16x9) ────────────────────────────────
const SLIDE_W = 10;
const SLIDE_H = 5.625;

// ─── PALETA CV ───────────────────────────────────────────────────────────────
const PALETTE_CV = {
  navy:      "1A2744",
  gold:      "B8862A",
  goldLight: "C9A84C",
  white:     "FFFFFF",
  text:      "1A2744",
  muted:     "6B7280",
  tblHeader: "1A2744",
  tblAlt:    "F2F4F7",
  tblTotal:  "E4E8EF",
  tblBorder: "D8DCE4",
  cardGold:  "C9A352",
  cardNavy:  "1A2744",
  cardGray:  "8FA3B1",
};

const FONT = { serif: "Cambria", sans: "Arial" };

/**
 * Deriva el grid MB completo a partir de parámetros base.
 * @param {object} opts
 *   bodySize: tamaño cuerpo de texto en pt (default 11)
 *   leadingRatio: multiplicador de leading (default 1.4)
 *   marginL/R: márgenes laterales en in (default 0.40)
 *   columns: "sym2" | "asym32" | "single" — tipo de grid de columnas
 *   footerH: altura footer en in (default 0.28)
 *   hasNotes: si reserva zona de notas al pie (default false)
 *   hasHeader: si reserva header navy arriba (default false)
 *   logoCV: si reserva espacio arriba derecha para logo CV (default true)
 */
function deriveGrid(opts = {}) {
  const bodySize     = opts.bodySize     ?? 11;
  const leadingRatio = opts.leadingRatio ?? 1.4;
  const ML           = opts.marginL      ?? 0.40;
  const MR           = opts.marginR      ?? 0.40;
  const columns      = opts.columns      ?? "asym32";
  const footerH       = opts.footerH      ?? 0.28;
  const hasNotes      = opts.hasNotes     ?? false;
  const hasHeader      = opts.hasHeader    ?? false;
  const headerH        = opts.headerH      ?? 0.28;
  // slideW/slideH son opcionales — por default usan las dimensiones base de
  // pptxgenjs (10x5.625, LAYOUT_16x9) que usa mb-generate. mb-format los
  // pasa explícitamente con el tamaño REAL del archivo que está leyendo,
  // que puede ser distinto (ej. 13.33x7.5, el widescreen estándar de
  // PowerPoint) — nunca asumir que todo archivo viene en el tamaño de
  // mb-generate.
  const SLIDE_W_LOCAL = opts.slideW ?? SLIDE_W;
  const SLIDE_H_LOCAL = opts.slideH ?? SLIDE_H;

  const MODULE = (bodySize * leadingRatio) / 72;  // leading en inches
  const GUTTER = MODULE * 2;                       // canaleta = 2 módulos (visible en pantalla)
  const WORK_W = SLIDE_W_LOCAL - ML - MR;

  // ── columnas ──────────────────────────────────────────────────────────────
  let cols;
  if (columns === "sym2") {
    const colW = (WORK_W - GUTTER) / 2;
    cols = {
      L: { x: ML, w: colW },
      R: { x: ML + colW + GUTTER, w: colW },
    };
  } else if (columns === "asym32") {
    const unit = (WORK_W - GUTTER) / 5;
    const clW  = unit * 3;
    const crW  = unit * 2;
    cols = {
      L: { x: ML, w: clW },
      R: { x: ML + clW + GUTTER, w: crW },
    };
  } else if (columns === "single") {
    cols = { L: { x: ML, w: WORK_W }, R: null };
  }

  // ── zonas verticales — de abajo hacia arriba ────────────────────────────────
  const FOOTER_Y = SLIDE_H_LOCAL - footerH;
  const GAP_FN   = MODULE * 1.5;          // respiro footer ↔ (notas o contenido)

  let NOTES_Y = null, NOTES_H = null, contBottom;
  if (hasNotes) {
    NOTES_H = 0.24;
    NOTES_Y = FOOTER_Y - GAP_FN - NOTES_H;
    const GAP_CN = MODULE * 1.0;
    contBottom = NOTES_Y - GAP_CN;
  } else {
    contBottom = FOOTER_Y - GAP_FN;
  }

  const headerH2 = hasHeader ? headerH : 0;
  const contTop  = (hasHeader ? headerH2 + 0.14 : 0.16);

  return {
    slideW: SLIDE_W_LOCAL, slideH: SLIDE_H_LOCAL,
    margins: { L: ML, R: MR },
    module: MODULE, gutter: GUTTER, workW: WORK_W,
    cols,
    zones: {
      headerH: headerH2, hasHeader,
      contentTop: contTop,
      contentBottom: contBottom,
      footerH, footerY: FOOTER_Y,
      notesY: NOTES_Y, notesH: NOTES_H, hasNotes,
    },
  };
}

/**
 * Calcula la cabecera (línea + título + subtítulo) y devuelve Y de inicio de bloques.
 * @param {object} grid - resultado de deriveGrid
 * @param {string} titleText - el texto REAL del título (obligatorio — se usa
 *   para estimar cuántas líneas ocupará y prevenir desborde)
 * @param {object} opts - { titleSize, gapAfterSubtitle }
 */
function deriveHeaderBlock(grid, titleText, opts = {}) {
  const M = grid.module;
  const requestedSize = opts.titleSize ?? 22;

  if (typeof titleText !== "string") {
    throw new Error("deriveHeaderBlock requiere titleText (string) como segundo argumento — es necesario para evitar que títulos largos se desborden de su caja.");
  }

  const titleWidthIn = grid.workW * 0.72;
  const { fontSize: titleSize, lines: titleLines } = fitTextLines(titleText, titleWidthIn, requestedSize, 2);
  const lineHeightIn = (titleSize / 72) * 1.25;

  const ruleY    = grid.zones.contentTop;
  const ruleH    = 0.018;
  const titleY   = ruleY + ruleH + M * 0.4;
  const titleH   = lineHeightIn * titleLines + 0.05;
  const subY     = titleY + titleH + M * 0.15;
  const subH     = M * 0.9;
  const blocksY  = subY + subH + M * (opts.gapAfterSubtitle ?? 0.9);
  return { ruleY, ruleH, titleY, titleH, subY, subH, blocksY, titleSize, titleLines };
}

/** Inserta línea dorada + título + subtítulo + (opcional) logo CV arriba derecha. */
function addHeader(slide, pres, grid, head, { title, subtitle, logoPath, logoH = 0.5 }) {
  const C = PALETTE_CV;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: grid.cols.L.x, y: head.ruleY, w: 0.38, h: head.ruleH,
    fill: { color: C.gold }, line: { color: C.gold },
  });
  slide.addText(title, {
    x: grid.cols.L.x, y: head.titleY, w: grid.workW * 0.72, h: head.titleH,
    fontSize: head.titleSize, color: C.navy, fontFace: FONT.serif,
    valign: "top", margin: 0, wrap: true,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: grid.cols.L.x, y: head.subY, w: grid.cols.L.w, h: head.subH,
      fontSize: 10, color: C.gold, fontFace: FONT.sans,
      valign: "top", margin: 0,
    });
  }
  if (logoPath && fs.existsSync(logoPath)) {
    slide.addImage({
      path: logoPath,
      x: grid.slideW - grid.margins.R - logoH,
      y: head.titleY + (head.titleH - logoH) / 2,
      w: logoH, h: logoH,
    });
  }
}

/** Inserta footer navy con 3 segmentos de texto. */
function addFooter(slide, pres, grid, { left, center, right }) {
  const C = PALETTE_CV;
  const { footerY, footerH } = grid.zones;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: footerY, w: grid.slideW, h: footerH,
    fill: { color: C.navy }, line: { color: C.navy },
  });
  // Segmento izquierdo: margen izq → 35% del slide
  const seg1W = grid.slideW * 0.35;
  const seg2X = grid.margins.L + seg1W;
  // Segmento central: 30% del slide, centrado
  const seg2W = grid.slideW * 0.30;
  const seg3X = seg2X + seg2W;
  const seg3W = grid.slideW - grid.margins.R - seg3X;

  slide.addText(left, {
    x: grid.margins.L, y: footerY, w: seg1W, h: footerH,
    fontSize: 5.5, bold: true, color: C.white, fontFace: FONT.sans,
    charSpacing: 0.8, valign: "middle", margin: 0,
  });
  slide.addText(center, {
    x: seg2X, y: footerY, w: seg2W, h: footerH,
    fontSize: 5.5, color: C.white, fontFace: FONT.sans,
    charSpacing: 0.5, valign: "middle", align: "center", margin: 0,
  });
  slide.addText(right, {
    x: seg3X, y: footerY, w: seg3W, h: footerH,
    fontSize: 5.5, bold: true, color: C.white, fontFace: FONT.sans,
    charSpacing: 0.8, valign: "middle", align: "right", margin: 0,
  });
}

/** Inserta notas al pie en su zona reservada (nunca toca el footer). */
function addNotes(slide, grid, notesArray) {
  if (!grid.zones.hasNotes) return;
  slide.addText(notesArray.join("\n"), {
    x: grid.margins.L, y: grid.zones.notesY, w: grid.workW * 0.75, h: grid.zones.notesH,
    fontSize: 5.5, color: PALETTE_CV.muted, fontFace: FONT.sans,
    valign: "top", margin: 0, lineSpacingMultiple: 1.2,
  });
}

/** Helper para construir celdas de tabla con estilo CV consistente. */
function tblCell(text, opts = {}) {
  const C = PALETTE_CV;
  return {
    text,
    options: {
      fontSize:  opts.fs    ?? 6.5,
      bold:      opts.bold  ?? false,
      underline: opts.ul    ? { style: "sng" } : undefined,
      color:     opts.color ?? C.text,
      fontFace:  FONT.sans,
      align:     opts.align ?? "left",
      valign:    "middle",
      fill:      opts.fill  ? { color: opts.fill } : undefined,
      margin:    [2, 5, 2, 5],
    }
  };
}

/** Infiere nombre corto de proyecto desde un título de slide. */
function inferProjectName(title) {
  if (!title) return "PROYECTO";
  // Busca patrón "... del/de/para X" y toma X
  const m = title.match(/(?:del?|para|de)\s+([A-ZÁÉÍÓÚÑ][\wÁÉÍÓÚÑáéíóúñ\s]+)$/);
  if (m) return m[1].trim().toUpperCase();
  // fallback: últimas 2-3 palabras capitalizadas
  const words = title.split(/\s+/);
  return words.slice(-2).join(" ").toUpperCase();
}

/**
 * Inserta una tarjeta de color con título + cuerpo, con protección de
 * desborde en el título (igual que deriveHeaderBlock/buildCover). El
 * título nunca puede invadir el área del cuerpo: se calcula su altura
 * real según el texto y se reduce la fuente si hace falta.
 *
 * @param {object} card - { x, y, w, h, color, title, body, titleSize, bodySize }
 */
function addCard(slide, pres, card) {
  const { x, y, w, h, color, title, body } = card;
  const requestedTitleSize = card.titleSize ?? 11;
  let bodySize = card.bodySize ?? 7.5;
  const M = card.module ?? ((11 * 1.4) / 72);

  // fondo de la tarjeta
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h, fill: { color }, line: { color },
  });

  const titlePadX = 0.1;
  const titleWidthIn = w - titlePadX * 2;
  const { fontSize: titleSize, lines: titleLines } = fitTextLines(title, titleWidthIn, requestedTitleSize, 2);
  const titleTopOffset = M * 0.6;
  const titleLineH = (titleSize / 72) * 1.3;
  const titleH = titleLineH * titleLines + M * 0.2;

  slide.addText(title, {
    x: x + titlePadX, y: y + titleTopOffset, w: titleWidthIn, h: titleH,
    fontSize: titleSize, bold: true, italic: true, color: PALETTE_CV.white,
    fontFace: FONT.sans, align: "center", valign: "top", margin: 0,
  });

  // ── PROTECCIÓN CONTRA DESBORDE DEL CUERPO ─────────────────────────────────
  // El cuerpo tiene una altura disponible FIJA (lo que queda de la tarjeta
  // tras el título). Si el texto necesita más líneas de las que esa altura
  // permite al tamaño solicitado, reducimos bodySize hasta que quepa, con un
  // piso mínimo de legibilidad. Si aun al piso no cabe, se acepta el
  // desborde mínimo necesario (mejor que truncar contenido silenciosamente).
  const bodyTopOffset = titleTopOffset + titleH + M * 0.3;
  const bodyPadX = 0.12;
  const bodyWidthIn = w - bodyPadX * 2;
  const bodyAvailH = Math.max(0.1, h - bodyTopOffset - 0.15);

  let bodyLineH = (bodySize / 72) * 1.3;
  let charsPerLine = (bodyWidthIn * 72) / (bodySize * 0.52);
  let bodyLines = Math.max(1, Math.ceil(body.length / charsPerLine));
  let neededH = bodyLineH * bodyLines;

  let attempts = 0;
  while (neededH > bodyAvailH && bodySize > 5.5 && attempts < 6) {
    const shrink = Math.max(0.85, bodyAvailH / neededH);
    bodySize = Math.max(5.5, bodySize * shrink);
    bodyLineH = (bodySize / 72) * 1.3;
    charsPerLine = (bodyWidthIn * 72) / (bodySize * 0.52);
    bodyLines = Math.max(1, Math.ceil(body.length / charsPerLine));
    neededH = bodyLineH * bodyLines;
    attempts++;
  }

  slide.addText(body, {
    x: x + bodyPadX, y: y + bodyTopOffset,
    w: bodyWidthIn, h: bodyAvailH,
    fontSize: bodySize, color: PALETTE_CV.white, fontFace: FONT.sans,
    valign: "top", margin: 0, lineSpacingMultiple: 1.3,
  });
}

/**
 * Dibuja un paso de flowchart (caja con label + body centrado) más la
 * flecha hacia el siguiente paso. La altura de la caja (h) debe venir
 * pre-calculada por layoutFlowchart — esta función solo dibuja, no decide
 * tamaños, para evitar tener dos lugares con lógica de cálculo de altura
 * que puedan desincronizarse.
 *
 * @param {object} opts
 *   x, y, w, h: posición y tamaño de la caja (h viene de layoutFlowchart)
 *   label, body: textos
 *   labelSize, bodySize: tamaños ya ajustados por layoutFlowchart
 *   isEdge: si es navy-filled (true) o outline (false)
 *   drawArrowBelow, arrowH: flecha hacia el siguiente paso
 */
function addFlowStep(slide, pres, opts) {
  const { x, y, w, h, label, body, isEdge, drawArrowBelow, arrowH, labelSize, bodySize } = opts;
  const C = PALETTE_CV;
  const padX = 0.15;
  const vPad = 0.08;

  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: isEdge ? C.navy : C.white },
    line: { color: C.navy, pt: isEdge ? 0 : 0.8 },
  });

  slide.addText([
    { text: label + "\n", options: { bold: true, fontSize: labelSize, breakLine: true } },
    { text: body, options: { fontSize: bodySize, bold: false } },
  ], {
    x: x + padX, y: y + vPad * 0.5, w: w - padX * 2, h: h - vPad,
    color: isEdge ? C.white : C.navy, fontFace: FONT.sans,
    align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.15,
  });

  if (drawArrowBelow) {
    const ay = y + h;
    const ax = x + w / 2;
    slide.addShape(pres.shapes.LINE, {
      x: ax, y: ay, w: 0, h: arrowH * 0.7,
      line: { color: C.navy, pt: 1.5 },
    });
    slide.addText("▼", {
      x: ax - 0.12, y: ay + arrowH * 0.6, w: 0.24, h: arrowH * 0.4,
      fontSize: 8, color: C.navy, fontFace: FONT.sans,
      align: "center", valign: "top", margin: 0,
    });
  }
}

/**
 * Calcula el layout completo de un flowchart vertical de N pasos, asegurando
 * que la suma de todas las cajas + flechas NUNCA exceda el espacio disponible.
 * Si el contenido es demasiado denso, reduce el tamaño de fuente de TODOS los
 * pasos proporcionalmente (consistencia visual) en vez de dejar que el último
 * paso se desborde fuera de la zona de contenido.
 *
 * @param {Array} steps - [{ label, body }]
 * @param {number} availableH - altura total disponible para el flowchart
 * @param {number} boxW - ancho de cada caja
 * @param {object} opts - { labelSize, bodySize, arrowH }
 * @returns {{ steps: Array<{label, body, h}>, labelSize, bodySize, arrowH }}
 */
function layoutFlowchart(stepsInput, availableH, boxW, opts = {}) {
  let labelSize = opts.labelSize ?? 8.5;
  let bodySize  = opts.bodySize  ?? 7;
  const arrowH  = opts.arrowH ?? ((11 * 1.4) / 72) * 0.75;
  const padX = 0.15;
  const textW = boxW - padX * 2;
  const vPad = 0.08;
  const gapLabelBody = 0.06;
  const minBoxH = 0.35; // altura mínima absoluta para que una caja siga siendo legible

  function computeHeights(lSize, bSize) {
    const labelLineH = (lSize / 72) * 1.25;
    const bodyLineH  = (bSize  / 72) * 1.25;
    return stepsInput.map((s) => {
      const { lines: labelLines } = fitTextLines(s.label, textW, lSize, 2);
      const charsPerLineBody = (textW * 72) / (bSize * 0.52);
      const bodyLines = Math.max(1, Math.ceil(s.body.length / charsPerLineBody));
      const h = vPad * 2 + labelLineH * labelLines + gapLabelBody + bodyLineH * bodyLines;
      return Math.max(h, minBoxH);
    });
  }

  const n = stepsInput.length;
  const totalArrows = (n - 1) * arrowH;

  let heights = computeHeights(labelSize, bodySize);
  let totalNeeded = heights.reduce((a, b) => a + b, 0) + totalArrows;

  // si no cabe, reducir fuente proporcionalmente (hasta un piso razonable) y recalcular
  let attempts = 0;
  while (totalNeeded > availableH && attempts < 10 && labelSize > 6 && bodySize > 5.5) {
    const shrinkFactor = Math.max(0.85, availableH / totalNeeded);
    labelSize = Math.max(6, labelSize * shrinkFactor);
    bodySize  = Math.max(5.5, bodySize * shrinkFactor);
    heights = computeHeights(labelSize, bodySize);
    totalNeeded = heights.reduce((a, b) => a + b, 0) + totalArrows;
    attempts++;
  }

  // Si AÚN no cabe tras reducir fuente al piso de legibilidad (5.5/6pt),
  // NO comprimimos las cajas por debajo de lo que el texto realmente
  // necesita — eso fue el bug original: producía cajas más pequeñas que
  // su contenido, causando que label y body se superpongan visualmente
  // (un problema invisible para el validador geométrico, que solo mira
  // las cajas declaradas, no el texto que se desborda de ellas).
  //
  // En su lugar, reportamos el overflow honestamente: el caller (el
  // skill / Claude) debe decidir — reducir el número de pasos, acortar
  // los textos, o dividir el flowchart en dos slides. No hay forma
  // segura de "inventar" espacio que no existe.
  const overflow = totalNeeded > availableH;
  const overflowAmount = overflow ? (totalNeeded - availableH) : 0;

  const resultSteps = stepsInput.map((s, i) => ({ ...s, h: heights[i] }));
  return {
    steps: resultSteps, labelSize, bodySize, arrowH,
    overflow, overflowAmount, totalNeeded, availableH,
  };
}

module.exports = {
  SLIDE_W, SLIDE_H, PALETTE_CV, FONT,
  deriveGrid, deriveHeaderBlock, addHeader, addFooter, addNotes,
  tblCell, inferProjectName, addCard, addFlowStep, layoutFlowchart,
};
