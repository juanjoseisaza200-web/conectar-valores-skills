/**
 * mb_reconstructor.js — Reconstruye una slide con identidad CV
 * a partir del JSON estructurado del extractor.
 *
 * Recibe: { slideW, slideH, elements[] } del extractor
 * Genera: una slide pptxgenjs con el grid MB y la paleta CV aplicada
 *
 * Este módulo maneja específicamente el template de "empresa subsidiaria"
 * (perfiles de subsidiaria), que es el primer caso real de mb-format.
 * Cuando se agreguen más familias de slides, cada una tendrá su propio
 * handler registrado en TEMPLATE_HANDLERS abajo.
 */

const path = require("path");
const E    = require("../../mb-generate/engine/mb_engine.js");

const C = E.PALETTE_CV;
const F = E.FONT;

const LOGO_CV = path.join(__dirname, "..", "..", "mb-generate", "assets", "logo_cv.png");

// ─── DETECCIÓN DE TEMPLATE ────────────────────────────────────────────────────

/**
 * Detecta qué template es esta slide basándose en sus elementos extraídos.
 * Devuelve un string de template o "unknown".
 */
function detectTemplate(extracted) {
  const roles = extracted.elements.map(e => e.role);
  const hasFintable  = roles.filter(r => r === "fin_table").length >= 1;
  const hasChart     = roles.includes("chart");
  const hasCompany   = roles.includes("company");
  const hasDesc      = roles.includes("description");
  const hasClosing   = roles.includes("closing_text");

  // Template "empresa subsidiaria": tiene nombre de empresa, descripción,
  // tabla(s) financiera(s), y texto de cierre. El chart es opcional
  // (slide 96 no tiene chart pero sigue siendo el mismo template).
  if (hasCompany && hasDesc && hasFintable && hasClosing) {
    return "subsidiary_profile";
  }

  // Más templates irán aquí a medida que se implementen
  return "unknown";
}

// ─── HELPERS DE CLASIFICACIÓN DE KPIs ────────────────────────────────────────

/**
 * Los KPIs en el archivo de referencia no están en un GROUP sino como shapes
 * sueltos con valores numéricos cortos. Los detectamos como elementos unknown
 * con font_size >= 12 y 1-2 palabras, agrupándolos por proximidad en Y.
 */
function extractKpisFromUnknowns(elements) {
  const candidates = elements.filter(e =>
    e.role === "unknown" &&
    e.meta.font_size >= 12 &&
    e.meta.words <= 2 &&
    /^[\d,.%x]+$/.test(e.text.trim())
  );
  const labels = elements.filter(e =>
    e.role === "unknown" &&
    e.meta.font_size <= 11 &&
    e.meta.words >= 1 && e.meta.words <= 4 &&
    !/^[\d,.%x]+$/.test(e.text.trim())
  );

  // emparejar cada valor con su label (label suele estar justo debajo del valor)
  const kpis = candidates.map(val => {
    const matchingLabel = labels.find(lbl =>
      Math.abs(lbl.x - val.x) < 0.5 &&
      lbl.y > val.y && lbl.y < val.y + 0.6
    );
    return {
      value: val.text.trim(),
      label: matchingLabel ? matchingLabel.text.trim().replace(/\n/g, " ") : "",
      x: val.x,
    };
  });

  // ordenar por posición X (izquierda a derecha)
  return kpis.sort((a, b) => a.x - b.x);
}

// ─── HANDLERS POR TEMPLATE ───────────────────────────────────────────────────

/**
 * Reconstruye una slide de perfil de empresa subsidiaria.
 * Layout: 2 columnas asym32, con el grid escalado al tamaño real del archivo.
 */
function reconstructSubsidiaryProfile(slide, pres, extracted) {
  const { slideW, slideH, elements } = extracted;

  const grid = E.deriveGrid({
    columns: "asym32",
    hasNotes: false,
    slideW,
    slideH,
  });

  // Extraer elementos por rol
  const title      = elements.find(e => e.role === "title");
  const subtitle   = elements.find(e => e.role === "subtitle");
  const company    = elements.find(e => e.role === "company");
  const desc       = elements.find(e => e.role === "description");
  const tables     = elements.filter(e => e.role === "fin_table");
  const chart      = elements.find(e => e.role === "chart");
  const closingTxts = elements.filter(e => e.role === "closing_text");
  const sectionHeaders = elements.filter(e => e.role === "section_header");
  const kpis       = extractKpisFromUnknowns(elements);

  // ── HEADER ────────────────────────────────────────────────────────────────
  const titleStr = title ? title.text : "Sin título";
  const head = E.deriveHeaderBlock(grid, titleStr, { titleSize: 26, gapAfterSubtitle: 0.3 });

  slide.background = { color: C.white };
  E.addHeader(slide, pres, grid, head, {
    title: titleStr,
    subtitle: subtitle ? subtitle.text : null,
    logoPath: LOGO_CV,
    logoH: 0.6,
  });

  // ── COLUMNA IZQUIERDA ─────────────────────────────────────────────────────
  let cursorY = head.blocksY;
  const lx = grid.cols.L.x;
  const lw = grid.cols.L.w;

  // nombre de empresa
  if (company) {
    slide.addText(company.text, {
      x: lx, y: cursorY, w: lw, h: 0.35,
      fontSize: 14, bold: true, underline: true,
      color: C.navy, fontFace: F.sans, valign: "top", margin: 0,
    });
    cursorY += 0.42;
  }

  // descripción
  if (desc) {
    const descH = 1.85;
    slide.addText(desc.text, {
      x: lx, y: cursorY, w: lw, h: descH,
      fontSize: 10, color: C.text, fontFace: F.sans,
      valign: "top", margin: 0, lineSpacingMultiple: 1.2,
    });
    cursorY += descH + 0.15;
  }

  // sección "Valoración"
  const valHeader = sectionHeaders.find(s => s.text.toLowerCase().includes("valorac"));
  if (valHeader) {
    slide.addText(valHeader.text, {
      x: lx, y: cursorY, w: lw, h: 0.3,
      fontSize: 13, bold: true, color: C.navy, fontFace: F.sans,
      valign: "top", margin: 0,
    });
    cursorY += 0.35;
  }

  // placeholder del chart (el chart original se preserva como imagen en el
  // archivo fuente; en esta versión de mb-format lo marcamos como pendiente
  // de recoloración — la recoloración real del chart nativo es trabajo
  // adicional que se implementará en la siguiente iteración)
  const chartH = Math.min(chart ? chart.h : 1.8, grid.zones.contentBottom - cursorY - 1.0);
  slide.addShape(pres.shapes.RECTANGLE, {
    x: lx, y: cursorY, w: lw, h: chartH,
    fill: { color: "F2F4F7" }, line: { color: C.tblBorder, pt: 0.5 },
  });
  slide.addText("[ waterfall chart — recolorear a paleta CV ]", {
    x: lx, y: cursorY, w: lw, h: chartH,
    fontSize: 9, italic: true, color: C.muted, fontFace: F.sans,
    align: "center", valign: "middle", margin: 0,
  });
  cursorY += chartH + 0.15;

  // texto de cierre izquierda
  const closingLeft = closingTxts.find(t => t.x < slideW / 2);
  if (closingLeft && cursorY < grid.zones.contentBottom) {
    const closingH = grid.zones.contentBottom - cursorY;
    slide.addText(closingLeft.text, {
      x: lx, y: cursorY, w: lw, h: closingH,
      fontSize: 9.5, color: C.text, fontFace: F.sans,
      valign: "top", margin: 0, lineSpacingMultiple: 1.15,
    });
  }

  // ── COLUMNA DERECHA ───────────────────────────────────────────────────────
  const rx  = grid.cols.R.x;
  const rw  = grid.cols.R.w;
  let rcursorY = head.blocksY;

  // sección "Información financiera"
  const infoHeader = sectionHeaders.find(s => s.text.toLowerCase().includes("informac"));
  if (infoHeader) {
    slide.addText(infoHeader.text, {
      x: rx, y: rcursorY, w: rw, h: 0.3,
      fontSize: 13, bold: true, color: C.navy, fontFace: F.sans,
      valign: "top", margin: 0,
    });
    rcursorY += 0.38;
  }

  // tablas financieras (hasta 2, lado a lado)
  let tblEndY = rcursorY;
  if (tables.length > 0) {
    const gap = 0.15;
    const tblW = tables.length >= 2
      ? (rw - gap) / 2
      : rw;

    tables.slice(0, 2).forEach((tbl, i) => {
      const tx = rx + i * (tblW + gap);
      const rows = tbl.meta.rows;
      if (!rows || rows.length === 0) return;
      const rowH = 0.215;
      const tblH = rows.length * rowH;
      const headerRow = rows[0];
      const dataRows  = rows.slice(1);
      const nCols = headerRow.length;
      const colW  = headerRow.map((_, j) => j === 0 ? tblW * 0.5 : tblW * (0.5 / (nCols - 1)));

      slide.addTable(
        [
          headerRow.map(cell => E.tblCell(cell, { fs: 8, bold: true, color: C.white, fill: C.tblHeader, align: "center" })),
          ...dataRows.map((row, ri) =>
            row.map((cell, ci) => E.tblCell(cell, {
              fs: 8,
              color: C.text,
              fill: ri % 2 === 0 ? C.tblAlt : C.white,
              align: ci === 0 ? "left" : "right",
            }))
          ),
        ],
        { x: tx, y: rcursorY, w: tblW, rowH, border: { pt: 0.3, color: C.tblBorder }, colW }
      );

      const thisTableBottom = rcursorY + tblH;
      if (thisTableBottom > tblEndY) tblEndY = thisTableBottom;
    });

    rcursorY = tblEndY + 0.2;
  }

  // tarjetas KPI — Versión B (outline minimalista, elegida por el usuario)
  if (kpis.length > 0) {
    const kpiGap = 0.1;
    const kpiW   = (rw - kpiGap * (kpis.length - 1)) / kpis.length;
    const kpiH   = 0.85;

    kpis.forEach((kpi, i) => {
      const kx = rx + i * (kpiW + kpiGap);
      // fondo claro + borde dorado (Versión B)
      slide.addShape(pres.shapes.RECTANGLE, {
        x: kx, y: rcursorY, w: kpiW, h: kpiH,
        fill: { color: "F7F8FA" }, line: { color: C.gold, width: 1 },
      });
      slide.addText(kpi.value, {
        x: kx, y: rcursorY + 0.08, w: kpiW, h: 0.35,
        fontSize: 13, bold: true, color: C.navy, fontFace: F.sans,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addText(kpi.label, {
        x: kx + 0.04, y: rcursorY + 0.44, w: kpiW - 0.08, h: 0.38,
        fontSize: 6.5, color: C.muted, fontFace: F.sans,
        align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.05,
      });
    });
    rcursorY += kpiH + 0.2;
  }

  // sección "Aspectos relevantes"
  const aspectHeader = sectionHeaders.find(s => s.text.toLowerCase().includes("aspectos"));
  if (aspectHeader) {
    slide.addText(aspectHeader.text, {
      x: rx, y: rcursorY, w: rw, h: 0.3,
      fontSize: 13, bold: true, color: C.navy, fontFace: F.sans,
      valign: "top", margin: 0,
    });
    rcursorY += 0.35;
  }

  // texto de aspectos (closing_text de la columna derecha)
  const closingRight = closingTxts.find(t => t.x >= slideW / 2);
  if (closingRight && rcursorY < grid.zones.contentBottom) {
    const bulletLines = closingRight.text
      .split("\n")
      .map(l => l.trim())
      .filter(l => l.length > 0);

    slide.addText(
      bulletLines.map((line, i) => ({
        text: line,
        options: { bullet: { indent: 10 }, breakLine: i < bulletLines.length - 1, paraSpaceAfter: 4 },
      })),
      {
        x: rx, y: rcursorY, w: rw, h: grid.zones.contentBottom - rcursorY,
        fontSize: 9, color: C.text, fontFace: F.sans,
        valign: "top", margin: 0, lineSpacingMultiple: 1.15,
      }
    );
  }

  // ── FOOTER ────────────────────────────────────────────────────────────────
  const companyShort = company ? company.text.replace(/^\d+\.\s*/, "").toUpperCase() : "EMPRESA";
  E.addFooter(slide, pres, grid, {
    left: "ESTRICTAMENTE PRIVADO Y CONFIDENCIAL",
    center: "VALORACIÓN GRUPO EJEMPLO",
    right: `${companyShort}`,
  });
}

// ─── ENTRY POINT ─────────────────────────────────────────────────────────────

/**
 * Reconstruye una slide a partir del JSON del extractor.
 * @param {object} slide   - slide de pptxgenjs ya añadida a la presentación
 * @param {object} pres    - presentación pptxgenjs
 * @param {object} extracted - JSON del extractor
 * @returns {{ template: string, ok: boolean, message: string }}
 */
function reconstructSlide(slide, pres, extracted) {
  const template = detectTemplate(extracted);

  if (template === "subsidiary_profile") {
    reconstructSubsidiaryProfile(slide, pres, extracted);
    return { template, ok: true, message: "OK" };
  }

  // template no reconocido — insertar placeholder con diagnóstico
  const roles = [...new Set(extracted.elements.map(e => e.role))].join(", ");
  slide.addText(
    `[ mb-format: template no reconocido ]\nRoles detectados: ${roles}\nFragmentación: ${extracted.fragRatio}`,
    {
      x: 0.5, y: 0.5, w: extracted.slideW - 1, h: 2,
      fontSize: 11, color: "CC0000", fontFace: "Arial",
      valign: "top", margin: 0,
    }
  );
  return { template, ok: false, message: `Template desconocido. Roles: ${roles}` };
}

module.exports = { reconstructSlide, detectTemplate };
