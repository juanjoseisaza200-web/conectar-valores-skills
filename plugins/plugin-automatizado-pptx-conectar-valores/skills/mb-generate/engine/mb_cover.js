/**
 * mb_cover.js — Template fijo de portada Conectar Valores
 * ============================================================
 * La portada NO se deriva del grid de contenido — es un template
 * de marca fijo que se aplica siempre igual (generate o format).
 *
 * Estructura:
 *   1. Header navy: "CONFIDENCIAL · USO EXCLUSIVO DEL DESTINATARIO" | "PREPARADO PARA  [CLIENTE]"
 *   2. Imagen hero full-width
 *   3. Línea dorada separadora
 *   4. Zona blanca: supratítulo dorado + línea vertical dorada + título + subtítulo + logo CV
 *   5. Footer: "[MES AÑO] · ESTRICTAMENTE PRIVADO Y CONFIDENCIAL · [CIUDAD] · [PAÍS]"
 */

const fs   = require("fs");
const path = require("path");
const E    = require("./mb_engine.js");
const { fitTextLines } = require("./mb_text_fit.js");

const C = E.PALETTE_CV;
const F = E.FONT;

/**
 * Construye la portada CV en el slide dado.
 * @param {object} opts
 *   company:    supratítulo, ej "AEROPUERTOS DEL PERÚ  ·  S.A."
 *   title:      título principal, ej "Pipeline de Financiamiento Estructurado"
 *   subtitle:   subtítulo italic, ej "Programa 2026 – 2027  ·  Tres tranches CRPAO para ADP"
 *   client:     destinatario, ej "MUFG"
 *   month:      ej "ABRIL 2026"
 *   city:       ej "LIMA"
 *   country:    ej "PERÚ"
 *   heroImagePath: ruta a imagen hero (si no existe, usa placeholder de color)
 *   logoCvPath: ruta al logo CV (default assets/logo_cv.png)
 */
function buildCover(slide, pres, opts) {
  const {
    company, title, subtitle, client,
    month, city, country,
    heroImagePath, logoCvPath,
  } = opts;

  const SLIDE_W = E.SLIDE_W;
  const SLIDE_H = E.SLIDE_H;
  const ML = 0.40, MR = 0.40;

  // ── ZONAS FIJAS DE LA PORTADA ───────────────────────────────────────────────
  const HEADER_H   = 0.30;
  const IMG_Y       = HEADER_H;
  const IMG_H       = 3.05;                  // imagen hero ocupa ~54% del slide
  const RULE_Y      = IMG_Y + IMG_H;
  const RULE_H      = 0.025;
  const WHITE_Y     = RULE_Y + RULE_H;
  const FOOTER_H    = 0.30;
  const FOOTER_Y    = SLIDE_H - FOOTER_H;
  const WHITE_H     = FOOTER_Y - WHITE_Y;

  slide.background = { color: C.white };

  // ── HEADER NAVY ──────────────────────────────────────────────────────────────
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: SLIDE_W, h: HEADER_H,
    fill: { color: C.navy }, line: { color: C.navy },
  });
  slide.addText("CONFIDENCIAL  ·  USO EXCLUSIVO DEL DESTINATARIO", {
    x: ML, y: 0, w: 5.5, h: HEADER_H,
    fontSize: 7, bold: true, color: C.white, fontFace: F.sans,
    charSpacing: 1.2, valign: "middle", margin: 0,
  });
  slide.addText(`PREPARADO PARA   ${client.toUpperCase()}`, {
    x: SLIDE_W - MR - 3.5, y: 0, w: 3.5, h: HEADER_H,
    fontSize: 7, bold: true, color: C.white, fontFace: F.sans,
    charSpacing: 1.2, valign: "middle", align: "right", margin: 0,
  });

  // ── IMAGEN HERO ──────────────────────────────────────────────────────────────
  if (heroImagePath && fs.existsSync(heroImagePath)) {
    slide.addImage({
      path: heroImagePath,
      x: 0, y: IMG_Y, w: SLIDE_W, h: IMG_H,
      sizing: { type: "cover", w: SLIDE_W, h: IMG_H },
    });
  } else {
    // placeholder: bloque sólido navy con texto indicativo
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: IMG_Y, w: SLIDE_W, h: IMG_H,
      fill: { color: C.navy }, line: { color: C.navy },
    });
    slide.addText("[ IMAGEN HERO — reemplazar con foto del sector ]", {
      x: 0, y: IMG_Y, w: SLIDE_W, h: IMG_H,
      fontSize: 10, color: C.goldLight, fontFace: F.sans,
      align: "center", valign: "middle", margin: 0,
    });
  }

  // ── LÍNEA DORADA SEPARADORA ──────────────────────────────────────────────────
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: RULE_Y, w: SLIDE_W, h: RULE_H,
    fill: { color: C.gold }, line: { color: C.gold },
  });

  // ── ZONA BLANCA: supratítulo + título + subtítulo + logo ────────────────────
  const M = (11 * 1.4) / 72;

  const supraY = WHITE_Y + WHITE_H * 0.12;
  slide.addText(company, {
    x: ML, y: supraY, w: SLIDE_W * 0.6, h: M * 1.0,
    fontSize: 9, bold: true, color: C.gold, fontFace: F.sans,
    charSpacing: 1.2, valign: "top", margin: 0,
  });

  // línea vertical dorada junto al título
  const titleY = supraY + M * 1.3;
  const requestedTitleSize = 24;
  const titleWidthIn = SLIDE_W * 0.62; // mismo ancho que usa addText para el título
  const { fontSize: titleFontSize, lines: titleLines } = fitTextLines(title, titleWidthIn, requestedTitleSize, 2);
  const titleH = (titleFontSize / 72) * 1.25 * titleLines + 0.08;

  slide.addShape(pres.shapes.RECTANGLE, {
    x: ML, y: titleY, w: 0.045, h: titleH * 0.92,
    fill: { color: C.gold }, line: { color: C.gold },
  });

  // título
  slide.addText(title, {
    x: ML + 0.18, y: titleY, w: SLIDE_W * 0.62, h: titleH,
    fontSize: titleFontSize, bold: true, color: C.navy, fontFace: F.sans,
    valign: "top", margin: 0,
  });

  // subtítulo — siempre debajo del título con respiro fijo
  const subY = titleY + titleH + M * 0.5;
  slide.addText(subtitle, {
    x: ML + 0.18, y: subY, w: SLIDE_W * 0.55, h: M * 1.2,
    fontSize: 10.5, italic: true, color: "4A5568", fontFace: F.sans,
    valign: "top", margin: 0,
  });

  // logo CV — grande, esquina derecha de la zona blanca
  if (logoCvPath && fs.existsSync(logoCvPath)) {
    const logoH = WHITE_H * 0.62;
    const logoW = logoH; // cuadrado
    slide.addImage({
      path: logoCvPath,
      x: SLIDE_W - MR - logoW - 0.1,
      y: WHITE_Y + (WHITE_H - logoH) / 2,
      w: logoW, h: logoH,
    });
  }

  // ── FOOTER ────────────────────────────────────────────────────────────────────
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: FOOTER_Y, w: SLIDE_W, h: 0.012,
    fill: { color: C.goldLight }, line: { color: C.goldLight },
  });
  slide.addText(`${month.toUpperCase()}  ·  ESTRICTAMENTE PRIVADO Y CONFIDENCIAL  ·  ${city.toUpperCase()}  ·  ${country.toUpperCase()}`, {
    x: 0, y: FOOTER_Y + 0.02, w: SLIDE_W, h: FOOTER_H - 0.02,
    fontSize: 7, bold: true, color: C.gold, fontFace: F.sans,
    charSpacing: 1, align: "center", valign: "middle", margin: 0,
  });
}

module.exports = { buildCover };
