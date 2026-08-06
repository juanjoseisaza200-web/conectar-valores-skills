const pptxgen = require("pptxgenjs");
const path    = require("path");
const fs      = require("fs");
const E       = require("../engine/mb_engine.js");
const { buildCover } = require("../engine/mb_cover.js");
const { validatePptx } = require("../engine/mb_validator.js");

const C = E.PALETTE_CV;
const F = E.FONT;
const OUT_DIR = path.join(__dirname, "stress_output");
const LOGO_CV = path.join(__dirname, "..", "assets", "logo_cv.png");
const HERO    = path.join(__dirname, "..", "assets", "hero_energia_real.jpg");

fs.mkdirSync(OUT_DIR, { recursive: true });

function bulletsOfCount(n, wordsEach = 20) {
  const sentence = "Este es un punto de análisis relevante que cubre aspectos operativos financieros y regulatorios del proyecto en cuestión durante el período evaluado completo.";
  const words = sentence.split(" ");
  const out = [];
  for (let i = 0; i < n; i++) out.push(words.slice(0, wordsEach).join(" ") + ` (punto ${i + 1})`);
  return out;
}

function tableRows(n) {
  const rows = [];
  for (let i = 1; i <= n; i++) rows.push([`Compañía con nombre largo número ${i}`, `${(Math.random() * 100000).toFixed(0)}`]);
  return rows;
}

const cases = [];

cases.push({
  name: "A1_titulo_largo_sym2",
  build: async (pres) => {
    const grid = E.deriveGrid({ columns: "sym2", hasNotes: false });
    const titleStr = "Análisis Estructurado Financiamiento Estratégico Regional Proyección Metodología Sostenible Operativo";
    const head = E.deriveHeaderBlock(grid, titleStr, { titleSize: 22 });
    const s = pres.addSlide();
    s.background = { color: C.white };
    E.addHeader(s, pres, grid, head, { title: titleStr, subtitle: "Subtítulo de prueba" });
    s.addText("Contenido de prueba.", {
      x: grid.cols.L.x, y: head.blocksY, w: grid.cols.L.w, h: 1,
      fontSize: 9, color: C.text, fontFace: F.sans, valign: "top", margin: 0,
    });
    E.addFooter(s, pres, grid, { left: "TEST", center: "TEST", right: "A1" });
  },
});

cases.push({
  name: "A2_bullets_excesivos",
  build: async (pres) => {
    const grid = E.deriveGrid({ columns: "asym32", hasNotes: false });
    const head = E.deriveHeaderBlock(grid, "Stress: 15 bullets", { titleSize: 20 });
    const s = pres.addSlide();
    s.background = { color: C.white };
    E.addHeader(s, pres, grid, head, { title: "Stress: 15 bullets", subtitle: "Caso límite" });
    const bullets = bulletsOfCount(15, 8);
    s.addText(
      bullets.map((b, i) => ({ text: b, options: { bullet: { indent: 8 }, breakLine: i < bullets.length - 1, paraSpaceAfter: 1 } })),
      { x: grid.cols.L.x, y: head.blocksY, w: grid.cols.L.w, h: grid.zones.contentBottom - head.blocksY,
        fontSize: 6, color: C.text, fontFace: F.sans, valign: "top", margin: 0, lineSpacingMultiple: 1.05 }
    );
    E.addFooter(s, pres, grid, { left: "TEST", center: "TEST", right: "A2" });
  },
});

cases.push({
  name: "A3_tabla_30_filas",
  build: async (pres) => {
    const grid = E.deriveGrid({ columns: "asym32", hasNotes: true });
    const head = E.deriveHeaderBlock(grid, "Stress: tabla de 30 filas", { titleSize: 20 });
    const s = pres.addSlide();
    s.background = { color: C.white };
    E.addHeader(s, pres, grid, head, { title: "Stress: tabla de 30 filas", subtitle: "Caso límite" });
    const rows = tableRows(30);
    const tblTitleH = grid.module * 0.85;
    const tblY = head.blocksY + tblTitleH + 0.04;
    const rowH = Math.max((grid.zones.contentBottom - tblY) / (rows.length + 1), 0.08);
    s.addText("Tabla con 30 filas", { x: grid.cols.R.x, y: head.blocksY, w: grid.cols.R.w, h: tblTitleH, fontSize: 7, bold: true, color: C.navy, fontFace: F.sans, valign: "top", margin: 0 });
    s.addTable([
      [E.tblCell("Compañía", { bold: true, color: C.white, fill: C.tblHeader }), E.tblCell("Valor", { bold: true, color: C.white, fill: C.tblHeader })],
      ...rows.map(([n, v]) => [E.tblCell(n, { fs: 5 }), E.tblCell(v, { fs: 5, align: "right" })]),
    ], { x: grid.cols.R.x, y: tblY, w: grid.cols.R.w, rowH, border: { pt: 0.3, color: C.tblBorder }, colW: [grid.cols.R.w * 0.65, grid.cols.R.w * 0.35] });
    E.addNotes(s, grid, ["Nota de prueba."]);
    E.addFooter(s, pres, grid, { left: "TEST", center: "TEST", right: "A3" });
  },
});

cases.push({
  name: "A4_sin_subtitulo",
  build: async (pres) => {
    const grid = E.deriveGrid({ columns: "single", hasNotes: false });
    const head = E.deriveHeaderBlock(grid, "Slide sin subtítulo", { titleSize: 22 });
    const s = pres.addSlide();
    s.background = { color: C.white };
    E.addHeader(s, pres, grid, head, { title: "Slide sin subtítulo", subtitle: null });
    s.addText("Contenido.", { x: grid.cols.L.x, y: head.blocksY, w: grid.cols.L.w, h: 1, fontSize: 10, color: C.text, fontFace: F.sans, valign: "top", margin: 0 });
    E.addFooter(s, pres, grid, { left: "TEST", center: "TEST", right: "A4" });
  },
});

cases.push({
  name: "B1_tarjetas_titulos_extremos",
  build: async (pres) => {
    const grid = E.deriveGrid({ columns: "single", hasNotes: false });
    const head = E.deriveHeaderBlock(grid, "Stress tarjetas: títulos extremos", { titleSize: 22, gapAfterSubtitle: 0.7 });
    const s = pres.addSlide();
    s.background = { color: C.white };
    E.addHeader(s, pres, grid, head, { title: "Stress tarjetas: títulos extremos", subtitle: "1 palabra vs título largo" });
    const cardsY = head.blocksY;
    const cardsH = grid.zones.contentBottom - cardsY;
    const gap = grid.gutter * 0.5;
    const cardW = (grid.workW - 2 * gap) / 3;
    const cards = [
      { color: C.cardGold, title: "X", body: "Cuerpo con título de una sola letra." },
      { color: C.cardNavy, title: "Generación Hidroeléctrica de Embalse Pasada y Bombeo Combinado Múltiple", body: "Cuerpo corto." },
      { color: C.cardGray, title: "Normal", body: "Cuerpo de prueba normal para la tercera tarjeta de este stress test." },
    ];
    cards.forEach((card, i) => {
      const cx = grid.margins.L + i * (cardW + gap);
      E.addCard(s, pres, { x: cx, y: cardsY, w: cardW, h: cardsH, color: card.color, title: card.title, body: card.body, module: grid.module });
    });
    E.addFooter(s, pres, grid, { left: "TEST", center: "TEST", right: "B1" });
  },
});

cases.push({
  name: "B2_tarjetas_cuerpo_vacio",
  build: async (pres) => {
    const grid = E.deriveGrid({ columns: "single", hasNotes: false });
    const head = E.deriveHeaderBlock(grid, "Stress tarjetas: cuerpo mínimo", { titleSize: 22, gapAfterSubtitle: 0.7 });
    const s = pres.addSlide();
    s.background = { color: C.white };
    E.addHeader(s, pres, grid, head, { title: "Stress tarjetas: cuerpo mínimo", subtitle: "Cuerpos muy cortos o casi vacíos" });
    const cardsY = head.blocksY;
    const cardsH = grid.zones.contentBottom - cardsY;
    const gap = grid.gutter * 0.5;
    const cardW = (grid.workW - 2 * gap) / 3;
    const cards = [
      { color: C.cardGold, title: "Card 1", body: "." },
      { color: C.cardNavy, title: "Card 2", body: "Ok." },
      { color: C.cardGray, title: "Card 3", body: "Cuerpo normal de prueba para esta tarjeta." },
    ];
    cards.forEach((card, i) => {
      const cx = grid.margins.L + i * (cardW + gap);
      E.addCard(s, pres, { x: cx, y: cardsY, w: cardW, h: cardsH, color: card.color, title: card.title, body: card.body, module: grid.module });
    });
    E.addFooter(s, pres, grid, { left: "TEST", center: "TEST", right: "B2" });
  },
});

cases.push({
  name: "B3_tarjetas_cuerpo_excesivo",
  build: async (pres) => {
    const grid = E.deriveGrid({ columns: "single", hasNotes: false });
    const head = E.deriveHeaderBlock(grid, "Stress tarjetas: cuerpo muy largo", { titleSize: 22, gapAfterSubtitle: 0.7 });
    const s = pres.addSlide();
    s.background = { color: C.white };
    E.addHeader(s, pres, grid, head, { title: "Stress tarjetas: cuerpo muy largo", subtitle: "Texto que podría desbordar verticalmente" });
    const cardsY = head.blocksY;
    const cardsH = grid.zones.contentBottom - cardsY;
    const gap = grid.gutter * 0.5;
    const cardW = (grid.workW - 2 * gap) / 3;
    const longBody = "Este cuerpo de tarjeta es deliberadamente muy largo, repitiendo contenido para forzar el caso límite donde el texto podría exceder la altura disponible de la tarjeta y salirse de su contenedor visual, lo cual debe evitarse incluso si esto significa que el texto se ve apretado. ".repeat(2);
    const cards = [
      { color: C.cardGold, title: "Card A", body: longBody },
      { color: C.cardNavy, title: "Card B", body: "Cuerpo corto normal." },
      { color: C.cardGray, title: "Card C", body: longBody },
    ];
    cards.forEach((card, i) => {
      const cx = grid.margins.L + i * (cardW + gap);
      E.addCard(s, pres, { x: cx, y: cardsY, w: cardW, h: cardsH, color: card.color, title: card.title, body: card.body, module: grid.module });
    });
    E.addFooter(s, pres, grid, { left: "TEST", center: "TEST", right: "B3" });
  },
});

cases.push({
  name: "C1_flowchart_6_pasos_uno_largo",
  build: async (pres) => {
    const grid = E.deriveGrid({ columns: "single", hasNotes: false });
    const head = E.deriveHeaderBlock(grid, "Stress flowchart: 6 pasos", { titleSize: 22, gapAfterSubtitle: 0.7 });
    const s = pres.addSlide();
    s.background = { color: C.white };
    E.addHeader(s, pres, grid, head, { title: "Stress flowchart: 6 pasos", subtitle: "Uno con cuerpo muy largo" });
    const steps = [
      { label: "Paso 1", body: "Cuerpo corto." },
      { label: "Paso 2 con nombre más largo de lo normal para este test", body: "Cuerpo extremadamente largo diseñado para forzar el caso límite de compresión proporcional automática del flowchart completo sin que nada se desborde fuera de su caja asignada." },
      { label: "Paso 3", body: "Cuerpo corto número tres." },
      { label: "Paso 4", body: "Cuerpo corto número cuatro con algo más de texto." },
      { label: "Paso 5", body: "Cuerpo corto cinco." },
      { label: "Paso 6 final", body: "Cuerpo final de cierre." },
    ];
    const boxW = grid.workW * 0.85;
    const boxX = grid.margins.L + (grid.workW - boxW) / 2;
    const availableH = grid.zones.contentBottom - head.blocksY;
    const layout = E.layoutFlowchart(steps, availableH, boxW);
    let cursorY = head.blocksY;
    layout.steps.forEach((step, i) => {
      const isEdge = i === 0 || i === layout.steps.length - 1;
      E.addFlowStep(s, pres, { x: boxX, y: cursorY, w: boxW, h: step.h, label: step.label, body: step.body, isEdge, labelSize: layout.labelSize, bodySize: layout.bodySize, drawArrowBelow: i < layout.steps.length - 1, arrowH: layout.arrowH });
      cursorY += step.h + (i < layout.steps.length - 1 ? layout.arrowH : 0);
    });
    E.addFooter(s, pres, grid, { left: "TEST", center: "TEST", right: "C1" });
  },
});

cases.push({
  name: "C2_flowchart_2_pasos_minimo",
  build: async (pres) => {
    const grid = E.deriveGrid({ columns: "single", hasNotes: false });
    const head = E.deriveHeaderBlock(grid, "Stress flowchart: 2 pasos", { titleSize: 22, gapAfterSubtitle: 0.7 });
    const s = pres.addSlide();
    s.background = { color: C.white };
    E.addHeader(s, pres, grid, head, { title: "Stress flowchart: 2 pasos", subtitle: "Caso mínimo de flowchart" });
    const steps = [
      { label: "Inicio", body: "Primer paso del proceso." },
      { label: "Fin", body: "Último paso del proceso." },
    ];
    const boxW = grid.workW * 0.85;
    const boxX = grid.margins.L + (grid.workW - boxW) / 2;
    const availableH = grid.zones.contentBottom - head.blocksY;
    const layout = E.layoutFlowchart(steps, availableH, boxW);
    let cursorY = head.blocksY;
    layout.steps.forEach((step, i) => {
      const isEdge = i === 0 || i === layout.steps.length - 1;
      E.addFlowStep(s, pres, { x: boxX, y: cursorY, w: boxW, h: step.h, label: step.label, body: step.body, isEdge, labelSize: layout.labelSize, bodySize: layout.bodySize, drawArrowBelow: i < layout.steps.length - 1, arrowH: layout.arrowH });
      cursorY += step.h + (i < layout.steps.length - 1 ? layout.arrowH : 0);
    });
    E.addFooter(s, pres, grid, { left: "TEST", center: "TEST", right: "C2" });
  },
});

cases.push({
  name: "C3_flowchart_9_pasos_overflow_esperado",
  build: async (pres) => {
    const grid = E.deriveGrid({ columns: "single", hasNotes: false });
    const head = E.deriveHeaderBlock(grid, "Stress flowchart: 9 pasos densos", { titleSize: 20, gapAfterSubtitle: 0.5 });
    const s = pres.addSlide();
    s.background = { color: C.white };
    E.addHeader(s, pres, grid, head, { title: "Stress flowchart: 9 pasos densos", subtitle: "Caso de overflow esperado y reportado" });
    const steps = [];
    for (let i = 1; i <= 9; i++) {
      steps.push({
        label: `Paso ${i} del proceso extendido`,
        body: `Descripción del paso número ${i} con contenido suficientemente largo para forzar que el motor deba comprimir el tamaño de fuente de todos los pasos proporcionalmente y así mantenerse dentro del espacio disponible del slide.`,
      });
    }
    const boxW = grid.workW * 0.85;
    const boxX = grid.margins.L + (grid.workW - boxW) / 2;
    const availableH = grid.zones.contentBottom - head.blocksY;
    const layout = E.layoutFlowchart(steps, availableH, boxW);

    // Este caso DEBE reportar overflow=true — 9 pasos con cuerpos largos
    // no caben en una sola slide incluso al piso de legibilidad. Si el
    // motor cambia y esto deja de reportar overflow, sería una señal de
    // que volvió el bug de compresión silenciosa.
    if (!layout.overflow) {
      throw new Error("Se esperaba overflow=true para este caso de 9 pasos densos — si esto falla, revisar si el motor volvió a comprimir silenciosamente.");
    }

    // con overflow detectado, el approach correcto es truncar a los pasos
    // que SÍ caben en vez de dibujar cajas rotas — aquí simulamos esa
    // decisión: tomamos solo los primeros N pasos que caben.
    let cursorY = head.blocksY;
    let stepsDrawn = 0;
    for (let i = 0; i < layout.steps.length; i++) {
      const step = layout.steps[i];
      const isLast = i === layout.steps.length - 1;
      const wouldExceed = (cursorY + step.h) > grid.zones.contentBottom;
      if (wouldExceed) break; // no dibujar cajas que no caben
      const isEdge = i === 0 || isLast;
      E.addFlowStep(s, pres, {
        x: boxX, y: cursorY, w: boxW, h: step.h,
        label: step.label, body: step.body,
        isEdge, labelSize: layout.labelSize, bodySize: layout.bodySize,
        drawArrowBelow: !isLast, arrowH: layout.arrowH,
      });
      cursorY += step.h + (!isLast ? layout.arrowH : 0);
      stepsDrawn++;
    }

    E.addFooter(s, pres, grid, { left: "TEST", center: `${stepsDrawn}/${steps.length} pasos mostrados`, right: "C3" });
  },
});

cases.push({
  name: "C4_flowchart_5_pasos_cuerpos_moderados",
  build: async (pres) => {
    const grid = E.deriveGrid({ columns: "single", hasNotes: false });
    const head = E.deriveHeaderBlock(grid, "Flowchart: 5 pasos moderados", { titleSize: 20, gapAfterSubtitle: 0.5 });
    const s = pres.addSlide();
    s.background = { color: C.white };
    E.addHeader(s, pres, grid, head, { title: "Flowchart: 5 pasos moderados", subtitle: "Caso que SÍ debe caber sin overflow" });
    const steps = [
      { label: "Inicio del proceso", body: "Descripción breve del primer paso." },
      { label: "Segundo paso", body: "Descripción breve del segundo paso del proceso." },
      { label: "Tercer paso intermedio", body: "Descripción del paso intermedio con algo más de detalle relevante." },
      { label: "Cuarto paso", body: "Descripción breve del cuarto paso." },
      { label: "Cierre del proceso", body: "Descripción final de cierre del proceso completo." },
    ];
    const boxW = grid.workW * 0.85;
    const boxX = grid.margins.L + (grid.workW - boxW) / 2;
    const availableH = grid.zones.contentBottom - head.blocksY;
    const layout = E.layoutFlowchart(steps, availableH, boxW);

    if (layout.overflow) {
      throw new Error(`Se esperaba que este caso moderado cupiera sin overflow, pero faltaron ${layout.overflowAmount.toFixed(3)}in — revisar layoutFlowchart.`);
    }

    let cursorY = head.blocksY;
    layout.steps.forEach((step, i) => {
      const isEdge = i === 0 || i === layout.steps.length - 1;
      E.addFlowStep(s, pres, {
        x: boxX, y: cursorY, w: boxW, h: step.h,
        label: step.label, body: step.body,
        isEdge, labelSize: layout.labelSize, bodySize: layout.bodySize,
        drawArrowBelow: i < layout.steps.length - 1, arrowH: layout.arrowH,
      });
      cursorY += step.h + (i < layout.steps.length - 1 ? layout.arrowH : 0);
    });
    E.addFooter(s, pres, grid, { left: "TEST", center: "TEST", right: "C4" });
  },
});

cases.push({
  name: "D1_portada_titulo_largo",
  build: async (pres) => {
    const s = pres.addSlide();
    buildCover(s, pres, {
      company: "ANÁLISIS SECTORIAL  ·  ENERGÍA",
      title: "Análisis Integral y Comprehensivo del Mercado Energético Latinoamericano Contemporáneo y sus Proyecciones",
      subtitle: "Subtítulo de prueba para portada con título extremo",
      client: "EQUIPO DE PRUEBA",
      month: "Junio 2026", city: "Medellín", country: "Colombia",
      heroImagePath: HERO, logoCvPath: LOGO_CV,
    });
  },
});

cases.push({
  name: "D2_portada_titulo_corto",
  build: async (pres) => {
    const s = pres.addSlide();
    buildCover(s, pres, {
      company: "TEST",
      title: "Energía",
      subtitle: "Subtítulo corto",
      client: "X",
      month: "Junio 2026", city: "Medellín", country: "Colombia",
      heroImagePath: HERO, logoCvPath: LOGO_CV,
    });
  },
});

cases.push({
  name: "D3_portada_sin_imagen",
  build: async (pres) => {
    const s = pres.addSlide();
    buildCover(s, pres, {
      company: "TEST SIN IMAGEN",
      title: "Portada con placeholder en lugar de imagen real",
      subtitle: "Verifica que el placeholder no rompe el layout",
      client: "EQUIPO",
      month: "Junio 2026", city: "Medellín", country: "Colombia",
      heroImagePath: null, logoCvPath: LOGO_CV,
    });
  },
});

cases.push({
  name: "D4_portada_cliente_largo",
  build: async (pres) => {
    const s = pres.addSlide();
    buildCover(s, pres, {
      company: "TEST",
      title: "Prueba de cliente con nombre muy largo en el header",
      subtitle: "Verifica el header navy con texto largo a la derecha",
      client: "Consorcio Internacional de Inversiones Financieras Asociadas S.A.",
      month: "Junio 2026", city: "Medellín", country: "Colombia",
      heroImagePath: HERO, logoCvPath: LOGO_CV,
    });
  },
});

async function runAll() {
  const results = [];

  for (const c of cases) {
    const pres = new pptxgen();
    pres.layout = "LAYOUT_16x9";
    try {
      await c.build(pres);
      const outPath = path.join(OUT_DIR, `${c.name}.pptx`);
      await pres.writeFile({ fileName: outPath });
      const issues = validatePptx(outPath, { default: { slideW: 10, slideH: 5.625, footerY: 5.345, notesY: null } });
      results.push({ name: c.name, path: outPath, issues });
    } catch (e) {
      results.push({ name: c.name, path: null, issues: [{ type: "BUILD_ERROR", detail: e.message }] });
    }
  }

  console.log("\n" + "═".repeat(74));
  console.log("  REPORTE COMPLETO DE STRESS TEST — /mb-generate (todos los layouts)");
  console.log("═".repeat(74) + "\n");

  let totalIssues = 0;
  for (const r of results) {
    const status = r.issues.length === 0 ? "✓ PASS" : `✗ FAIL (${r.issues.length})`;
    console.log(`  ${status.padEnd(12)} ${r.name}`);
    if (r.issues.length > 0) {
      r.issues.forEach(iss => console.log(`      - ${iss.type}: ${iss.detail || iss.shape}`));
      totalIssues += r.issues.length;
    }
  }

  console.log("\n" + "─".repeat(74));
  console.log(`  Total: ${results.length} casos | ${results.filter(r => r.issues.length === 0).length} pasaron | ${totalIssues} problemas totales`);
  console.log("─".repeat(74) + "\n");

  return results;
}

runAll().catch(e => { console.error(e); process.exit(1); });
