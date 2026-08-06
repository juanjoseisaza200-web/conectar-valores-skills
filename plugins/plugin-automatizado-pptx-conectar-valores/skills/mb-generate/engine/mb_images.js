/**
 * mb_images.js — Búsqueda y descarga de imágenes libres de derechos
 * =====================================================================
 * Fuente: Pexels API (gratis, sin atribución obligatoria pero
 * apreciada). Unsplash queda como fallback futuro si se agrega su key.
 *
 * La API key vive en config/api_keys.json — NUNCA hardcodeada en este
 * archivo. Si el skill se redistribuye, cada copia lleva su propio
 * config/api_keys.json (no se versiona si el proyecto usa Git).
 */

const https = require("https");
const fs    = require("fs");
const path  = require("path");
const sharp = require("sharp");

const CONFIG_PATH = path.join(__dirname, "..", "config", "api_keys.json");

function loadApiKeys() {
  if (!fs.existsSync(CONFIG_PATH)) {
    throw new Error(
      `No se encontró ${CONFIG_PATH}. Crea el archivo con: {"pexels": "TU_KEY", "unsplash": null}`
    );
  }
  return JSON.parse(fs.readFileSync(CONFIG_PATH, "utf-8"));
}

/** Búsqueda en Pexels. Devuelve array de candidatos {id, width, height, url, photographer}. */
function searchPexels(query, perPage = 5) {
  const keys = loadApiKeys();
  if (!keys.pexels) throw new Error("No hay API key de Pexels configurada.");

  return new Promise((resolve, reject) => {
    const options = {
      hostname: "api.pexels.com",
      path: `/v1/search?query=${encodeURIComponent(query)}&per_page=${perPage}&orientation=landscape`,
      headers: { Authorization: keys.pexels },
    };
    https.get(options, (res) => {
      let data = "";
      res.on("data", (chunk) => (data += chunk));
      res.on("end", () => {
        if (res.statusCode !== 200) {
          return reject(new Error(`Pexels API error ${res.statusCode}: ${data}`));
        }
        try {
          const json = JSON.parse(data);
          const results = (json.photos || []).map((p) => ({
            id: p.id,
            width: p.width,
            height: p.height,
            url: p.src.original,
            urlLarge: p.src.large2x || p.src.large,
            photographer: p.photographer,
            photographerUrl: p.photographer_url,
          }));
          resolve(results);
        } catch (e) {
          reject(e);
        }
      });
    }).on("error", reject);
  });
}

/** Descarga una URL de imagen a un path local. Sigue hasta 5 redirects. */
function downloadImage(url, destPath, _redirectCount = 0) {
  return new Promise((resolve, reject) => {
    if (_redirectCount > 5) {
      return reject(new Error("Too many redirects downloading image"));
    }
    const file = fs.createWriteStream(destPath);
    const req = https.get(url, (res) => {
      if (res.statusCode === 301 || res.statusCode === 302 || res.statusCode === 307 || res.statusCode === 308) {
        file.close(() => fs.unlink(destPath, () => {}));
        const redirectUrl = res.headers["location"];
        if (!redirectUrl) {
          return reject(new Error(`Redirect sin Location header (${res.statusCode})`));
        }
        downloadImage(redirectUrl, destPath, _redirectCount + 1).then(resolve).catch(reject);
        return;
      }
      if (res.statusCode !== 200) {
        file.close(() => fs.unlink(destPath, () => {}));
        return reject(new Error(`Download failed: ${res.statusCode}`));
      }
      res.pipe(file);
      file.on("finish", () => file.close(() => resolve(destPath)));
      file.on("error", (e) => {
        file.close(() => fs.unlink(destPath, () => {}));
        reject(e);
      });
    });
    req.on("error", (e) => {
      file.close(() => fs.unlink(destPath, () => {}));
      reject(e);
    });
  });
}

/**
 * Recorta/ajusta una imagen al aspect ratio requerido (cover, sin distorsión).
 * @param {string} srcPath
 * @param {string} destPath
 * @param {number} targetW - ancho objetivo en px (para el cálculo de aspect ratio)
 * @param {number} targetH - alto objetivo en px
 */
async function fitToAspect(srcPath, destPath, targetW, targetH) {
  await sharp(srcPath)
    .resize(targetW, targetH, { fit: "cover", position: "centre" })
    .jpeg({ quality: 88 })
    .toFile(destPath);
  return destPath;
}

/**
 * Flujo completo: busca, descarga la mejor opción, recorta al aspect ratio
 * del hero de portada (10in x ~2.6in → ratio ancho), y devuelve la ruta local.
 *
 * @param {string} query - término de búsqueda, ej "energía renovable latinoamérica"
 * @param {string} outDir - carpeta destino (típicamente assets/)
 * @param {object} opts
 *   aspectW, aspectH: relación de aspecto deseada (default 10:2.95, el hero de portada)
 *   filenameHint: nombre base del archivo, ej "hero_energia"
 */
async function fetchHeroImage(query, outDir, opts = {}) {
  const aspectW = opts.aspectW ?? 10;
  const aspectH = opts.aspectH ?? 2.95;
  const filenameHint = opts.filenameHint ?? "hero";

  const results = await searchPexels(query, 5);
  if (results.length === 0) {
    throw new Error(`Sin resultados en Pexels para: "${query}"`);
  }

  // elegir el primer resultado con buena resolución
  const best = results[0];

  const rawPath = path.join(outDir, `_raw_${filenameHint}.jpg`);
  await downloadImage(best.urlLarge, rawPath);

  // resolución objetivo: usar un múltiplo razonable para impresión nítida (2x)
  const targetW = 2000;
  const targetH = Math.round(targetW * (aspectH / aspectW));

  const finalPath = path.join(outDir, `${filenameHint}.jpg`);
  await fitToAspect(rawPath, finalPath, targetW, targetH);

  fs.unlinkSync(rawPath); // limpiar raw

  return {
    path: finalPath,
    photographer: best.photographer,
    photographerUrl: best.photographerUrl,
    sourceId: best.id,
  };
}

module.exports = { searchPexels, downloadImage, fitToAspect, fetchHeroImage };
