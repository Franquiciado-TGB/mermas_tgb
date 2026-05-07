/**
 * TGB STOCK CONTROL - Backend Apps Script v3.0
 * Nuevas funciones: OCR albaranes (Gemini Vision), subida PDF ventas, cálculo de mermas
 */

// ── Nombres de hojas ──────────────────────────────────────────────────────────
const SHEET_CONTEOS          = "Conteos";
const SHEET_ALBARANES        = "Albaranes";
const SHEET_ALBARANES_DETALLE = "Albaranes_Detalle";
const SHEET_VENTAS           = "Ventas";
const SHEET_MERMAS           = "Mermas";
const SHEET_LOG              = "Log";

// ── Drive ─────────────────────────────────────────────────────────────────────
const DRIVE_FOLDER_ALBARANES = "TGB Albaranes";
const DRIVE_FOLDER_VENTAS    = "TGB Ventas";

// ── Gemini ────────────────────────────────────────────────────────────────────
// Usa el modelo gemini-1.5-flash a través de la API de IA generativa de Google.
// La clave se obtiene en https://aistudio.google.com/app/apikey (gratis).
// Pégala aquí o guárdala en Archivo > Propiedades del proyecto (más seguro).
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY") || "";

// =============================================================================
// ENTRADA PRINCIPAL
// =============================================================================
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    log_("recibido", payload.type + " · " + (payload.venue || "?"));
    let result;
    if      (payload.type === "conteo")   result = handleConteo(payload);
    else if (payload.type === "albaran")  result = handleAlbaran(payload);
    else if (payload.type === "ventas")   result = handleVentas(payload);
    else throw new Error("Tipo desconocido: " + payload.type);
    return jsonResponse({ ok: true, result: result });
  } catch (err) {
    log_("error", String(err && err.message ? err.message : err));
    return jsonResponse({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function doGet() {
  return jsonResponse({ ok: true, message: "TGB Stock backend activo", version: "3.0", timestamp: new Date().toISOString() });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// =============================================================================
// HANDLER: CONTEO
// =============================================================================
function handleConteo(p) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureSheet_(ss, SHEET_CONTEOS, [
    "Fecha envío","Local","Semana inicio","Semana fin","Responsable",
    "Producto","Categoría","Unidad","Cantidad","ID producto","Versión app"
  ]);
  const ts   = new Date(p.timestamp || new Date());
  const rows = [];
  for (const it of p.items) {
    rows.push([
      ts, p.venue, p.weekStart, p.weekEnd, p.responsable || "",
      it.nombre, it.categoria, it.und,
      (it.cantidad === null || it.cantidad === undefined) ? "" : it.cantidad,
      it.id, p.appVersion || ""
    ]);
  }
  if (rows.length > 0)
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

  // Recalcular mermas tras nuevo conteo
  try { calcularMermas_(); } catch(e) { log_("warn", "mermas: " + e.message); }

  return { rows: rows.length, venue: p.venue, week: p.weekStart };
}

// =============================================================================
// HANDLER: ALBARÁN (foto → OCR → detalle)
// =============================================================================
function handleAlbaran(p) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = ensureSheet_(ss, SHEET_ALBARANES, [
    "Fecha envío","Local","Semana inicio","Semana fin",
    "Archivo (Drive)","URL Drive","Tamaño KB","Estado OCR","Nº líneas OCR","Versión app"
  ]);
  const folder = ensureFolder_(DRIVE_FOLDER_ALBARANES);
  const ts     = new Date(p.timestamp || new Date());
  const tsStr  = Utilities.formatDate(ts, "Europe/Madrid", "yyyyMMdd_HHmmss");
  const fname  = "albaran_" + p.venue + "_" + tsStr + ".jpg";
  const base64 = (p.photoBase64 || "").replace(/^data:image\/[a-zA-Z]+;base64,/, "");
  const blob   = Utilities.newBlob(Utilities.base64Decode(base64), "image/jpeg", fname);
  const file   = folder.createFile(blob);

  // ── OCR con Gemini Vision ──────────────────────────────────────────────────
  let ocrStatus = "Sin clave Gemini";
  let ocrLines  = 0;
  let ocrItems  = [];

  if (GEMINI_API_KEY) {
    try {
      ocrItems  = ocrAlbaran_(base64, p.venue, p.weekStart);
      ocrLines  = ocrItems.length;
      ocrStatus = ocrLines > 0 ? "OK (" + ocrLines + " líneas)" : "Sin líneas detectadas";

      // Guardar detalle en hoja Albaranes_Detalle
      if (ocrItems.length > 0) {
        const detalle = ensureSheet_(ss, SHEET_ALBARANES_DETALLE, [
          "Fecha","Local","Semana inicio","Archivo","Código artículo","Producto","Cantidad","P.U.","Total"
        ]);
        const dRows = ocrItems.map(it => [
          ts, p.venue, p.weekStart, fname,
          it.codigo || "", it.producto || "", it.cantidad || 0,
          it.pu || 0, it.total || 0
        ]);
        detalle.getRange(detalle.getLastRow() + 1, 1, dRows.length, dRows[0].length).setValues(dRows);
      }

      // Recalcular mermas con nuevas entradas
      try { calcularMermas_(); } catch(e) { log_("warn", "mermas: " + e.message); }

    } catch (ocrErr) {
      ocrStatus = "Error OCR: " + ocrErr.message;
      log_("error", "OCR: " + ocrErr.message);
    }
  }

  sheet.appendRow([ts, p.venue, p.weekStart, p.weekEnd, fname, file.getUrl(), p.sizeKB || "", ocrStatus, ocrLines, p.appVersion || ""]);
  return { fileName: fname, url: file.getUrl(), venue: p.venue, ocrLines: ocrLines, ocrStatus: ocrStatus, items: ocrItems };
}

// =============================================================================
// HANDLER: VENTAS (PDF de TPV)
// =============================================================================
function handleVentas(p) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = ensureSheet_(ss, SHEET_VENTAS, [
    "Fecha envío","Local","Semana inicio","Semana fin",
    "Archivo (Drive)","URL Drive","Tamaño KB","Versión app"
  ]);
  const folder = ensureFolder_(DRIVE_FOLDER_VENTAS);
  const ts     = new Date(p.timestamp || new Date());
  const tsStr  = Utilities.formatDate(ts, "Europe/Madrid", "yyyyMMdd_HHmmss");
  const fname  = "ventas_" + (p.venue || "ADMIN") + "_" + tsStr + ".pdf";
  const base64 = (p.pdfBase64 || "").replace(/^data:application\/pdf;base64,/, "");
  const blob   = Utilities.newBlob(Utilities.base64Decode(base64), "application/pdf", fname);
  const file   = folder.createFile(blob);
  sheet.appendRow([ts, p.venue || "ADMIN", p.weekStart, p.weekEnd, fname, file.getUrl(), p.sizeKB || "", p.appVersion || ""]);
  return { fileName: fname, url: file.getUrl() };
}

// =============================================================================
// OCR: GEMINI VISION
// =============================================================================
function ocrAlbaran_(base64, venue, weekStart) {
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" + GEMINI_API_KEY;

  const prompt = `Eres un asistente de extracción de datos. Esta es una foto de un albarán de entrega de productos de alimentación.
Extrae TODAS las líneas de producto de la tabla del albarán.
Para cada línea devuelve un objeto JSON con estos campos exactos:
- "codigo": código de artículo (ej: ERE100002427)
- "producto": nombre del producto en texto
- "cantidad": cantidad numérica (número entero o decimal, sin unidad)
- "pu": precio unitario como número (sin símbolo €)
- "total": importe total de la línea como número (sin símbolo €)

Devuelve SOLO un array JSON válido, sin texto adicional, sin markdown, sin bloques de código.
Si no puedes leer algún campo, ponlo como null.
Ejemplo de salida: [{"codigo":"ERE100002427","producto":"HAMBURGUESA VACUNO 100G","cantidad":15,"pu":101.59,"total":1523.82}]`;

  const body = {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: "image/jpeg", data: base64 } }
      ]
    }],
    generationConfig: { temperature: 0, maxOutputTokens: 4096 }
  };

  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) {
    throw new Error("Gemini HTTP " + resp.getResponseCode() + ": " + resp.getContentText().substring(0, 200));
  }

  const data = JSON.parse(resp.getContentText());
  const text = data.candidates?.[0]?.content?.parts?.[0]?.text || "";

  // Limpiar markdown si Gemini lo envuelve
  const clean = text.replace(/```json\n?/g, "").replace(/```\n?/g, "").trim();
  const items = JSON.parse(clean);
  if (!Array.isArray(items)) throw new Error("OCR no devolvió array");
  return items;
}

// =============================================================================
// CÁLCULO DE MERMAS
// Merma = Stock_anterior + Entradas_albarán - Conteo_actual
// Se recalcula cada vez que llega un conteo o albarán nuevo.
// =============================================================================
function calcularMermas_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shConteos  = ss.getSheetByName(SHEET_CONTEOS);
  const shDetalle  = ss.getSheetByName(SHEET_ALBARANES_DETALLE);
  const shMermas   = ensureSheet_(ss, SHEET_MERMAS, [
    "Semana inicio","Local","Producto","Stock anterior","Entradas albarán","Conteo final","Merma","% Merma"
  ]);

  if (!shConteos) return;

  // Leer conteos (saltar cabecera)
  const contData = shConteos.getDataRange().getValues();
  if (contData.length < 2) return;
  const contHeaders = contData[0]; // Fecha envío,Local,Semana inicio,...,Producto,Categoría,Unidad,Cantidad,...
  const iLocal      = contHeaders.indexOf("Local");
  const iWeekStart  = contHeaders.indexOf("Semana inicio");
  const iProducto   = contHeaders.indexOf("Producto");
  const iCantidad   = contHeaders.indexOf("Cantidad");

  // Agrupar conteos por semana+local+producto
  // Tomamos el ÚLTIMO conteo de cada combinación (más reciente)
  const conteoMap = {}; // key: "semana|local|producto" -> cantidad
  for (let r = 1; r < contData.length; r++) {
    const row    = contData[r];
    const local  = row[iLocal];
    const semana = row[iWeekStart] instanceof Date
      ? Utilities.formatDate(row[iWeekStart], "Europe/Madrid", "yyyy-MM-dd")
      : String(row[iWeekStart]);
    const prod   = row[iProducto];
    const cant   = Number(row[iCantidad]) || 0;
    const key    = semana + "|" + local + "|" + prod;
    conteoMap[key] = { semana, local, prod, cant, fecha: row[0] };
  }

  // Leer entradas de albarán por semana+local+producto
  const entradasMap = {}; // key -> cantidad total entrada
  if (shDetalle) {
    const detData    = shDetalle.getDataRange().getValues();
    const detHeaders = detData[0];
    const dLocal     = detHeaders.indexOf("Local");
    const dSemana    = detHeaders.indexOf("Semana inicio");
    const dProd      = detHeaders.indexOf("Producto");
    const dCant      = detHeaders.indexOf("Cantidad");
    for (let r = 1; r < detData.length; r++) {
      const row    = detData[r];
      const local  = row[dLocal];
      const semana = row[dSemana] instanceof Date
        ? Utilities.formatDate(row[dSemana], "Europe/Madrid", "yyyy-MM-dd")
        : String(row[dSemana]);
      const prod   = row[dProd];
      const cant   = Number(row[dCant]) || 0;
      const key    = semana + "|" + local + "|" + prod;
      entradasMap[key] = (entradasMap[key] || 0) + cant;
    }
  }

  // Obtener el conteo ANTERIOR por local+producto (semana previa)
  // Ordenar semanas
  const semanas = [...new Set(Object.values(conteoMap).map(v => v.semana))].sort();

  const mermaRows = [];
  for (const key of Object.keys(conteoMap)) {
    const { semana, local, prod, cant: conteoFinal } = conteoMap[key];
    const idxSemana = semanas.indexOf(semana);

    // Stock anterior = conteo de la semana previa para este local+producto
    let stockAnterior = 0;
    if (idxSemana > 0) {
      const semAnt  = semanas[idxSemana - 1];
      const keyAnt  = semAnt + "|" + local + "|" + prod;
      stockAnterior = conteoMap[keyAnt] ? conteoMap[keyAnt].cant : 0;
    }

    const entradas = entradasMap[key] || 0;
    const merma    = stockAnterior + entradas - conteoFinal;
    const pctMerma = (stockAnterior + entradas) > 0
      ? Math.round((merma / (stockAnterior + entradas)) * 10000) / 100
      : 0;

    mermaRows.push([semana, local, prod, stockAnterior, entradas, conteoFinal, merma, pctMerma]);
  }

  // Reescribir hoja Mermas desde fila 2
  const existingRows = shMermas.getLastRow();
  if (existingRows > 1) {
    shMermas.getRange(2, 1, existingRows - 1, 8).clearContent();
  }
  if (mermaRows.length > 0) {
    shMermas.getRange(2, 1, mermaRows.length, 8).setValues(mermaRows);
  }
}

// =============================================================================
// HELPERS
// =============================================================================
function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold").setBackground("#000000").setFontColor("#bef264");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function ensureFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

function log_(level, message) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ensureSheet_(ss, SHEET_LOG, ["Fecha","Nivel","Mensaje"]);
    sheet.appendRow([new Date(), level, message]);
  } catch(e) {}
}

// =============================================================================
// SETUP INICIAL — ejecutar una vez para crear todas las hojas
// =============================================================================
function testSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEET_CONTEOS, [
    "Fecha envío","Local","Semana inicio","Semana fin","Responsable",
    "Producto","Categoría","Unidad","Cantidad","ID producto","Versión app"
  ]);
  ensureSheet_(ss, SHEET_ALBARANES, [
    "Fecha envío","Local","Semana inicio","Semana fin",
    "Archivo (Drive)","URL Drive","Tamaño KB","Estado OCR","Nº líneas OCR","Versión app"
  ]);
  ensureSheet_(ss, SHEET_ALBARANES_DETALLE, [
    "Fecha","Local","Semana inicio","Archivo","Código artículo","Producto","Cantidad","P.U.","Total"
  ]);
  ensureSheet_(ss, SHEET_VENTAS, [
    "Fecha envío","Local","Semana inicio","Semana fin",
    "Archivo (Drive)","URL Drive","Tamaño KB","Versión app"
  ]);
  ensureSheet_(ss, SHEET_MERMAS, [
    "Semana inicio","Local","Producto","Stock anterior","Entradas albarán","Conteo final","Merma","% Merma"
  ]);
  ensureSheet_(ss, SHEET_LOG, ["Fecha","Nivel","Mensaje"]);
  ensureFolder_(DRIVE_FOLDER_ALBARANES);
  ensureFolder_(DRIVE_FOLDER_VENTAS);
  Logger.log("✓ Setup v3.0 completado: hojas y carpetas creadas");
}

// =============================================================================
// CONFIGURAR CLAVE GEMINI (ejecutar una vez en el editor)
// =============================================================================
function setGeminiKey() {
  // Cambia el valor por tu clave real de Google AI Studio
  PropertiesService.getScriptProperties().setProperty("GEMINI_API_KEY", "PEGA_AQUI_TU_CLAVE_GEMINI");
  Logger.log("Clave Gemini guardada");
}
