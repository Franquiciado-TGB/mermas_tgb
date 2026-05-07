/**
 * TGB STOCK CONTROL - Backend Apps Script v4.0
 * OCR albaranes: Drive nativo (gratis, sin API keys)
 * Flujo: subir PDF/foto → OCR → devolver líneas al encargado → revisar → confirmar
 * Desviación: Consumo Real vs Consumo Teórico (PDF TPV)
 */

// ── Nombres de hojas ──────────────────────────────────────────────────────────
const SHEET_CONTEOS          = "Conteos";
const SHEET_ALBARANES        = "Albaranes";
const SHEET_ALBARANES_DETALLE= "Albaranes_Detalle";
const SHEET_CONSUMO_TEORICO  = "Consumo_Teorico";
const SHEET_DESVIACION       = "Desviacion";
const SHEET_LOG              = "Log";

// ── Drive ─────────────────────────────────────────────────────────────────────
const DRIVE_FOLDER_ALBARANES = "TGB Albaranes";
const DRIVE_FOLDER_VENTAS    = "TGB Ventas";

// ── Tabla de mapeo: código ERE → nombre maestro app ───────────────────────────
// Añadir/modificar según catálogo Logirest real
const MAPA_PRODUCTOS = {
  "ERE100002236": "Hamburguesa Vacuno 135g",
  "ERE100000225": "Lagrimas de Pollo",
  "ERE100002427": "Crispy Chicken",
  "ERE100000184": "Alitas Barbacoa 2kg",
  "ERE100002135": "Pulled Pork TGB",
  "ERE100002005": "Bacon Loncheado",
  "ERE100001931": "Bacon Sello",
  "ERE100002215": "Pan Burger XL 64g",
  "ERE100002504": "Pan Burger Doble Disco",
  "ERE100000220": "Pan Lobster 75g",
  "ERE100002186": "Bites Cheddar Jalapeño",
  "ERE100000213": "Queso Cabra Empanado",
  "ERE100001838": "Salsa Cheddar 1kg",
  "ERE100000254": "Salsa Barbacoa Heinz",
  "ERE100002154": "Salsa Mostaza Miel",
  "ERE100000200": "Salsa TGB",
  "ERE100001696": "Patata Really Crunchy",
  "ERE100002350": "Sweet Potato Fries",
  "ERE100000310": "Harina Rebozar",
  "ERE100002301": "Aceite Fritura 5L",
  "ERE100002401": "Lechuga Batavia",
  "ERE100002402": "Tomate Maduro",
  "ERE100002403": "Cebolla Pelada",
  "ERE100002501": "Helado Vainilla C-10",
  "ERE100002502": "B&J Cookie 100ml",
  "ERE100002503": "B&J Chocolate 100ml"
};

// Mapeo por texto (fallback si no hay código ERE conocido)
const MAPA_TEXTO = {
  "HAMBURGUESA VACUNO":      "Hamburguesa Vacuno 135g",
  "LAGRIMAS DE POLLO":       "Lagrimas de Pollo",
  "CRISPY CHICKEN":          "Crispy Chicken",
  "ALITA BARBACOA":          "Alitas Barbacoa 2kg",
  "PULLED PORK TGB":         "Pulled Pork TGB",
  "BACON LONCHEADO":         "Bacon Loncheado",
  "BACON SELLO":             "Bacon Sello",
  "PAN BURGER TGB XL":       "Pan Burger XL 64g",
  "PAN BURGER XL":           "Pan Burger XL 64g",
  "PAN BURGER DOBLE":        "Pan Burger Doble Disco",
  "PAN LOBSTER":             "Pan Lobster 75g",
  "BITES CHEEDAR":           "Bites Cheddar Jalapeño",
  "BITES CHEDDAR":           "Bites Cheddar Jalapeño",
  "QUESO CABRA EMPANADO":    "Queso Cabra Empanado",
  "SALSA CHEDDAR":           "Salsa Cheddar 1kg",
  "SALSA KETCHUP HEINZ":     "Salsa Barbacoa Heinz",
  "SALSA BARBACOA":          "Salsa Barbacoa Heinz",
  "SALSA MAYONESA":          "Salsa Mayonesa Chovi",
  "SALSA TGB":               "Salsa TGB",
  "SALSA MOSTAZA":           "Salsa Mostaza Miel",
  "PATATA REALLY CRUNCHY":   "Patata Really Crunchy",
  "SWEET POTATO":            "Sweet Potato Fries",
  "HARINA TRIGOMAIZ":        "Harina Rebozar",
  "HARINA REBOZAR":          "Harina Rebozar",
  "ACEITE":                  "Aceite Fritura 5L",
  "LECHUGA BATAVIA":         "Lechuga Batavia",
  "TOMATE MADURO":           "Tomate Maduro",
  "CEBOLLA ENTERA":          "Cebolla Pelada",
  "CEBOLLA PELADA":          "Cebolla Pelada",
  "HELADO VAINILLA":         "Helado Vainilla C-10",
  "COOKIE DOUGH":            "B&J Cookie 100ml",
  "CHOCOLATE FUDGE":         "B&J Chocolate 100ml"
};

// =============================================================================
// ENTRADA PRINCIPAL
// =============================================================================
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    log_("recibido", payload.type + " · " + (payload.venue || "?"));
    let result;
    if      (payload.type === "conteo")           result = handleConteo(payload);
    else if (payload.type === "albaran_ocr")      result = handleAlbaranOCR(payload);
    else if (payload.type === "albaran_confirmar")result = handleAlbaranConfirmar(payload);
    else if (payload.type === "ventas")           result = handleVentas(payload);
    else throw new Error("Tipo desconocido: " + payload.type);
    return jsonResponse({ ok: true, result: result });
  } catch (err) {
    log_("error", String(err && err.message ? err.message : err));
    return jsonResponse({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function doGet() {
  return jsonResponse({ ok: true, message: "TGB Stock backend activo", version: "4.0",
                        timestamp: new Date().toISOString() });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// =============================================================================
// HANDLER: CONTEO DE INVENTARIO
// =============================================================================
function handleConteo(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureSheet_(ss, SHEET_CONTEOS, [
    "Fecha envío","Local","Semana inicio","Semana fin","Tipo conteo",
    "Responsable","Producto","Categoría","Unidad","Cantidad","ID producto","Versión app"
  ]);

  const ts = new Date(p.timestamp || new Date());
  const rows = [];
  for (const it of p.items) {
    rows.push([
      ts, p.venue, p.weekStart, p.weekEnd,
      p.tipoConteo || "inicio",          // "inicio" | "fin"
      p.responsable || "",
      it.nombre, it.categoria, it.und,
      (it.cantidad === null || it.cantidad === undefined) ? "" : it.cantidad,
      it.id, p.appVersion || ""
    ]);
  }
  if (rows.length > 0)
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

  // Si es conteo FIN, recalcular desviación
  if (p.tipoConteo === "fin") {
    try { calcularDesviacion_(); } catch(e) { log_("warn", "desviacion: " + e.message); }
  }

  return { rows: rows.length, venue: p.venue, week: p.weekStart, tipoConteo: p.tipoConteo };
}

// =============================================================================
// HANDLER: ALBARÁN FASE 1 — subir foto/PDF y hacer OCR
// Devuelve las líneas detectadas para que el encargado las revise en la app
// =============================================================================
function handleAlbaranOCR(p) {
  const folder  = ensureFolder_(DRIVE_FOLDER_ALBARANES);
  const ts      = new Date(p.timestamp || new Date());
  const tsStr   = Utilities.formatDate(ts, "Europe/Madrid", "yyyyMMdd_HHmmss");

  // Determinar si es imagen o PDF
  const isImage = (p.photoBase64 && p.photoBase64.length > 0);
  const base64raw = isImage
    ? (p.photoBase64 || "").replace(/^data:[^;]+;base64,/, "")
    : (p.pdfBase64  || "").replace(/^data:[^;]+;base64,/, "");
  const mimeType = isImage ? "image/jpeg" : "application/pdf";
  const ext      = isImage ? "jpg" : "pdf";
  const fname    = "albaran_" + p.venue + "_" + tsStr + "." + ext;

  // Guardar archivo original en Drive
  const blob = Utilities.newBlob(Utilities.base64Decode(base64raw), mimeType, fname);
  const file  = folder.createFile(blob);

  // ── OCR con Drive nativo ──────────────────────────────────────────────────
  // Drive convierte a Google Doc haciendo OCR automáticamente
  let ocrItems  = [];
  let ocrStatus = "OK";
  try {
    ocrItems = ocrConDriveNativo_(blob, p.venue);
  } catch (ocrErr) {
    ocrStatus = "Error OCR: " + ocrErr.message;
    log_("error", "OCR Drive: " + ocrErr.message);
  }

  // Registrar en hoja Albaranes (pendiente de confirmación)
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureSheet_(ss, SHEET_ALBARANES, [
    "Fecha envío","Local","Semana inicio","Semana fin",
    "Archivo (Drive)","URL Drive","Tamaño KB","Estado OCR","Nº líneas OCR","Versión app"
  ]);
  sheet.appendRow([
    ts, p.venue, p.weekStart, p.weekEnd,
    fname, file.getUrl(), p.sizeKB || "",
    "Pendiente confirmación (" + ocrItems.length + " líneas)",
    ocrItems.length, p.appVersion || ""
  ]);

  return {
    fileName : fname,
    url      : file.getUrl(),
    venue    : p.venue,
    weekStart: p.weekStart,
    ocrStatus: ocrStatus,
    items    : ocrItems   // devolver al encargado para revisión
  };
}

// =============================================================================
// HANDLER: ALBARÁN FASE 2 — encargado confirma/corrige las líneas
// =============================================================================
function handleAlbaranConfirmar(p) {
  // p.items = array corregido por el encargado: [{codigo, producto, productoMaestro, cantidad}, ...]
  // p.fileName = nombre del archivo ya guardado en Drive

  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const detalle = ensureSheet_(ss, SHEET_ALBARANES_DETALLE, [
    "Fecha","Local","Semana inicio","Archivo",
    "Código artículo","Producto Logirest","Producto Maestro","Cantidad"
  ]);

  const ts = new Date(p.timestamp || new Date());
  const rows = (p.items || []).map(it => [
    ts, p.venue, p.weekStart, p.fileName || "",
    it.codigo || "", it.producto || "",
    it.productoMaestro || it.producto || "",
    Number(it.cantidad) || 0
  ]);

  if (rows.length > 0)
    detalle.getRange(detalle.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

  // Actualizar estado en hoja Albaranes
  const shAlb = ss.getSheetByName(SHEET_ALBARANES);
  if (shAlb && p.fileName) {
    const data = shAlb.getDataRange().getValues();
    for (let r = 1; r < data.length; r++) {
      if (data[r][4] === p.fileName) {
        shAlb.getRange(r + 1, 8).setValue("Confirmado (" + rows.length + " líneas)");
        break;
      }
    }
  }

  // Recalcular desviación con nuevas entradas
  try { calcularDesviacion_(); } catch(e) { log_("warn", "desviacion: " + e.message); }

  log_("info", "Albarán confirmado: " + p.fileName + " · " + rows.length + " líneas · " + p.venue);
  return { rows: rows.length, venue: p.venue, weekStart: p.weekStart };
}

// =============================================================================
// HANDLER: VENTAS — subir PDF de TPV "Consumos de Materia Prima"
// =============================================================================
function handleVentas(p) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const shVentas= ensureSheet_(ss, SHEET_CONSUMO_TEORICO, [
    "Fecha envío","Local","Semana inicio","Semana fin",
    "Archivo (Drive)","URL Drive","Tamaño KB","Versión app"
  ]);
  const folder = ensureFolder_(DRIVE_FOLDER_VENTAS);
  const ts     = new Date(p.timestamp || new Date());
  const tsStr  = Utilities.formatDate(ts, "Europe/Madrid", "yyyyMMdd_HHmmss");
  const fname  = "ventas_" + (p.venue || "ADMIN") + "_" + tsStr + ".pdf";
  const base64 = (p.pdfBase64 || "").replace(/^data:[^;]+;base64,/, "");
  const blob   = Utilities.newBlob(Utilities.base64Decode(base64), "application/pdf", fname);
  const file   = folder.createFile(blob);
  shVentas.appendRow([
    ts, p.venue || "ADMIN", p.weekStart, p.weekEnd,
    fname, file.getUrl(), p.sizeKB || "", p.appVersion || ""
  ]);
  return { fileName: fname, url: file.getUrl() };
}

// =============================================================================
// OCR CON DRIVE NATIVO (gratis, sin API key)
// Sube la imagen/PDF como Google Doc con OCR activado → extrae texto → parsea
// =============================================================================
function ocrConDriveNativo_(blob, venue) {
  // Convertir a Google Doc (Drive hace OCR automáticamente)
  const resource = { title: "ocr_tmp_" + new Date().getTime(), mimeType: "application/vnd.google-apps.document" };
  const options  = { ocr: true, ocrLanguage: "es" };

  const file = Drive.Files.insert(resource, blob, options);
  const docId = file.id;

  // Leer el texto del Doc generado
  const doc  = DocumentApp.openById(docId);
  const text = doc.getBody().getText();

  // Eliminar el Doc temporal
  try { DriveApp.getFileById(docId).setTrashed(true); } catch(e) {}

  // Parsear el texto según formato Logirest
  return parsearAlbaranLogirest_(text);
}

// =============================================================================
// PARSER ALBARÁN LOGIREST
// Formato de línea: ERE1000XXXXX  DESCRIPCION  CANTIDAD  PESO  PU  IVA%  IMPORTE
// =============================================================================
function parsearAlbaranLogirest_(text) {
  const items = [];
  const lines = text.split("\n");

  // Regex para líneas de producto Logirest
  // Patrón: ERE100XXXXXX seguido de descripción, luego números
  const reLinea = /^(ERE\d{9,12})\s+(.+?)\s+(\d+)\s+[\d,]+\s+[\d,]+\s+[\d,]+%?\s+[\d,]+/;
  // Patrón alternativo más simple: ERE + texto + número al inicio de nums
  const reSimple = /^(ERE\d{9,12})\s+(.+)/;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line.startsWith("ERE")) continue;

    let codigo = "", producto = "", cantidad = 0;

    const m = reLinea.exec(line);
    if (m) {
      codigo   = m[1].trim();
      // El texto puede tener la cantidad al final — extraer la primera secuencia numérica después del texto
      const resto = line.slice(codigo.length).trim();
      // Separar texto de números: buscar primer número entero después del nombre
      const partes = resto.match(/^(.*?)\s+(\d+)\s+[\d,]+(\s+.*)?$/);
      if (partes) {
        producto = partes[1].trim().replace(/BONIFICADO/gi, "").trim();
        cantidad = parseInt(partes[2], 10);
      } else {
        producto = m[2].trim();
        cantidad = parseInt(m[3], 10);
      }
    } else {
      const ms = reSimple.exec(line);
      if (!ms) continue;
      codigo = ms[1].trim();
      const resto = ms[2].trim();
      // Intentar extraer nombre y cantidad del resto
      // El primer número entero suelto = cantidad
      const numMatch = resto.match(/^(.*?)\s+(\d+)\s/);
      if (numMatch) {
        producto = numMatch[1].trim().replace(/BONIFICADO/gi, "").trim();
        cantidad = parseInt(numMatch[2], 10);
      } else {
        producto = resto;
        cantidad = 0;
      }
    }

    // Limpiar nombre del producto
    producto = producto.replace(/\s+/g, " ").trim();

    // Mapear a nombre maestro
    const productoMaestro = mapearProducto_(codigo, producto);

    items.push({
      codigo         : codigo,
      producto       : producto,          // nombre tal como viene en Logirest
      productoMaestro: productoMaestro,   // nombre en la app (puede ser igual si no hay mapeo)
      cantidad       : cantidad,
      mapeado        : productoMaestro !== producto
    });
  }

  return items;
}

// =============================================================================
// MAPEO PRODUCTO: código ERE → nombre maestro, con fallback por texto
// =============================================================================
function mapearProducto_(codigo, textoLogirest) {
  // 1. Buscar por código ERE exacto
  if (MAPA_PRODUCTOS[codigo]) return MAPA_PRODUCTOS[codigo];

  // 2. Buscar por texto (palabras clave)
  const textoUp = textoLogirest.toUpperCase();
  for (const [clave, nombre] of Object.entries(MAPA_TEXTO)) {
    if (textoUp.includes(clave.toUpperCase())) return nombre;
  }

  // 3. Sin mapeo: devolver el texto original capitalizado
  return textoLogirest.split(" ")
    .map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
    .join(" ");
}

// =============================================================================
// CÁLCULO DE DESVIACIÓN
// Consumo Real    = Stock_inicio + Entradas_albarán - Stock_fin
// Consumo Teórico = datos del PDF de TPV (hoja Consumo_Teorico_Detalle) — pendiente
// Desviación      = Teórico - Real
// =============================================================================
function calcularDesviacion_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shConteos = ss.getSheetByName(SHEET_CONTEOS);
  const shDetalle = ss.getSheetByName(SHEET_ALBARANES_DETALLE);
  const shDesv    = ensureSheet_(ss, SHEET_DESVIACION, [
    "Semana inicio","Local","Producto",
    "Stock inicio","Entradas albarán","Stock fin",
    "Consumo Real","Consumo Teórico","Desviación","% Desviación"
  ]);

  if (!shConteos || shConteos.getLastRow() < 2) return;

  // Leer conteos
  const contData    = shConteos.getDataRange().getValues();
  const contHeaders = contData[0];
  const iLocal      = contHeaders.indexOf("Local");
  const iWeekStart  = contHeaders.indexOf("Semana inicio");
  const iTipo       = contHeaders.indexOf("Tipo conteo");
  const iProducto   = contHeaders.indexOf("Producto");
  const iCantidad   = contHeaders.indexOf("Cantidad");

  // Agrupar: por semana+local+producto, separar inicio y fin
  // key = "semana|local|producto"
  const mapInicio = {};
  const mapFin    = {};

  for (let r = 1; r < contData.length; r++) {
    const row    = contData[r];
    const local  = row[iLocal];
    const semana = fmtFecha_(row[iWeekStart]);
    const tipo   = (row[iTipo] || "inicio").toString().toLowerCase();
    const prod   = row[iProducto];
    const cant   = Number(row[iCantidad]) || 0;
    const key    = semana + "|" + local + "|" + prod;

    if (tipo === "fin") {
      // Último fin gana
      if (!mapFin[key] || row[0] > mapFin[key].fecha)
        mapFin[key] = { semana, local, prod, cant, fecha: row[0] };
    } else {
      // "inicio" — último inicio gana
      if (!mapInicio[key] || row[0] > mapInicio[key].fecha)
        mapInicio[key] = { semana, local, prod, cant, fecha: row[0] };
    }
  }

  // REGLA CLAVE: si no hay conteo inicio para esta semana,
  // usar el conteo fin de la semana ANTERIOR como inicio (carry-forward)
  const semanas = [...new Set([
    ...Object.values(mapInicio).map(v => v.semana),
    ...Object.values(mapFin   ).map(v => v.semana)
  ])].sort();

  // Completar mapInicio con carry-forward
  for (const key of Object.keys(mapFin)) {
    const { semana, local, prod, cant } = mapFin[key];
    const idxSem = semanas.indexOf(semana);
    if (idxSem > 0) {
      const semSig = semanas[idxSig = idxSem]; // esta semana
      // El FIN de semana N se convierte en INICIO de semana N+1
      if (idxSem + 1 < semanas.length) {
        const semNext = semanas[idxSem + 1];
        const keyNext = semNext + "|" + local + "|" + prod;
        if (!mapInicio[keyNext]) {
          mapInicio[keyNext] = { semana: semNext, local, prod, cant, fecha: mapFin[key].fecha };
        }
      }
    }
  }

  // Leer entradas de albarán por semana+local+productoMaestro
  const entradasMap = {};
  if (shDetalle && shDetalle.getLastRow() > 1) {
    const detData    = shDetalle.getDataRange().getValues();
    const detHeaders = detData[0];
    const dLocal  = detHeaders.indexOf("Local");
    const dSemana = detHeaders.indexOf("Semana inicio");
    const dProd   = detHeaders.indexOf("Producto Maestro");
    const dCant   = detHeaders.indexOf("Cantidad");
    for (let r = 1; r < detData.length; r++) {
      const row    = detData[r];
      const local  = row[dLocal];
      const semana = fmtFecha_(row[dSemana]);
      const prod   = row[dProd];
      const cant   = Number(row[dCant]) || 0;
      const key    = semana + "|" + local + "|" + prod;
      entradasMap[key] = (entradasMap[key] || 0) + cant;
    }
  }

  // Calcular consumo real para cada semana+local+producto con AMBOS conteos
  const desvRows = [];
  for (const key of Object.keys(mapFin)) {
    const { semana, local, prod, cant: stockFin } = mapFin[key];
    const inicio   = mapInicio[key];
    const stockIni = inicio ? inicio.cant : 0;
    const entradas = entradasMap[key] || 0;

    const consumoReal = stockIni + entradas - stockFin;
    // Consumo teórico: pendiente — se leerá de hoja futura
    const consumoTeorico = "";
    const desviacion     = "";
    const pctDesv        = "";

    desvRows.push([
      semana, local, prod,
      stockIni, entradas, stockFin,
      consumoReal, consumoTeorico, desviacion, pctDesv
    ]);
  }

  // Reescribir hoja Desviacion
  if (shDesv.getLastRow() > 1)
    shDesv.getRange(2, 1, shDesv.getLastRow() - 1, 10).clearContent();
  if (desvRows.length > 0)
    shDesv.getRange(2, 1, desvRows.length, 10).setValues(desvRows);
}

// =============================================================================
// HELPERS
// =============================================================================
function fmtFecha_(val) {
  if (val instanceof Date)
    return Utilities.formatDate(val, "Europe/Madrid", "yyyy-MM-dd");
  return String(val);
}

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
// SETUP INICIAL — ejecutar una vez para crear hojas y carpetas
// =============================================================================
function testSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEET_CONTEOS, [
    "Fecha envío","Local","Semana inicio","Semana fin","Tipo conteo",
    "Responsable","Producto","Categoría","Unidad","Cantidad","ID producto","Versión app"
  ]);
  ensureSheet_(ss, SHEET_ALBARANES, [
    "Fecha envío","Local","Semana inicio","Semana fin",
    "Archivo (Drive)","URL Drive","Tamaño KB","Estado OCR","Nº líneas OCR","Versión app"
  ]);
  ensureSheet_(ss, SHEET_ALBARANES_DETALLE, [
    "Fecha","Local","Semana inicio","Archivo",
    "Código artículo","Producto Logirest","Producto Maestro","Cantidad"
  ]);
  ensureSheet_(ss, SHEET_CONSUMO_TEORICO, [
    "Fecha envío","Local","Semana inicio","Semana fin",
    "Archivo (Drive)","URL Drive","Tamaño KB","Versión app"
  ]);
  ensureSheet_(ss, SHEET_DESVIACION, [
    "Semana inicio","Local","Producto",
    "Stock inicio","Entradas albarán","Stock fin",
    "Consumo Real","Consumo Teórico","Desviación","% Desviación"
  ]);
  ensureSheet_(ss, SHEET_LOG, ["Fecha","Nivel","Mensaje"]);
  ensureFolder_(DRIVE_FOLDER_ALBARANES);
  ensureFolder_(DRIVE_FOLDER_VENTAS);
  Logger.log("✓ Setup v4.0 completado");
}
