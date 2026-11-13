/**
 * ==============================================================================
 * SISTEMA CENTRALIZADO DE AUDITORÍA FDM - BACKEND GOOGLE DRIVE & SHEETS v3.2
 * ==============================================================================
 * TABLA 1: "CENSO_CAJAS_FOTOS" (Inventario de cada caja física, ambiente, contenido y fotos)
 * TABLA 2: "MATRIZ_69_OBRAS"   (Ficha contractual del proyecto: Fechas y 11 Checks de Auditoría)
 * ==============================================================================
 */

const FOLDER_NAME = "AUDITORIA_FDM_EVIDENCIAS_2022_2025";
const SHEET_CAJAS = "CENSO_CAJAS_FOTOS";
const SHEET_OBRAS = "MATRIZ_69_OBRAS";

const CHECK_KEYS = [
  'chk_exp', 'chk_cont', 'chk_acta_ent', 'chk_cuad', 
  'chk_val', 'chk_adic', 'chk_sup', 'chk_rec', 
  'chk_liq', 'chk_res', 'chk_transf'
];

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCajas = getOrCreateSheetCajas(ss);
    const sheetObras = getOrCreateSheetObras(ss);

    // 1. Leer Cajas y Fotos
    const dataCajas = sheetCajas.getDataRange().getValues();
    const cajas = [];
    for (let i = 1; i < dataCajas.length; i++) {
      const row = dataCajas[i];
      if (!row[0]) continue;
      cajas.push({
        timestamp: row[0],
        obra_id: parseInt(row[1]),
        codigo_caja: row[2],
        ambiente: row[3],
        contenido_caja: row[4],
        tipo_foto: row[5],
        drive_url: row[6],
        file_id: row[7] || "",
        ocr_acervo: row[8] || "",
        ocr_raw: row[9] || ""
      });
    }

    // 2. Leer Matriz de 69 Obras (Solo Fechas y 11 Checks)
    const dataObras = sheetObras.getDataRange().getValues();
    const obras = {};
    for (let i = 1; i < dataObras.length; i++) {
      const row = dataObras[i];
      const oId = parseInt(row[0]);
      if (!oId) continue;

      const checks = {};
      CHECK_KEYS.forEach((k, idx) => {
        checks[k] = (row[6 + idx] === "SÍ" || row[6 + idx] === true);
      });

      obras[oId] = {
        obra_id: oId,
        desc: row[1],
        cui: row[2],
        provincia: row[3],
        fecha_inicio: row[4],
        fecha_culm: row[5],
        checks: checks,
        notas_auditoria: row[17] || ""
      };
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      total_cajas: cajas.length,
      cajas: cajas,
      obras: obras
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const contents = e.postData.contents;
    const payload = JSON.parse(contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const action = payload.action || "save_box_photo";

    // ==============================================================
    // CASO 1: GUARDAR O ACTUALIZAR CHECKLIST DE AUDITORÍA (MATRIZ_69_OBRAS)
    // ==============================================================
    if (action === "save_obra_checks") {
      const sheetObras = getOrCreateSheetObras(ss);
      const obraId = parseInt(payload.obra_id);
      upsertObraRow(sheetObras, obraId, payload);

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        action: "save_obra_checks",
        obra_id: obraId,
        message: "Checklist y fechas actualizadas en MATRIZ_69_OBRAS"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ==============================================================
    // CASO 2: GUARDAR FOTO DE CAJA / ADICIONAL (CENSO_CAJAS_FOTOS)
    // ==============================================================
    if (action === "save_box_photo") {
      const sheetCajas = getOrCreateSheetCajas(ss);

      const obraId = parseInt(payload.obra_id);
      const desc = payload.desc || "";
      const codigoCaja = String(payload.codigo_caja || "").trim();
      const ambiente = payload.ambiente || payload.ubicacion || "";
      const contenidoCaja = payload.contenido_caja || payload.nota || "";
      const tipoFoto = payload.tipo_foto || payload.slot || "Rótulo de Caja";

      let driveUrl = "";
      let fileId = "";

      if (payload.photo && payload.photo.base64) {
        const rootFolder = getOrCreateFolder(FOLDER_NAME);
        const safeDesc = desc.replace(/[^a-zA-Z0-9_-]/g, '_').substring(0, 30);
        const obraFolderName = `OBRA_${String(obraId).padStart(2, '0')}_${safeDesc}`;
        const obraFolder = getOrCreateSubFolder(rootFolder, obraFolderName);

        const base64Data = payload.photo.base64.replace(/^data:image\/\w+;base64,/, "");
        const decodedBytes = Utilities.base64Decode(base64Data);
        const filename = payload.photo.filename || `OBRA_${obraId}_${codigoCaja}.jpg`;
        const blob = Utilities.newBlob(decodedBytes, "image/jpeg", filename);
        
        const file = obraFolder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        driveUrl = file.getUrl();
        fileId = file.getId();
      }

      // OCR Automático en la Nube sobre la imagen guardada
      let ocrAcervo = "";
      let ocrRaw = "";
      if (fileId) {
        try {
          const ocrRes = extraerTextoOcrDeArchivo(fileId);
          ocrAcervo = ocrRes.limpio || "";
          ocrRaw = ocrRes.raw || "";
        } catch (ocrErr) {
          Logger.log("Error OCR automático: " + ocrErr);
        }
      }

      sheetCajas.appendRow([
        new Date().toLocaleString(),
        obraId,
        codigoCaja,
        ambiente,
        contenidoCaja,
        tipoFoto,
        driveUrl,
        fileId,
        ocrAcervo,
        ocrRaw
      ]);

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        action: "save_box_photo",
        obra_id: obraId,
        codigo_caja: codigoCaja,
        drive_url: driveUrl,
        file_id: fileId,
        ocr_acervo: ocrAcervo
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ==============================================================
    // CASO 2.B: EJECUTAR OCR POR LOTES SOBRE FOTOS YA SUBIDAS
    // ==============================================================
    if (action === "run_batch_ocr") {
      const sheetCajas = getOrCreateSheetCajas(ss);
      const targetObraId = payload.obra_id ? parseInt(payload.obra_id) : null;
      const forceAll = payload.force === true;
      const data = sheetCajas.getDataRange().getValues();
      let processed = 0;

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowObraId = parseInt(row[1]);
        let fileId = String(row[7] || "").trim();
        const existingOcr = String(row[8] || "").trim();

        if (targetObraId && rowObraId !== targetObraId) continue;

        // Si no tiene fileId en col 8 pero tiene URL de Drive en col 7, extraer fileId
        if (!fileId && row[6]) {
          const m = String(row[6]).match(/[-\w]{25,}/);
          if (m) {
            fileId = m[0];
            sheetCajas.getRange(i + 1, 8).setValue(fileId);
          }
        }

        if (!fileId) continue;
        if (existingOcr && !forceAll) continue; // Si ya tiene OCR y no forzamos, omitir

        try {
          const ocrRes = extraerTextoOcrDeArchivo(fileId);
          if (ocrRes.limpio || ocrRes.raw) {
            sheetCajas.getRange(i + 1, 9).setValue(ocrRes.limpio);
            sheetCajas.getRange(i + 1, 10).setValue(ocrRes.raw);
            processed++;
          }
        } catch (err) {
          Logger.log("Error procesando lote fila " + (i + 1) + ": " + err);
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        action: "run_batch_ocr",
        obra_id: targetObraId,
        total_procesados: processed,
        message: `OCR completado para ${processed} fotos`
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ==============================================================
    // CASO 2.C: ACTUALIZAR CONTENIDO/RÓTULO EN GABINETE (NO DESTRUCTIVO)
    // ==============================================================
    if (action === "update_box_content") {
      const sheetCajas = getOrCreateSheetCajas(ss);
      const obraId = parseInt(payload.obra_id);
      const codigoCaja = String(payload.codigo_caja || "").trim();
      const nuevoContenido = String(payload.contenido_caja || "").trim();

      if (!obraId || !codigoCaja) {
        return ContentService.createTextOutput(JSON.stringify({
          status: "error",
          message: "Faltan parámetros obra_id o codigo_caja"
        })).setMimeType(ContentService.MimeType.JSON);
      }

      // Respaldo de seguridad antes de editar texto de rótulo
      try {
        crearRespaldoAutomaticoCajas(ss, sheetCajas, "RESPALDO_CAJAS_AUTO");
      } catch (bakErr) {
        Logger.log("Aviso respaldo automático: " + bakErr);
      }

      const data = sheetCajas.getDataRange().getValues();
      let updatedRows = 0;
      for (let i = 1; i < data.length; i++) {
        const rowObraId = parseInt(data[i][1]);
        const rowCod = String(data[i][2] || "").trim();
        if (rowObraId === obraId && rowCod === codigoCaja) {
          // Actualizar UNICAMENTE la columna 5 (CONTENIDO_DE_ESTA_CAJA)
          sheetCajas.getRange(i + 1, 5).setValue(nuevoContenido);
          updatedRows++;
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        action: "update_box_content",
        obra_id: obraId,
        codigo_caja: codigoCaja,
        filas_actualizadas: updatedRows,
        message: `Rótulo actualizado en ${updatedRows} filas sin alterar archivos de Drive`
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ==============================================================
    // CASO 2.D: GENERAR RESPALDO COMPLETO EN HOJA DE CÁLCULO
    // ==============================================================
    if (action === "backup_sheet") {
      const sheetCajas = getOrCreateSheetCajas(ss);
      const now = new Date();
      const fechaTag = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy_MM_dd");
      const backupName = "BACKUP_CAJAS_" + fechaTag;
      const totalRows = crearRespaldoAutomaticoCajas(ss, sheetCajas, backupName);

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        action: "backup_sheet",
        backup_sheet_name: backupName,
        total_filas: totalRows,
        message: `Respaldo guardado en pestaña ${backupName} (${totalRows} filas)`
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ==============================================================
    // CASO 3: BORRAR CAJA O FOTO ESPECÍFICA (CENSO_CAJAS_FOTOS)
    // ==============================================================
    if (action === "delete_record") {
      const sheetCajas = getOrCreateSheetCajas(ss);

      const obraId = parseInt(payload.obra_id);
      const codigoCaja = String(payload.codigo_caja || "").trim();
      const driveUrl = String(payload.drive_url || "").trim();
      let fileId = payload.file_id || "";

      if (!fileId && driveUrl) {
        const m = driveUrl.match(/[-\w]{25,}/);
        if (m) fileId = m[0];
      }

      let fileDeleted = false;
      if (fileId) {
        try {
          const file = DriveApp.getFileById(fileId);
          if (file) { file.setTrashed(true); fileDeleted = true; }
        } catch (e) {}
      }

      const data = sheetCajas.getDataRange().getValues();
      let rowDeleted = false;
      for (let i = data.length - 1; i >= 1; i--) {
        const row = data[i];
        const rowObraId = parseInt(row[1]);
        const rowCod = String(row[2] || "").trim();
        const rowUrl = String(row[6] || "").trim();
        const rowFileId = String(row[7] || "").trim();

        const matchObra = (rowObraId === obraId);
        const matchFile = (fileId && rowFileId === fileId) || (driveUrl && rowUrl === driveUrl);
        const matchCod = (codigoCaja && rowCod === codigoCaja);

        // Si se provee link de archivo, borrar exactamente esa foto
        if (matchObra && (matchFile || (!driveUrl && matchCod))) {
          sheetCajas.deleteRow(i + 1);
          rowDeleted = true;
          break;
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        action: "delete_record",
        obra_id: obraId,
        codigo_caja: codigoCaja,
        row_deleted: rowDeleted,
        file_deleted: fileDeleted
      })).setMimeType(ContentService.MimeType.JSON);
    }

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function upsertObraRow(sheetObras, obraId, payload) {
  const data = sheetObras.getDataRange().getValues();
  let targetRowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if (parseInt(data[i][0]) === obraId) {
      targetRowIndex = i + 1;
      break;
    }
  }

  const desc = payload.desc || "";
  const cui = payload.cui || "";
  const prov = payload.provincia || "";
  const fIni = payload.fecha_inicio || "";
  const fCulm = payload.fecha_culm || "";
  const notasAuditoria = payload.notas_auditoria || payload.notas || "";
  const checks = payload.checks || {};

  const checkCols = CHECK_KEYS.map(k => checks[k] ? "SÍ" : "");

  const rowValues = [
    obraId, desc, cui, prov, fIni, fCulm,
    checkCols[0], checkCols[1], checkCols[2], checkCols[3],
    checkCols[4], checkCols[5], checkCols[6], checkCols[7],
    checkCols[8], checkCols[9], checkCols[10],
    notasAuditoria, new Date().toLocaleString()
  ];

  if (targetRowIndex !== -1) {
    sheetObras.getRange(targetRowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheetObras.appendRow(rowValues);
  }
}

function getOrCreateSheetCajas(ss) {
  let sheet = ss.getSheetByName(SHEET_CAJAS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CAJAS, 0);
    const headers = [
      "TIMESTAMP", "OBRA_ID", "CODIGO_CAJA", "AMBIENTE", 
      "CONTENIDO_DE_ESTA_CAJA", "TIPO_FOTO", "ENLACE_FOTO_DRIVE", "FILE_ID_DRIVE",
      "ACERVO_DETECTADO_OCR", "TEXTO_COMPLETO_OCR"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground("#0F766E").setFontColor("#FFFFFF").setFontWeight("bold");
    sheet.setFrozenRows(1);
  } else {
    // Si la hoja ya existe pero no tiene las columnas OCR, agregarlas
    const lastCol = sheet.getLastColumn();
    if (lastCol <= 8) {
      sheet.getRange(1, 9).setValue("ACERVO_DETECTADO_OCR").setBackground("#0F766E").setFontColor("#FFFFFF").setFontWeight("bold");
      sheet.getRange(1, 10).setValue("TEXTO_COMPLETO_OCR").setBackground("#0F766E").setFontColor("#FFFFFF").setFontWeight("bold");
    }
  }
  return sheet;
}

function crearRespaldoAutomaticoCajas(ss, sheetCajas, backupName) {
  const sourceRange = sheetCajas.getDataRange();
  const values = sourceRange.getValues();
  if (!values || values.length === 0) return 0;

  let backupSheet = ss.getSheetByName(backupName);
  if (!backupSheet) {
    backupSheet = ss.insertSheet(backupName);
  } else {
    backupSheet.clearContents();
  }

  backupSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
  backupSheet.getRange(1, 1, 1, values[0].length).setBackground("#1E293B").setFontColor("#FFFFFF").setFontWeight("bold");
  backupSheet.setFrozenRows(1);
  return Math.max(0, values.length - 1);
}

function getOrCreateSheetObras(ss) {
  let sheet = ss.getSheetByName(SHEET_OBRAS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_OBRAS, 1);
    const headers = [
      "OBRA_N°", "DESCRIPCION_PROYECTO", "CUI", "PROVINCIA", 
      "FECHA_INICIO_REAL", "FECHA_CULMINACION_REAL",
      "1_EXP_TECNICO", "2_CONTRATO", "3_ACTA_INICIO", "4_CUADERNO",
      "5_VALORIZACIONES", "6_ADICIONALES", "7_SUPERVISION", "8_ACTA_RECEPCION",
      "9_LIQ_OBRA", "10_RES_LIQ_FDM", "11_TRANSFERENCIA",
      "NOTAS_AUDITORIA", "ULTIMA_ACTUALIZACION"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground("#1E293B").setFontColor("#FFFFFF").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
}

function getOrCreateSubFolder(parentFolder, subFolderName) {
  const folders = parentFolder.getFoldersByName(subFolderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(subFolderName);
}

// ==========================================
// MOTOR DE OCR NATIVO DE GOOGLE DRIVE Y FILTRO
// ==========================================

/**
 * Realiza OCR nativo en Google Drive y extrae solo elementos físicos clave de auditoría.
 */
function extraerTextoOcrDeArchivo(fileId) {
  if (!fileId) return { limpio: "", raw: "" };

  try {
    const url = "https://www.googleapis.com/drive/v2/files/" + fileId + "/copy?ocr=true&ocrLanguage=es";
    const res = UrlFetchApp.fetch(url, {
      method: "post",
      headers: {
        "Authorization": "Bearer " + ScriptApp.getOAuthToken()
      },
      contentType: "application/json",
      payload: JSON.stringify({
        title: "TEMP_OCR_" + fileId
      }),
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      Logger.log("Drive OCR API code: " + res.getResponseCode());
      return { limpio: "", raw: "" };
    }

    const json = JSON.parse(res.getContentText());
    const tempDocId = json.id;
    if (!tempDocId) return { limpio: "", raw: "" };

    const doc = DocumentApp.openById(tempDocId);
    const rawText = doc.getBody().getText() || "";

    // Eliminar el archivo Google Doc temporal para mantener limpio Drive
    try {
      DriveApp.getFileById(tempDocId).setTrashed(true);
    } catch (e) {}

    const limpio = filtrarAcervoAuditoria(rawText);
    return { limpio: limpio, raw: rawText };

  } catch (err) {
    Logger.log("Error en extraerTextoOcrDeArchivo: " + err);
    return { limpio: "", raw: "" };
  }
}

/**
 * Filtra el texto detectado por OCR para omitir ruido y extraer únicamente
 * archivos físicos relevantes de auditoría (Expedientes, Valorizaciones, Cuadernos, Liquidaciones, etc.)
 */
function filtrarAcervoAuditoria(rawText) {
  if (!rawText) return "";

  const lines = rawText.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 3);
  const foundItems = [];
  const seenPatterns = new Set();

  const DOCUMENT_PATTERNS = [
    {
      tipo: "Expediente Técnico",
      regex: /(?:(?:tomo|volumen|vol\.?|item|n[°o]?)\s*[0-9ivx\.-]+\s*[-–:]?\s*)?(?:expediente\s+t[eé]cnico|exp\.?\s*t[eé]c\.?)(?:\s*(?:aprobado|original|copia))?(?:\s*[-–:]?\s*(?:tomo\s*[0-9ivx]+|volumen\s*[0-9ivx]+|[a-z0-9\s,-]+))?/i
    },
    {
      tipo: "Valorizaciones de Obra",
      regex: /(?:(?:tomo|volumen|vol\.?|item|n[°o]?)\s*[0-9ivx\.-]+\s*[-–:]?\s*)?(?:valorizaci[oó]n(?:es)?\s*(?:de\s*obra|principal)?|val\.?\s*(?:de\s*obra|n[°o]))(?:\s*(?:n[°o]?\s*[0-9]+(?:\s*(?:al?|[-–])\s*[0-9]+)?|[0-9ivx\s,-]+))?/i
    },
    {
      tipo: "Valorizaciones de Supervisión",
      regex: /(?:(?:tomo|volumen|vol\.?|item|n[°o]?)\s*[0-9ivx\.-]+\s*[-–:]?\s*)?(?:valorizaci[oó]n(?:es)?\s*(?:de\s*)?supervisi[oó]n|val\.?\s*superv\.?)(?:\s*(?:n[°o]?\s*[0-9]+(?:\s*(?:al?|[-–])\s*[0-9]+)?|[0-9ivx\s,-]+))?/i
    },
    {
      tipo: "Cuaderno de Obra",
      regex: /(?:(?:tomo|volumen|vol\.?|item|n[°o]?)\s*[0-9ivx\.-]+\s*[-–:]?\s*)?(?:cuaderno\s*de\s*obra|c\.?\s*de\s*obra)(?:\s*(?:n[°o]?\s*[0-9]+|tomo\s*[0-9ivx]+|original|copia|asientos?\s*[0-9\s–-]+))?/i
    },
    {
      tipo: "Informes Mensuales",
      regex: /(?:(?:tomo|volumen|vol\.?|item|n[°o]?)\s*[0-9ivx\.-]+\s*[-–:]?\s*)?(?:informe(?:s)?\s*mensual(?:es)?|inf\.?\s*mensual)(?:\s*(?:de\s*supervisi[oó]n|de\s*residencia|n[°o]?\s*[0-9]+|[a-z0-9\s,-]+))?/i
    },
    {
      tipo: "Reportes Semanales",
      regex: /(?:reporte(?:s)?\s*semanal(?:es)?|informe(?:s)?\s*semanal(?:es)?)(?:\s*(?:n[°o]?\s*[0-9]+|[a-z0-9\s,-]+))?/i
    },
    {
      tipo: "Informe Final / Pre-liquidación",
      regex: /(?:informe\s*final(?:\s*de\s*obra)?|pre\s*[-–]?\s*liquidaci[oó]n)/i
    },
    {
      tipo: "Liquidación Técnica / Financiera",
      regex: /(?:(?:tomo|volumen|vol\.?|item|n[°o]?)\s*[0-9ivx\.-]+\s*[-–:]?\s*)?(?:liquidaci[oó]n\s*(?:t[eé]cnica(?:\s*y\s*financiera)?|financiera|de\s*obra|de\s*contrato)?)(?:\s*[-–:]?\s*(?:tomo\s*[0-9ivx]+|[a-z0-9\s,-]+))?/i
    },
    {
      tipo: "Contrato y Bases",
      regex: /(?:contrato(?:\s*de\s*obra|\s*n[°o]\s*[0-9-]+)?|bases\s*integradas|adenda\s*n[°o]?\s*[0-9]+)/i
    },
    {
      tipo: "Adicionales / Ampliaciones de Plazo",
      regex: /(?:adicional\s*(?:de\s*obra)?\s*n[°o]?\s*[0-9]+|ampliaci[oó]n\s*de\s*plazo\s*n[°o]?\s*[0-9]+|deductivo)/i
    },
    {
      tipo: "Kárdex / Almacén / Guías",
      regex: /(?:k[aá]rdex(?:\s*de\s*almac[eé]n)?|gu[ií]as?\s*de\s*remisi[oó]n|pecoas?|orden(?:es)?\s*de\s*compra)/i
    },
    {
      tipo: "Actas de Obra",
      regex: /(?:acta\s*de\s*entrega\s*(?:de\s*terreno)?|acta\s*de\s*recepci[oó]n(?:\s*de\s*obra)?|acta\s*de\s*constataci[oó]n)/i
    },
    {
      tipo: "Pruebas de Calidad / Ensayos",
      regex: /(?:ensayos?\s*de\s*laboratorio|roturas?\s*de\s*probetas?|densidad\s*de\s*campo|control\s*de\s*calidad)/i
    },
    {
      tipo: "Planos de Replanteo",
      regex: /(?:planos?\s*de\s*replanteo|planos?\s*post\s*construcci[oó]n|as\s*[-–]?\s*built)/i
    }
  ];

  lines.forEach(line => {
    const cleanLine = line.replace(/\s+/g, ' ').trim();
    if (cleanLine.length < 4 || cleanLine.length > 100) return;

    DOCUMENT_PATTERNS.forEach(pat => {
      const match = cleanLine.match(pat.regex);
      if (match) {
        let extracted = match[0].trim();
        extracted = extracted.charAt(0).toUpperCase() + extracted.slice(1);
        const normKey = extracted.toLowerCase().replace(/[^a-z0-9]/g, '');
        if (!seenPatterns.has(normKey)) {
          seenPatterns.add(normKey);
          foundItems.push(extracted);
        }
      }
    });
  });

  // Si no coincidió ningún patrón específico pero hay líneas con contenido documental de obra
  if (foundItems.length === 0) {
    const fallbackKeywords = /(?:tomo|volumen|expediente|cuaderno|valorizaci|informe|liquidaci|contrato|acta|adicional|obra|archivo)/i;
    lines.forEach(line => {
      const cleanLine = line.replace(/\s+/g, ' ').trim();
      if (cleanLine.length >= 6 && cleanLine.length <= 80 && fallbackKeywords.test(cleanLine)) {
        const normKey = cleanLine.toLowerCase().replace(/[^a-z0-9]/g, '');
        if (!seenPatterns.has(normKey) && foundItems.length < 5) {
          seenPatterns.add(normKey);
          foundItems.push(cleanLine);
        }
      }
    });
  }

  return foundItems.join(' | ');
}
