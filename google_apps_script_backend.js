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
const SHEET_INFORME = "INFORME_SITUACIONAL_OBRAS";

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("🏛️ Auditoría FDM")
      .addItem("🖨️ Preparar Vista Impresión Horizontal", "formatearImpresionInforme")
      .addToUi();
  } catch (e) {}
}

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

    // 2. Leer Matriz de Obras (Fechas Inicio/Fin, Clasificación y Notas)
    const dataObras = sheetObras.getDataRange().getValues();
    const headerRow = dataObras.length > 0 ? dataObras[0] : [];
    const isNewSchema = String(headerRow[6] || "").trim() === "TIPO_PROYECTO";
    const obras = {};
    for (let i = 1; i < dataObras.length; i++) {
      const row = dataObras[i];
      const oId = parseInt(row[0]);
      if (!oId) continue;

      const fIni = formatSheetCellDateStr(row[4]);
      const fCulm = formatSheetCellDateStr(row[5]);
      const notas = isNewSchema ? (row[13] || "") : (row[17] || row[13] || "");

      obras[oId] = {
        obra_id: oId,
        desc: row[1],
        cui: row[2],
        provincia: row[3],
        fecha_inicio: fIni,
        fecha_culm: fCulm,
        tipo_proyecto: isNewSchema ? (row[6] || "") : "",
        convenio: isNewSchema ? (row[7] || "") : "",
        estatus: isNewSchema ? (row[8] || "") : "",
        estado_situacional: isNewSchema ? (row[9] || "") : "",
        notas_auditoria: notas
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

    // ==============================================================
    // CASO 4: CREAR / ACTUALIZAR INFORME SITUACIONAL EN GOOGLE SHEETS
    // ==============================================================
    if (action === "sync_informe_situacional") {
      const sheetInfo = crearHojaInformeSituacional(ss, payload);
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        action: "sync_informe_situacional",
        sheet_name: SHEET_INFORME,
        sheet_url: sheetInfo.sheetUrl,
        total_obras: sheetInfo.totalObras,
        message: `Hoja "${SHEET_INFORME}" actualizada con ${sheetInfo.totalObras} obras civiles y ubicaciones en vivo.`
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ==============================================================
    // CASO 5: POBLAR MATRIZ_69_OBRAS COMPLETA CON CLASIFICACIÓN (SIN CHECKLISTS)
    // ==============================================================
    if (action === "sync_matriz_obras_completa") {
      const resSync = poblarMatrizObrasCompleta(ss, payload);
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        action: "sync_matriz_obras_completa",
        total_obras: resSync.totalObras,
        backup_sheet: resSync.backupSheet,
        message: `MATRIZ_69_OBRAS poblada con ${resSync.totalObras} proyectos y respaldo en ${resSync.backupSheet}.`
      })).setMimeType(ContentService.MimeType.JSON);
    }

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function formatSheetCellDateStr(val) {
  if (!val) return "";
  if (val instanceof Date) {
    try {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
    } catch (e) {
      return val.toISOString().slice(0, 10);
    }
  }
  return String(val).trim();
}

function upsertObraRow(sheetObras, obraId, payload) {
  const data = sheetObras.getDataRange().getValues();
  const headerRow = data.length > 0 ? data[0] : [];
  const isNewSchema = String(headerRow[6] || "").trim() === "TIPO_PROYECTO";
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

  if (isNewSchema) {
    const rowValues = [
      obraId,
      desc,
      cui,
      prov,
      fIni,
      fCulm,
      payload.es_obra_civil || "OBRA CIVIL OBJETIVO",
      payload.convenio || "",
      payload.estatus || "",
      payload.estado_situacional || "",
      Number(payload.prio || 0),
      Number(payload.costo_ejec || 0),
      payload.ubicacion_ambientes || "",
      notasAuditoria,
      new Date().toLocaleString()
    ];
    if (targetRowIndex !== -1) {
      sheetObras.getRange(targetRowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      sheetObras.appendRow(rowValues);
    }
  } else {
    // Hoja con esquema anterior: actualizar columnas 1-6 y notas/fecha actualización (18-19) sin tocar 7-17
    if (targetRowIndex !== -1) {
      sheetObras.getRange(targetRowIndex, 1, 1, 6).setValues([[obraId, desc, cui, prov, fIni, fCulm]]);
      sheetObras.getRange(targetRowIndex, 18, 1, 2).setValues([[notasAuditoria, new Date().toLocaleString()]]);
    } else {
      const rowLegacy = [obraId, desc, cui, prov, fIni, fCulm, "", "", "", "", "", "", "", "", "", "", "", notasAuditoria, new Date().toLocaleString()];
      sheetObras.appendRow(rowLegacy);
    }
  }
}

function poblarMatrizObrasCompleta(ss, payload) {
  let sheetObras = getOrCreateSheetObras(ss);
  const dataActual = sheetObras.getDataRange().getValues();
  const isLegacy = dataActual.length > 0 && String(dataActual[0][6] || "").trim() !== "TIPO_PROYECTO";

  // 1. Respaldo de seguridad de lo que ya tenía MATRIZ_69_OBRAS
  const backupName = "BACKUP_MATRIZ_OBRAS_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
  if (dataActual.length > 1) {
    crearRespaldoAutomaticoCajas(ss, sheetObras, backupName);
  }

  // 2. Preservar fechas y notas ya registradas manualmente en Google Sheets
  const existentes = {};
  for (let i = 1; i < dataActual.length; i++) {
    const r = dataActual[i];
    const id = parseInt(r[0]);
    if (!id) continue;
    const iniEx = formatSheetCellDateStr(r[4]);
    const finEx = formatSheetCellDateStr(r[5]);
    const notaEx = isLegacy ? String(r[17] || "").trim() : String(r[13] || "").trim();
    existentes[id] = { ini: iniEx, fin: finEx, nota: notaEx };
  }

  // 3. Indexar ubicaciones por ambiente desde CENSO_CAJAS_FOTOS
  const sheetCajas = getOrCreateSheetCajas(ss);
  const dataCajas = sheetCajas.getDataRange().getValues();
  const ubicMap = {};
  for (let i = 1; i < dataCajas.length; i++) {
    const oId = parseInt(dataCajas[i][1]);
    if (!oId) continue;
    const amb = String(dataCajas[i][3] || "").trim() || "Ambiente";
    const cod = String(dataCajas[i][2] || "").trim() || "Caja";
    if (!ubicMap[oId]) ubicMap[oId] = {};
    if (!ubicMap[oId][amb]) ubicMap[oId][amb] = {};
    ubicMap[oId][amb][cod] = true;
  }

  // 4. Reescribir MATRIZ_69_OBRAS con el nuevo esquema limpio y clasificable
  sheetObras.clear();
  const headers = [
    "OBRA_N°", "DESCRIPCION_PROYECTO", "CUI", "PROVINCIA",
    "FECHA_INICIO_REAL", "FECHA_CULMINACION_REAL",
    "TIPO_PROYECTO", "CONVENIO", "STATUS_EXCEL", "SITUACION_ACTUAL",
    "PTO_PRIORIZADO", "PTO_EJECUTADO", "UBICACION_AMBIENTES",
    "NOTAS_AUDITORIA", "ULTIMA_ACTUALIZACION"
  ];
  sheetObras.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground("#1E293B").setFontColor("#FFFFFF").setFontWeight("bold").setWrap(true);
  sheetObras.setFrozenRows(1);

  const obras = Array.isArray(payload.obras) ? payload.obras : [];
  const rows = [];
  const bgColors = [];
  const stamp = new Date().toLocaleString();

  obras.forEach((o, idx) => {
    const id = parseInt(o.obra_id);
    const ex = existentes[id] || {};
    const fIni = ex.ini || o.fecha_inicio || "";
    const fFin = ex.fin || o.fecha_culm || "";
    const nota = ex.nota || o.notas_auditoria || "";

    let ubicStr = o.ubicacion_ambientes || "";
    if (ubicMap[id]) {
      ubicStr = Object.keys(ubicMap[id]).map(amb => `${amb} (${Object.keys(ubicMap[id][amb]).length})`).join(" | ");
    }

    rows.push([
      id,
      o.desc || "",
      o.cui || "S/N",
      o.provincia || "",
      fIni,
      fFin,
      o.es_obra_civil || "ADQUISICIÓN / OTROS",
      o.convenio || "",
      o.estatus || "",
      o.estado_situacional || "",
      Number(o.prio || 0),
      Number(o.costo_ejec || 0),
      ubicStr,
      nota,
      stamp
    ]);

    const baseBg = (o.es_obra_civil === "OBRA CIVIL OBJETIVO") ? ((idx % 2 === 1) ? "#F8FAFC" : "#FFFFFF") : "#F1F5F9";
    const rowBg = new Array(headers.length).fill(baseBg);
    if (o.es_obra_civil === "OBRA CIVIL OBJETIVO") {
      rowBg[6] = "#DBEAFE";
    }
    bgColors.push(rowBg);
  });

  if (rows.length > 0) {
    const rng = sheetObras.getRange(2, 1, rows.length, headers.length);
    rng.setValues(rows).setWrap(true).setVerticalAlignment("middle").setBackgrounds(bgColors);
    sheetObras.getRange(2, 5, rows.length, 2).setNumberFormat("@");
    sheetObras.getRange(2, 11, rows.length, 2).setNumberFormat("#,##0.00");
  }

  return { totalObras: rows.length, backupSheet: backupName };
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
      "TIPO_PROYECTO", "CONVENIO", "STATUS_EXCEL", "SITUACION_ACTUAL",
      "PTO_PRIORIZADO", "PTO_EJECUTADO", "UBICACION_AMBIENTES",
      "NOTAS_AUDITORIA", "ULTIMA_ACTUALIZACION"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground("#1E293B").setFontColor("#FFFFFF").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function formatearImpresionInforme() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_INFORME);
  if (!sheet) return;
  sheet.setFrozenRows(4);
  sheet.autoResizeRows(5, Math.max(1, sheet.getLastRow() - 4));
}

/**
 * Crea o regenera la pestaña INFORME_SITUACIONAL_OBRAS combinando las obras civiles
 * enviadas desde el visor web con el censo en vivo de CENSO_CAJAS_FOTOS (por Ambiente).
 */
function crearHojaInformeSituacional(ss, payload) {
  const sheetCajas = getOrCreateSheetCajas(ss);
  const dataCajas = sheetCajas.getDataRange().getValues();

  // 1. Indexar cajas y documentos sueltos por OBRA_ID y AMBIENTE en vivo
  const censoPorObra = {};
  for (let i = 1; i < dataCajas.length; i++) {
    const r = dataCajas[i];
    const oId = parseInt(r[1]);
    if (!oId) continue;
    const cod = String(r[2] || "").trim() || "SIN_CODIGO";
    const amb = String(r[3] || "").trim() || "AMBIENTE";
    const cont = String(r[4] || "").trim();
    const ocr = String(r[8] || "").trim();
    if (!censoPorObra[oId]) censoPorObra[oId] = {};
    if (!censoPorObra[oId][amb]) censoPorObra[oId][amb] = {};
    if (!censoPorObra[oId][amb][cod]) censoPorObra[oId][amb][cod] = [];
    if (cont && cont !== cod && cont !== "Contenido no especificado" && censoPorObra[oId][amb][cod].indexOf(cont) === -1) {
      censoPorObra[oId][amb][cod].push(cont);
    }
    if (ocr && censoPorObra[oId][amb][cod].length === 0) {
      censoPorObra[oId][amb][cod].push("OCR: " + ocr);
    }
  }

  let sheet = ss.getSheetByName(SHEET_INFORME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_INFORME, 0);
  } else {
    sheet.clear();
  }

  const headers = [
    "N°", "CONVENIO", "PROVINCIA", "CUI",
    "ACUERDO MUNICIPAL", "ACUERDO CONSEJO DIRECTIVO FDM",
    "ESTATUS DEL PY", "DESCRIPCIÓN DE PROYECTOS (OBRA CIVIL)",
    "FECHA INICIO", "FECHA DE TÉRMINO", "ESTADO SITUACIONAL ACTUAL",
    "PTO. PRIORIZADO (S/)", "AMPLIACIÓN (S/)", "COSTE TOTAL PY (S/)", "PTO. EJECUTADO (S/)",
    "UBICACIÓN POR AMBIENTES", "CONTENIDO DOCUMENTAL ENCONTRADO (CAJAS / SUELTOS)", "COMENTARIO / NOTA"
  ];
  const maxCols = headers.length;

  sheet.getRange(1, 1, 1, maxCols).merge()
    .setValue("FONDO DE DESARROLLO DE MOQUEGUA (FDM) — INFORME SITUACIONAL DE OBRAS CIVILES (2021 - 2025)")
    .setBackground("#0F172A").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(12)
    .setVerticalAlignment("middle").setHorizontalAlignment("left");
  sheet.setRowHeight(1, 28);

  sheet.getRange(2, 1, 1, maxCols).merge()
    .setValue("Actualizado al " + new Date().toLocaleString() + " • Para imprimir: Archivo -> Imprimir -> Horizontal -> Ajustar al ancho")
    .setBackground("#EFF6FF").setFontColor("#1E3A8A").setFontStyle("italic").setFontSize(9)
    .setVerticalAlignment("middle").setHorizontalAlignment("left");
  sheet.setRowHeight(2, 20);

  sheet.getRange(4, 1, 1, maxCols).setValues([headers])
    .setBackground("#1E3A8A").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(9)
    .setWrap(true).setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(4, 34);
  sheet.setFrozenRows(4);

  const obras = Array.isArray(payload.obras) ? payload.obras : [];
  const rows = [];
  const bgColors = [];
  const txtColors = [];

  obras.forEach((p, idx) => {
    const oId = parseInt(p.id);
    const liveAmb = censoPorObra[oId] || null;
    let ubicStr = p.ubicacion_ambientes || "";
    let contStr = p.contenido_cajas || "";

    if (liveAmb) {
      const ambNames = [];
      const ambParts = [];
      Object.keys(liveAmb).forEach(amb => {
        const codObj = liveAmb[amb];
        const codKeys = Object.keys(codObj);
        ambNames.push(amb + " (" + codKeys.length + ")");
        const lines = codKeys.map(cod => {
          const det = codObj[cod].join("; ");
          return cod + (det ? ": " + det : "");
        });
        ambParts.push("[" + amb + "] " + lines.join(" | "));
      });
      ubicStr = ambNames.join(" | ");
      contStr = ambParts.join("\n");
    }

    const est = String(p.estatus || "").toUpperCase();
    const sit = String(p.estado_situacional || "").toUpperCase();
    const ini = p.fecha_inicio || "";
    const fin = p.fecha_termino || "";
    const isEjec = est === "EJECUTADO" || sit.indexOf("LIQUID") !== -1;

    rows.push([
      oId,
      p.convenio || "",
      p.provincia || "",
      p.cui || "S/N",
      p.acuerdo_municipal || "",
      p.acuerdo_fdm || "",
      est || "SIN STATUS",
      p.desc || "",
      ini || (isEjec ? "FALTA INICIO" : "—"),
      fin || (isEjec ? "FALTA FIN" : "—"),
      sit || "—",
      Number(p.prio || 0),
      Number(p.ampliacion || 0),
      Number(p.costo_total || 0),
      Number(p.costo_ejec || 0),
      ubicStr || (isEjec ? "FALTA CENSAR (PRIORIDAD)" : "Por censar"),
      contStr || (isEjec ? "Pendiente ubicar cajas en campo" : "—"),
      p.comentario || ""
    ]);

    const baseBg = (idx % 2 === 1) ? "#F8FAFC" : "#FFFFFF";
    const rBg = new Array(maxCols).fill(baseBg);
    const rFg = new Array(maxCols).fill("#0F172A");

    // Col 7 ESTATUS
    if (est === "EJECUTADO") { rBg[6] = "#DCFCE7"; rFg[6] = "#166534"; }
    else if (est.indexOf("EN EJECUCION") !== -1) { rBg[6] = "#FEF3C7"; rFg[6] = "#B45309"; }
    else if (est.indexOf("CAMBIO") !== -1 || est.indexOf("EJECUTADO POR") !== -1) { rBg[6] = "#FFEDD5"; rFg[6] = "#C2410C"; }
    else if (est.indexOf("OBSERV") !== -1 || est.indexOf("FALTA") !== -1) { rBg[6] = "#FEE2E2"; rFg[6] = "#B91C1C"; }
    else { rBg[6] = "#F1F5F9"; rFg[6] = "#475569"; }

    // Col 11 SITUACION
    if (sit === "LIQUIDADO") { rBg[10] = "#059669"; rFg[10] = "#FFFFFF"; }
    else if (sit.indexOf("LIQUIDACION") !== -1 || sit.indexOf("FINALIZAD") !== -1) { rBg[10] = "#E0F2FE"; rFg[10] = "#0369A1"; }

    // Alerta fechas / censo
    if (!ini && isEjec) { rBg[8] = "#FEE2E2"; rFg[8] = "#B91C1C"; }
    if (!fin && isEjec) { rBg[9] = "#FEE2E2"; rFg[9] = "#B91C1C"; }
    if (!ubicStr && isEjec) { rBg[15] = "#FEE2E2"; rFg[15] = "#B91C1C"; }
    else if (ubicStr) { rBg[15] = "#DCFCE7"; rFg[15] = "#166534"; }

    bgColors.push(rBg);
    txtColors.push(rFg);
  });

  if (rows.length > 0) {
    const startRow = 5;
    const rng = sheet.getRange(startRow, 1, rows.length, maxCols);
    rng.setValues(rows)
      .setWrap(true)
      .setVerticalAlignment("middle")
      .setFontSize(9)
      .setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
    rng.setBackgrounds(bgColors);
    rng.setFontColors(txtColors);

    sheet.getRange(startRow, 1, rows.length, 4).setHorizontalAlignment("center");
    sheet.getRange(startRow, 7, rows.length, 1).setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange(startRow, 9, rows.length, 3).setHorizontalAlignment("center");
    sheet.getRange(startRow, 12, rows.length, 4).setHorizontalAlignment("right").setNumberFormat("#,##0.00");

    // Fila de Totales
    const totRow = startRow + rows.length;
    sheet.getRange(totRow, 1, 1, 11).merge().setValue("TOTAL (" + rows.length + " OBRAS CIVILES)")
      .setHorizontalAlignment("right").setFontWeight("bold");
    for (let c = 12; c <= 15; c++) {
      const colLetter = String.fromCharCode(64 + c);
      sheet.getRange(totRow, c).setFormula(`=SUM(${colLetter}${startRow}:${colLetter}${totRow - 1})`)
        .setNumberFormat("#,##0.00").setFontWeight("bold").setHorizontalAlignment("right");
    }
    sheet.getRange(totRow, 1, 1, maxCols).setBackground("#E2E8F0").setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true, "#1E3A8A", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.setRowHeight(totRow, 24);
  }

  const widths = [45, 78, 85, 72, 130, 150, 115, 290, 92, 92, 125, 105, 95, 105, 105, 170, 270, 180];
  widths.forEach((w, idx) => sheet.setColumnWidth(idx + 1, w));

  return {
    totalObras: rows.length,
    sheetUrl: ss.getUrl() + "#gid=" + sheet.getSheetId()
  };
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
