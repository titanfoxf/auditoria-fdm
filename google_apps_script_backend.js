/**
 * ==============================================================================
 * SISTEMA CENTRALIZADO DE AUDITORÍA FDM - BACKEND GOOGLE DRIVE & SHEETS v3.0
 * ARQUITECTURA MODULAR RELACIONAL (2 HOJAS SEPARADAS):
 * 1. "MATRIZ_69_OBRAS": 1 fila por cada proyecto (Fechas, 11 Checks, Ubicación)
 * 2. "CENSO_CAJAS_FOTOS": Registro individual de cada caja física y foto en Drive
 * ==============================================================================
 */

const FOLDER_NAME = "AUDITORIA_FDM_EVIDENCIAS_2022_2025";
const SHEET_OBRAS = "MATRIZ_69_OBRAS";
const SHEET_CAJAS = "CENSO_CAJAS_FOTOS";

const CHECK_KEYS = [
  'chk_exp', 'chk_cont', 'chk_acta_ent', 'chk_cuad', 
  'chk_val', 'chk_adic', 'chk_sup', 'chk_rec', 
  'chk_liq', 'chk_res', 'chk_transf'
];

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetObras = getOrCreateSheetObras(ss);
    const sheetCajas = getOrCreateSheetCajas(ss);

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
        ubicacion: row[3],
        nota: row[4],
        drive_url: row[5],
        slot: row[6],
        file_id: row[7] || ""
      });
    }

    // 2. Leer Matriz de Obras y Checklist
    const dataObras = sheetObras.getDataRange().getValues();
    const obras = {};
    for (let i = 1; i < dataObras.length; i++) {
      const row = dataObras[i];
      const oId = parseInt(row[0]);
      if (!oId) continue;

      const checks = {};
      CHECK_KEYS.forEach((k, idx) => {
        checks[k] = (row[7 + idx] === "SÍ" || row[7 + idx] === true);
      });

      obras[oId] = {
        obra_id: oId,
        desc: row[1],
        cui: row[2],
        provincia: row[3],
        fecha_inicio: row[4],
        fecha_culm: row[5],
        ubicacion_compilada: row[6],
        checks: checks,
        notas: row[18] || ""
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
    // CASO 1: GUARDAR O ACTUALIZAR CHECKLIST & DATOS GENERALES DE OBRA
    // ==============================================================
    if (action === "save_obra_checks") {
      const sheetObras = getOrCreateSheetObras(ss);
      const obraId = parseInt(payload.obra_id);
      upsertObraRow(sheetObras, obraId, payload);

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        action: "save_obra_checks",
        obra_id: obraId,
        message: "Checklist y datos de obra actualizados en MATRIZ_69_OBRAS"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ==============================================================
    // CASO 2: GUARDAR FOTO DE CAJA (INVENTARIO FÍSICO)
    // ==============================================================
    if (action === "save_box_photo") {
      const sheetCajas = getOrCreateSheetCajas(ss);
      const sheetObras = getOrCreateSheetObras(ss);

      const obraId = parseInt(payload.obra_id);
      const desc = payload.desc || "";
      const codigoCaja = String(payload.codigo_caja || "").trim();
      const ubicacion = payload.ubicacion || "";
      const nota = payload.nota || "";
      const slot = payload.slot || "rotulo";

      let driveUrl = "";
      let fileId = "";

      // Guardar foto en Google Drive
      if (payload.photo && payload.photo.base64) {
        const rootFolder = getOrCreateFolder(FOLDER_NAME);
        const safeDesc = desc.replace(/[^a-zA-Z0-9_-]/g, '_').substring(0, 30);
        const obraFolderName = `OBRA_${String(obraId).padStart(2, '0')}_${safeDesc}`;
        const obraFolder = getOrCreateSubFolder(rootFolder, obraFolderName);

        const base64Data = payload.photo.base64.replace(/^data:image\/\w+;base64,/, "");
        const decodedBytes = Utilities.base64Decode(base64Data);
        const blob = Utilities.newBlob(decodedBytes, "image/jpeg", payload.photo.filename || `OBRA_${obraId}_${codigoCaja}.jpg`);
        
        const file = obraFolder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        driveUrl = file.getUrl();
        fileId = file.getId();
      }

      // 1. Agregar a CENSO_CAJAS_FOTOS
      sheetCajas.appendRow([
        new Date().toLocaleString(),
        obraId,
        codigoCaja,
        ubicacion,
        nota,
        driveUrl,
        slot,
        fileId
      ]);

      // 2. Actualizar también la ubicación compilada en la Matriz de Obras
      upsertObraRow(sheetObras, obraId, payload, true);

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        action: "save_box_photo",
        obra_id: obraId,
        codigo_caja: codigoCaja,
        drive_url: driveUrl,
        file_id: fileId,
        message: "Caja y foto registradas en CENSO_CAJAS_FOTOS"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ==============================================================
    // CASO 3: BORRADO DE CAJA Y FOTO
    // ==============================================================
    if (action === "delete_record") {
      const sheetCajas = getOrCreateSheetCajas(ss);
      const sheetObras = getOrCreateSheetObras(ss);

      const obraId = parseInt(payload.obra_id);
      const codigoCaja = String(payload.codigo_caja || "").trim();
      const driveUrl = String(payload.drive_url || "").trim();
      let fileId = payload.file_id || "";

      if (!fileId && driveUrl) {
        const m = driveUrl.match(/[-\w]{25,}/);
        if (m) fileId = m[0];
      }

      // Borrar archivo en Drive
      let fileDeleted = false;
      if (fileId) {
        try {
          const file = DriveApp.getFileById(fileId);
          if (file) {
            file.setTrashed(true);
            fileDeleted = true;
          }
        } catch (e) {}
      }

      // Borrar fila en CENSO_CAJAS_FOTOS
      const data = sheetCajas.getDataRange().getValues();
      let rowDeleted = false;
      for (let i = data.length - 1; i >= 1; i--) {
        const row = data[i];
        const rowObraId = parseInt(row[1]);
        const rowCod = String(row[2] || "").trim();
        const rowUrl = String(row[5] || "").trim();

        if (rowObraId === obraId && (rowCod === codigoCaja || (driveUrl && rowUrl === driveUrl))) {
          if (!fileDeleted && rowUrl) {
            const m = rowUrl.match(/[-\w]{25,}/);
            if (m) {
              try { DriveApp.getFileById(m[0]).setTrashed(true); fileDeleted = true; } catch (err) {}
            }
          }
          sheetCajas.deleteRow(i + 1);
          rowDeleted = true;
          break;
        }
      }

      // Actualizar conteo en Matriz de Obras
      upsertObraRow(sheetObras, obraId, payload, true);

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

// ==========================================
// FUNCIONES AUXILIARES DE HOJAS
// ==========================================
function upsertObraRow(sheetObras, obraId, payload, soloUbicacion = false) {
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
  const ubi = payload.ubicacion || "";
  const notas = payload.notas || payload.nota || "";
  const checks = payload.checks || {};

  const checkCols = CHECK_KEYS.map(k => checks[k] ? "SÍ" : "");

  if (targetRowIndex !== -1) {
    // Fila existente -> Actualizar
    if (soloUbicacion) {
      if (ubi) sheetObras.getRange(targetRowIndex, 7).setValue(ubi);
    } else {
      sheetObras.getRange(targetRowIndex, 1, 1, 20).setValues([[
        obraId, desc, cui, prov, fIni, fCulm, ubi,
        checkCols[0], checkCols[1], checkCols[2], checkCols[3],
        checkCols[4], checkCols[5], checkCols[6], checkCols[7],
        checkCols[8], checkCols[9], checkCols[10],
        notas, new Date().toLocaleString()
      ]]);
    }
  } else {
    // Fila nueva -> Insertar
    sheetObras.appendRow([
      obraId, desc, cui, prov, fIni, fCulm, ubi,
      checkCols[0], checkCols[1], checkCols[2], checkCols[3],
      checkCols[4], checkCols[5], checkCols[6], checkCols[7],
      checkCols[8], checkCols[9], checkCols[10],
      notas, new Date().toLocaleString()
    ]);
  }
}

function getOrCreateSheetObras(ss) {
  let sheet = ss.getSheetByName(SHEET_OBRAS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_OBRAS, 0);
    const headers = [
      "OBRA_N°", "DESCRIPCION_PROYECTO", "CUI", "PROVINCIA", 
      "FECHA_INICIO_REAL", "FECHA_CULMINACION_REAL", "UBICACION_FISICA_COMPILADA",
      "1_EXP_TECNICO", "2_CONTRATO", "3_ACTA_INICIO", "4_CUADERNO",
      "5_VALORIZACIONES", "6_ADICIONALES", "7_SUPERVISION", "8_ACTA_RECEPCION",
      "9_LIQ_OBRA", "10_RES_LIQ_FDM", "11_TRANSFERENCIA",
      "NOTAS_OBSERVACIONES", "ULTIMA_ACTUALIZACION"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground("#1E293B").setFontColor("#FFFFFF").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateSheetCajas(ss) {
  let sheet = ss.getSheetByName(SHEET_CAJAS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CAJAS, 1);
    const headers = [
      "TIMESTAMP", "OBRA_ID", "CODIGO_CAJA", "AMBIENTE", 
      "CONTENIDO_ROTULO", "ENLACE_FOTO_DRIVE", "SLOT", "FILE_ID_DRIVE"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground("#0F766E").setFontColor("#FFFFFF").setFontWeight("bold");
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
