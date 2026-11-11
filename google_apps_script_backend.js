/**
 * ==============================================================================
 * SISTEMA CENTRALIZADO DE AUDITORÍA FDM - BACKEND GOOGLE DRIVE & SHEETS
 * ==============================================================================
 * 
 * Versión 2.0: Con soporte para guardado, sincronización y BORRADO en Drive/Sheets.
 */

const FOLDER_NAME = "AUDITORIA_FDM_EVIDENCIAS_2022_2025";
const SHEET_NAME = "REGISTRO_CAMPO";

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = initSheet(ss);
    }

    const data = sheet.getDataRange().getValues();
    const records = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      records.push({
        timestamp: row[0],
        obra_id: parseInt(row[1]),
        desc: row[2],
        cui: row[3],
        codigo_caja: row[4],
        ubicacion: row[5],
        nota: row[6],
        drive_url: row[7],
        slot: row[8],
        fecha_inicio: row[9],
        fecha_culm: row[10],
        checks: row[11] ? JSON.parse(row[11]) : {}
      });
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      total_registros: records.length,
      records: records
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
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = initSheet(ss);
    }

    // ==========================================
    // CASO 1: ACCIÓN DE BORRADO DE REGISTRO Y FOTO
    // ==========================================
    if (payload.action === "delete_record") {
      const obraId = parseInt(payload.obra_id);
      const codigoCaja = String(payload.codigo_caja || "").trim();
      const driveUrl = String(payload.drive_url || "").trim();
      let fileId = payload.file_id || "";

      if (!fileId && driveUrl) {
        const m = driveUrl.match(/[-\w]{25,}/);
        if (m) fileId = m[0];
      }

      // 1. Borrar archivo de Google Drive (mover a papelera)
      let fileDeleted = false;
      if (fileId) {
        try {
          const file = DriveApp.getFileById(fileId);
          if (file) {
            file.setTrashed(true);
            fileDeleted = true;
          }
        } catch (e) {
          // Ignorar si el archivo ya no existe
        }
      }

      // 2. Borrar la fila correspondiente en Google Sheets
      const data = sheet.getDataRange().getValues();
      let rowDeleted = false;

      for (let i = data.length - 1; i >= 1; i--) {
        const row = data[i];
        const rowObraId = parseInt(row[1]);
        const rowCod = String(row[4] || "").trim();
        const rowUrl = String(row[7] || "").trim();

        const matchObra = (rowObraId === obraId);
        const matchCod = codigoCaja && (rowCod === codigoCaja);
        const matchUrl = driveUrl && (rowUrl === driveUrl);

        if (matchObra && (matchCod || matchUrl)) {
          if (!fileDeleted && rowUrl) {
            const m = rowUrl.match(/[-\w]{25,}/);
            if (m) {
              try {
                DriveApp.getFileById(m[0]).setTrashed(true);
                fileDeleted = true;
              } catch (e) {}
            }
          }
          sheet.deleteRow(i + 1);
          rowDeleted = true;
          break; // Borrar el registro más reciente que coincida
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        action: "delete_record",
        obra_id: obraId,
        codigo_caja: codigoCaja,
        row_deleted: rowDeleted,
        file_deleted: fileDeleted,
        message: "Registro eliminado de Google Sheets y Drive"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ==========================================
    // CASO 2: GUARDADO NORMAL DE REGISTRO Y FOTO
    // ==========================================
    const obraId = payload.obra_id;
    const desc = payload.desc || "";
    const cui = payload.cui || "";
    const codigoCaja = payload.codigo_caja || "";
    const ubicacion = payload.ubicacion || "";
    const nota = payload.nota || "";
    const slot = payload.slot || "rotulo";
    const fechaInicio = payload.fecha_inicio || "";
    const fechaCulm = payload.fecha_culm || "";
    const checks = payload.checks ? JSON.stringify(payload.checks) : "{}";

    let driveUrl = "";
    let fileId = "";

    // Si viene foto en base64, guardarla en Google Drive
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

    // Registrar fila en el Google Sheet
    sheet.appendRow([
      new Date().toLocaleString(),
      obraId,
      desc,
      cui,
      codigoCaja,
      ubicacion,
      nota,
      driveUrl,
      slot,
      fechaInicio,
      fechaCulm,
      checks
    ]);

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      obra_id: obraId,
      codigo_caja: codigoCaja,
      drive_url: driveUrl,
      file_id: fileId,
      message: "Registrado en Google Sheets y Google Drive"
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function initSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  const headers = [
    "TIMESTAMP", "OBRA_ID", "DESCRIPCION", "CUI", "CODIGO_CAJA", 
    "UBICACION_AMBIENTE", "CONTENIDO_NOTAS", "ENLACE_FOTO_DRIVE", 
    "SLOT", "FECHA_INICIO_REAL", "FECHA_CULMINACION_REAL", "CHECKS_DOCUMENTALES"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setBackground("#1E293B").setFontColor("#FFFFFF").setFontWeight("bold");
  sheet.setFrozenRows(1);
  return sheet;
}

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

function getOrCreateSubFolder(parentFolder, subFolderName) {
  const folders = parentFolder.getFoldersByName(subFolderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(subFolderName);
}
