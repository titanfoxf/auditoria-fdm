/**
 * ==============================================================================
 * SISTEMA CENTRALIZADO DE AUDITORÍA FDM - BACKEND GOOGLE DRIVE & SHEETS
 * ==============================================================================
 * 
 * INSTRUCCIONES DE INSTALACIÓN (5 MINUTOS):
 * 1. Entra a https://sheets.new para crear una nueva Hoja de Cálculo en Google Drive.
 *    Nómbrala: "AUDITORIA_FDM_REGISTRO_CENTRAL"
 * 2. En el menú superior de la hoja, haz clic en:
 *    Extensiones > Apps Script
 * 3. Borra todo el código que aparezca y pega este código completo.
 * 4. Arriba a la derecha, haz clic en el botón azul: "Implementar" > "Nueva implementación".
 * 5. Configura:
 *    - Tipo: Aplicación web
 *    - Descripción: Backend Auditoría FDM
 *    - Ejecutar como: "Yo" (tu cuenta)
 *    - Quién tiene acceso: "Cualquier usuario" (Anyone)  <--- ¡MUY IMPORTANTE!
 * 6. Haz clic en "Implementar", autoriza los permisos de Google.
 * 7. Copia la "URL de la aplicación web" (termina en /exec) y pégala en el chat.
 * ==============================================================================
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
    const headers = data[0];
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
