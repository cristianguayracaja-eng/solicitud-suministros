/**
 * Backend tipo "clave-valor" para el sistema de Solicitud de Suministros.
 * Se ejecuta dentro de una Google Sheet (Extensiones > Apps Script) y se
 * publica como Web App. No requiere ninguna cuenta de quien llena los
 * formularios: solo tú necesitas tu cuenta de Google para instalarlo.
 *
 * Guarda cada "clave" del sistema (config, listas de artículos, tiendas,
 * solicitudes enviadas, etc.) como una fila en la hoja "KV" de este archivo.
 */

const SHEET_NAME = "KV";

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["key", "value"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

function findValue_(sheet, key) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return null;
}

function upsert_(sheet, key, value) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function deleteKey_(sheet, key) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.deleteRow(i + 1);
      return;
    }
  }
}

/** Peticiones de lectura: ?action=get&key=... o ?action=list&prefix=... */
function doGet(e) {
  const sheet = getSheet_();
  const action = e.parameter.action;

  if (action === "get") {
    const value = findValue_(sheet, e.parameter.key || "");
    return jsonResponse_({ value: value });
  }

  if (action === "list") {
    const prefix = e.parameter.prefix || "";
    const data = sheet.getDataRange().getValues();
    const keys = [];
    for (let i = 1; i < data.length; i++) {
      const k = data[i][0];
      if (typeof k === "string" && k.indexOf(prefix) === 0) keys.push(k);
    }
    return jsonResponse_({ keys: keys });
  }

  return jsonResponse_({ error: "acción desconocida" });
}

/** Peticiones de escritura: body JSON { action: "set"|"delete", key, value } */
function doPost(e) {
  const sheet = getSheet_();
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse_({ error: "JSON inválido" });
  }

  if (body.action === "set") {
    upsert_(sheet, body.key, body.value);
    return jsonResponse_({ ok: true });
  }

  if (body.action === "delete") {
    deleteKey_(sheet, body.key);
    return jsonResponse_({ ok: true });
  }

  return jsonResponse_({ error: "acción desconocida" });
}
