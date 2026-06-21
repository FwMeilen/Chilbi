// ============================================================
// Chilbi Herrliberg – Schichtplanung Backend
// Google Apps Script  |  Cl1.017
// Schema Konfiguration: ID|Datum|Von|Bis|Schicht|Aufgabe|Max Personen|Farbe|Informationen|Geschlossen
// Schema Anmeldungen:   Tag|Name|Schicht|Aufgabe|Timestamp
// Schema Tage:          Datum|Typ
// ============================================================

const SHEET_ID  = '1XqTNfgONmHX9GvmfOVvb94BUo_oQ04Uy7R97FVfdWyo';
const ADMIN_PW  = 'chilbi2025';
const SS        = SpreadsheetApp.openById(SHEET_ID);
const SH_CONFIG = 'Konfiguration';
const SH_SIGNUP = 'Anmeldungen';
const SH_TAGE   = 'Tage';

function doGet(e) {
  return jsonResponse({ config: getConfig(), signups: getSignups(), tage: getTage() });
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    if (p.action === 'signup')     return jsonResponse(saveSignup(p));
    if (p.action === 'unsignup')   return jsonResponse(removeSignup(p));
    if (p.action === 'editSignup') return jsonResponse(editSignup(p));
    if (p.action === 'saveConfig') {
      if (p.password !== ADMIN_PW) return jsonResponse({ ok: false, error: 'Falsches Passwort' });
      return jsonResponse(saveConfig(p.rows));
    }
    if (p.action === 'saveTage') {
      if (p.password !== ADMIN_PW) return jsonResponse({ ok: false, error: 'Falsches Passwort' });
      return jsonResponse(saveTage(p.rows));
    }
    return jsonResponse({ ok: false, error: 'Unbekannte Aktion' });
  } catch(err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

function getConfig() {
  const sh = SS.getSheetByName(SH_CONFIG);
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  if (rows.length < 2) return [];
  const headers = rows[0];
  return rows.slice(1).filter(r => r[0] !== '').map(r => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });
}

function getSignups() {
  const sh = SS.getSheetByName(SH_SIGNUP);
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  if (rows.length < 2) return [];
  const headers = rows[0];
  return rows.slice(1).filter(r => r[0] !== '').map(r => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });
}

function getTage() {
  const sh = SS.getSheetByName(SH_TAGE);
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  if (rows.length < 2) return [];
  const headers = rows[0];
  return rows.slice(1).filter(r => r[0] !== '').map(r => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });
}

function saveSignup(p) {
  const sh = SS.getSheetByName(SH_SIGNUP);
  sh.appendRow([p.tag, p.name, p.schicht, p.aufgabe, new Date()]);
  return { ok: true };
}

function removeSignup(p) {
  const sh = SS.getSheetByName(SH_SIGNUP);
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(p.tag) && String(rows[i][1]) === String(p.name)) {
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Eintrag nicht gefunden' };
}

function editSignup(p) {
  const sh = SS.getSheetByName(SH_SIGNUP);
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(p.oldTag) && String(rows[i][1]) === String(p.oldName)) {
      sh.getRange(i + 1, 1, 1, 5).setValues([[p.tag, p.name, p.schicht, p.aufgabe, rows[i][4]]]);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Eintrag nicht gefunden' };
}

function saveConfig(rows) {
  const sh = SS.getSheetByName(SH_CONFIG);
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, 10).clearContent();
  if (rows.length > 0) {
    const data = rows.map(r => [
      r.Tag, r.Datum || r.TagLabel, r.Von, r.Bis,
      r.Schicht, r.Aufgabe, r.MaxPersonen, r.Farbe, r.Informationen || '', r.Geschlossen || '0'
    ]);
    sh.getRange(2, 1, data.length, 10).setValues(data);
  }
  return { ok: true };
}

function saveTage(rows) {
  let sh = SS.getSheetByName(SH_TAGE);
  if (!sh) {
    sh = SS.insertSheet(SH_TAGE);
    sh.getRange(1, 1, 1, 2).setValues([['Datum', 'Typ']]);
  }
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, 2).clearContent();
  if (rows.length > 0) {
    const data = rows.map(r => [r.Datum, r.Typ]);
    sh.getRange(2, 1, data.length, 2).setValues(data);
  }
  return { ok: true };
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
