/**
 * Comptador de reproduccions - Ràdio Institut Sabadell
 * Enganxa aquest codi a Extensions > Apps Script del full de càlcul.
 *
 *   GET ?action=get          -> {"1": 12, "2": 5, ...}
 *   GET ?action=hit&id=3     -> {"id": "3", "count": 13}
 */
const SHEET_NAME = 'Reproduccions';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['id', 'reproduccions']);
  }
  return sh;
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'get';
  const sh = getSheet_();

  if (action === 'hit') {
    const id = String(e.parameter.id || '').trim();
    // Només ids numèrics curts, per evitar que s'omplin files amb brossa
    if (!/^\d{1,4}$/.test(id)) return json_({ error: 'id invàlid' });

    const lock = LockService.getScriptLock();
    lock.waitLock(10000); // evita perdre comptes si dues persones fan play alhora
    try {
      const data = sh.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === id) {
          const n = Number(data[i][1] || 0) + 1;
          sh.getRange(i + 1, 2).setValue(n);
          return json_({ id: id, count: n });
        }
      }
      sh.appendRow([id, 1]);
      return json_({ id: id, count: 1 });
    } finally {
      lock.releaseLock();
    }
  }

  // action=get: retorna tots els comptadors
  const data = sh.getDataRange().getValues();
  const counts = {};
  for (let i = 1; i < data.length; i++) {
    counts[String(data[i][0])] = Number(data[i][1] || 0);
  }
  return json_(counts);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
