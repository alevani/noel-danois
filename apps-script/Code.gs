// À coller dans la feuille Google : Extensions > Apps Script.
const SHEET_NAME = 'Réponses';
const LABELS = ['ven. 20 nov.', 'sam. 21 nov.', 'ven. 27 nov.', 'sam. 28 nov.', 'ven. 4 déc.', 'sam. 5 déc.', 'ven. 11 déc.', 'sam. 12 déc.', 'ven. 18 déc.', 'sam. 19 déc.'];
const DAYS = ['d1', 'd2', 'd3', 'd4', 'd5', 'd6', 'd7', 'd8', 'd9', 'd10'];
const CHOICES = { oui: 'Dispo', peutetre: 'Peut-être', non: 'Pas dispo' };
const COLUMNS = DAYS.length + 3; // id, nom, dates, mise à jour
const CODE_COLUMN = COLUMNS + 1; // rows_() stops before it, so codes never leave the sheet
const MAX_ROWS = 200;

function newCode_() {
  const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  const hex = Utilities.getUuid().replace(/-/g, '').slice(-16);
  let code = '';
  for (let i = 0; i < 16; i += 2) {
    code += alphabet[parseInt(hex.substr(i, 2), 16) % 32];
    if (i === 6) code += '-';
  }
  return code;
}

function writeCode_(sheet, row, code) {
  sheet.getRange(row, CODE_COLUMN).setNumberFormat('@').setValue(code);
}

function sheet_() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = book.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = book.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, COLUMNS).setNumberFormat('@')
      .setValues([['id', 'Nom'].concat(LABELS, ['Mis à jour'])]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  // First run with codes: add the column and give existing answers a code.
  if (sheet.getRange(1, CODE_COLUMN).getDisplayValue() !== 'Code') {
    sheet.getRange(1, CODE_COLUMN).setNumberFormat('@').setValue('Code').setFontWeight('bold');
    for (let row = 2; row <= sheet.getLastRow(); row++) {
      if (sheet.getRange(row, 1).getDisplayValue()) writeCode_(sheet, row, newCode_());
    }
  }
  return sheet;
}

function rows_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const codes = {};
  Object.keys(CHOICES).forEach(function (code) { codes[CHOICES[code]] = code; });
  return sheet.getRange(2, 1, last - 1, COLUMNS).getDisplayValues()
    .filter(function (row) { return row[0] && row[1]; })
    .map(function (row) {
      const answers = {};
      DAYS.forEach(function (day, i) {
        if (codes[row[i + 2]]) answers[day] = codes[row[i + 2]];
      });
      return { id: row[0], name: row[1], answers: answers, at: row[COLUMNS - 1] };
    });
}

function reply_(body) {
  return ContentService.createTextOutput(JSON.stringify(body)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return reply_({ ok: true, reponses: rows_(sheet_()) });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const input = JSON.parse(e.postData.contents);
    const id = String(input.id || '');
    if (!/^[a-z0-9-]{1,200}$/.test(id)) return reply_({ ok: false, error: 'id invalide' });

    const sheet = sheet_();
    const last = sheet.getLastRow();
    const ids = last > 1 ? sheet.getRange(2, 1, last - 1, 1).getDisplayValues().map(function (r) { return r[0]; }) : [];
    const index = ids.indexOf(id);
    if (index >= 0) {
      const stored = sheet.getRange(index + 2, CODE_COLUMN).getDisplayValue();
      const given = String(input.code || '').trim().toUpperCase();
      if (!stored || stored !== given) return reply_({ ok: false, error: 'code' });
    }
    let created = '';

    if (input.action === 'delete') {
      if (index >= 0) sheet.deleteRow(index + 2);
    } else if (input.action === 'set') {
      const name = String(input.name || '').trim().slice(0, 60);
      const answers = input.answers || {};
      const labels = DAYS.map(function (day) { return CHOICES[answers[day]]; });
      if (!name || labels.indexOf(undefined) >= 0) return reply_({ ok: false, error: 'réponse incomplète' });
      if (index < 0 && ids.length >= MAX_ROWS) return reply_({ ok: false, error: 'liste pleine' });
      const row = index >= 0 ? index + 2 : last + 1;
      // A leading apostrophe keeps a name like =1+1 as text: Sheets would run it as a formula.
      const safeName = /^[=+\-@']/.test(name) ? "'" + name : name;
      sheet.getRange(row, 1, 1, COLUMNS).setNumberFormat('@')
        .setValues([[id, safeName].concat(labels, [new Date().toISOString()])]);
      if (index < 0) {
        created = newCode_();
        writeCode_(sheet, row, created);
      }
    } else {
      return reply_({ ok: false, error: 'action inconnue' });
    }
    return reply_({ ok: true, code: created || undefined, reponses: rows_(sheet) });
  } catch (err) {
    return reply_({ ok: false, error: 'requête invalide' });
  } finally {
    lock.releaseLock();
  }
}
