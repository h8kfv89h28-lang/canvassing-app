/**
 * Canvassing – raccolta esiti in tempo reale su Foglio Google.
 * Incolla tutto questo testo in un nuovo progetto Apps Script (script.new),
 * poi Esegui il deployment › Nuovo deployment › App web
 * (Esegui come: Me · Chi ha accesso: Chiunque) e copia l'URL che termina con /exec.
 * Al primo collegamento lo script crea nel tuo Drive il foglio "Esiti canvassing".
 */
const SHEET = 'Esiti';
const HEAD = ['Aggiornato', 'Giornata', 'Coppia', 'Cliente', 'Insegna', 'Indirizzo', 'Lista', 'Esito', 'Nota', 'Ora', 'Telefono', 'ID cliente', 'ID piano', 'ts'];

function book_() {
  // Se lo script è dentro un foglio usa quello, altrimenti crea (una volta sola) il foglio "Esiti canvassing" nel tuo Drive.
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (ss) return ss;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  ss = SpreadsheetApp.create('Esiti canvassing');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function sheet_() {
  const ss = book_();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold');
  }
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const items = [].concat(JSON.parse(e.postData.contents));
    const sh = sheet_();
    const data = sh.getDataRange().getValues();
    const idx = {};
    for (let r = 1; r < data.length; r++) idx[data[r][12] + '|' + data[r][11]] = r + 1;
    items.forEach(function (d) {
      if (!d || !d.plan || !d.id) return;
      const key = d.plan + '|' + d.id;
      const row = [new Date(), d.date || '', d.team || '', d.rs || '', d.ins || '', d.addr || '', d.src || '', d.esito || '', d.nota || '', d.at || '', d.dev || '', d.id, d.plan, d.ts || Date.now()];
      if (idx[key]) {
        const old = sh.getRange(idx[key], 14).getValue();
        if (Number(old) > Number(d.ts || 0)) return; // dato più vecchio: ignoro
        sh.getRange(idx[key], 1, 1, row.length).setValues([row]);
      } else {
        sh.appendRow(row);
        idx[key] = sh.getLastRow();
      }
    });
    return out_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const plan = e.parameter.plan, team = e.parameter.team;
  if (!plan) { sheet_(); return out_({ ok: true, info: 'Canvassing: collegamento attivo', sheet: book_().getUrl() }); }
  const data = sheet_().getDataRange().getValues();
  const rows = [];
  for (let r = 1; r < data.length; r++) {
    const d = data[r];
    if (String(d[12]) !== plan) continue;
    if (team && String(d[2]) !== team) continue;
    rows.push({ id: d[11], team: d[2], esito: d[7], nota: d[8], at: d[9], ts: Number(d[13]) || 0 });
  }
  return out_({ ok: true, rows: rows });
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
