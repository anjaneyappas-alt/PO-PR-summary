// DC inward Damages – webhook for the Streamlit "Push to Google Sheet" button.
// Columns in Damage_Short:
// A Date | B vendor name | C PO NO | D Item | E Qty | F reason | G Status / Action |
// H Actioning from Category | I Timeline for Replacement | J Category | K PR done on

const SHEET_ID = '1j4kCgAbstRvkFevjusah-4erIuwtN_HbwEsY9cAc1HM';
const TZ = 'Asia/Kolkata';

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function norm(v) {
  return String(v == null ? '' : v).trim().toLowerCase().replace(/\s+/g, ' ');
}

// "20/11/2026", "20-11-2026" or a Date -> real Date (avoids locale mix-ups)
function toDate(v) {
  if (v instanceof Date) return v;
  const m = String(v || '').trim().match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
  if (!m) return v;
  let y = Number(m[3]); if (y < 100) y += 2000;
  return new Date(y, Number(m[2]) - 1, Number(m[1]));
}

function dateKey(v) {
  const d = toDate(v);
  return d instanceof Date ? Utilities.formatDate(d, TZ, 'dd/MM/yyyy') : norm(v);
}

function rowKey(date, po, item, qty) {
  return [dateKey(date), norm(po), norm(item), String(Number(qty) || 0)].join('|');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return out({ status: 'error', message: 'Sheet busy, try again' });
  try {
    const body = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(body.sheet_name || 'Damage_Short');
    if (!sh) return out({ status: 'error', message: 'Tab not found: ' + body.sheet_name });

    // Header for the extra column
    if (!sh.getRange(1, 11).getValue()) sh.getRange(1, 11).setValue('PR done on');

    // Last row that really has an Item (column D) – ignores empty/formatted rows
    const colD = sh.getRange(1, 4, sh.getMaxRows(), 1).getValues();
    let last = 1;
    for (let i = colD.length - 1; i >= 0; i--) {
      if (String(colD[i][0]).trim()) { last = i + 1; break; }
    }

    // Keys of rows already in the sheet (read live, so deleted rows can be re-pushed)
    const existing = new Set();
    if (last > 1) {
      sh.getRange(2, 1, last - 1, 5).getValues()
        .forEach(r => { if (r[3]) existing.add(rowKey(r[0], r[2], r[3], r[4])); });
    }

    const toAdd = [];
    const prDoneOn = [];
    let skipped = 0;
    (body.rows || []).forEach(r => {
      const k = rowKey(r['Date'], r['PO NO'], r['Item'], r['Qty']);
      if (existing.has(k)) { skipped++; return; }
      existing.add(k);
      // A-F only. G-I are for the team, J (Category) is an ARRAYFORMULA in J2 –
      // writing anything into J breaks it (#REF!), so it is never touched.
      toAdd.push([
        toDate(r['Date']),
        r['vendor name'] || '',
        r['PO NO'] || '',
        r['Item'] || '',
        Number(r['Qty']) || r['Qty'] || '',
        r['reason'] || ''
      ]);
      prDoneOn.push([r['PR done on'] || '']);         // K
    });

    if (toAdd.length) {
      const start = last + 1;
      const need = start + toAdd.length - 1 - sh.getMaxRows();
      if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need);
      sh.getRange(start, 1, toAdd.length, 6).setValues(toAdd);
      sh.getRange(start, 11, prDoneOn.length, 1).setValues(prDoneOn);
      sh.getRange(start, 1, toAdd.length, 1).setNumberFormat('dd/MM/yyyy');
    }
    return out({ status: 'success', appended: toAdd.length, skipped: skipped });
  } catch (err) {
    return out({ status: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Open the web app URL in a browser to confirm the deployment is live
function doGet() {
  return out({ status: 'ok', message: 'Damage webhook is live' });
}
