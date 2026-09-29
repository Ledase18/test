/**
 * Pilotage du fichier "Suivi des dépenses" (onglet Situation_).
 * Script lié au Google Sheets : Extensions > Apps Script.
 * Les zones (charges fixes, dépenses, revenus) sont détectées via les formules SUM du fichier,
 * donc insérer une ligne dans une zone ne casse rien.
 */
const SHEET_NAME = 'Situation_';
const MONTH_NAMES = ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre'];
const ANNUAL_FIRST_ROW = 9; // colonnes G (mois), H (libellé), I (montant)

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Comptes')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}

function num_(v) {
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).replace(/\s/g, '').replace(',', '.'));
  return isFinite(n) ? n : 0;
}

function blank_(v) { return v === '' || v === null || v === undefined; }

function norm_(s) {
  return String(s).trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
}

function layout_(sh) {
  const L = {
    fixe: { from: 5, to: 23, total: 24 },
    dep: { from: 25, to: 37, total: 38 },
    rev: { from: 39, to: 41, total: 42 }
  };
  const last = Math.min(sh.getLastRow(), 300);
  const dF = sh.getRange(1, 4, last, 1).getFormulas();
  const eF = sh.getRange(1, 5, last, 1).getFormulas();
  dF.forEach((r, i) => {
    const m = /^=SUM\(D(\d+):D(\d+)\)$/i.exec(r[0]);
    if (m) L.fixe = { from: +m[1], to: +m[2], total: i + 1 };
  });
  eF.forEach((r, i) => {
    const m = /^=SUM\(D(\d+):E(\d+)\)$/i.exec(r[0]);
    if (m) L.dep = { from: +m[1], to: +m[2], total: i + 1 };
    if (/^=D\d+\+/i.test(r[0])) {
      const ns = (r[0].match(/D(\d+)/g) || []).map(x => +x.slice(1));
      L.rev = { from: Math.min.apply(null, ns), to: Math.max.apply(null, ns), total: i + 1 };
    }
  });
  return L;
}

function getState() {
  const sh = sheet_();
  const L = layout_(sh);
  const block = (z, c1, c2) => sh.getRange(z.from, c1, z.to - z.from + 1, c2 - c1 + 1).getValues();

  const base = sh.getRange('B2:B4').getValues();
  const solde = num_(base[0][0]), ca = num_(base[1][0]), pret = num_(base[2][0]);

  const fixes = [];
  let fixePending = 0;
  block(L.fixe, 3, 5).forEach((r, i) => {
    if (blank_(r[0]) && blank_(r[1]) && blank_(r[2])) return;
    const p = num_(r[1]);
    fixePending += p;
    fixes.push({ row: L.fixe.from + i, label: String(r[0]), pending: p, model: num_(r[2]) });
  });

  const cats = [], ops = [];
  let depTotal = 0;
  block(L.dep, 2, 5).forEach((r, i) => {
    const row = L.dep.from + i;
    if (blank_(r[0]) && blank_(r[1])) return;
    const label = String(r[0]) + (blank_(r[1]) ? '' : ' – ' + r[1]);
    cats.push({ row, label });
    const d = num_(r[2]), e = num_(r[3]);
    depTotal += d + e;
    if (d !== 0) ops.push({ row, slot: 1, label: String(r[0]), amount: d });
    if (e !== 0) ops.push({ row, slot: 2, label, amount: e });
  });

  const revs = [];
  let revTotal = 0;
  block(L.rev, 2, 5).forEach((r, i) => {
    const d = num_(r[2]), e = num_(r[3]);
    revTotal += d + e;
    if (!blank_(r[0])) revs.push({ row: L.rev.from + i, label: String(r[0]), amount: d, extra: e });
  });

  const annual = [];
  sh.getRange(ANNUAL_FIRST_ROW, 7, L.rev.to - ANNUAL_FIRST_ROW + 1, 3).getValues().forEach((r, i) => {
    if (blank_(r[0]) && blank_(r[1]) && blank_(r[2])) return;
    annual.push({
      row: ANNUAL_FIRST_ROW + i,
      month: String(r[0]).trim(),
      monthIdx: MONTH_NAMES.map(norm_).indexOf(norm_(r[0])),
      label: String(r[1]),
      amount: num_(r[2])
    });
  });

  const month = Number(Utilities.formatDate(new Date(), sh.getParent().getSpreadsheetTimeZone(), 'M')) - 1;
  const reste = solde + ca - pret - depTotal - revTotal - fixePending;
  return { solde, ca, pret, fixes, fixePending, cats, ops, depTotal, revs, revTotal, annual, month, reste };
}

/* ---------- écriture ---------- */

function run_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sheet_();
    fn(sh, layout_(sh));
    SpreadsheetApp.flush();
    return getState();
  } finally {
    lock.releaseLock();
  }
}

function amt_(a, allowZero) {
  if (typeof a !== 'number' || !isFinite(a) || a < 0 || (a === 0 && !allowZero)) throw new Error('Montant invalide.');
  return a;
}

function zone_(row, z) {
  if (!(row >= z.from && row <= z.to)) throw new Error('Ligne hors zone.');
}

function saveBase(o) {
  return run_(sh => {
    sh.getRange('B2').setValue(amt_(o.solde, true));
    sh.getRange('B3').setValue(amt_(o.ca, true));
    sh.getRange('B4').setValue(amt_(o.pret, true));
  });
}

function addExpense(row, amount) {
  return run_((sh, L) => {
    zone_(row, L.dep);
    amount = amt_(amount);
    const r = sh.getRange(row, 4, 1, 2).getValues()[0];
    if (num_(r[0]) === 0) sh.getRange(row, 4).setValue(amount);
    else if (num_(r[1]) === 0) sh.getRange(row, 5).setValue(amount);
    else throw new Error('Les deux cases de cette ligne sont déjà occupées : pointe-en une ou choisis une autre ligne.');
  });
}

function clearExpense(row, slot) {
  return run_((sh, L) => {
    zone_(row, L.dep);
    sh.getRange(row, slot === 2 ? 5 : 4).setValue(0);
  });
}

function addFixe(label, amount, thisMonth) {
  return run_((sh, L) => {
    label = String(label || '').trim();
    if (!label) throw new Error('Libellé manquant.');
    amount = amt_(amount);
    const v = sh.getRange(L.fixe.from, 3, L.fixe.to - L.fixe.from + 1, 3).getValues();
    const i = v.findIndex(r => blank_(r[0]) && blank_(r[1]) && blank_(r[2]));
    if (i < 0) throw new Error("Plus de ligne libre dans les charges : insère une ligne dans le fichier (au-dessus de la dernière ligne de la zone).");
    const row = L.fixe.from + i;
    sh.getRange(row, 3).setValue(label);
    sh.getRange(row, 5).setValue(amount);
    if (thisMonth) sh.getRange(row, 4).setValue(amount);
  });
}

function payFixe(row, paid) {
  return run_((sh, L) => {
    zone_(row, L.fixe);
    sh.getRange(row, 4).setValue(paid ? '' : num_(sh.getRange(row, 5).getValue()));
  });
}

function editFixe(row, amount) {
  return run_((sh, L) => {
    zone_(row, L.fixe);
    amount = amt_(amount, true);
    sh.getRange(row, 5).setValue(amount);
    if (num_(sh.getRange(row, 4).getValue()) !== 0) sh.getRange(row, 4).setValue(amount);
  });
}

function deleteFixe(row) {
  return run_((sh, L) => {
    zone_(row, L.fixe);
    sh.getRange(row, 3, 1, 3).clearContent();
  });
}

function setRevenu(row, amount) {
  return run_((sh, L) => {
    zone_(row, L.rev);
    sh.getRange(row, 4).setValue(amt_(amount, true));
  });
}

/** Début de mois : recopie E -> D (charges fixes) et efface les revenus mis de côté. */
function newMonth() {
  return run_((sh, L) => {
    const n = L.fixe.to - L.fixe.from + 1;
    const e = sh.getRange(L.fixe.from, 5, n, 1).getValues();
    sh.getRange(L.fixe.from, 4, n, 1).setValues(e);
    sh.getRange(L.rev.from, 4, L.rev.to - L.rev.from + 1, 2).clearContent();
  });
}

function addAnnual(monthIdx, label, amount) {
  return run_((sh, L) => {
    label = String(label || '').trim();
    if (!label || !(monthIdx >= 0 && monthIdx < 12)) throw new Error('Mois ou libellé manquant.');
    amount = amt_(amount);
    const n = L.rev.to - ANNUAL_FIRST_ROW + 1;
    const v = sh.getRange(ANNUAL_FIRST_ROW, 7, n, 3).getValues();
    const i = v.findIndex(r => r.every(blank_));
    if (i < 0) throw new Error('Plus de ligne libre dans les charges annuelles.');
    sh.getRange(ANNUAL_FIRST_ROW + i, 7, 1, 3).setValues([[MONTH_NAMES[monthIdx], label, amount]]);
  });
}

function deleteAnnual(row) {
  return run_((sh, L) => {
    if (!(row >= ANNUAL_FIRST_ROW && row <= L.rev.to)) throw new Error('Ligne hors zone.');
    sh.getRange(row, 7, 1, 3).clearContent();
  });
}
