/**
 * CAULE HUNTER 2027 – Classement partagé en ligne
 * À coller dans Google Sheets > Extensions > Apps Script, puis déployer en "Application Web".
 * Les scores sont stockés dans l'onglet "Scores" de la feuille.
 */
const ONGLET = 'Scores';
const TOP_N = 10;
const SCORE_MAX = 150000;   // garde-fou contre les scores farfelus

function feuille_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let f = ss.getSheetByName(ONGLET);
  if (!f) {
    f = ss.insertSheet(ONGLET);
    f.appendRow(['Date', 'Nom', 'Score', 'Stages réussis', 'Boss vaincu']);
    f.setFrozenRows(1);
  }
  return f;
}

function top_() {
  const lignes = feuille_().getDataRange().getValues().slice(1)
    .filter(r => r[1] !== '' && Number(r[2]) > 0);
  lignes.sort((a, b) => Number(b[2]) - Number(a[2]));
  return lignes.slice(0, TOP_N).map(r => ({
    n: String(r[1]),
    s: Number(r[2]),
    st: Number(r[3]) || 0,
    b: r[4] === 'OUI',
    d: r[0] instanceof Date ? Utilities.formatDate(r[0], 'Europe/Paris', 'dd/MM/yyyy') : ''
  }));
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** Lecture du classement */
function doGet() {
  return json_({ ok: true, top: top_() });
}

/** Enregistrement d'un score */
function doPost(e) {
  const verrou = LockService.getScriptLock();
  verrou.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const nom = String(d.n || '').toUpperCase().replace(/[^A-ZÀ-Ý \-]/g, '').trim().slice(0, 10) || 'ANONYME';
    const score = Math.floor(Number(d.s));
    if (!(score > 0 && score <= SCORE_MAX)) return json_({ ok: false, erreur: 'score invalide', top: top_() });
    const stages = Math.max(0, Math.min(5, Math.floor(Number(d.st) || 0)));
    feuille_().appendRow([new Date(), nom, score, stages, d.b ? 'OUI' : 'NON']);
    return json_({ ok: true, top: top_() });
  } finally {
    verrou.releaseLock();
  }
}
