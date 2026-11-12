/**
 * CAULE DUCK HUNT - Classement partage en ligne
 * A coller dans Google Sheets > Extensions > Apps Script, puis deployer en "Application Web".
 * Les scores sont stockes dans l'onglet "Scores" de la feuille.
 * (Fichier volontairement sans accents pour eviter tout probleme d'encodage.)
 */
var ONGLET = 'Scores';
var TOP_N = 10;
var SCORE_MAX = 150000;   // garde-fou contre les scores farfelus

function feuille_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var f = ss.getSheetByName(ONGLET);
  if (!f) {
    f = ss.insertSheet(ONGLET);
    f.appendRow(['Date', 'Nom', 'Score', 'Stages reussis', 'Boss vaincu']);
    f.setFrozenRows(1);
  }
  return f;
}

function top_() {
  var lignes = feuille_().getDataRange().getValues().slice(1).filter(function (r) {
    return r[1] !== '' && Number(r[2]) > 0;
  });
  lignes.sort(function (a, b) { return Number(b[2]) - Number(a[2]); });
  return lignes.slice(0, TOP_N).map(function (r) {
    return {
      n: String(r[1]),
      s: Number(r[2]),
      st: Number(r[3]) || 0,
      b: r[4] === 'OUI',
      d: (r[0] instanceof Date) ? Utilities.formatDate(r[0], 'Europe/Paris', 'dd/MM/yyyy') : ''
    };
  });
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
  var verrou = LockService.getScriptLock();
  verrou.waitLock(10000);
  try {
    var d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    // lettres A-Z + lettres accentuees majuscules (\u00C0-\u00DD), espace et tiret
    var nom = String(d.n || '').toUpperCase().replace(/[^A-Z\u00C0-\u00DD \-]/g, '').trim().slice(0, 10) || 'ANONYME';
    var score = Math.floor(Number(d.s));
    if (!(score > 0 && score <= SCORE_MAX)) return json_({ ok: false, erreur: 'score invalide', top: top_() });
    var stages = Math.max(0, Math.min(12, Math.floor(Number(d.st) || 0)));
    feuille_().appendRow([new Date(), nom, score, stages, d.b ? 'OUI' : 'NON']);
    return json_({ ok: true, top: top_() });
  } finally {
    verrou.releaseLock();
  }
}
