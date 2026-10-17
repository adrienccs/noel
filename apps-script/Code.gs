/**
 * Tirage de Noël — stockage des listes d'idées dans ce Google Sheet.
 *
 * Installation :
 *  1. Extensions > Apps Script, colle ce fichier à la place de Code.gs, enregistre.
 *  2. Déployer > Nouveau déploiement > type « Application Web »
 *     - Exécuter en tant que : Moi
 *     - Qui a accès : Tout le monde
 *  3. Autorise l'accès, puis copie l'URL qui finit par /exec.
 *
 * Feuille « Listes » : Famille | Prénom | Idées | Mise à jour | Clé
 * La clé empêche de modifier la liste de quelqu'un d'autre.
 */

var SHEET = 'Listes';
var MAX_IDEAS = 2000;

function sheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(['Famille', 'Prénom', 'Idées', 'Mise à jour', 'Clé']);
    sh.setFrozenRows(1);
    sh.getRange('A1:E1').setFontWeight('bold');
    sh.setColumnWidth(3, 420);
    sh.getRange('C:C').setWrap(true);
  }
  return sh;
}

function clean_(s, max) {
  return String(s == null ? '' : s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim().slice(0, max);
}

function findRow_(sh, fam, name) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var vals = sh.getRange(2, 1, last - 1, 2).getValues();
  var f = fam.toLowerCase(), n = name.toLowerCase();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).toLowerCase() === f && String(vals[i][1]).toLowerCase() === n) return i + 2;
  }
  return -1;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Lecture : ?action=get&f=Famille&n=Prénom */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'ping') return json_({ ok: true });
  if (p.action !== 'get') return json_({ ok: false, error: 'action inconnue' });
  var fam = clean_(p.f, 60), name = clean_(p.n, 40);
  if (!name) return json_({ ok: false, error: 'prénom manquant' });
  var sh = sheet_();
  var row = findRow_(sh, fam, name);
  if (row < 0) return json_({ ok: true, ideas: '', updated: null });
  var r = sh.getRange(row, 1, 1, 4).getValues()[0];
  return json_({ ok: true, ideas: String(r[2] || ''), updated: r[3] ? new Date(r[3]).toISOString() : null });
}

/** Écriture : POST (text/plain) {action:"save", f, n, k, ideas} */
function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'requête invalide' }); }
  if (body.action !== 'save') return json_({ ok: false, error: 'action inconnue' });
  var fam = clean_(body.f, 60), name = clean_(body.n, 40), key = clean_(body.k, 40);
  var ideas = clean_(body.ideas, MAX_IDEAS);
  if (!name || !key) return json_({ ok: false, error: 'lien incomplet' });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = sheet_();
    var row = findRow_(sh, fam, name);
    var now = new Date();
    if (row < 0) {
      sh.appendRow([fam, name, ideas, now, key]);
    } else {
      var stored = String(sh.getRange(row, 5).getValue() || '');
      if (stored && stored !== key) return json_({ ok: false, error: 'Ce lien ne permet pas de modifier cette liste.' });
      sh.getRange(row, 3, 1, 3).setValues([[ideas, now, stored || key]]);
    }
    return json_({ ok: true, updated: now.toISOString() });
  } finally {
    lock.releaseLock();
  }
}
