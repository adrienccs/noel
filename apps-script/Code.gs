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
var OPENS = 'Ouvertures';
var STATUS_KEY = 'iD7u0KIOMH1PxcZS'; // protège la page de suivi
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
  if (p.action === 'open') return recordOpen_(clean_(p.f, 60), clean_(p.n, 40));
  if (p.action === 'status') return status_(p);
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


/* ---------- Suivi des ouvertures ---------- */

function opensSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(OPENS);
  if (!sh) {
    sh = ss.insertSheet(OPENS);
    sh.appendRow(['Famille', 'Prénom', 'Première ouverture', 'Dernière ouverture', 'Nombre']);
    sh.setFrozenRows(1);
    sh.getRange('A1:E1').setFontWeight('bold');
  }
  return sh;
}

/** Appelé par la page de révélation : ?action=open&f=Famille&n=Prénom */
function recordOpen_(fam, name) {
  if (!name) return json_({ ok: false });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = opensSheet_();
    var row = findRow_(sh, fam, name);
    var now = new Date();
    if (row < 0) sh.appendRow([fam, name, now, now, 1]);
    else {
      var n = Number(sh.getRange(row, 5).getValue()) || 0;
      sh.getRange(row, 4, 1, 2).setValues([[now, n + 1]]);
    }
  } finally { lock.releaseLock(); }
  return json_({ ok: true });
}

function esc_(t) {
  return String(t == null ? '' : t).replace(/[&<>"]/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
  });
}

function fmt_(d) {
  return d ? Utilities.formatDate(new Date(d), 'Europe/Paris', "dd/MM 'à' HH'h'mm") : '';
}

function index_(sh, col) {
  var out = {}, last = sh.getLastRow();
  if (last < 2) return out;
  var vals = sh.getRange(2, 1, last - 1, 5).getValues();
  vals.forEach(function (r) { out[(String(r[0]) + '|' + String(r[1])).toLowerCase()] = r; });
  return out;
}

/** Page de suivi : ?action=status&k=CLÉ&d=Famille:Prénom1,Prénom2;Famille2:... */
function status_(p) {
  if (p.k !== STATUS_KEY) return HtmlService.createHtmlOutput('<p>Accès refusé.</p>');
  var opens = index_(opensSheet_()), lists = index_(sheet_());
  var fams = String(p.d || '').split(';').filter(String).map(function (chunk) {
    var i = chunk.indexOf(':');
    return { name: chunk.slice(0, i), people: chunk.slice(i + 1).split(',').filter(String) };
  });
  var total = 0, opened = 0, body = '';
  fams.forEach(function (f) {
    var rows = '', o = 0;
    f.people.forEach(function (n) {
      var key = (f.name + '|' + n).toLowerCase(), r = opens[key], l = lists[key];
      var hasList = l && String(l[2]).trim();
      if (r) o++;
      rows += '<li class="' + (r ? 'ok' : 'no') + '"><span class="n">' + esc_(n) + '</span>' +
        '<span class="s">' + (r ? 'Ouvert le ' + fmt_(r[2]) + (r[4] > 1 ? ' · ' + r[4] + ' fois' : '') : 'Pas encore ouvert') + '</span>' +
        '<span class="l">' + (hasList ? 'Liste d\'idées remplie' : '') + '</span></li>';
    });
    total += f.people.length; opened += o;
    body += '<section><h2>' + esc_(f.name) + ' <small>' + o + ' / ' + f.people.length + ' ouverts</small></h2><ul>' + rows + '</ul></section>';
  });
  var html = '<!doctype html><html lang="fr"><head><meta charset="utf-8"><title>Qui a ouvert ?</title><style>' +
    'body{margin:0;background:#f3f6f4;color:#14231c;font:16px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif;padding:24px 16px}' +
    'main{max-width:560px;margin:0 auto;display:flex;flex-direction:column;gap:16px}' +
    'h1{font-size:1.8rem;margin:0}h1 span{color:#c8243f}p{margin:0;color:#5d6f66}' +
    'section{background:#fff;border:1px solid #d6e0da;border-radius:14px;padding:16px}' +
    'h2{margin:0 0 10px;font-size:1.15rem}h2 small{font-weight:500;color:#5d6f66;font-size:.8rem;letter-spacing:.05em;text-transform:uppercase}' +
    'ul{list-style:none;margin:0;padding:0;display:flex;flex-direction:column;gap:6px}' +
    'li{display:flex;flex-wrap:wrap;gap:4px 10px;align-items:baseline;padding:8px 12px;border-radius:10px;background:#f3f6f4;border-left:4px solid #d6e0da}' +
    'li.ok{border-left-color:#1f7a4d}li.no{border-left-color:#c8243f}' +
    '.n{font-weight:700;min-width:7em}.s{color:#5d6f66;font-size:.9rem}li.ok .s{color:#1f7a4d}.l{font-size:.8rem;color:#1f4d3a;margin-left:auto}' +
    '</style></head><body><main><h1>Qui a <span>ouvert ?</span></h1>' +
    '<p>' + opened + ' / ' + total + ' personnes ont ouvert leur lien. Recharge la page pour actualiser.</p>' +
    body + '</main></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle('Qui a ouvert ?')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
