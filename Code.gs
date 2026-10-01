/**
 * Getränke-Bestellbuch – Datenquelle für die App auf dem iPhone
 * SV Blau Weiß Etteln 1923 e.V.
 *
 * Was dieses Skript macht:
 *  - liest jede Stunde deine Octopus-Bestellmails aus Gmail (nur lesen)
 *  - speichert die Bestellungen in einer Google-Tabelle „Getränke-Bestellbuch“
 *  - liefert die Bestellungen an deine App (nur mit deinem geheimen Schlüssel)
 *
 * Einrichtung: siehe ANLEITUNG. Kurz: als Web-App bereitstellen, dann „einrichten“ ausführen.
 */

var SUCHE = '(from:octopusorder.com OR "Octopus Bestellung" OR "Einzelheiten der Bestellung") newer_than:400d';

/* ---------- Einmal ausführen ---------- */
function einrichten() {
  var p = PropertiesService.getScriptProperties();
  if (!p.getProperty('KEY')) p.setProperty('KEY', Utilities.getUuid().replace(/-/g, ''));
  tabelle_();
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'mailsAbrufen') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('mailsAbrufen').timeBased().everyHours(1).create();
  var neu = mailsAbrufen();
  var url = ScriptApp.getService().getUrl();
  Logger.log('Fertig. ' + neu + ' Bestellung(en) übernommen.');
  if (url) {
    Logger.log('Diesen Link kopierst du in die App:');
    Logger.log(url + '?k=' + p.getProperty('KEY'));
  } else {
    Logger.log('Noch keine Web-App bereitgestellt: erst „Bereitstellen → Neue Bereitstellung → Web-App“, dann „einrichten“ noch einmal ausführen.');
  }
}

/* ---------- Läuft jede Stunde ---------- */
function mailsAbrufen() {
  var sh = tabelle_(), vorhanden = {}, gmailIds = {};
  var werte = sh.getDataRange().getValues();
  for (var r = 1; r < werte.length; r++) { vorhanden[werte[r][0]] = true; if (werte[r][4]) gmailIds[werte[r][4]] = true; }
  var neu = 0;
  GmailApp.search(SUCHE, 0, 200).forEach(function (thread) {
    thread.getMessages().forEach(function (msg) {
      var mid = msg.getId();
      if (gmailIds[mid]) return;
      var o = parseMail_(msg.getPlainBody() || '', msg.getDate());
      if (!o.items.length) return;
      o.id = docId_(o, mid);
      if (vorhanden[o.id]) return;
      sh.appendRow([o.id, o.date, o.orderNo, JSON.stringify(o.items), mid, 'E-Mail', new Date().toISOString()]);
      vorhanden[o.id] = true; gmailIds[mid] = true; neu++;
    });
  });
  return neu;
}

/* ---------- Schnittstelle für die App ---------- */
function doGet(e) {
  if (!erlaubt_(e.parameter.k)) return json_({ ok: false, error: 'Falscher Schlüssel' });
  if (e.parameter.a === 'sync') mailsAbrufen();
  return json_({ ok: true, orders: alle_() });
}

function doPost(e) {
  var b = {};
  try { b = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Ungültige Anfrage' }); }
  if (!erlaubt_(b.k)) return json_({ ok: false, error: 'Falscher Schlüssel' });
  var sh = tabelle_(), werte = sh.getDataRange().getValues();
  if (b.a === 'save' && b.order && b.order.items && b.order.items.length) {
    var o = b.order, id = String(o.id || docId_(o, 'man-' + Date.now()));
    for (var r = 1; r < werte.length; r++) if (werte[r][0] === id) return json_({ ok: true, exists: true, orders: alle_() });
    sh.appendRow([id, o.date, o.orderNo || '', JSON.stringify(o.items), '', 'manuell', new Date().toISOString()]);
  } else if (b.a === 'delete' && b.id) {
    for (var i = werte.length - 1; i >= 1; i--) if (werte[i][0] === b.id) sh.deleteRow(i + 1);
  }
  return json_({ ok: true, orders: alle_() });
}

/* ---------- Hilfsfunktionen ---------- */
function erlaubt_(k) { var key = PropertiesService.getScriptProperties().getProperty('KEY'); return !!key && k === key; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function tabelle_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('SHEET'), ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('Getränke-Bestellbuch');
    p.setProperty('SHEET', ss.getId());
    ss.getSheets()[0].setName('Bestellungen').appendRow(['ID', 'Datum', 'Bestellnummer', 'Artikel (JSON)', 'Gmail-ID', 'Quelle', 'Eingetragen']);
    ss.getSheets()[0].setFrozenRows(1);
  }
  return ss.getSheets()[0];
}

function alle_() {
  var werte = tabelle_().getDataRange().getValues(), out = [];
  for (var r = 1; r < werte.length; r++) {
    var w = werte[r], items = [];
    try { items = JSON.parse(w[3]); } catch (e) {}
    var d = w[1] instanceof Date ? Utilities.formatDate(w[1], 'Europe/Berlin', 'yyyy-MM-dd') : String(w[1]);
    out.push({ id: String(w[0]), date: d, orderNo: String(w[2] || ''), items: items, source: String(w[5] || '') });
  }
  out.sort(function (a, b) { return a.date < b.date ? 1 : -1; });
  return out;
}

function docId_(o, fallback) {
  return o.orderNo ? 'b-' + String(o.orderNo).replace(/[^A-Za-z0-9.-]/g, '-') : 'm-' + fallback;
}

function iso_(d) { return Utilities.formatDate(d, 'Europe/Berlin', 'yyyy-MM-dd'); }

function parseMail_(raw, fallbackDate) {
  var t = raw.replace(/ /g, ' ').replace(/\r/g, '');
  var items = [], re = /\(\s*Einheit:\s*([^;)]*);?\s*(?:Gebinde:\s*([^)]*))?\)[\s|*]*(\d{1,4})(?!\d)/g, m, last = 0;
  while ((m = re.exec(t))) {
    var seg = t.slice(last, m.index); last = re.lastIndex;
    var rn = /(?:^|[\s|*])(\d{5,8})(?=[\s|*])/g, n = null, x;
    while ((x = rn.exec(seg))) n = x;
    if (!n) continue;
    var name = seg.slice(n.index + n[0].length).replace(/[|*]/g, ' ').replace(/\s+/g, ' ').trim();
    if (!name) continue;
    items.push({ nr: n[1], name: name, einheit: (m[1] || '').trim(), gebinde: (m[2] || '').replace(/\s+/g, '').trim(), qty: +m[3] });
  }
  var pats = [/Erstellung[^\n\d]{0,40}(\d{1,2})\.(\d{1,2})\.(\d{2,4})/i, /Bestell(?:datum|ung vom)[^\n\d]{0,20}(\d{1,2})\.(\d{1,2})\.(\d{2,4})/i, /Lieferzeitpunkt[^\n\d]{0,20}(\d{1,2})\.(\d{1,2})\.(\d{2,4})/i];
  var dm = null; for (var i = 0; i < pats.length && !dm; i++) dm = t.match(pats[i]);
  var date = iso_(fallbackDate || new Date());
  if (dm) { var y = +dm[3]; if (y < 100) y += 2000; date = iso_(new Date(y, +dm[2] - 1, +dm[1])); }
  var on = t.match(/(?:Octopus ID|Bestell(?:nummer|-?nr\.?|ungsnummer))[\s:|*]*([A-Za-z0-9][A-Za-z0-9./-]{3,}[A-Za-z0-9])/i);
  return { items: items, date: date, orderNo: on ? on[1] : '' };
}
