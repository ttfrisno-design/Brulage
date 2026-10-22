// ==================================================================================
// RebonTT — script du COMITÉ DÉPARTEMENTAL (ex. CD76TT)
// ==================================================================================
// À installer par le comité, dans SON Google Drive (indépendant des clubs) :
//   1. Créer un Google Sheet « Inscriptions compétitions » sur le Drive du
//      comité, puis Extensions › Apps Script, et coller ce fichier.
//   2. ⚙️ Paramètres du projet › Propriétés du script :
//        COMITE_CODE  = le code comité (le même que dans chaque club) ;
//        COMITE_CLUBS = la liste des clubs, en JSON :
//          [{"nom":"FR Isneauville TT","url":"https://script.google.com/macros/s/…/exec"}]
//   3. Recharger le classeur : menu « Comité » › « Mettre à jour les
//      inscriptions » (Google demande l'autorisation la première fois), puis
//      « Mise à jour automatique (toutes les heures) ».
// Le classeur contient alors les onglets « Par compétition », « Par club »,
// « Résumé » et « Clubs » (état de chaque club), à partager avec qui de droit.

const COMITE_ONGLETS_ = { comp: 'Par compétition', club: 'Par club', resume: 'Résumé', clubs: 'Clubs' };

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Comité')
    .addItem('Mettre à jour les inscriptions', 'mettreAJourInscriptions')
    .addItem('Mise à jour automatique (toutes les heures)', 'installerMiseAJourAuto')
    .addToUi();
}

function comiteClubs_() {
  const raw = PropertiesService.getScriptProperties().getProperty('COMITE_CLUBS') || '[]';
  let list;
  try { list = JSON.parse(raw); } catch (e) { throw new Error('Propriété COMITE_CLUBS : JSON invalide.'); }
  return (list || []).filter(function (c) { return c && c.url; }).map(function (c) {
    return { nom: String(c.nom || c.url), url: String(c.url), ok: /^https:\/\/script\.google\.com\/\S+\/exec$/.test(String(c.url)) };
  });
}

function dateKey_(d) { const m = String(d || '').match(/^(\d{2})\/(\d{2})\/(\d{4})$/); return m ? m[3] + m[2] + m[1] : '99999999'; }
function compLabel_(r) { return r.competition + (r.epreuve ? ' – ' + r.epreuve : ''); }

/** Interroge le script de chaque club et réécrit les onglets du classeur. */
function mettreAJourInscriptions() {
  const code = PropertiesService.getScriptProperties().getProperty('COMITE_CODE') || '';
  if (code.length < 12) throw new Error('Propriété COMITE_CODE absente ou trop courte (12 caractères minimum).');
  const all = comiteClubs_();
  if (!all.length) throw new Error('Propriété COMITE_CLUBS vide : ajoutez les adresses /exec des clubs.');
  const clubs = all.filter(function (c) { return c.ok; });
  const responses = UrlFetchApp.fetchAll(clubs.map(function (c) {
    return { url: c.url, method: 'post', contentType: 'text/plain', muteHttpExceptions: true, followRedirects: true,
      payload: JSON.stringify({ fn: 'getComiteExport', args: [code] }) };
  }));
  let rows = [];
  const etat = [['Club', 'État', 'Inscriptions', 'Données du']];
  all.filter(function (c) { return !c.ok; }).forEach(function (c) {
    etat.push([c.nom, 'Erreur : adresse invalide (https://script.google.com/…/exec attendu)', 0, '']);
  });
  responses.forEach(function (res, i) {
    const c = clubs[i];
    let p = null;
    try { p = JSON.parse(res.getContentText()); } catch (e) { p = null; }
    if (!p || !p.ok) {
      etat.push([c.nom, 'Erreur : ' + ((p && p.error) || ('HTTP ' + res.getResponseCode())), 0, '']);
      return;
    }
    const nom = (p.result.club && p.result.club.nom) || c.nom;
    p.result.inscriptions.forEach(function (r) { r.club = nom; rows.push(r); });
    etat.push([nom, 'OK', p.result.inscriptions.length, p.result.generatedAt]);
  });
  const line = function (r) { return [r.date, r.type, r.competition, r.epreuve, r.lieu, r.club, r.nom, r.prenom, r.categorie, r.sexe, r.licence, r.points, r.equipe, r.participation]; };
  const cols = ['Date', 'Type', 'Compétition', 'Épreuve', 'Lieu', 'Club', 'Nom', 'Prénom', 'Catégorie', 'Sexe', 'Licence', 'Points', 'Équipe', 'Participation'];
  const byComp = rows.slice().sort(function (a, b) {
    return dateKey_(a.date).localeCompare(dateKey_(b.date)) || compLabel_(a).localeCompare(compLabel_(b)) || a.club.localeCompare(b.club) || a.nom.localeCompare(b.nom);
  });
  const byClub = rows.slice().sort(function (a, b) {
    return a.club.localeCompare(b.club) || dateKey_(a.date).localeCompare(dateKey_(b.date)) || compLabel_(a).localeCompare(compLabel_(b)) || a.nom.localeCompare(b.nom);
  });
  ecrireOnglet_(COMITE_ONGLETS_.comp, [cols].concat(byComp.map(line)));
  ecrireOnglet_(COMITE_ONGLETS_.club, [['Club', 'Date', 'Type', 'Compétition', 'Épreuve', 'Lieu', 'Nom', 'Prénom', 'Catégorie', 'Sexe', 'Licence', 'Points', 'Équipe', 'Participation']]
    .concat(byClub.map(function (r) { const l = line(r); return [l[5]].concat(l.slice(0, 5), l.slice(6)); })));
  // Résumé : nombre d'inscrits par épreuve et par club.
  const clubNames = etat.slice(1).filter(function (e) { return e[1] === 'OK'; }).map(function (e) { return e[0]; });
  const resume = [['Date', 'Type', 'Épreuve'].concat(clubNames, ['Total'])];
  const seen = {};
  byComp.forEach(function (r) {
    const k = r.date + '|' + compLabel_(r);
    if (seen[k] === undefined) { seen[k] = resume.length; resume.push([r.date, r.type, compLabel_(r)].concat(clubNames.map(function () { return 0; }), [0])); }
    const row = resume[seen[k]];
    row[3 + clubNames.indexOf(r.club)]++;
    row[row.length - 1]++;
  });
  ecrireOnglet_(COMITE_ONGLETS_.resume, resume);
  etat.push(['', '', '', ''], ['Mise à jour', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'), rows.length, '']);
  ecrireOnglet_(COMITE_ONGLETS_.clubs, etat);
  return rows.length;
}

function ecrireOnglet_(nom, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(nom) || ss.insertSheet(nom);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.clearContents();
  const w = data[0].length;
  const all = data.map(function (r) { const x = r.slice(0, w); while (x.length < w) x.push(''); return x; });
  // Colonnes entièrement numériques (comptes du Résumé) gardées en nombres ;
  // les autres en texte (dates JJ/MM/AAAA, n° de licence…).
  for (let j = 0; j < w; j++) {
    const num = all.slice(1).every(function (r) { return typeof r[j] === 'number' || r[j] === '' || r[j] === null || r[j] === undefined; });
    if (!num) sh.getRange(1, j + 1, all.length, 1).setNumberFormat('@');
  }
  sh.getRange(1, 1, all.length, w).setValues(all.map(function (r) {
    return r.map(function (v) { return v === null || v === undefined ? '' : (typeof v === 'number' ? v : String(v)); });
  }));
  sh.getRange(1, 1, 1, w).setFontWeight('bold').setBackground('#e8f1fb');
  sh.setFrozenRows(1);
  if (all.length > 1) sh.getRange(1, 1, all.length, w).createFilter();
  sh.autoResizeColumns(1, w);
}

/** Mise à jour automatique toutes les heures. */
function installerMiseAJourAuto() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'mettreAJourInscriptions') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('mettreAJourInscriptions').timeBased().everyHours(1).create();
  SpreadsheetApp.getActive().toast('Mise à jour automatique installée (toutes les heures).');
}
