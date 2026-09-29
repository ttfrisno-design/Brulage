/**
 * Gestion des brûlages – FRI Tennis de Table
 * ============================================
 * Application web (Google Apps Script + Google Sheets) permettant :
 *  - aux joueurs de renseigner leurs disponibilités (vert = dispo, rouge = indispo)
 *  - à l'administrateur de renseigner l'équipe jouée par chaque joueur à chaque date
 *  - le calcul automatique du "brûlage" : un joueur ayant cumulé 2 matchs dans
 *    les équipes au-dessus d'une équipe donnée ne peut plus jouer dans cette équipe.
 *
 * Toutes les données sont stockées dans le classeur Google Sheets lié à ce
 * projet Apps Script. Les feuilles sont créées et initialisées automatiquement
 * au premier lancement (voir ensureSheets_).
 */

// ----------------------------------------------------------------------------------
// CONFIGURATION DES ÉQUIPES ET DES DATES
// ----------------------------------------------------------------------------------

// Ordre 1 -> 8 = ordre hiérarchique utilisé pour le calcul du brûlage.
// "group" identifie l'ensemble de dates (jour de match) partagé par les équipes.
const TEAMS = [
  { id: 1, name: 'FRI1', division: 'Régional 3',      day: 'Dimanche', time: '',      group: 'A' },
  { id: 2, name: 'FRI2', division: 'Départemental 1',  day: 'Vendredi', time: '20:00', group: 'B' },
  { id: 3, name: 'FRI3', division: 'Départemental 2',  day: 'Vendredi', time: '20:00', group: 'B' },
  { id: 4, name: 'FRI4', division: 'Départemental 2',  day: 'Vendredi', time: '20:00', group: 'B' },
  { id: 5, name: 'FRI5', division: 'Départemental 3',  day: 'Vendredi', time: '20:00', group: 'C' },
  { id: 6, name: 'FRI6', division: 'Départemental 4',  day: 'Vendredi', time: '20:00', group: 'C' },
  { id: 7, name: 'FRI7', division: 'Départemental 4',  day: 'Vendredi', time: '20:00', group: 'C' },
  { id: 8, name: 'FRI8', division: 'Départemental 4',  day: 'Vendredi', time: '20:00', group: 'C' },
];

// Dates de phase 1, par groupe (fournies par le club).
// Les dates de phase 2 sont éditables par l'administrateur (initialement vides).
const DEFAULT_DATES = {
  1: { // Phase 1
    A: ['20/09/2025', '04/10/2025', '18/10/2025', '08/11/2025', '22/11/2025', '06/12/2025', '13/12/2025'],
    B: ['25/09/2025', '09/10/2025', '30/10/2025', '13/11/2025', '27/11/2025', '11/12/2025', '18/12/2025'],
    C: ['02/10/2025', '16/10/2025', '06/11/2025', '20/11/2025', '04/12/2025', '08/01/2026', '15/01/2026'],
  },
  2: { // Phase 2 - à compléter par l'administrateur
    A: ['', '', '', '', '', '', ''],
    B: ['', '', '', '', '', '', ''],
    C: ['', '', '', '', '', '', ''],
  },
};

const DEFAULT_ADMIN_PIN = '0000'; // À changer immédiatement depuis l'onglet admin.

// ----------------------------------------------------------------------------------
// WEB APP ENTRY POINT
// ----------------------------------------------------------------------------------

function doGet(e) {
  ensureSheets_();
  const tpl = HtmlService.createTemplateFromFile('Index');
  // URL de ce déploiement, transmise à la page pour ses appels fetch() vers
  // doPost (voir plus bas) — bien plus fiable que le pont google.script.run.
  tpl.scriptUrl = ScriptApp.getService().getUrl();
  return tpl.evaluate()
    .setTitle('Brûlages FRI - Tennis de Table')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ----------------------------------------------------------------------------------
// POINT D'ENTRÉE POST — API JSON appelée via fetch() depuis le navigateur
// ----------------------------------------------------------------------------------
// Le pont interne de google.script.run (iframe caché + postMessage) s'est
// montré peu fiable dans les tests (le serveur termine correctement — voir
// l'onglet Exécutions — mais le navigateur ne reçoit jamais la réponse),
// et ce sur plusieurs navigateurs et réseaux. On expose donc les mêmes
// fonctions via une petite API JSON (doPost), appelée en fetch() classique
// depuis Index.html, qui suit le même chemin réseau fiable que le
// chargement de la page elle-même.
const API_FUNCTIONS_ = {
  login: login,
  getStaticConfig: getStaticConfig,
  getPlayerData: getPlayerData,
  setAvailability: setAvailability,
  getAdminData: getAdminData,
  setAssignment: setAssignment,
  setAssignments: setAssignments,
  resetPhaseAssignments: resetPhaseAssignments,
  listPlayers: listPlayers,
  addPlayer: addPlayer,
  updatePlayer: updatePlayer,
  deletePlayer: deletePlayer,
  updateDates: updateDates,
  setAdminPin: setAdminPin,
};

function doPost(e) {
  let out;
  try {
    const body = JSON.parse(e.postData.contents);
    const fn = API_FUNCTIONS_[body.fn];
    if (!fn) {
      throw new Error('Fonction inconnue : ' + body.fn);
    }
    const args = body.args || [];
    const result = fn.apply(null, args);
    out = { ok: true, result: result };
  } catch (err) {
    out = { ok: false, error: err && err.message ? err.message : String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------------------------
// PROVISIONING DES FEUILLES
// ----------------------------------------------------------------------------------

// Classeur ouvert une seule fois par exécution (openById est lent et était
// appelé à chaque lecture de feuille).
let ssCache_ = null;

function ss_() {
  if (!ssCache_) ssCache_ = openSpreadsheet_();
  return ssCache_;
}

function openSpreadsheet_() {
  // SpreadsheetApp.getActiveSpreadsheet() renvoie null lorsque le script
  // s'exécute en tant qu'application web (il n'y a pas de classeur "actif"
  // ouvert par la personne qui déclenche la requête). On ouvre donc
  // toujours le classeur par un ID fixe, mémorisé une fois pour toutes
  // dans les propriétés du script.
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SPREADSHEET_ID');
  if (id) {
    return SpreadsheetApp.openById(id);
  }

  // Premier lancement : si le script est lié à un classeur (Extensions >
  // Apps Script depuis un Google Sheet), on le récupère et on retient son ID.
  const bound = SpreadsheetApp.getActiveSpreadsheet();
  if (bound) {
    props.setProperty('SPREADSHEET_ID', bound.getId());
    return bound;
  }

  // Sinon (script autonome créé depuis script.new), on crée un classeur
  // dédié et on retient son ID pour tous les prochains appels.
  const created = SpreadsheetApp.create('Brûlages FRI - Données');
  props.setProperty('SPREADSHEET_ID', created.getId());
  return created;
}

/** Utilitaire : à lancer manuellement depuis l'éditeur Apps Script (menu
 * Exécuter, fonction "getDataSpreadsheetUrl") pour retrouver l'URL du
 * classeur de données si besoin. Le lien est aussi affiché dans l'onglet
 * admin "Dates & Réglages". */
function getDataSpreadsheetUrl() {
  ensureSheets_();
  const url = ss_().getUrl();
  Logger.log(url);
  return url;
}

// Vérification/réparation des feuilles : coûteuse (dizaines d'écritures de
// format, relecture et dédoublonnage de toutes les feuilles), elle était
// relancée à CHAQUE appel — et même deux fois pour login(). Avec un
// démarrage à froid d'Apps Script, une connexion pouvait dépasser le délai
// d'attente de l'application et échouer ("Le serveur met trop de temps à
// répondre"). On ne la fait donc plus qu'une fois par exécution, et au plus
// une fois par heure (mémorisé dans le cache du script).
let sheetsChecked_ = false;
const SHEETS_CHECK_CACHE_KEY_ = 'SHEETS_CHECKED_V1';

function ensureSheets_() {
  if (sheetsChecked_) return;
  const cache = CacheService.getScriptCache();
  if (cache.get(SHEETS_CHECK_CACHE_KEY_)) {
    sheetsChecked_ = true;
    return;
  }
  ensureSheetsNow_();
  sheetsChecked_ = true;
  cache.put(SHEETS_CHECK_CACHE_KEY_, '1', 3600);
}

/** À lancer manuellement depuis l'éditeur Apps Script (menu Exécuter) après
 * avoir modifié ou supprimé une feuille à la main, pour forcer tout de suite
 * la vérification/réparation au lieu d'attendre l'heure suivante. */
function forceSheetsCheck() {
  CacheService.getScriptCache().remove(SHEETS_CHECK_CACHE_KEY_);
  sheetsChecked_ = false;
  ensureSheets_();
}

function ensureSheetsNow_() {
  const ss = ss_();

  // --- Config ---
  let cfg = ss.getSheetByName('Config');
  if (!cfg) {
    cfg = ss.insertSheet('Config');
    // Format texte AVANT d'écrire les valeurs : sinon Sheets convertit un
    // PIN comme "0000" en nombre 0 et fait perdre les zéros de tête.
    cfg.getRange(1, 1, 10, 2).setNumberFormat('@');
    cfg.getRange(1, 1, 3, 2).setValues([
      ['Clé', 'Valeur'],
      ['AdminPIN', DEFAULT_ADMIN_PIN],
      ['PhaseActive', '1'],
    ]);
    cfg.setFrozenRows(1);
  }

  // --- Dates ---
  let datesSheet = ss.getSheetByName('Dates');
  if (!datesSheet) {
    datesSheet = ss.insertSheet('Dates');
    // Colonne F "Journée" (optionnelle) : quand elle est renseignée, deux
    // dates portant la même valeur (ex. "J1") sont considérées comme la
    // même journée de championnat — un joueur ne peut alors être affecté
    // qu'à une seule d'entre elles (voir getStaticConfig / journeeByDate).
    datesSheet.getRange(1, 1, 1, 6).setValues([['Phase', 'Groupe', 'Index', 'Date', 'Libellé', 'Journée']]);
    // Colonne Date (D) toujours en texte : sinon Sheets convertit "20/09/2025"
    // en date interne (avec un risque d'inversion jour/mois et de décalage
    // d'un jour selon le fuseau horaire lors de la relecture).
    datesSheet.getRange(2, 4, 998, 1).setNumberFormat('@');
    datesSheet.setFrozenRows(1);
    const rows = [];
    [1, 2].forEach(function (phase) {
      ['A', 'B', 'C'].forEach(function (grp) {
        DEFAULT_DATES[phase][grp].forEach(function (d, idx) {
          rows.push([phase, grp, idx + 1, d, '', '']);
        });
      });
    });
    datesSheet.getRange(2, 1, rows.length, 6).setValues(rows);
  }
  // Réparation : classeur créé par une version antérieure sans la colonne F
  // "Journée" — on ajoute juste l'en-tête, sans toucher aux données.
  if (datesSheet.getLastColumn() < 6) {
    datesSheet.getRange(1, 6).setValue('Journée');
  }

  // --- Players ---
  let players = ss.getSheetByName('Joueurs');
  if (!players) {
    players = ss.insertSheet('Joueurs');
    players.getRange(1, 1, 1, 7).setValues([['ID', 'Nom', 'Prénom', 'EquipeDomicile', 'PIN', 'Actif', 'Capitaine']]);
    // Colonne PIN (E) toujours en texte, pour ne jamais perdre un zéro de tête (ex. "0042").
    players.getRange(2, 5, 998, 1).setNumberFormat('@');
    players.setFrozenRows(1);
  }
  // Réparation : classeur créé par une version antérieure sans la colonne
  // "Capitaine" (7e colonne) — on l'ajoute sans toucher aux données existantes.
  if (players.getLastColumn() < 7) {
    players.getRange(1, 7).setValue('Capitaine');
  }

  // --- Availability ---
  let avail = ss.getSheetByName('Disponibilites');
  if (!avail) {
    avail = ss.insertSheet('Disponibilites');
    avail.getRange(1, 1, 1, 4).setValues([['JoueurID', 'Phase', 'Date', 'Statut']]);
    // Colonne Date (C) en texte dès la création — même raison que sur
    // Affectations (voir plus bas) et Dates.
    avail.getRange(2, 3, 998, 1).setNumberFormat('@');
    avail.setFrozenRows(1);
  }
  // Réparation idempotente de la colonne Date (C), comme pour Affectations.
  if (avail.getLastRow() > 1) {
    const availDateRange = avail.getRange(2, 3, avail.getLastRow() - 1, 1);
    const availDateValues = availDateRange.getValues();
    let availDatesChanged = false;
    const availDatesFixed = availDateValues.map(function (row) {
      const v = row[0];
      if (v instanceof Date) {
        availDatesChanged = true;
        return [Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy')];
      }
      return [v];
    });
    availDateRange.setNumberFormat('@');
    if (availDatesChanged) {
      availDateRange.setValues(availDatesFixed);
    }
  }
  avail.getRange(2, 3, 998, 1).setNumberFormat('@');
  // Réparation : lignes en double ou contradictoires (même joueur/phase/
  // date, statuts différents) — même cause que sur Affectations (écritures
  // retentées automatiquement). On ne garde que la dernière.
  dedupeDisponibilites_();

  // --- Assignments ---
  let assign = ss.getSheetByName('Affectations');
  if (!assign) {
    assign = ss.insertSheet('Affectations');
    assign.getRange(1, 1, 1, 4).setValues([['JoueurID', 'Phase', 'Date', 'EquipeJouee']]);
    // Colonne Date (C) en texte dès la création, comme sur l'onglet Dates :
    // sinon Sheets peut convertir "02/10/2026" tapé directement en date
    // interne, avec un risque de décalage d'un jour selon le fuseau horaire
    // à la relecture — l'affectation existe alors sous une date légèrement
    // différente de celle affichée, et semble "invisible" dans l'appli.
    assign.getRange(2, 3, 998, 1).setNumberFormat('@');
    assign.setFrozenRows(1);
  }
  // Réparation idempotente de la colonne Date (C) : si Sheets a déjà
  // converti certaines valeurs en dates internes (lignes tapées directement
  // dans le sheet plutôt que via l'appli), on les reconvertit en texte
  // "jj/mm/aaaa" puis on force définitivement le format texte.
  if (assign.getLastRow() > 1) {
    const assignDateRange = assign.getRange(2, 3, assign.getLastRow() - 1, 1);
    const assignDateValues = assignDateRange.getValues();
    let assignDatesChanged = false;
    const assignDatesFixed = assignDateValues.map(function (row) {
      const v = row[0];
      if (v instanceof Date) {
        assignDatesChanged = true;
        return [Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy')];
      }
      return [v];
    });
    assignDateRange.setNumberFormat('@');
    if (assignDatesChanged) {
      assignDateRange.setValues(assignDatesFixed);
    }
  }
  assign.getRange(2, 3, 998, 1).setNumberFormat('@');

  // Réparation : des lignes en double (même joueur/phase/date) ont pu
  // s'accumuler avec d'anciennes versions du script (écritures concurrentes
  // lors des tentatives automatiques réseau) — on nettoie systématiquement,
  // sans effet ni écriture si tout est déjà propre.
  dedupeAffectations_();

  // Réparation idempotente : si un classeur existant a été créé par une
  // version antérieure du script (avant le passage en format texte), on
  // force le format texte sur les cellules de PIN et on ré-écrit les
  // valeurs telles quelles pour restaurer les zéros de tête déjà perdus
  // (ex. AdminPIN affiché "0" -> remis à "0000" uniquement s'il vaut "0").
  const cfgRange = cfg.getRange(2, 2, Math.max(cfg.getLastRow() - 1, 1), 1);
  cfgRange.setNumberFormat('@');
  const adminPinCell = cfg.getRange(2, 2);
  if (String(adminPinCell.getValue()) === '0') {
    adminPinCell.setValue(DEFAULT_ADMIN_PIN);
  }
  if (players.getLastRow() > 1) {
    players.getRange(2, 5, players.getLastRow() - 1, 1).setNumberFormat('@');
  }
  players.getRange(2, 5, 998, 1).setNumberFormat('@');

  // Réparation idempotente de la colonne Date : si Sheets a déjà converti
  // certaines valeurs en dates internes (classeur créé avant ce correctif),
  // on les reconvertit en texte "jj/mm/aaaa" puis on force définitivement
  // le format texte pour que ça ne se reproduise plus.
  if (datesSheet.getLastRow() > 1) {
    const dateRange = datesSheet.getRange(2, 4, datesSheet.getLastRow() - 1, 1);
    const dateValues = dateRange.getValues();
    let changed = false;
    const fixed = dateValues.map(function (row) {
      const v = row[0];
      if (v instanceof Date) {
        changed = true;
        return [Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy')];
      }
      return [v];
    });
    dateRange.setNumberFormat('@');
    if (changed) {
      dateRange.setValues(fixed);
    }
  }
  datesSheet.getRange(2, 4, 998, 1).setNumberFormat('@');

  // Remove default "Feuille 1" if empty and unused
  const sheet1 = ss.getSheetByName('Feuille 1') || ss.getSheetByName('Sheet1');
  if (sheet1 && ss.getSheets().length > 1 && sheet1.getLastRow() === 0) {
    ss.deleteSheet(sheet1);
  }
}

// ----------------------------------------------------------------------------------
// HELPERS DE LECTURE DE FEUILLE (retournent des tableaux d'objets)
// ----------------------------------------------------------------------------------

function readTable_(sheetName) {
  const sheet = ss_().getSheetByName(sheetName);
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  return values
    .filter(function (row) { return row.some(function (c) { return c !== '' && c !== null; }); })
    .map(function (row) {
      const obj = {};
      headers.forEach(function (h, i) {
        let v = row[i];
        // Filet de sécurité : si Sheets renvoie malgré tout un objet Date
        // (ex. classeur ancien pas encore réparé), on le reconvertit en
        // texte "jj/mm/aaaa" plutôt que de laisser passer un ISO/UTC brut.
        if (v instanceof Date) {
          v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        }
        obj[h] = v;
      });
      return obj;
    });
}

function normalize_(str) {
  return String(str || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '');
}

function pad2_(n) {
  n = String(n);
  return n.length < 2 ? '0' + n : n;
}

/**
 * Force toute date saisie au format strict "JJ/MM/AAAA", quel que soit le
 * séparateur utilisé (/, -, .) ou l'ordre ISO (AAAA-MM-JJ). Lève une erreur
 * explicite si la valeur ne peut pas être interprétée comme une date —
 * c'est ce qui garantit que le format reste toujours JJ/MM/AAAA dans la
 * feuille "Dates", quoi que l'administrateur ait tapé.
 */
function normalizeDateFr_(raw) {
  const s = String(raw || '').trim();
  if (!s) return '';

  let dd, mm, yyyy;
  let m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (m) {
    dd = m[1]; mm = m[2]; yyyy = m[3];
  } else {
    m = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})$/); // format ISO aaaa-mm-jj
    if (m) {
      yyyy = m[1]; mm = m[2]; dd = m[3];
    }
  }
  if (!m) {
    throw new Error('Date invalide : "' + raw + '" — le format attendu est JJ/MM/AAAA (ex. 20/09/2025).');
  }
  const ddNum = Number(dd), mmNum = Number(mm);
  if (ddNum < 1 || ddNum > 31 || mmNum < 1 || mmNum > 12) {
    throw new Error('Date invalide : "' + raw + '" — le format attendu est JJ/MM/AAAA (ex. 20/09/2025).');
  }
  return pad2_(dd) + '/' + pad2_(mm) + '/' + yyyy;
}

// ----------------------------------------------------------------------------------
// AUTHENTIFICATION
// ----------------------------------------------------------------------------------

function login(pin) {
  ensureSheets_();

  // Connexion admin : le code correspond au PIN administrateur.
  const cfg = readTable_('Config');
  const adminPin = (cfg.find(function (r) { return r.Clé === 'AdminPIN'; }) || {}).Valeur;
  if (String(pin) === String(adminPin)) {
    return { ok: true, isAdmin: true, name: 'Administrateur', config: getStaticConfig() };
  }

  // Connexion joueur : identification uniquement par le code personnel
  // (plus besoin de saisir nom/prénom). Chaque joueur doit donc avoir un
  // PIN unique — voir la vérification anti-doublon dans addPlayer/updatePlayer.
  const players = readTable_('Joueurs');
  const match = players.find(function (p) {
    return String(p.PIN) === String(pin) &&
      p.Actif !== false && p.Actif !== 'FAUX' && p.Actif !== 'NON';
  });
  if (!match) {
    return { ok: false, error: "Code incorrect." };
  }
  const isCaptain = match.Capitaine === true || match.Capitaine === 'VRAI' || match.Capitaine === 'TRUE';
  return {
    ok: true,
    isAdmin: false,
    isCaptain: isCaptain,
    playerId: match.ID,
    name: match.Prénom + ' ' + match.Nom,
    homeTeam: Number(match.EquipeDomicile) || null,
    captainTeamId: isCaptain ? (Number(match.EquipeDomicile) || null) : null,
    config: getStaticConfig(),
  };
}

// ----------------------------------------------------------------------------------
// DONNÉES DE RÉFÉRENCE (équipes / dates)
// ----------------------------------------------------------------------------------

function getStaticConfig() {
  ensureSheets_();
  const sheet = ss_().getSheetByName('Dates');
  const values = sheet.getDataRange().getValues();
  values.shift(); // en-têtes
  const dates = { 1: { A: [], B: [], C: [] }, 2: { A: [], B: [], C: [] } };
  // journeeByDate : date (JJ/MM/AAAA) -> libellé de journée (ex. "J1"), lu
  // directement en colonne F par position (et non par nom d'en-tête, pour
  // rester robuste quel que soit l'intitulé exact tapé dans la feuille).
  // Uniquement rempli pour les dates où cette colonne a été renseignée.
  const journeeByDate = {};
  values.forEach(function (row) {
    if (!row.some(function (c) { return c !== '' && c !== null; })) return;
    const phase = Number(row[0]);
    const grp = row[1];
    const idx = Number(row[2]) - 1;
    let dateVal = row[3];
    if (dateVal instanceof Date) {
      dateVal = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }
    if (dates[phase] && dates[phase][grp] !== undefined) {
      dates[phase][grp][idx] = dateVal;
    }
    const journee = row[5];
    if (dateVal && journee !== undefined && journee !== null && String(journee).trim() !== '') {
      journeeByDate[dateVal] = String(journee).trim();
    }
  });
  return { teams: TEAMS, dates: dates, journeeByDate: journeeByDate };
}

function updateDates(phase, group, dates) {
  ensureSheets_();
  // Chaque date saisie est forcée au format JJ/MM/AAAA (ou vidée si vide) —
  // une erreur explicite est renvoyée si l'une d'elles est incompréhensible,
  // avant toute écriture, pour que la feuille ne contienne jamais autre
  // chose que ce format.
  const normalized = dates.map(normalizeDateFr_);
  const sheet = ss_().getSheetByName('Dates');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (Number(values[i][0]) === Number(phase) && values[i][1] === group) {
      const idx = Number(values[i][2]) - 1;
      if (normalized[idx] !== undefined) {
        const cell = sheet.getRange(i + 1, 4);
        cell.setNumberFormat('@'); // toujours en texte
        cell.setValue(normalized[idx]);
      }
    }
  }
  return getStaticConfig();
}

function setAdminPin(newPin) {
  ensureSheets_();
  const players = readTable_('Joueurs');
  if (players.some(function (p) { return String(p.PIN) === String(newPin); })) {
    throw new Error('Ce code est déjà utilisé par un joueur. Choisissez-en un autre.');
  }
  const sheet = ss_().getSheetByName('Config');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === 'AdminPIN') {
      const cell = sheet.getRange(i + 1, 2);
      cell.setNumberFormat('@'); // toujours en texte, pour ne pas perdre un zéro de tête
      cell.setValue(String(newPin));
      return { ok: true };
    }
  }
  return { ok: false };
}

// ----------------------------------------------------------------------------------
// CALCUL DU BRÛLAGE
// ----------------------------------------------------------------------------------

/**
 * matchesByTeam: { teamId: nombre de matchs joués dans cette équipe (phase en cours) }
 * Retourne { teamId: bool } -> true si le joueur est "brûlé" pour cette équipe
 * (= il ne peut plus y jouer), calculé selon la règle :
 *   un joueur est brûlé pour l'équipe N si la somme des matchs joués dans les
 *   équipes classées AU-DESSUS de N (numéros < N) atteint 2.
 */
function computeBurnedTeams_(matchesByTeam) {
  const burned = {};
  let cumulative = 0;
  TEAMS.forEach(function (team) {
    burned[team.id] = cumulative >= 2;
    cumulative += Number(matchesByTeam[team.id] || 0);
  });
  return burned;
}

// ----------------------------------------------------------------------------------
// VUE JOUEUR
// ----------------------------------------------------------------------------------

function getPlayerData(playerId, phase) {
  ensureSheets_();
  const avail = readTable_('Disponibilites').filter(function (r) {
    return String(r.JoueurID) === String(playerId) && Number(r.Phase) === Number(phase);
  });
  const assign = readTable_('Affectations').filter(function (r) {
    return String(r.JoueurID) === String(playerId) && Number(r.Phase) === Number(phase);
  });

  const availMap = {};
  avail.forEach(function (r) { availMap[r.Date] = r.Statut; });

  const assignMap = {};
  const matchesByTeam = {};
  assign.forEach(function (r) {
    assignMap[r.Date] = Number(r.EquipeJouee);
    matchesByTeam[r.EquipeJouee] = (matchesByTeam[r.EquipeJouee] || 0) + 1;
  });

  return {
    availability: availMap,
    assignments: assignMap,
    burned: computeBurnedTeams_(matchesByTeam),
    matchesByTeam: matchesByTeam,
  };
}

function setAvailability(playerId, phase, date, status) {
  ensureSheets_();
  const sheet = ss_().getSheetByName('Disponibilites');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(playerId) && Number(values[i][1]) === Number(phase) && values[i][2] === date) {
      if (!status) {
        sheet.deleteRow(i + 1);
      } else {
        sheet.getRange(i + 1, 4).setValue(status);
      }
      return { ok: true };
    }
  }
  if (status) {
    sheet.appendRow([playerId, phase, date, status]);
  }
  return { ok: true };
}

// ----------------------------------------------------------------------------------
// VUE ADMINISTRATEUR
// ----------------------------------------------------------------------------------

function getAdminData(phase) {
  ensureSheets_();
  const players = readTable_('Joueurs');
  const avail = readTable_('Disponibilites').filter(function (r) { return Number(r.Phase) === Number(phase); });
  const assign = readTable_('Affectations').filter(function (r) { return Number(r.Phase) === Number(phase); });

  const availByPlayer = {};
  avail.forEach(function (r) {
    availByPlayer[r.JoueurID] = availByPlayer[r.JoueurID] || {};
    availByPlayer[r.JoueurID][r.Date] = r.Statut;
  });

  const assignByPlayer = {};
  const matchesByPlayerTeam = {};
  assign.forEach(function (r) {
    assignByPlayer[r.JoueurID] = assignByPlayer[r.JoueurID] || {};
    assignByPlayer[r.JoueurID][r.Date] = Number(r.EquipeJouee);
    matchesByPlayerTeam[r.JoueurID] = matchesByPlayerTeam[r.JoueurID] || {};
    matchesByPlayerTeam[r.JoueurID][r.EquipeJouee] = (matchesByPlayerTeam[r.JoueurID][r.EquipeJouee] || 0) + 1;
  });

  const result = players
    .filter(function (p) { return p.Actif !== false && p.Actif !== 'FAUX' && p.Actif !== 'NON'; })
    .map(function (p) {
      return {
        id: p.ID,
        nom: p.Nom,
        prenom: p.Prénom,
        homeTeam: Number(p.EquipeDomicile) || null,
        availability: availByPlayer[p.ID] || {},
        assignments: assignByPlayer[p.ID] || {},
        matchesByTeam: matchesByPlayerTeam[p.ID] || {},
        burned: computeBurnedTeams_(matchesByPlayerTeam[p.ID] || {}),
      };
    });

  return result;
}

/**
 * Nettoie les lignes en double dans la feuille "Affectations" (plusieurs
 * lignes pour le même JoueurID+Phase+Date, apparues avec d'anciennes
 * versions du script lors d'écritures concurrentes — ex. une requête
 * retentée automatiquement après un délai dépassé côté client, alors que
 * le serveur avait déjà terminé). Ne garde que la DERNIÈRE ligne
 * rencontrée pour chaque combinaison, ce qui correspond à la valeur la
 * plus récemment écrite. Ne réécrit la feuille QUE si un doublon a été
 * trouvé (aucun effet, aucune écriture, si tout est déjà propre).
 */
function dedupeAffectations_() {
  const sheet = ss_().getSheetByName('Affectations');
  if (!sheet || sheet.getLastRow() < 2) return false;
  const values = sheet.getDataRange().getValues();
  const header = values.shift();

  const byKey = {};
  let hadDuplicates = false;
  values.forEach(function (row) {
    if (!row.some(function (c) { return c !== '' && c !== null; })) return;
    const key = row[0] + '|' + row[1] + '|' + row[2];
    if (Object.prototype.hasOwnProperty.call(byKey, key)) { hadDuplicates = true; }
    byKey[key] = row; // garde la dernière occurrence rencontrée
  });
  if (!hadDuplicates) return false;

  const kept = Object.keys(byKey)
    .map(function (k) { return byKey[k]; })
    .filter(function (row) { return row[3] !== '' && row[3] !== null && row[3] !== undefined; });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (kept.length) {
    sheet.getRange(2, 1, kept.length, 4).setValues(kept);
  }
  return true;
}

/**
 * Même nettoyage que dedupeAffectations_, pour la feuille "Disponibilites" :
 * plusieurs lignes pour le même JoueurID+Phase+Date, parfois avec des
 * statuts contradictoires ("Disponible" ET "Indisponible" en même temps),
 * causées par les mêmes écritures retentées automatiquement côté client.
 * Ne garde que la DERNIÈRE ligne rencontrée pour chaque combinaison.
 */
function dedupeDisponibilites_() {
  const sheet = ss_().getSheetByName('Disponibilites');
  if (!sheet || sheet.getLastRow() < 2) return false;
  const values = sheet.getDataRange().getValues();
  const header = values.shift();

  const byKey = {};
  let hadDuplicates = false;
  values.forEach(function (row) {
    if (!row.some(function (c) { return c !== '' && c !== null; })) return;
    const key = row[0] + '|' + row[1] + '|' + row[2];
    if (Object.prototype.hasOwnProperty.call(byKey, key)) { hadDuplicates = true; }
    byKey[key] = row; // garde la dernière occurrence rencontrée
  });
  if (!hadDuplicates) return false;

  const kept = Object.keys(byKey)
    .map(function (k) { return byKey[k]; })
    .filter(function (row) { return row[3] !== '' && row[3] !== null && row[3] !== undefined; });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (kept.length) {
    sheet.getRange(2, 1, kept.length, 4).setValues(kept);
  }
  return true;
}

function setAssignment(playerId, phase, date, teamId) {
  ensureSheets_();

  // Un joueur ne peut être affecté qu'une seule fois par "journée" (colonne
  // F de la feuille Dates, ex. "J1") : si une autre date porte la même
  // journée et que ce joueur y est déjà affecté, on refuse — l'interface
  // grise normalement déjà ces cases, ceci est un filet de sécurité côté
  // serveur (ex. en cas de double saisie concurrente).
  if (teamId) {
    const journeeByDate = getStaticConfig().journeeByDate || {};
    const targetJournee = journeeByDate[date];
    if (targetJournee) {
      const existing = readTable_('Affectations').filter(function (r) {
        return String(r.JoueurID) === String(playerId) && Number(r.Phase) === Number(phase) && r.Date !== date;
      });
      const conflict = existing.find(function (r) { return journeeByDate[r.Date] === targetJournee; });
      if (conflict) {
        throw new Error('Ce joueur est déjà affecté le ' + conflict.Date + ', qui est la même journée (' + targetJournee + ') que le ' + date + '.');
      }
    }
  }

  const sheet = ss_().getSheetByName('Affectations');
  const values = sheet.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(playerId) && Number(values[i][1]) === Number(phase) && values[i][2] === date) {
      found = true;
      if (!teamId) {
        sheet.deleteRow(i + 1);
      } else {
        sheet.getRange(i + 1, 4).setValue(Number(teamId));
      }
      break;
    }
  }
  if (!found && teamId) {
    sheet.appendRow([playerId, phase, date, Number(teamId)]);
  }

  // Renvoie l'état de brûlage à jour pour CE joueur (pour rafraîchir l'UI immédiatement)
  const assign = readTable_('Affectations').filter(function (r) {
    return String(r.JoueurID) === String(playerId) && Number(r.Phase) === Number(phase);
  });
  const matchesByTeam = {};
  assign.forEach(function (r) { matchesByTeam[r.EquipeJouee] = (matchesByTeam[r.EquipeJouee] || 0) + 1; });
  return { ok: true, matchesByTeam: matchesByTeam, burned: computeBurnedTeams_(matchesByTeam) };
}

/**
 * Version "en lot" de setAssignment : applique plusieurs affectations en
 * UN SEUL appel serveur (une seule lecture + une seule écriture de la
 * feuille Affectations), au lieu d'un aller-retour réseau par case cochée.
 * Utilisée par le bouton "Valider et enregistrer" côté client, qui pouvait
 * mettre plus d'une minute à tout envoyer case par case dès qu'il y avait
 * beaucoup de modifications.
 * entries: [{ playerId, phase, date, teamId }, ...] (teamId vide/0 = retire
 * l'affectation).
 */
function setAssignments(entries) {
  ensureSheets_();
  entries = entries || [];
  if (!entries.length) return { ok: true };

  const journeeByDate = getStaticConfig().journeeByDate || {};
  const sheet = ss_().getSheetByName('Affectations');
  const values = sheet.getDataRange().getValues();
  const header = values.shift();

  const rowIndex = {};
  values.forEach(function (row, i) {
    rowIndex[row[0] + '|' + row[1] + '|' + row[2]] = i;
  });

  // Applique chaque entrée en mémoire, avec la même vérification "une seule
  // date par journée" que setAssignment — mais en tenant compte, pour
  // chaque entrée, des affectations déjà appliquées plus tôt dans CE lot
  // (values reflète l'état au fur et à mesure).
  entries.forEach(function (e) {
    if (e.teamId) {
      const targetJournee = journeeByDate[e.date];
      if (targetJournee) {
        const conflict = values.some(function (row) {
          return String(row[0]) === String(e.playerId) && Number(row[1]) === Number(e.phase) &&
            row[2] !== e.date && row[3] !== '' && row[3] !== null && row[3] !== undefined &&
            journeeByDate[row[2]] === targetJournee;
        });
        if (conflict) {
          throw new Error('Un joueur est déjà affecté à une autre date de la même journée (' + targetJournee + ', le ' + e.date + ').');
        }
      }
    }
    const key = e.playerId + '|' + e.phase + '|' + e.date;
    const idx = rowIndex[key];
    if (idx !== undefined) {
      values[idx][3] = e.teamId ? Number(e.teamId) : '';
    } else if (e.teamId) {
      values.push([e.playerId, e.phase, e.date, Number(e.teamId)]);
      rowIndex[key] = values.length - 1;
    }
  });

  const kept = values.filter(function (row) {
    return row[3] !== '' && row[3] !== null && row[3] !== undefined;
  });

  // Réécrit toute la feuille en une seule opération (bien plus rapide que
  // modifier ligne par ligne, surtout pour un gros lot de modifications).
  sheet.clearContents();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (kept.length) {
    sheet.getRange(2, 1, kept.length, 4).setValues(kept);
  }

  return { ok: true };
}

function resetPhaseAssignments(phase) {
  ensureSheets_();
  const sheet = ss_().getSheetByName('Affectations');
  const values = sheet.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) {
    if (Number(values[i][1]) === Number(phase)) {
      sheet.deleteRow(i + 1);
    }
  }
  return { ok: true };
}

// ----------------------------------------------------------------------------------
// GESTION DES JOUEURS (effectif)
// ----------------------------------------------------------------------------------

function listPlayers() {
  ensureSheets_();
  return readTable_('Joueurs');
}

function assertPinAvailable_(pin, excludeId) {
  // Le code (PIN) est désormais le SEUL identifiant de connexion (plus de
  // nom/prénom) : il doit donc être unique parmi les joueurs actifs, et
  // différent du code administrateur.
  const cfg = readTable_('Config');
  const adminPin = (cfg.find(function (r) { return r.Clé === 'AdminPIN'; }) || {}).Valeur;
  if (String(pin) === String(adminPin)) {
    throw new Error("Ce code est déjà utilisé comme code administrateur. Choisissez-en un autre.");
  }
  const players = readTable_('Joueurs');
  const clash = players.some(function (p) {
    return String(p.PIN) === String(pin) && (excludeId === undefined || String(p.ID) !== String(excludeId));
  });
  if (clash) {
    throw new Error('Ce code est déjà utilisé par un autre joueur. Choisissez-en un autre.');
  }
}

function addPlayer(nom, prenom, teamId, pin, capitaine) {
  ensureSheets_();
  assertPinAvailable_(pin);
  const sheet = ss_().getSheetByName('Joueurs');
  const players = readTable_('Joueurs');
  const maxId = players.reduce(function (m, p) { return Math.max(m, Number(p.ID) || 0); }, 0);
  const newId = maxId + 1;
  sheet.appendRow([newId, nom, prenom, teamId, pin, true, !!capitaine]);
  return { ok: true, id: newId };
}

function updatePlayer(id, nom, prenom, teamId, pin, actif, capitaine) {
  ensureSheets_();
  assertPinAvailable_(pin, id);
  const sheet = ss_().getSheetByName('Joueurs');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sheet.getRange(i + 1, 2, 1, 6).setValues([[nom, prenom, teamId, pin, actif, !!capitaine]]);
      return { ok: true };
    }
  }
  return { ok: false };
}

function deletePlayer(id) {
  ensureSheets_();
  const sheet = ss_().getSheetByName('Joueurs');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false };
}

// ----------------------------------------------------------------------------------
// IMPORT PONCTUEL — Effectif FRI1 (Régional 3)
// ----------------------------------------------------------------------------------

/**
 * À lancer UNE FOIS depuis l'éditeur Apps Script (menu Exécuter, fonction
 * "seedR3Players") pour ajouter d'un coup les joueurs de l'équipe FRI1.
 * PIN provisoires 1001-1006, à changer ensuite depuis l'onglet Effectif.
 * Relancer la fonction sans risque : les joueurs déjà présents (même nom +
 * prénom) ne sont pas dupliqués.
 */
function seedR3Players() {
  return seedTeam_(1, [
    ['DUMAINE', 'Fabrice'],
    ['BENON', 'Nathan'],
    ['BOUWYN', 'Jean-Paul'],
    ['FRIDA', 'Mohamed'],
    ['LANGLET', 'Marius'],
    ['MAINEMARE', 'Richard'],
  ], 1001);
}

/**
 * À lancer UNE FOIS depuis l'éditeur Apps Script (menu Exécuter, fonction
 * "seedD1Players") pour ajouter d'un coup les joueurs de l'équipe FRI2 (D1).
 * PIN provisoires 1007-1011, à changer ensuite depuis l'onglet Effectif.
 */
function seedD1Players() {
  return seedTeam_(2, [
    ['LANGLOIS', 'Thomas'],
    ['HAVE', 'Dominique'],
    ['BLOQUET', 'Perig'],
    ['JAUMOTTE', 'Lucien'],
    ['COUTURE', 'Stephane'],
  ], 1007);
}

/**
 * À lancer UNE FOIS depuis l'éditeur Apps Script (menu Exécuter, fonction
 * "seedD2aPlayers") pour ajouter d'un coup les joueurs de l'équipe FRI3 (D2).
 * PIN provisoires 1012-1017, à changer ensuite depuis l'onglet Effectif.
 */
function seedD2aPlayers() {
  return seedTeam_(3, [
    ['ESCARGUEIL', 'Sebastien'],
    ['LEDREAU', 'Alex'],
    ['GRULEY', 'Tony'],
    ['FONTAINE', 'Joseph'],
    ['MAUSSION', 'Frédéric'],
    ['ZERBIB', 'Laurent'],
  ], 1012);
}

/**
 * À lancer UNE FOIS depuis l'éditeur Apps Script (menu Exécuter, fonction
 * "seedD2bPlayers") pour ajouter d'un coup les joueurs de l'équipe FRI4 (D2).
 * PIN provisoires 1018-1022, à changer ensuite depuis l'onglet Effectif.
 */
function seedD2bPlayers() {
  return seedTeam_(4, [
    ['MARCHESI', 'Laurent'],
    ['BLOQUET', 'Thierry'],
    ['BLOQUET', 'Aesane'],
    ['HUE', 'Jeanne'],
    ['LARCHEVEQUE', 'Sacha'],
  ], 1018);
}

/**
 * À lancer UNE FOIS depuis l'éditeur Apps Script (menu Exécuter, fonction
 * "seedD3Players") pour ajouter d'un coup les joueurs de l'équipe FRI5 (D3).
 * PIN provisoires 1023-1026, à changer ensuite depuis l'onglet Effectif.
 */
function seedD3Players() {
  return seedTeam_(5, [
    ['DUJARDIN', 'Francois'],
    ['PESSON', 'Philippe'],
    ['LEMARCHAND', 'Jean-Noël'],
    ['TRECOURT', 'François'],
  ], 1023);
}

/**
 * À lancer UNE FOIS depuis l'éditeur Apps Script (menu Exécuter, fonction
 * "seedD4aPlayers") pour ajouter d'un coup les joueurs de l'équipe FRI6 (D4).
 * PIN provisoires 1027-1031, à changer ensuite depuis l'onglet Effectif.
 */
function seedD4aPlayers() {
  return seedTeam_(6, [
    ['HUE', 'Patrice'],
    ['PAPILLON-LEROUX', 'Robin'],
    ['LARCHEVEQUE', 'Gabin'],
    ['SCHMIT', 'Robin'],
    ['MATELOT', 'Matthis'],
  ], 1027);
}

/**
 * À lancer UNE FOIS depuis l'éditeur Apps Script (menu Exécuter, fonction
 * "seedD4bPlayers") pour ajouter d'un coup les joueurs de l'équipe FRI7 (D4).
 * PIN provisoires 1032-1036, à changer ensuite depuis l'onglet Effectif.
 */
function seedD4bPlayers() {
  return seedTeam_(7, [
    ['DUCHEMIN', 'Pierre-Louis'],
    ['TRUSSART', 'Xavier'],
    ['BUREL', 'Cyrille'],
    ['DELEAU', 'Christian'],
    ['BENARD', 'Fabien'],
  ], 1032);
}

/**
 * À lancer UNE FOIS depuis l'éditeur Apps Script (menu Exécuter, fonction
 * "seedD4cPlayers") pour ajouter d'un coup les joueurs de l'équipe FRI8 (D4).
 * PIN provisoires 1037-1039, à changer ensuite depuis l'onglet Effectif.
 */
function seedD4cPlayers() {
  return seedTeam_(8, [
    ['RADENAC', 'Alexandre'],
    ['MARUITTE', 'Kévin'],
    ['TOUPAIN', 'Allan'],
  ], 1037);
}

/**
 * Ajoute une liste de joueurs [ [Nom, Prénom], ... ] à l'équipe teamId,
 * avec des PIN provisoires séquentiels à partir de pinStart. Ne duplique
 * pas un joueur déjà présent (comparaison Nom+Prénom insensible aux
 * accents/casse) — on peut relancer la fonction sans risque.
 */
function seedTeam_(teamId, roster, pinStart) {
  ensureSheets_();
  const existing = readTable_('Joueurs');
  let added = 0;
  const pins = [];
  roster.forEach(function (row, idx) {
    const nom = row[0];
    const prenom = row[1];
    const pin = String(pinStart + idx);
    const already = existing.some(function (p) {
      return normalize_(p.Nom) === normalize_(nom) && normalize_(p.Prénom) === normalize_(prenom);
    });
    if (already) return;
    addPlayer(nom, prenom, teamId, pin);
    added++;
    pins.push(prenom + ' ' + nom + ' = ' + pin);
  });
  const msg = added + ' joueur(s) ajouté(s) sur ' + roster.length +
    ' (les autres existaient déjà).' + (pins.length ? ' PIN provisoires -> ' + pins.join(', ') : '');
  Logger.log(msg);
  return msg;
}
