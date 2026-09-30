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
  // Championnat Jeunes : équipes SANS brûlage (kind différent de 'FRI'),
  // avec leur propre jeu de dates (groupe E), réservées aux jeunes
  // (catégorie Poussin à Junior, ou catégorie non renseignée).
  // (Les identifiants 9 à 12 restent réservés : la Coupe de Rouen, d'abord
  // gérée ici, se gère désormais dans les compétitions individuelles, avec
  // choix de l'équipe CDR1-CDR4 à la validation de l'inscription.)
  { id: 13, name: 'CJ1',  division: 'Championnat Jeunes', day: 'Samedi',   time: '', group: 'E', kind: 'CJ', youthOnly: true },
  { id: 14, name: 'CJ2',  division: 'Championnat Jeunes', day: 'Samedi',   time: '', group: 'E', kind: 'CJ', youthOnly: true },
  { id: 15, name: 'CJ3',  division: 'Championnat Jeunes', day: 'Samedi',   time: '', group: 'E', kind: 'CJ', youthOnly: true },
  { id: 16, name: 'CJ4',  division: 'Championnat Jeunes', day: 'Samedi',   time: '', group: 'E', kind: 'CJ', youthOnly: true },
];
TEAMS.forEach(function (t) { if (!t.kind) t.kind = 'FRI'; });
const DATE_GROUPS_ = ['A', 'B', 'C', 'E'];

// Dates de phase 1, par groupe (fournies par le club).
// Les dates de phase 2 sont éditables par l'administrateur (initialement vides).
const DEFAULT_DATES = {
  1: { // Phase 1
    A: ['20/09/2025', '04/10/2025', '18/10/2025', '08/11/2025', '22/11/2025', '06/12/2025', '13/12/2025'],
    B: ['25/09/2025', '09/10/2025', '30/10/2025', '13/11/2025', '27/11/2025', '11/12/2025', '18/12/2025'],
    C: ['02/10/2025', '16/10/2025', '06/11/2025', '20/11/2025', '04/12/2025', '08/01/2026', '15/01/2026'],
    // Championnat Jeunes (samedi) 2026-2027, d'après la plaquette du club.
    E: ['07/11/2026', '05/12/2026', '', '', '', '', ''],
  },
  2: { // Phase 2 - à compléter par l'administrateur
    A: ['', '', '', '', '', '', ''],
    B: ['', '', '', '', '', '', ''],
    C: ['', '', '', '', '', '', ''],
    E: ['16/01/2027', '06/02/2027', '13/03/2027', '22/05/2027', '20/06/2027', '', ''],
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
  sendConvocations: sendConvocations,
  answerConvocation: answerConvocation,
  getCompetitions: getCompetitions,
  requestInscriptions: requestInscriptions,
  cancelInscription: cancelInscription,
  confirmParticipation: confirmParticipation,
  getCompetitionsAdmin: getCompetitionsAdmin,
  setInscriptionStatus: setInscriptionStatus,
  setInscriptionEquipe: setInscriptionEquipe,
  saveCompetition: saveCompetition,
  deleteCompetition: deleteCompetition,
  setCompetitionSettings: setCompetitionSettings,
  getDataVersion: getDataVersion,
  saveRencontre: saveRencontre,
};

// Fonctions en lecture seule : toutes les autres modifient les données et
// font avancer le "numéro de version" des données (voir getDataVersion).
const READ_ONLY_FUNCTIONS_ = {
  login: true, getStaticConfig: true, getPlayerData: true, getAdminData: true, listPlayers: true,
  getCompetitions: true, getCompetitionsAdmin: true, getDataVersion: true,
};

/** Numéro de version des données, changé à chaque modification faite via
 * l'appli. Très rapide (aucune lecture de feuille) : l'appli l'interroge
 * toutes les 20 s et ne retélécharge les données que s'il a changé, au lieu
 * de tout retélécharger à chaque fois. (Une modification faite directement
 * dans le classeur n'est pas détectée : l'appli refait de toute façon un
 * téléchargement complet toutes les 5 minutes.) */
function getDataVersion() {
  return PropertiesService.getScriptProperties().getProperty('DATA_VERSION') || '0';
}

// Verrou global du script, réentrant (une fonction déjà sous verrou peut en
// appeler une autre qui le demande aussi).
let lockDepth_ = 0;
function withScriptLock_(fn) {
  if (lockDepth_ > 0) return fn();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  lockDepth_++;
  try {
    return fn();
  } finally {
    lockDepth_--;
    lock.releaseLock();
  }
}

/** Réécrit une feuille (en-tête + lignes) SANS la vider d'abord : les
 * nouvelles lignes remplacent les anciennes, puis seules les lignes en trop
 * à la fin sont effacées. Une lecture simultanée ne voit donc jamais une
 * feuille vide (ce qui faisait « disparaître » puis réapparaître des
 * affectations au hasard dans le tableau). */
function rewriteTable_(sheet, header, rows) {
  const n = header.length;
  const all = [header].concat(rows.map(function (r) {
    const row = r.slice(0, n);
    while (row.length < n) row.push('');
    return row;
  }));
  sheet.getRange(1, 1, all.length, n).setValues(all);
  const last = sheet.getLastRow();
  if (last > all.length) {
    sheet.getRange(all.length + 1, 1, last - all.length, Math.max(n, sheet.getLastColumn())).clearContent();
  }
}

function bumpDataVersion_() {
  PropertiesService.getScriptProperties().setProperty('DATA_VERSION', String(Date.now()));
}

function doPost(e) {
  let out;
  try {
    const body = JSON.parse(e.postData.contents);
    const fn = API_FUNCTIONS_[body.fn];
    if (!fn) {
      throw new Error('Fonction inconnue : ' + body.fn);
    }
    const args = body.args || [];
    // Toutes les modifications passent l'une après l'autre (verrou du
    // script) : deux capitaines / admins qui enregistrent au même moment ne
    // peuvent plus écraser mutuellement leurs saisies.
    const result = READ_ONLY_FUNCTIONS_[body.fn]
      ? fn.apply(null, args)
      : withScriptLock_(function () { return fn.apply(null, args); });
    if (!READ_ONLY_FUNCTIONS_[body.fn]) bumpDataVersion_();
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
const SHEETS_CHECK_CACHE_KEY_ = 'SHEETS_CHECKED_V7';

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
      DATE_GROUPS_.forEach(function (grp) {
        DEFAULT_DATES[phase][grp].forEach(function (d, idx) {
          rows.push([phase, grp, idx + 1, d, '', '']);
        });
      });
    });
    datesSheet.getRange(2, 1, rows.length, 6).setValues(rows);
  }
  // Ajout des groupes de dates D (Coupe de Rouen) et E (Championnat Jeunes)
  // dans un classeur créé avant leur existence : lignes ajoutées à la fin,
  // sans modifier les dates existantes.
  if (datesSheet.getLastRow() > 1) {
    const presentGroups = {};
    datesSheet.getRange(2, 1, datesSheet.getLastRow() - 1, 2).getValues().forEach(function (r) {
      presentGroups[Number(r[0]) + '|' + r[1]] = true;
    });
    const extra = [];
    [1, 2].forEach(function (phase) {
      DATE_GROUPS_.forEach(function (grp) {
        if (presentGroups[phase + '|' + grp]) return;
        DEFAULT_DATES[phase][grp].forEach(function (d, idx) { extra.push([phase, grp, idx + 1, d, '', '']); });
      });
    });
    if (extra.length) {
      const start = datesSheet.getLastRow() + 1;
      datesSheet.getRange(start, 4, extra.length, 1).setNumberFormat('@');
      datesSheet.getRange(start, 1, extra.length, 6).setValues(extra);
    }
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
    players.getRange(1, 1, 1, 8).setValues([['ID', 'Nom', 'Prénom', 'EquipeDomicile', 'PIN', 'Actif', 'Capitaine', 'Téléphone']]);
    // Colonne PIN (E) toujours en texte, pour ne jamais perdre un zéro de tête (ex. "0042").
    players.getRange(2, 5, 998, 1).setNumberFormat('@');
    players.setFrozenRows(1);
  }
  // Réparation : classeur créé par une version antérieure sans la colonne
  // "Capitaine" (7e colonne) — on l'ajoute sans toucher aux données existantes.
  if (players.getLastColumn() < 7) {
    players.getRange(1, 7).setValue('Capitaine');
  }
  // Réparation : colonne "Téléphone" (8e colonne, pour les convocations
  // par SMS) absente des classeurs créés par une version antérieure.
  if (players.getLastColumn() < 8) {
    players.getRange(1, 8).setValue('Téléphone');
  }
  // Toujours en texte : sinon Sheets transforme "0612345678" en nombre et
  // supprime le zéro de tête.
  players.getRange(2, 8, 998, 1).setNumberFormat('@');

  // Réparation : colonne "Catégorie" (9e colonne : Poussin, Benjamin...,
  // Senior, V40...) pour les compétitions individuelles. On ajoute
  // uniquement l'en-tête : les lignes des joueurs existants ne sont jamais
  // modifiées.
  if (players.getLastColumn() < 9) {
    players.getRange(1, 9).setValue('Catégorie');
  }

  // Réglages des compétitions individuelles, ajoutés à la feuille Config
  // s'ils n'existent pas encore (sans toucher aux autres réglages).
  const cfgKeys = cfg.getRange(1, 1, Math.max(cfg.getLastRow(), 1), 1).getValues().map(function (r) { return r[0]; });
  [['DelaiInscriptionJours', '15'], ['RappelJours', '10']].forEach(function (kv) {
    if (cfgKeys.indexOf(kv[0]) < 0) {
      const row = cfg.getLastRow() + 1;
      cfg.getRange(row, 1, 1, 2).setNumberFormat('@').setValues([kv]);
    }
  });

  // --- Compétitions individuelles ---
  // Une ligne par épreuve (tour, journée...). DateLimite vide = date de
  // l'épreuve moins DelaiInscriptionJours (feuille Config).
  let compet = ss.getSheetByName('Competitions');
  if (!compet) {
    compet = ss.insertSheet('Competitions');
    compet.getRange(1, 1, 1, 10).setValues([COMPET_HEADERS_]);
    compet.getRange(2, 5, 998, 1).setNumberFormat('@');
    compet.getRange(2, 7, 998, 1).setNumberFormat('@');
    compet.setFrozenRows(1);
    const seed = defaultCompetitions_();
    compet.getRange(2, 1, seed.length, 10).setValues(seed);
  }

  // --- Inscriptions ---
  let insc = ss.getSheetByName('Inscriptions');
  if (!insc) {
    insc = ss.insertSheet('Inscriptions');
    insc.getRange(1, 1, 1, 7).setValues([['CompetitionID', 'JoueurID', 'Statut', 'DemandeLe', 'Participation', 'ParticipationLe', 'Equipe']]);
    insc.setFrozenRows(1);
  }
  // Colonne G "Equipe" (ex. CDR2 pour la Coupe de Rouen) ajoutée aux
  // classeurs existants : en-tête seulement.
  if (insc.getLastColumn() < 7) {
    insc.getRange(1, 7).setValue('Equipe');
  }

  // --- Rencontres ---
  // Adversaire et lieu (Domicile = Isneauville / Extérieur) de chaque match
  // des équipes FRI, utilisés dans le message de convocation. Pré-rempli
  // avec les calendriers de poule de la phase 1 2026-2027 (FRI2 à FRI6) ;
  // complété / corrigé depuis la carte « Envoyer les convocations ».
  let renc = ss.getSheetByName('Rencontres');
  if (!renc) {
    renc = ss.insertSheet('Rencontres');
    renc.getRange(1, 1, 1, 6).setValues([['Phase', 'Equipe', 'Date', 'Heure', 'Adversaire', 'Lieu']]);
    renc.getRange(2, 3, 998, 2).setNumberFormat('@');
    renc.setFrozenRows(1);
    const seedRenc = defaultRencontres_();
    renc.getRange(2, 1, seedRenc.length, 6).setValues(seedRenc);
  } else {
    // Feuille déjà existante : on ajoute seulement les matchs du calendrier
    // intégré qui n'y sont pas encore (nouvelles poules), sans jamais
    // modifier ni écraser les lignes existantes.
    const existing = {};
    if (renc.getLastRow() > 1) {
      renc.getRange(2, 1, renc.getLastRow() - 1, 3).getValues().forEach(function (r) {
        let d = r[2];
        if (d instanceof Date) d = Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        existing[Number(r[0]) + '|' + Number(r[1]) + '|' + d] = true;
      });
    }
    const missing = defaultRencontres_().filter(function (r) { return !existing[r[0] + '|' + r[1] + '|' + r[2]]; });
    if (missing.length) {
      const start = renc.getLastRow() + 1;
      renc.getRange(start, 3, missing.length, 2).setNumberFormat('@');
      renc.getRange(start, 1, missing.length, 6).setValues(missing);
    }
  }

  // --- Convocations ---
  // Une ligne par joueur convoqué à un match (JoueurID + Phase + Date),
  // avec l'équipe, le message envoyé et la réponse du joueur.
  let convoc = ss.getSheetByName('Convocations');
  if (!convoc) {
    convoc = ss.insertSheet('Convocations');
    convoc.getRange(1, 1, 1, 8).setValues([['JoueurID', 'Phase', 'Date', 'Equipe', 'EnvoyeeLe', 'Message', 'Reponse', 'ReponduLe']]);
    convoc.getRange(2, 3, 998, 1).setNumberFormat('@');
    convoc.setFrozenRows(1);
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
  const dates = { 1: { A: [], B: [], C: [], E: [] }, 2: { A: [], B: [], C: [], E: [] } };
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
  // rencontres : "phase|équipe|date" -> { adversaire, lieu, heure }
  const rencontres = {};
  const rencSheet = ss_().getSheetByName('Rencontres');
  if (rencSheet && rencSheet.getLastRow() > 1) {
    rencSheet.getRange(2, 1, rencSheet.getLastRow() - 1, 6).getValues().forEach(function (r) {
      let d = r[2];
      if (d instanceof Date) d = Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      if (!r[0] || !r[1] || !d) return;
      let h = r[3];
      if (h instanceof Date) h = Utilities.formatDate(h, Session.getScriptTimeZone(), 'HH:mm');
      rencontres[Number(r[0]) + '|' + Number(r[1]) + '|' + d] = { adversaire: String(r[4] || ''), lieu: String(r[5] || ''), heure: String(h || '') };
    });
  }
  return { teams: TEAMS, dates: dates, journeeByDate: journeeByDate, rencontres: rencontres };
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
    // Coupe de Rouen / Championnat Jeunes : pas de brûlage, et leurs matchs
    // ne comptent pas pour le brûlage des équipes FRI.
    if (team.kind !== 'FRI') { burned[team.id] = false; return; }
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
    categorie: String((readTable_('Joueurs').find(function (p) { return String(p.ID) === String(playerId); }) || {}).Catégorie || ''),
    // Toutes phases confondues : le joueur doit voir sa convocation quel
    // que soit l'onglet de phase affiché.
    convocations: readTable_('Convocations')
      .filter(function (r) { return String(r.JoueurID) === String(playerId); })
      .map(convocationToClient_),
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

  const convocByPlayer = {};
  readTable_('Convocations')
    .filter(function (r) { return Number(r.Phase) === Number(phase); })
    .forEach(function (r) {
      convocByPlayer[r.JoueurID] = convocByPlayer[r.JoueurID] || {};
      convocByPlayer[r.JoueurID][r.Date] = convocationToClient_(r);
    });

  const result = players
    .filter(function (p) { return p.Actif !== false && p.Actif !== 'FAUX' && p.Actif !== 'NON'; })
    .map(function (p) {
      return {
        id: p.ID,
        nom: p.Nom,
        prenom: p.Prénom,
        homeTeam: Number(p.EquipeDomicile) || null,
        categorie: String(p.Catégorie || ''),
        availability: availByPlayer[p.ID] || {},
        assignments: assignByPlayer[p.ID] || {},
        matchesByTeam: matchesByPlayerTeam[p.ID] || {},
        burned: computeBurnedTeams_(matchesByPlayerTeam[p.ID] || {}),
        phone: String(p.Téléphone || ''),
        convocations: convocByPlayer[p.ID] || {},
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
  return withScriptLock_(dedupeAffectations_Now_);
}

function dedupeAffectations_Now_() {
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

  rewriteTable_(sheet, header.slice(0, 4), kept);
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
  return withScriptLock_(dedupeDisponibilites_Now_);
}

function dedupeDisponibilites_Now_() {
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

  rewriteTable_(sheet, header.slice(0, 4), kept);
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
  rewriteTable_(sheet, header.slice(0, 4), kept);

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
// RENCONTRES (adversaire / domicile ou extérieur)
// ----------------------------------------------------------------------------------

/** Calendriers de poule phase 1 2026-2027 (FFTT) des équipes FRI1 à FRI7
 * (FRI8 : poule pas encore connue). */
function defaultRencontres_() {
  return [
    [1, 1, '20/09/2026', '09:00', 'CP QUEVILLAIS 5', 'Extérieur'],
    [1, 1, '04/10/2026', '09:00', 'EVREUX EC 2', 'Domicile'],
    [1, 1, '18/10/2026', '09:00', 'Safran NS 1', 'Extérieur'],
    [1, 1, '08/11/2026', '09:00', 'RACING CLUB PORT HAVRE 1', 'Domicile'],
    [1, 1, '22/11/2026', '09:00', 'AS ST ETIENNE ROUVRAY 1', 'Extérieur'],
    [1, 1, '06/12/2026', '09:00', 'ALCL GD QUEVILLY 4', 'Extérieur'],
    [1, 1, '13/12/2026', '09:00', 'AS HONG LANDIN 3', 'Domicile'],
    [1, 2, '25/09/2026', '20:00', 'BLAINVILLE CREV 1', 'Extérieur'],
    [1, 2, '09/10/2026', '20:00', 'ASC BONSECOURS 3', 'Domicile'],
    [1, 2, '30/10/2026', '20:00', 'US C BOIS GUILLAUME 5', 'Extérieur'],
    [1, 2, '13/11/2026', '20:00', 'S SOTTEVILLAIS 2', 'Domicile'],
    [1, 2, '27/11/2026', '20:00', 'FRANQUEVILLE ST 4', 'Extérieur'],
    [1, 2, '11/12/2026', '20:00', 'ASM AMFREVILLE 1', 'Domicile'],
    [1, 2, '18/12/2026', '20:00', 'CP QUEVILLAIS 7', 'Extérieur'],
    [1, 3, '25/09/2026', '20:00', 'AA COURONNE 3', 'Extérieur'],
    [1, 3, '09/10/2026', '20:00', 'FRANQUEVILLE ST 6', 'Domicile'],
    [1, 3, '30/10/2026', '20:00', 'LE TRAIT YAINVI 2', 'Extérieur'],
    [1, 3, '13/11/2026', '20:00', 'S SOTTEVILLAIS 4', 'Domicile'],
    [1, 3, '27/11/2026', '20:00', 'ASTT 3', 'Extérieur'],
    [1, 3, '11/12/2026', '20:00', 'US C BOIS GUILL 9', 'Domicile'],
    [1, 3, '18/12/2026', '20:00', 'MT ST AIGNAN TT 5', 'Extérieur'],
    [1, 4, '25/09/2026', '20:00', 'JP VALLIQUERVIL 2', 'Domicile'],
    [1, 4, '09/10/2026', '20:00', 'TT DU CAILLY 1', 'Extérieur'],
    [1, 4, '30/10/2026', '20:00', 'ASM AMFREVILLE 2', 'Domicile'],
    [1, 4, '13/11/2026', '20:00', 'SAINT PIERRAISE ENT 10', 'Extérieur'],
    [1, 4, '27/11/2026', '20:00', 'CP YVETOT 5', 'Domicile'],
    [1, 4, '11/12/2026', '20:00', 'FJ FAUVILLE 2', 'Extérieur'],
    [1, 4, '18/12/2026', '20:00', 'SPO ROUEN 10', 'Domicile'],
    [1, 5, '02/10/2026', '20:00', 'ASC BONSECOURS 5', 'Domicile'],
    [1, 5, '16/10/2026', '20:00', 'TT DU CAILLY 3', 'Extérieur'],
    [1, 5, '06/11/2026', '20:00', 'AMFTT 4', 'Domicile'],
    [1, 5, '20/11/2026', '20:00', 'TT BULLY 5', 'Extérieur'],
    [1, 5, '04/12/2026', '20:00', 'LA CRIQUE TT 2', 'Domicile'],
    [1, 5, '08/01/2027', '20:00', 'AL DARNETAL 3', 'Domicile'],
    [1, 5, '15/01/2027', '20:00', 'FRANQUEVILLE ST 7', 'Extérieur'],
    [1, 6, '02/10/2026', '20:00', 'CP BUCHY (6)', 'Extérieur'],
    [1, 6, '16/10/2026', '20:00', 'CP QUEVILLAIS 14', 'Domicile'],
    [1, 6, '06/11/2026', '20:00', 'G C O BIHOREL 2', 'Extérieur'],
    [1, 6, '20/11/2026', '20:00', 'LA CRIQUE TT 4', 'Domicile'],
    [1, 6, '04/12/2026', '20:00', 'TTT 1', 'Extérieur'],
    [1, 6, '08/01/2027', '20:00', 'AMFTT 7', 'Extérieur'],
    [1, 6, '15/01/2027', '20:00', 'MESNIL-ESNARDTT (3)', 'Domicile'],
    [1, 7, '02/10/2026', '20:00', 'CP BUCHY 7', 'Domicile'],
    [1, 7, '16/10/2026', '20:00', 'TT DU CAILLY 5', 'Extérieur'],
    [1, 7, '06/11/2026', '20:00', 'G C O BIHOREL 4', 'Domicile'],
    [1, 7, '20/11/2026', '20:00', 'TT BULLY 6', 'Extérieur'],
    [1, 7, '04/12/2026', '20:00', 'CAMA TT 8', 'Domicile'],
    [1, 7, '08/01/2027', '20:00', 'RAQUETTE NEUFCH 8', 'Domicile'],
    [1, 7, '15/01/2027', '20:00', 'AMFTT 6', 'Extérieur'],
  ];
}

/** Enregistre (ou corrige) l'adversaire et le lieu d'un match. */
function saveRencontre(phase, teamId, date, adversaire, lieu) {
  ensureSheets_();
  if (lieu && lieu !== 'Domicile' && lieu !== 'Extérieur') throw new Error('Lieu invalide.');
  const sheet = ss_().getSheetByName('Rencontres');
  const values = sheet.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < values.length; i++) {
    let d = values[i][2];
    if (d instanceof Date) d = Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    if (Number(values[i][0]) === Number(phase) && Number(values[i][1]) === Number(teamId) && d === date) { row = i + 1; break; }
  }
  if (row < 0) {
    row = sheet.getLastRow() + 1;
    const team = TEAMS.find(function (t) { return t.id === Number(teamId); }) || {};
    sheet.getRange(row, 3, 1, 2).setNumberFormat('@');
    sheet.getRange(row, 1, 1, 4).setValues([[Number(phase), Number(teamId), date, team.time || '']]);
  }
  sheet.getRange(row, 5, 1, 2).setValues([[String(adversaire || ''), String(lieu || '')]]);
  return getStaticConfig();
}

// ----------------------------------------------------------------------------------
// CONVOCATIONS
// ----------------------------------------------------------------------------------

function convocationToClient_(r) {
  return {
    phase: Number(r.Phase),
    date: r.Date,
    teamId: Number(r.Equipe),
    sentAt: r.EnvoyeeLe instanceof Date ? r.EnvoyeeLe.toISOString() : String(r.EnvoyeeLe || ''),
    message: String(r.Message || ''),
    response: String(r.Reponse || ''),
  };
}

/**
 * Convoque à un match (phase + date + équipe) tous les joueurs affectés à
 * cette équipe ce jour-là (feuille Affectations). Enregistre une ligne par
 * joueur dans "Convocations" (visible dans l'appli du joueur) et renvoie la
 * liste des joueurs avec leur téléphone, pour l'envoi des SMS depuis le
 * téléphone du capitaine / de l'administrateur.
 * Une nouvelle convocation pour le même joueur à la même date remplace la
 * précédente (et efface sa réponse).
 */
function sendConvocations(phase, date, teamId, message) {
  ensureSheets_();
  let selected;
  withScriptLock_(function () {
    selected = readTable_('Affectations').filter(function (r) {
      return Number(r.Phase) === Number(phase) && r.Date === date && Number(r.EquipeJouee) === Number(teamId);
    });
    if (!selected.length) {
      throw new Error('Aucun joueur affecté à cette équipe le ' + date + '. Enregistrez d\'abord la feuille de match.');
    }
    const selectedIds = {};
    selected.forEach(function (r) { selectedIds[String(r.JoueurID)] = true; });

    const sheet = ss_().getSheetByName('Convocations');
    const values = sheet.getDataRange().getValues();
    const header = values.shift();
    const now = new Date();
    // On retire les anciennes convocations de cette date pour cette équipe
    // (joueurs retirés de la feuille de match depuis) et celles des joueurs
    // sélectionnés (remplacées ci-dessous).
    const kept = values.filter(function (row) {
      if (!row.some(function (c) { return c !== '' && c !== null; })) return false;
      const sameMatch = Number(row[1]) === Number(phase) && row[2] === date;
      if (!sameMatch) return true;
      return !selectedIds[String(row[0])] && Number(row[3]) !== Number(teamId);
    });
    selected.forEach(function (r) {
      kept.push([r.JoueurID, Number(phase), date, Number(teamId), now, String(message || ''), '', '']);
    });
    sheet.getRange(2, 3, Math.max(kept.length, 1), 1).setNumberFormat('@');
    rewriteTable_(sheet, header.slice(0, 8), kept);
  });

  return selectedIdsToPlayers_(readTable_('Joueurs'), selected);
}

function selectedIdsToPlayers_(players, rows) {
  return rows.map(function (r) {
    const p = players.find(function (pl) { return String(pl.ID) === String(r.JoueurID); }) || {};
    return { id: r.JoueurID, name: (p.Prénom || '') + ' ' + (p.Nom || ''), phone: String(p.Téléphone || '') };
  });
}

/** Réponse du joueur à sa convocation : "Présent" ou "Absent". */
function answerConvocation(playerId, phase, date, response) {
  ensureSheets_();
  if (response !== 'Présent' && response !== 'Absent') {
    throw new Error('Réponse invalide.');
  }
  const sheet = ss_().getSheetByName('Convocations');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(playerId) && Number(values[i][1]) === Number(phase) && values[i][2] === date) {
      sheet.getRange(i + 1, 7, 1, 2).setValues([[response, new Date()]]);
      return { ok: true };
    }
  }
  throw new Error('Convocation introuvable (elle a peut-être été annulée).');
}

// ----------------------------------------------------------------------------------
// COMPÉTITIONS INDIVIDUELLES (jeunes / adultes)
// ----------------------------------------------------------------------------------

const COMPET_HEADERS_ = ['ID', 'Type', 'Nom', 'Epreuve', 'Date', 'Lieu', 'DateLimite', 'Categories', 'Tarif', 'Infos'];
const YOUTH_CATEGORIES_ = ['Poussin', 'Benjamin', 'Minime', 'Cadet', 'Junior'];
const ADULT_CATEGORIES_ = ['Senior', 'V40', 'V50', 'V60', 'V70', 'V80'];

/** Calendrier 2026-2027 (plaquettes « Compétitions jeunes / adultes » du
 * club), utilisé uniquement à la création de la feuille Competitions. */
function defaultCompetitions_() {
  const rows = [];
  let id = 0;
  function add(type, nom, epreuves, lieu, cats, tarif, infos) {
    epreuves.forEach(function (e) {
      rows.push([++id, type, nom, e[1], e[0], e[2] || lieu, '', cats, tarif, infos]);
    });
  }
  // --- Adultes ---
  add('Adultes', 'Critérium Fédéral', [
    ['11/10/2026', 'Tour 1'], ['15/11/2026', 'Tour 2'], ['31/01/2027', 'Tour 3'],
    ['07/03/2027', 'Tour 4'], ['28/03/2027', 'Finales départementales'],
  ], 'Communiqué avant chaque tour', '', '35 € la saison',
    'Le dimanche. 4 tours en poules puis tableau, montées et descentes de division (départementale, régionale, nationale).');
  add('Adultes', 'Coupe de Rouen', [
    ['14/10/2026', 'T1'], ['04/11/2026', 'T2'], ['25/11/2026', 'T3'], ['16/12/2026', 'T4'],
    ['06/01/2027', 'T5'], ['27/01/2027', 'T6'], ['17/02/2027', 'T7'], ['10/03/2027', 'T8'],
    ['31/03/2027', 'T9'], ['14/04/2027', 'T10'], ['05/05/2027', 'T11'], ['19/05/2027', 'Tournoi individuel'],
  ], 'District rouennais', '', '20 € la saison',
    'Le mercredi soir, clubs du district rouennais. 11 tours puis tournoi individuel en mai.');
  add('Adultes', 'Championnat individuel Vétérans', [
    ['19/12/2026', 'Tour 1', 'Rouen'], ['13/02/2027', 'Tour 2', 'Le Havre'],
    ['15/05/2027', 'Tour 3', 'Dieppe'], ['19/06/2027', 'Finale', 'Rouen'],
  ], '', 'V40,V50,V60,V70,V80', '',
    'Le samedi, 40 ans et plus. 3 tours en poules de 5-6 joueurs (2 tours minimum pour la finale). Double possible.');
  add('Adultes', 'Challenge SERANO', [
    ['06/10/2026', 'T1'], ['10/11/2026', 'T2'], ['15/12/2026', 'T3'], ['12/01/2027', 'T4'],
    ['09/02/2027', 'T5'], ['09/03/2027', 'T6'], ['06/04/2027', 'T7'], ['18/05/2027', 'T8'],
  ], 'District rouennais', 'V50,V60,V70,V80', '25 € la saison',
    'Le mardi, vétérans de plus de 55 ans du district rouennais. 8 tours, classement et podiums en fin de saison.');
  // --- Jeunes ---
  const c500 = [
    ['03/10/2026', 'T1', 'Le Havre'], ['28/11/2026', 'T2', 'Rouen'], ['12/12/2026', 'T3', 'Dieppe'],
    ['13/02/2027', 'T4', 'Rouen'], ['20/03/2027', 'T5', 'Le Havre'], ['10/04/2027', 'T6', 'Dieppe'],
  ];
  add('Jeunes', 'Challenge 500', c500, '', '', '6 € par tour',
    'Le samedi. Pour les débutants classés 500 points : beaucoup de matchs contre des joueurs de ton niveau.');
  add('Jeunes', 'Circuit Jeunes', c500, '', 'Poussin,Benjamin', '6 € par tour',
    'Le samedi. Poussins et benjamins 1re année (nés en 2017), tous niveaux.');
  add('Jeunes', 'Championnat Jeunes (équipe de 2)', [
    ['07/11/2026', 'J1'], ['05/12/2026', 'J2'], ['16/01/2027', 'J3'], ['06/02/2027', 'J4'],
    ['13/03/2027', 'J5'], ['22/05/2027', 'Finale district'], ['20/06/2027', 'Finale départementale'],
  ], 'Amfreville-la-Mi-Voie, Sotteville ou Saint-Étienne-du-Rouvray', '', '25 € par équipe',
    'Le samedi. Équipe de 2 (ou 3 avec un remplaçant) : 4 simples et 1 double.');
  add('Jeunes', 'Critérium Fédéral', [
    ['10/10/2026', 'Tour 1'], ['14/11/2026', 'Tour 2'], ['30/01/2027', 'Tour 3'],
    ['06/03/2027', 'Tour 4'], ['27/03/2027', 'Finales départementales'],
  ], 'Communiqué avant chaque tour', '', '20 € (poussin à cadet), 25 € (junior)',
    'Le samedi, par catégorie d\'âge, tous niveaux. Montées et descentes de division, porte d\'entrée vers les Championnats de Normandie.');
  add('Jeunes', 'Championnat de Seine-Maritime', [['19/12/2026', 'Journée']], 'Communiqué ultérieurement', '', '', '');
  // --- Tous ---
  add('Tous', 'Rassemblements féminins', [
    ['01/11/2026', 'Solidarité Cancer'], ['21/03/2027', 'Tournoi par équipes'], ['06/06/2027', 'Tournoi féminin'],
  ], 'Communiqué ultérieurement', '', '', 'Rendez-vous du CD76TT pour les joueuses, tous âges et tous niveaux.');
  return rows;
}

function readConfigValue_(key, fallback) {
  const row = readTable_('Config').find(function (r) { return r.Clé === key; });
  return row && String(row.Valeur) !== '' ? row.Valeur : fallback;
}

function competitionSettings_() {
  return {
    delaiJours: Number(readConfigValue_('DelaiInscriptionJours', 15)) || 0,
    rappelJours: Number(readConfigValue_('RappelJours', 10)) || 0,
  };
}

function parseFrDate_(s) {
  const m = String(s || '').match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  return m ? new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])) : null;
}

function formatFrDate_(d) {
  return pad2_(d.getDate()) + '/' + pad2_(d.getMonth() + 1) + '/' + d.getFullYear();
}

function readCompetitions_(settings) {
  return readTable_('Competitions').filter(function (r) { return r.ID !== '' && r.Date; }).map(function (r) {
    const date = parseFrDate_(r.Date);
    let deadline = r.DateLimite ? String(r.DateLimite) : '';
    let deadlineAuto = false;
    if (!deadline && date) {
      const d = new Date(date.getTime());
      d.setDate(d.getDate() - settings.delaiJours);
      deadline = formatFrDate_(d);
      deadlineAuto = true;
    }
    return {
      id: Number(r.ID), type: String(r.Type), nom: String(r.Nom), epreuve: String(r.Epreuve || ''),
      date: String(r.Date), lieu: String(r.Lieu || ''), dateLimite: deadline, dateLimiteAuto: deadlineAuto,
      categories: String(r.Categories || '').split(',').map(function (c) { return c.trim(); }).filter(Boolean),
      tarif: String(r.Tarif || ''), infos: String(r.Infos || ''),
      teams: competitionTeams_(String(r.Nom)),
    };
  });
}

// Compétitions individuelles jouées par équipes : l'administrateur choisit
// l'équipe de chaque inscrit validé.
const COMPETITION_TEAMS_ = {
  'Coupe de Rouen': ['CDR1', 'CDR2', 'CDR3', 'CDR4'],
};
function competitionTeams_(nom) {
  return COMPETITION_TEAMS_[nom] || [];
}

/** Admin : équipe (ex. CDR2) d'un inscrit, pour une épreuve. */
function setInscriptionEquipe(competitionId, playerId, equipe) {
  ensureSheets_();
  withInscriptionsSheet_(function (rows) {
    const row = rows.find(function (r) { return Number(r[0]) === Number(competitionId) && String(r[1]) === String(playerId); });
    if (!row) throw new Error('Inscription introuvable.');
    while (row.length < 7) row.push('');
    row[6] = String(equipe || '');
  });
  return getCompetitionsAdmin();
}

function readInscriptions_() {
  return readTable_('Inscriptions').map(function (r) {
    return {
      competitionId: Number(r.CompetitionID), playerId: r.JoueurID, statut: String(r.Statut || ''),
      demandeLe: r.DemandeLe instanceof Date ? r.DemandeLe.toISOString() : String(r.DemandeLe || ''),
      participation: String(r.Participation || ''),
      equipe: String(r.Equipe || ''),
    };
  });
}

function isEligible_(compet, categorie) {
  if (!categorie) return true; // catégorie non renseignée : on ne bloque pas
  const allowed = compet.categories.length ? compet.categories
    : compet.type === 'Jeunes' ? YOUTH_CATEGORIES_
    : compet.type === 'Adultes' ? ADULT_CATEGORIES_
    : YOUTH_CATEGORIES_.concat(ADULT_CATEGORIES_);
  return allowed.indexOf(categorie) >= 0;
}

function todayStart_() {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  return d;
}

/** Vue joueur : toutes les épreuves + ses inscriptions + réglages. */
function getCompetitions(playerId) {
  ensureSheets_();
  const settings = competitionSettings_();
  const player = readTable_('Joueurs').find(function (p) { return String(p.ID) === String(playerId); }) || {};
  const categorie = String(player.Catégorie || '');
  const mine = {};
  readInscriptions_().forEach(function (i) {
    if (String(i.playerId) === String(playerId)) mine[i.competitionId] = i;
  });
  return {
    settings: settings,
    categorie: categorie,
    competitions: readCompetitions_(settings).map(function (c) {
      c.eligible = isEligible_(c, categorie);
      return c;
    }),
    inscriptions: mine,
  };
}

function withInscriptionsSheet_(fn) {
  return withScriptLock_(function () {
    const sheet = ss_().getSheetByName('Inscriptions');
    const values = sheet.getDataRange().getValues();
    const header = values.shift();
    const rows = values.filter(function (row) { return row.some(function (c) { return c !== '' && c !== null; }); });
    const result = fn(rows);
    rewriteTable_(sheet, header, rows);
    return result;
  });
}

function findInscriptionRow_(rows, competitionId, playerId) {
  return rows.find(function (r) { return Number(r[0]) === Number(competitionId) && String(r[1]) === String(playerId); });
}

/** Demande d'inscription du joueur à une ou plusieurs épreuves (avant la
 * date limite de chacune). */
function requestInscriptions(playerId, competitionIds) {
  ensureSheets_();
  const data = getCompetitions(playerId);
  const byId = {};
  data.competitions.forEach(function (c) { byId[c.id] = c; });
  const today = todayStart_();
  const refused = [];
  withInscriptionsSheet_(function (rows) {
    (competitionIds || []).forEach(function (cid) {
      const c = byId[Number(cid)];
      if (!c) return;
      const limit = parseFrDate_(c.dateLimite);
      if (!c.eligible || (limit && limit < today)) { refused.push(c.nom + ' ' + c.epreuve); return; }
      const row = findInscriptionRow_(rows, c.id, playerId);
      if (row) {
        if (row[2] === 'Refusée' || row[2] === 'Annulée') { row[2] = 'Demandée'; row[3] = new Date(); row[4] = ''; row[5] = ''; }
      } else {
        rows.push([c.id, playerId, 'Demandée', new Date(), '', '']);
      }
    });
  });
  if (refused.length && refused.length === (competitionIds || []).length) {
    throw new Error('Inscription impossible (date limite dépassée ou catégorie non concernée) : ' + refused.join(', '));
  }
  return getCompetitions(playerId);
}

/** Le joueur retire sa demande (avant la date limite). */
function cancelInscription(playerId, competitionId) {
  ensureSheets_();
  withInscriptionsSheet_(function (rows) {
    const idx = rows.findIndex(function (r) { return Number(r[0]) === Number(competitionId) && String(r[1]) === String(playerId); });
    if (idx >= 0) rows.splice(idx, 1);
  });
  return getCompetitions(playerId);
}

/** Confirmation de participation par le joueur, une fois l'inscription
 * validée par le club : "Confirmée" ou "Annulée" (désistement). */
function confirmParticipation(playerId, competitionId, value) {
  ensureSheets_();
  if (value !== 'Confirmée' && value !== 'Annulée') throw new Error('Réponse invalide.');
  withInscriptionsSheet_(function (rows) {
    const row = findInscriptionRow_(rows, competitionId, playerId);
    if (!row) throw new Error('Inscription introuvable.');
    row[4] = value;
    row[5] = new Date();
  });
  return getCompetitions(playerId);
}

/** Vue administrateur : épreuves, toutes les inscriptions, joueurs. */
function getCompetitionsAdmin() {
  ensureSheets_();
  const settings = competitionSettings_();
  return {
    settings: settings,
    competitions: readCompetitions_(settings),
    inscriptions: readInscriptions_(),
    players: readTable_('Joueurs')
      .filter(function (p) { return p.Actif !== false && p.Actif !== 'FAUX' && p.Actif !== 'NON'; })
      .map(function (p) {
        return { id: p.ID, nom: p.Nom, prenom: p.Prénom, categorie: String(p.Catégorie || ''), phone: String(p.Téléphone || '') };
      }),
  };
}

/** Admin : valide / refuse une demande, ou inscrit directement un joueur
 * (statut "Validée"). statut vide = supprime l'inscription. */
function setInscriptionStatus(competitionId, playerId, statut) {
  ensureSheets_();
  if (['Demandée', 'Validée', 'Refusée', ''].indexOf(statut) < 0) throw new Error('Statut invalide.');
  withInscriptionsSheet_(function (rows) {
    const idx = rows.findIndex(function (r) { return Number(r[0]) === Number(competitionId) && String(r[1]) === String(playerId); });
    if (!statut) { if (idx >= 0) rows.splice(idx, 1); return; }
    if (idx >= 0) { rows[idx][2] = statut; }
    else { rows.push([Number(competitionId), playerId, statut, new Date(), '', '']); }
  });
  return getCompetitionsAdmin();
}

/** Admin : crée (id vide) ou modifie une épreuve. */
function saveCompetition(c) {
  ensureSheets_();
  const date = normalizeDateFr_(c.date);
  if (!date) throw new Error('La date de l\'épreuve est obligatoire.');
  const limite = normalizeDateFr_(c.dateLimite || '');
  if (!c.nom) throw new Error('Le nom de la compétition est obligatoire.');
  const sheet = ss_().getSheetByName('Competitions');
  const values = sheet.getDataRange().getValues();
  let rowIdx = -1;
  let id = Number(c.id) || 0;
  if (id) {
    for (let i = 1; i < values.length; i++) { if (Number(values[i][0]) === id) { rowIdx = i + 1; break; } }
  }
  if (rowIdx < 0) {
    id = values.slice(1).reduce(function (m, r) { return Math.max(m, Number(r[0]) || 0); }, 0) + 1;
    rowIdx = sheet.getLastRow() + 1;
  }
  const range = sheet.getRange(rowIdx, 1, 1, 10);
  range.setNumberFormat('@');
  range.setValues([[String(id), c.type || 'Adultes', c.nom, c.epreuve || '', date, c.lieu || '', limite,
    (c.categories || []).join(','), c.tarif || '', c.infos || '']]);
  return getCompetitionsAdmin();
}

function deleteCompetition(id) {
  ensureSheets_();
  const sheet = ss_().getSheetByName('Competitions');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (Number(values[i][0]) === Number(id)) { sheet.deleteRow(i + 1); break; }
  }
  withInscriptionsSheet_(function (rows) {
    for (let i = rows.length - 1; i >= 0; i--) { if (Number(rows[i][0]) === Number(id)) rows.splice(i, 1); }
  });
  return getCompetitionsAdmin();
}

function setCompetitionSettings(delaiJours, rappelJours) {
  ensureSheets_();
  const sheet = ss_().getSheetByName('Config');
  const values = sheet.getDataRange().getValues();
  const wanted = { DelaiInscriptionJours: String(Number(delaiJours) || 0), RappelJours: String(Number(rappelJours) || 0) };
  for (let i = 1; i < values.length; i++) {
    if (wanted[values[i][0]] !== undefined) {
      const cell = sheet.getRange(i + 1, 2);
      cell.setNumberFormat('@');
      cell.setValue(wanted[values[i][0]]);
    }
  }
  return getCompetitionsAdmin();
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

function addPlayer(nom, prenom, teamId, pin, capitaine, telephone, categorie) {
  ensureSheets_();
  assertPinAvailable_(pin);
  const sheet = ss_().getSheetByName('Joueurs');
  const players = readTable_('Joueurs');
  const maxId = players.reduce(function (m, p) { return Math.max(m, Number(p.ID) || 0); }, 0);
  const newId = maxId + 1;
  sheet.appendRow([newId, nom, prenom, teamId, pin, true, !!capitaine, String(telephone || ''), String(categorie || '')]);
  return { ok: true, id: newId };
}

function updatePlayer(id, nom, prenom, teamId, pin, actif, capitaine, telephone, categorie) {
  ensureSheets_();
  assertPinAvailable_(pin, id);
  const sheet = ss_().getSheetByName('Joueurs');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sheet.getRange(i + 1, 2, 1, 6).setValues([[nom, prenom, teamId, pin, actif, !!capitaine]]);
      // Téléphone (colonne H) : on ne l'écrase que s'il est fourni, pour ne
      // pas l'effacer si une ancienne version de l'appli appelle updatePlayer
      // sans ce paramètre.
      if (telephone !== undefined) {
        const phoneCell = sheet.getRange(i + 1, 8);
        phoneCell.setNumberFormat('@');
        phoneCell.setValue(String(telephone || ''));
      }
      if (categorie !== undefined) {
        sheet.getRange(i + 1, 9).setValue(String(categorie || ''));
      }
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
