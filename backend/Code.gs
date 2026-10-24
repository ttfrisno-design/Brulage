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
// Version du script (affichée au comité et dans l'admin) : permet de savoir
// quels clubs ont recollé la dernière version de Code.gs.
const REBONTT_VERSION = '2026.10.22';

// Équipes par défaut (celles du FRI). Chaque club règle les siennes dans
// l'onglet « Equipes » de son classeur (créé avec ces valeurs) : ID, Nom,
// Division, Jour, Heure, Groupe de dates (A, B ou C).
const DEFAULT_TEAMS_ = [
  { id: 1, name: 'FRI1', division: 'Régional 3',      day: 'Dimanche', time: '',      group: 'A' },
  { id: 2, name: 'FRI2', division: 'Départemental 1',  day: 'Vendredi', time: '20:00', group: 'B' },
  { id: 3, name: 'FRI3', division: 'Départemental 2',  day: 'Vendredi', time: '20:00', group: 'B' },
  { id: 4, name: 'FRI4', division: 'Départemental 2',  day: 'Vendredi', time: '20:00', group: 'B' },
  { id: 5, name: 'FRI5', division: 'Départemental 3',  day: 'Vendredi', time: '20:00', group: 'C' },
  { id: 6, name: 'FRI6', division: 'Départemental 4',  day: 'Vendredi', time: '20:00', group: 'C' },
  { id: 7, name: 'FRI7', division: 'Départemental 4',  day: 'Vendredi', time: '20:00', group: 'C' },
  { id: 8, name: 'FRI8', division: 'Départemental 4',  day: 'Vendredi', time: '20:00', group: 'C' },
  // (Identifiants 9 à 16 réservés : la Coupe de Rouen (CDR1-4) et le
  // Championnat Jeunes (CJ1-4), d'abord gérés ici, se gèrent désormais dans
  // les compétitions individuelles, avec choix de l'équipe à la validation
  // de l'inscription — voir COMPETITION_TEAMS_.)
];
const TEAMS = DEFAULT_TEAMS_.map(function (t) { return Object.assign({ kind: 'FRI' }, t); });
const TEAMS_HEADERS_ = ['ID', 'Nom', 'Division', 'Jour', 'Heure', 'Groupe'];
let teamsLoaded_ = false;
/** Équipes du club lues dans l'onglet « Equipes » (une fois par exécution) ;
 * à défaut, celles par défaut. */
function loadTeams_() {
  if (teamsLoaded_) return;
  teamsLoaded_ = true;
  const sh = ss_().getSheetByName('Equipes');
  if (!sh || sh.getLastRow() < 2) return;
  const list = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues().map(function (r) {
    return { id: Number(r[0]), name: String(r[1] || '').trim(), division: String(r[2] || '').trim(), day: String(r[3] || '').trim(),
      time: r[4] instanceof Date ? Utilities.formatDate(r[4], Session.getScriptTimeZone(), 'HH:mm') : String(r[4] || '').trim(),
      group: String(r[5] || '').trim().toUpperCase(), kind: 'FRI' };
  }).filter(function (t) { return t.id > 0 && t.name && ['A', 'B', 'C'].indexOf(t.group) >= 0; });
  if (!list.length) return;
  TEAMS.length = 0;
  list.sort(function (a, b) { return a.id - b.id; }).forEach(function (t) { TEAMS.push(t); });
}
const DATE_GROUPS_ = ['A', 'B', 'C'];

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
    .setTitle('RebonTT - Tennis de Table')
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
  setSeasonInscription: setSeasonInscription,
  saveCompetition: saveCompetition,
  setSeriesDeadline: setSeriesDeadline,
  deleteCompetition: deleteCompetition,
  setCompetitionSettings: setCompetitionSettings,
  getDataVersion: getDataVersion,
  getProfile: getProfile,
  setFfttLink: setFfttLink,
  setSetting: setSetting,
  setPlayerLink: setPlayerLink,
  getClubInfo: getClubInfo,
  uploadClubPhoto: uploadClubPhoto,
  setClubInfo: setClubInfo,
  saveRencontre: saveRencontre,
  registerPushToken: registerPushToken,
  createCheckout: createCheckout,
  checkMyPayments: checkMyPayments,
  getHelloAssoStatus: getHelloAssoStatus,
  unregisterPushToken: unregisterPushToken,
  getPushStatus: getPushStatus,
  sendTestPush: sendTestPush,
  // Site d'inscription au club (inscription.html)
  getAdhesionInfo: getAdhesionInfo,
  findAdherent: findAdherent,
  createAdhesion: createAdhesion,
  getAdhesion: getAdhesion,
  submitQS: submitQS,
  acceptReglement: acceptReglement,
  uploadAdhesionDocument: uploadAdhesionDocument,
  requestAdhesionChange: requestAdhesionChange,
  sendAdhesionMessage: sendAdhesionMessage,
  createAdhesionCheckout: createAdhesionCheckout,
  getAdhesionsAdmin: getAdhesionsAdmin,
  saveAdhesionSettings: saveAdhesionSettings,
  setAdhesionAdmin: setAdhesionAdmin,
  setAdhesionMessageStatut: setAdhesionMessageStatut,
  createPlayerFromAdhesion: createPlayerFromAdhesion,
  setChequesDeposes: setChequesDeposes,
  // Comité départemental
  getComiteExport: getComiteExport,
};

// Fonctions en lecture seule : toutes les autres modifient les données et
// font avancer le "numéro de version" des données (voir getDataVersion).
const READ_ONLY_FUNCTIONS_ = {
  login: true, getStaticConfig: true, getPlayerData: true, getAdminData: true, listPlayers: true,
  getCompetitions: true, getCompetitionsAdmin: true, getDataVersion: true, getProfile: true,
  getPushStatus: true, getHelloAssoStatus: true, getClubInfo: true,
  getAdhesionInfo: true, findAdherent: true, getAdhesionsAdmin: true, getComiteExport: true,
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
    // Notification envoyée par HelloAsso après un paiement : on ne fait pas
    // confiance à son contenu, on revérifie les paiements en attente
    // directement auprès de HelloAsso.
    if (!body.fn && body.eventType) {
      withScriptLock_(function () { checkPayments_(null); });
      bumpDataVersion_();
      return ContentService.createTextOutput('ok');
    }
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
const SHEETS_CHECK_CACHE_KEY_ = 'SHEETS_CHECKED_V21';

function ensureSheets_() {
  loadTeams_();
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

  // --- Equipes (ajout d'un onglet seulement, avec les équipes par défaut) ---
  if (!ss.getSheetByName('Equipes')) {
    const eq = ss.insertSheet('Equipes');
    eq.getRange(1, 1, 30, 6).setNumberFormat('@');
    eq.getRange(1, 1, DEFAULT_TEAMS_.length + 1, 6).setValues([TEAMS_HEADERS_].concat(DEFAULT_TEAMS_.map(function (t) {
      return [String(t.id), t.name, t.division, t.day, t.time, t.group];
    })));
    eq.setFrozenRows(1);
  }

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
  // Colonnes J "Licence" (n° de licence FFTT) et K "Points" (points FFTT),
  // pour le profil du joueur : en-têtes seulement.
  if (players.getLastColumn() < 10) {
    players.getRange(1, 10).setValue('Licence');
  }
  if (players.getLastColumn() < 11) {
    players.getRange(1, 11).setValue('Points');
  }
  players.getRange(2, 10, 998, 1).setNumberFormat('@');
  // Colonne L "Sexe" (F / M), pour les rassemblements féminins : en-tête
  // seulement, renseignée par l'admin dans l'effectif.
  if (players.getLastColumn() < 12) {
    players.getRange(1, 12).setValue('Sexe');
  }
  // Colonne M "LienEspace" : lien personnel du joueur vers son espace sur
  // l'application FFTT ou PingPocket (prioritaire sur le modèle commun).
  if (players.getLastColumn() < 13) {
    players.getRange(1, 13).setValue('LienEspace');
  }

  // Réglages des compétitions individuelles, ajoutés à la feuille Config
  // s'ils n'existent pas encore (sans toucher aux autres réglages).
  const cfgKeys = cfg.getRange(1, 1, Math.max(cfg.getLastRow(), 1), 1).getValues().map(function (r) { return r[0]; });
  [['DelaiInscriptionJours', '15'], ['RappelJours', '10'], ['LienFicheFFTT', ''],
    ['SiteClub', 'https://ttfrisno.wixsite.com/ttfri'],
    // Présentation publique du club (page affichée avant la connexion).
    ['ClubNom', 'FR Isneauville Tennis de Table'], ['ClubVille', 'Isneauville'],
    ['ClubAdresse', ''], ['ClubPhoto', ''], ['ClubEntrainements', ''], ['ClubTarifs', ''],
    ['ClubInscription', ''], ['ClubCreneaux', ''], ['ClubEntrainementsPhoto', ''],
    // Adresse de l'organisateur à qui envoyer les inscriptions, par
    // compétition (clé "MailInscription:<nom de la compétition>").
    ['MailInscription:Challenge SERANO', 'girard76@wanadoo.fr'],
    ['MailInscription:Coupe de Rouen', 'coupederouen@gmail.com'],
    ['MailInscription:Championnat Jeunes (équipe de 2)', 'lefebvrejmichel@hotmail.com'],
    // Tarifs du Critérium fédéral selon la catégorie du joueur (paiement
    // HelloAsso), clés "PrixHelloAsso:<compétition>@<catégorie>".
    ['PrixHelloAsso:Critérium Fédéral@Poussin', '20'], ['PrixHelloAsso:Critérium Fédéral@Benjamin', '20'],
    ['PrixHelloAsso:Critérium Fédéral@Minime', '20'], ['PrixHelloAsso:Critérium Fédéral@Cadet', '20'],
    ['PrixHelloAsso:Critérium Fédéral@Junior', '25'], ['PrixHelloAsso:Critérium Fédéral@Senior', '35'],
    ['PrixHelloAsso:Critérium Fédéral@V40', '35'], ['PrixHelloAsso:Critérium Fédéral@V50', '35'],
    ['PrixHelloAsso:Critérium Fédéral@V60', '35'], ['PrixHelloAsso:Critérium Fédéral@V70', '35'],
    ['PrixHelloAsso:Critérium Fédéral@V80', '35'],
  ].forEach(function (kv) {
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
  // Stages jeunes pendant les vacances (CDTT76) : ajoutés une seule fois, y
  // compris à une feuille existante (lignes ajoutées en bas, les autres ne
  // sont pas touchées) ; l'admin peut ensuite les modifier ou les supprimer.
  if (!readConfigValue_('StagesAjoutes', '')) {
    const existing = compet.getDataRange().getValues();
    if (!existing.slice(1).some(function (r) { return r[1] === 'Stages'; })) {
      let nextId = existing.slice(1).reduce(function (m, r) { return Math.max(m, Number(r[0]) || 0); }, 0);
      const stages = defaultStages_().map(function (r) { return [String(++nextId)].concat(r); });
      const range = compet.getRange(compet.getLastRow() + 1, 1, stages.length, 10);
      range.setNumberFormat('@');
      range.setValues(stages);
    }
    cfg.getRange(cfg.getLastRow() + 1, 1, 1, 2).setNumberFormat('@').setValues([['StagesAjoutes', 'oui']]);
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
  // Colonne H "Reglement" : date du règlement en ligne (inscription
  // automatique). En-tête seulement.
  if (insc.getLastColumn() < 8) {
    insc.getRange(1, 8).setValue('Reglement');
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
        if (d instanceof Date) d = cellDate_(d);
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

  // --- Appareils (notifications push) ---
  // Un téléphone par ligne : jeton Firebase de l'appli installée, rattaché
  // au joueur connecté dessus.
  if (!ss.getSheetByName('Appareils')) {
    const dev = ss.insertSheet('Appareils');
    dev.getRange(1, 1, 1, 4).setValues([['JoueurID', 'Jeton', 'Plateforme', 'MajLe']]);
    dev.setFrozenRows(1);
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
        return [cellDate_(v)];
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
  repairSwappedDates_();
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
        return [cellDate_(v)];
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
        return [cellDate_(v)];
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
          v = cellDate_(v);
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

/** Date d'une cellule sous forme "jj/mm/aaaa", que Sheets l'ait gardée en
 * texte ou convertie en vraie date. Sans cette normalisation, une ligne dont
 * la date avait été convertie n'était jamais retrouvée : chaque clic
 * ajoutait une nouvelle ligne (doublons dans Disponibilites), et
 * l'affichage, qui lit la dernière ligne, ne changeait plus. */
function cellDate_(v) {
  if (v instanceof Date) {
    // Une cellule devenue "vraie date" vient d'un texte jj/mm/aaaa que Google
    // Sheets a interprété selon les paramètres régionaux du classeur : en
    // format américain (mois en premier), "08/11/2026" a été lu comme le
    // 11 août. On reconstitue donc le texte d'origine dans le même ordre, et
    // dans le fuseau horaire du classeur (pas celui du script).
    const fmt = sheetIsMonthFirst_() ? 'MM/dd/yyyy' : 'dd/MM/yyyy';
    return Utilities.formatDate(v, sheetTimeZone_(), fmt);
  }
  return String(v === null || v === undefined ? '' : v).trim();
}

let sheetLocaleCache_ = null;
function sheetLocaleInfo_() {
  if (!sheetLocaleCache_) {
    const ss = ss_();
    const locale = String((ss.getSpreadsheetLocale && ss.getSpreadsheetLocale()) || '');
    sheetLocaleCache_ = {
      // Paramètres régionaux qui écrivent les dates mois / jour / année.
      monthFirst: /^en(_US|_PH|_CA)?$/.test(locale) || locale === 'en_US',
      tz: (ss.getSpreadsheetTimeZone && ss.getSpreadsheetTimeZone()) || Session.getScriptTimeZone(),
    };
  }
  return sheetLocaleCache_;
}
function sheetIsMonthFirst_() { return sheetLocaleInfo_().monthFirst; }
function sheetTimeZone_() { return sheetLocaleInfo_().tz; }

/** Répare les dates jour/mois inversées (ex. "11/08/2026" au lieu de
 * "08/11/2026") laissées dans Disponibilites, Affectations et Convocations
 * par d'anciennes conversions : une date qui n'est aucune date de match de
 * la feuille Dates, mais dont l'inverse l'est, est remplacée par l'inverse.
 * Sans effet (ni écriture) si tout est déjà correct. */
function repairSwappedDates_() {
  const ss = ss_();
  const datesSheet = ss.getSheetByName('Dates');
  if (!datesSheet || datesSheet.getLastRow() < 2) return;
  const known = {};
  datesSheet.getRange(2, 4, datesSheet.getLastRow() - 1, 1).getValues().forEach(function (r) {
    const d = cellDate_(r[0]);
    if (d) known[d] = true;
  });
  ['Disponibilites', 'Affectations', 'Convocations'].forEach(function (name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return;
    const range = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1);
    const vals = range.getValues();
    let changed = false;
    const fixed = vals.map(function (r) {
      const d = cellDate_(r[0]);
      const m = d.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
      if (m && !known[d]) {
        const swapped = m[2] + '/' + m[1] + '/' + m[3];
        if (known[swapped]) { changed = true; return [swapped]; }
      }
      if (r[0] instanceof Date) { changed = true; return [d]; }
      return [r[0]];
    });
    if (changed) {
      withScriptLock_(function () {
        range.setNumberFormat('@');
        range.setValues(fixed);
      });
    }
  });
}

function rowKey_(row) {
  return String(row[0]) + '|' + Number(row[1]) + '|' + cellDate_(row[2]);
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
    return { ok: true, isAdmin: true, name: 'Administrateur', config: getStaticConfig(), adminToken: adminToken_() };
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

// ----------------------------------------------------------------------------------
// PRÉSENTATION DU CLUB (publique, avant connexion)
// ----------------------------------------------------------------------------------
const CLUB_INFO_KEYS_ = ['ClubNom', 'ClubVille', 'ClubAdresse', 'ClubPhoto', 'ClubEntrainements', 'ClubTarifs', 'ClubInscription',
  'ClubCreneaux', 'ClubEntrainementsPhoto'];
const JOURS_ = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi', 'Dimanche'];

/** Créneaux d'entraînement (JSON dans Config) : jour, horaires, niveau
 * (= activité, ex. « Jeunes confirmés »), et pour le site d'inscription :
 * public (Jeunes / Adultes / Tous), lieu, encadrant, tarif, séances par
 * semaine (2 = créneau lié à l'autre créneau du même niveau : on s'inscrit
 * aux deux), part licence FFTT et places. */
function parseCreneaux_(raw) {
  let list = [];
  try { list = JSON.parse(String(raw || '[]')); } catch (e) { list = []; }
  const num = function (v) { return Number(String(v === undefined || v === null ? '' : v).replace(',', '.')) || 0; };
  return (Array.isArray(list) ? list : []).filter(function (c) { return c && (c.jour || c.debut || c.niveau); }).map(function (c, i) {
    return {
      id: String(c.id || ('s' + i)), jour: String(c.jour || ''), debut: String(c.debut || ''), fin: String(c.fin || ''), niveau: String(c.niveau || ''),
      public: c.public === 'Jeunes' || c.public === 'Adultes' ? c.public : 'Tous',
      lieu: String(c.lieu || ''), encadrant: String(c.encadrant || ''),
      tarif: Math.max(0, num(c.tarif)), seances: Number(c.seances) === 2 ? 2 : 1, licence: Math.max(0, num(c.licence)),
      places: Math.max(0, Math.floor(num(c.places))),
      // Créneau de compétition (adultes : PPS au lieu du questionnaire
      // « licence découverte »). Non renseigné : déduit du nom du niveau.
      competition: c.competition === undefined || c.competition === null || c.competition === ''
        ? /comp[ée]t/i.test(String(c.niveau || '')) : (c.competition === true || c.competition === 'oui' || c.competition === 'true'),
    };
  });
}

/** Présentation du club : nom, ville, adresse et photo de la salle,
 * créneaux d'entraînement, tarifs, lien d'inscription et équipes engagées.
 * Accessible sans code (aucune donnée personnelle). */
function getClubInfo() {
  ensureSheets_();
  const v = function (k) { return String(readConfigValue_(k, '') || ''); };
  return {
    nom: v('ClubNom'), ville: v('ClubVille'), adresse: v('ClubAdresse'), photo: v('ClubPhoto'),
    entrainements: v('ClubEntrainements'), tarifs: v('ClubTarifs'), inscription: v('ClubInscription'),
    creneaux: parseCreneaux_(v('ClubCreneaux')), entrainementsPhoto: v('ClubEntrainementsPhoto'),
    site: v('SiteClub'),
    version: REBONTT_VERSION,
    adhesion: { ouvert: !/^(non|0|false|fermé|ferme)$/i.test(v('AdhOuvert') || 'oui'), saison: v('AdhSaison') || defaultSaison_() },
    teams: TEAMS.filter(function (t) { return t.kind === 'FRI'; })
      .map(function (t) { return { name: t.name, division: t.division, day: t.day, time: t.time }; }),
  };
}

/** Admin : enregistre la présentation du club (clés Club* de Config). */
function setClubInfo(info) {
  ensureSheets_();
  info = info || {};
  const map = { nom: 'ClubNom', ville: 'ClubVille', adresse: 'ClubAdresse', photo: 'ClubPhoto',
    entrainements: 'ClubEntrainements', tarifs: 'ClubTarifs', inscription: 'ClubInscription',
    entrainementsPhoto: 'ClubEntrainementsPhoto' };
  if (info.creneaux !== undefined) {
    const stamp = Date.now();
    const list = parseCreneaux_(JSON.stringify((info.creneaux || []).map(function (c, i) {
      return c && !c.id ? Object.assign({}, c, { id: 'c' + (stamp + i).toString(36) }) : c;
    })));
    list.forEach(function (c) { if (c.jour && JOURS_.indexOf(c.jour) < 0) throw new Error('Jour invalide : ' + c.jour); });
    setSetting('ClubCreneaux', list.length ? JSON.stringify(list) : '');
  }
  Object.keys(map).forEach(function (k) {
    if (info[k] === undefined) return;
    const v = String(info[k] || '').trim();
    if ((k === 'photo' || k === 'inscription' || k === 'entrainementsPhoto') && v && !/^https?:\/\/\S+$/i.test(v)) throw new Error('Lien invalide (https://…) : ' + v);
    setSetting(map[k], v);
  });
  return getClubInfo();
}

/** Admin : photo de la salle (JPEG en base64, déjà réduite par l'appli),
 * enregistrée dans le Drive du club et partagée en lecture par lien. */
function uploadClubPhoto(base64, kind) {
  ensureSheets_();
  const data = String(base64 || '').replace(/^data:image\/\w+;base64,/, '');
  if (!data) throw new Error('Photo vide.');
  if (data.length > 4 * 1024 * 1024) throw new Error('Photo trop volumineuse.');
  const planning = kind === 'entrainements';
  const blob = Utilities.newBlob(Utilities.base64Decode(data), 'image/jpeg', planning ? 'planning-entrainements.jpg' : 'salle-club.jpg');
  const file = DriveApp.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1600';
  setSetting(planning ? 'ClubEntrainementsPhoto' : 'ClubPhoto', url);
  return { ok: true, photo: url };
}

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
      dateVal = cellDate_(dateVal);
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
      if (d instanceof Date) d = cellDate_(d);
      if (!r[0] || !r[1] || !d) return;
      let h = r[3];
      if (h instanceof Date) h = Utilities.formatDate(h, Session.getScriptTimeZone(), 'HH:mm');
      rencontres[Number(r[0]) + '|' + Number(r[1]) + '|' + d] = { adversaire: String(r[4] || ''), lieu: String(r[5] || ''), heure: String(h || '') };
    });
  }
  // Liens des poules (FFTT / PingPocket), par équipe et par phase :
  // clés Config « LienPoule:<équipe>:<phase> » -> { '<phase>|<équipe>': url }.
  const poolLinks = {};
  readTable_('Config').forEach(function (r) {
    const m = /^LienPoule:(\d+):(\d+)$/.exec(String(r.Clé || ''));
    if (m && String(r.Valeur || '').trim()) poolLinks[Number(m[2]) + '|' + Number(m[1])] = String(r.Valeur).trim();
  });
  return { teams: TEAMS, dates: dates, journeeByDate: journeeByDate, rencontres: rencontres, ffttLink: String(readConfigValue_('LienFicheFFTT', '') || ''),
    siteClub: String(readConfigValue_('SiteClub', '') || ''), poolLinks: poolLinks };
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
      // Nouveau code : les sessions administrateur ouvertes avec l'ancien
      // doivent se reconnecter pour les données des adhésions.
      PropertiesService.getScriptProperties().deleteProperty('ADMIN_TOKEN');
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
  return withScriptLock_(function () {
    upsertDatedRow_('Disponibilites', playerId, phase, date, status || '');
    return { ok: true };
  });
}

/** Écrit (ou supprime si value est vide) la ligne JoueurID+Phase+Date d'une
 * feuille à 4 colonnes (Disponibilites, Affectations). Toutes les lignes en
 * double pour cette clé sont supprimées, et la date est toujours écrite en
 * texte. */
function upsertDatedRow_(sheetName, playerId, phase, date, value) {
  const sheet = ss_().getSheetByName(sheetName);
  const values = sheet.getDataRange().getValues();
  const key = rowKey_([playerId, phase, date]);
  const matches = [];
  for (let i = 1; i < values.length; i++) {
    if (rowKey_(values[i]) === key) matches.push(i + 1);
  }
  // Doublons : on supprime en partant du bas pour ne pas décaler les autres.
  for (let k = matches.length - 1; k >= 1; k--) sheet.deleteRow(matches[k]);
  if (matches.length) {
    if (value === '' || value === null || value === undefined) {
      sheet.deleteRow(matches[0]);
    } else {
      const r = sheet.getRange(matches[0], 3, 1, 2);
      r.setNumberFormat('@');
      r.setValues([[date, value]]);
    }
  } else if (value !== '' && value !== null && value !== undefined) {
    const row = sheet.getLastRow() + 1;
    sheet.getRange(row, 3).setNumberFormat('@');
    sheet.getRange(row, 1, 1, 4).setValues([[playerId, Number(phase), date, value]]);
  }
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
        points: p.Points === '' || p.Points === undefined || p.Points === null ? '' : Number(p.Points) || '',
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
    row[2] = cellDate_(row[2]);
    const key = rowKey_(row);
    if (Object.prototype.hasOwnProperty.call(byKey, key)) { hadDuplicates = true; }
    byKey[key] = row; // garde la dernière occurrence rencontrée
  });
  if (!hadDuplicates) return false;

  const kept = Object.keys(byKey)
    .map(function (k) { return byKey[k]; })
    .filter(function (row) { return row[3] !== '' && row[3] !== null && row[3] !== undefined; });

  sheet.getRange(2, 3, Math.max(kept.length, 1), 1).setNumberFormat('@');
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
    row[2] = cellDate_(row[2]);
    const key = rowKey_(row);
    if (Object.prototype.hasOwnProperty.call(byKey, key)) { hadDuplicates = true; }
    byKey[key] = row; // garde la dernière occurrence rencontrée
  });
  if (!hadDuplicates) return false;

  const kept = Object.keys(byKey)
    .map(function (k) { return byKey[k]; })
    .filter(function (row) { return row[3] !== '' && row[3] !== null && row[3] !== undefined; });

  sheet.getRange(2, 3, Math.max(kept.length, 1), 1).setNumberFormat('@');
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

  withScriptLock_(function () {
    upsertDatedRow_('Affectations', playerId, phase, date, teamId ? Number(teamId) : '');
  });

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
  const raw = sheet.getDataRange().getValues();
  const header = raw.shift();

  // Une seule ligne par joueur + phase + date (la dernière l'emporte, comme
  // à l'affichage) : les doublons éventuels sont supprimés au passage.
  const rowIndex = {};
  const values = [];
  raw.forEach(function (row) {
    if (!row.some(function (c) { return c !== '' && c !== null; })) return;
    row[2] = cellDate_(row[2]);
    const key = rowKey_(row);
    if (rowIndex[key] !== undefined) { values[rowIndex[key]] = row; return; }
    rowIndex[key] = values.length;
    values.push(row);
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
    const key = rowKey_([e.playerId, e.phase, e.date]);
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
  sheet.getRange(2, 3, Math.max(kept.length, 1), 1).setNumberFormat('@');
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
    if (d instanceof Date) d = cellDate_(d);
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
  let answered = [];
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
    // Réponse déjà donnée par un joueur à une convocation précédente pour ce
    // même match (même date, même équipe) : on la garde quand la convocation
    // est renvoyée (composition modifiée, autre joueur ajouté…).
    const previous = {};
    values.forEach(function (row) {
      if (Number(row[1]) === Number(phase) && cellDate_(row[2]) === date && Number(row[3]) === Number(teamId) && row[6]) {
        previous[String(row[0])] = { reponse: row[6], le: row[7] };
      }
    });
    // On retire les anciennes convocations de cette date pour cette équipe
    // (joueurs retirés de la feuille de match depuis) et celles des joueurs
    // sélectionnés (remplacées ci-dessous).
    const kept = values.filter(function (row) {
      if (!row.some(function (c) { return c !== '' && c !== null; })) return false;
      const sameMatch = Number(row[1]) === Number(phase) && cellDate_(row[2]) === date;
      if (!sameMatch) return true;
      return !selectedIds[String(row[0])] && Number(row[3]) !== Number(teamId);
    });
    selected.forEach(function (r) {
      const prev = previous[String(r.JoueurID)];
      kept.push([r.JoueurID, Number(phase), date, Number(teamId), now, String(message || ''), prev ? prev.reponse : '', prev ? prev.le : '']);
    });
    answered = selected.filter(function (r) { return previous[String(r.JoueurID)]; }).map(function (r) { return String(r.JoueurID); });
    sheet.getRange(2, 3, Math.max(kept.length, 1), 1).setNumberFormat('@');
    rewriteTable_(sheet, header.slice(0, 8), kept);
  });

  const team = TEAMS.find(function (t) { return t.id === Number(teamId); }) || {};
  // Notification seulement aux joueurs qui n'ont pas encore répondu pour ce
  // match (ceux qui ont déjà répondu gardent leur réponse).
  sendPush_(selected.map(function (r) { return r.JoueurID; }).filter(function (id) { return answered.indexOf(String(id)) < 0; }),
    '📣 Convocation ' + (team.name || '') + ' — ' + date,
    String(message || '').split('\n')[0] || 'Vous êtes convoqué(e). Répondez dans l\'appli.');

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
    if (String(values[i][0]) === String(playerId) && Number(values[i][1]) === Number(phase) && cellDate_(values[i][2]) === date) {
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

/** Stages de tennis de table du comité (CDTT76) pendant les vacances
 * scolaires, réservés aux jeunes. Date = 1er jour du stage. */
function defaultStages_() {
  const nom = 'Stage de tennis de table';
  const infos = 'Stage du comité départemental (CDTT76) pendant les vacances scolaires, de 9h00 à 16h00 : ' +
    'enchaînements de travail technique et montées de table. Prévoir un panier repas pour le midi, une gourde et une serviette.';
  return [
    ['Vacances de la Toussaint : lundi 26, mardi 27 et mercredi 28 octobre (3 jours)', '26/10/2026', '19/10/2026'],
    ['Vacances de Noël : dimanche 20 et lundi 21 décembre (2 jours)', '20/12/2026', '13/12/2026'],
    ['Vacances d\'hiver : lundi 1er, mardi 2 et mercredi 3 mars (3 jours)', '01/03/2027', '22/02/2027'],
    ['Vacances de Pâques : lundi 26, mardi 27 et mercredi 28 avril (3 jours)', '26/04/2027', '19/04/2027'],
    ['Vacances d\'été : lundi 5, mardi 6, mercredi 7 et jeudi 8 juillet (4 jours)', '05/07/2027', '28/06/2027'],
  ].map(function (e) { return ['Stages', nom, e[0], e[1], '', e[2], '', '', infos]; });
}

function readConfigValue_(key, fallback) {
  const row = readTable_('Config').find(function (r) { return r.Clé === key; });
  return row && String(row.Valeur) !== '' ? row.Valeur : fallback;
}

function competitionSettings_() {
  const mails = {};
  const prices = { serie: {}, epreuve: {}, categorie: {} };
  const deadlines = {};
  readTable_('Config').forEach(function (r) {
    const k = String(r.Clé || '');
    // Date limite d'inscription d'une compétition (toutes ses épreuves) :
    // « DateLimite:<compétition> » = jj/mm/aaaa.
    if (k.indexOf('DateLimite:') === 0) {
      const v = r.Valeur instanceof Date ? formatFrDate_(r.Valeur) : normalizeDateFr_(String(r.Valeur || ''));
      if (v) deadlines[k.slice('DateLimite:'.length)] = v;
    }
    if (k.indexOf('MailInscription:') === 0) mails[k.slice('MailInscription:'.length)] = String(r.Valeur || '');
    // Montants du paiement HelloAsso (en euros) : « PrixHelloAsso:<compétition> »
    // (règlement unique pour la saison, ou montant par défaut d'une épreuve)
    // ou « PrixHelloAsso:#<id> » (une épreuve).
    if (k.indexOf('PrixHelloAsso:') === 0) {
      const v = Number(String(r.Valeur || '').replace(',', '.'));
      if (v > 0) {
        const rest = k.slice('PrixHelloAsso:'.length);
        const at = rest.lastIndexOf('@');
        if (rest.charAt(0) === '#') prices.epreuve[Number(rest.slice(1))] = v;
        else if (at > 0) { // tarif selon la catégorie du joueur
          const nom = rest.slice(0, at);
          (prices.categorie[nom] = prices.categorie[nom] || {})[rest.slice(at + 1)] = v;
        } else prices.serie[rest] = v;
      }
    }
  });
  return {
    delaiJours: Number(readConfigValue_('DelaiInscriptionJours', 15)) || 0,
    rappelJours: Number(readConfigValue_('RappelJours', 10)) || 0,
    mails: mails,
    prices: prices,
    deadlines: deadlines,
    helloasso: !!helloAssoConfig_(),
  };
}

// ----------------------------------------------------------------------------------
// PAIEMENT DES INSCRIPTIONS PAR HELLOASSO (Checkout API v5)
// ----------------------------------------------------------------------------------
// Le script crée un paiement HelloAsso (montant + nom de la compétition +
// date) et renvoie la page de paiement. C'est HelloAsso qui valide : tant
// que HelloAsso n'a pas confirmé le paiement, le joueur n'est pas inscrit.
// Configuration (Paramètres du projet › Propriétés du script) :
//   HELLOASSO_CLIENT_ID, HELLOASSO_CLIENT_SECRET : clés API (HelloAsso ›
//     Mon compte › Intégrations et API) ;
//   HELLOASSO_ORG : identifiant (slug) de l'association dans HelloAsso ;
//   HELLOASSO_SANDBOX = oui : environnement de test (facultatif) ;
//   APP_URL : page de retour après paiement (par défaut la version web).

function helloAssoConfig_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('HELLOASSO_CLIENT_ID');
  const secret = p.getProperty('HELLOASSO_CLIENT_SECRET');
  const org = p.getProperty('HELLOASSO_ORG');
  if (!id || !secret || !org) return null;
  const sandbox = /^(oui|1|true|yes)$/i.test(String(p.getProperty('HELLOASSO_SANDBOX') || ''));
  return {
    id: id, secret: secret, org: String(org).trim(),
    api: sandbox ? 'https://api.helloasso-sandbox.com' : 'https://api.helloasso.com',
    sandbox: sandbox,
    appUrl: p.getProperty('APP_URL') || 'https://ttfrisno-design.github.io/Brulage/',
  };
}

// Clé de cache liée aux identifiants : un changement de clés ou d'association
// (ex. passage sandbox → officiel) n'utilise jamais un ancien jeton.
function helloAssoTokenKey_(cfg) {
  return 'HA_TOKEN_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5, cfg.api + '|' + cfg.id + '|' + cfg.org));
}

function helloAssoToken_(cfg) {
  const cache = CacheService.getScriptCache();
  // Clé liée aux identifiants : un changement de clés ou d'association
  // (ex. passage sandbox → officiel) n'utilise jamais un ancien jeton.
  const key = helloAssoTokenKey_(cfg);
  const cached = cache.get(key);
  if (cached) return cached;
  const res = UrlFetchApp.fetch(cfg.api + '/oauth2/token', {
    method: 'post', muteHttpExceptions: true,
    payload: { grant_type: 'client_credentials', client_id: cfg.id, client_secret: cfg.secret },
  });
  const data = JSON.parse(res.getContentText() || '{}');
  if (!data.access_token) throw new Error('HelloAsso : connexion refusée (vérifiez les clés API).');
  cache.put(key, data.access_token, 1500); // valable 30 min
  return data.access_token;
}

function helloAssoApi_(cfg, method, path, body) {
  const opts = {
    method: method, muteHttpExceptions: true, contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + helloAssoToken_(cfg) },
  };
  if (body) opts.payload = JSON.stringify(body);
  const res = UrlFetchApp.fetch(cfg.api + '/v5/organizations/' + encodeURIComponent(cfg.org) + path, opts);
  const code = res.getResponseCode();
  const text = res.getContentText() || '{}';
  if (code === 401) CacheService.getScriptCache().remove(helloAssoTokenKey_(cfg));
  if (code === 401 || code === 403) {
    throw new Error('HelloAsso (' + code + ') : les clés API sont acceptées mais n\'ont pas accès à l\'association « ' + cfg.org +
      ' ». Vérifiez HELLOASSO_ORG (le nom après helloasso.com/associations/ dans l\'adresse de la page de l\'association) et que les clés ont été créées sur ce même compte association.');
  }
  if (code >= 300) throw new Error('HelloAsso (' + code + ') : ' + text.slice(0, 300));
  return JSON.parse(text);
}

/** Compétitions réglées épreuve par épreuve (Challenge 500, Circuit Jeunes,
 * stages ; les autres : règlement unique
 * qui vaut inscription à toutes les épreuves de la saison). */
function perEpreuvePayment_(nom) {
  return /challenge\s*500|circuit\s*jeune|stage/i.test(String(nom || ''));
}

/** Inscription unique pour toute la saison (Coupe de Rouen, Challenge SERANO) : impossible
 * dès que la date du 1er tour est passée. */
function singleInscription_(nom) {
  return /coupe\s*de\s*rouen|serano/i.test(String(nom || ''));
}
/** Date du 1er tour de la compétition (même nom, même type). */
function seriesFirstDate_(competitions, c) {
  let first = null;
  competitions.forEach(function (x) {
    if (x.nom !== c.nom || x.type !== c.type) return;
    const d = parseFrDate_(x.date);
    if (d && (!first || d < first)) first = d;
  });
  return first;
}
/** Inscription encore possible à cette épreuve aujourd'hui. */
function inscriptionOpen_(competitions, c, today) {
  const date = parseFrDate_(c.date);
  if (!c.eligible || !date || date < today) return false;
  if (lateInscriptionAllowed_(c.nom)) return true;
  if (singleInscription_(c.nom)) {
    // Inscription unique : date limite de la compétition si elle est
    // réglée, sinon le jour du 1er tour.
    const lim = parseFrDate_(c.dateLimiteSerie || '') || seriesFirstDate_(competitions, c);
    return !lim || lim >= today;
  }
  const limit = parseFrDate_(c.dateLimite);
  return !limit || limit >= today;
}

/** Inscription encore possible après la date limite (jusqu'au jour de
 * l'épreuve) : Critérium fédéral, Challenge 500, Circuit Jeunes. */
function lateInscriptionAllowed_(nom) {
  return /crit[ée]rium|challenge\s*500|circuit\s*jeune/i.test(String(nom || ''));
}

/** Montant à payer : épreuve (Challenge 500 / Circuit Jeunes), sinon tarif
 * de la catégorie du joueur (ex. Critérium fédéral), sinon montant de la
 * compétition. */
function priceOf_(c, prices, categorie) {
  const p = prices || {};
  return (perEpreuvePayment_(c.nom) && (p.epreuve || {})[c.id]) ||
    (((p.categorie || {})[c.nom] || {})[categorie]) ||
    (p.serie || {})[c.nom] || 0;
}

function ensurePaymentsSheet_() {
  let sh = ss_().getSheetByName('Paiements');
  if (!sh) {
    sh = ss_().insertSheet('Paiements');
    sh.getRange(1, 1, 1, 9).setValues([['IntentID', 'JoueurID', 'CompetitionIDs', 'Montant', 'Libelle', 'Statut', 'CreeLe', 'PayeLe', 'OrderID']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Le joueur demande à payer son inscription : crée le paiement HelloAsso
 * et renvoie l'adresse de la page de paiement. */
function createCheckout(playerId, competitionId) {
  ensureSheets_();
  const cfg = helloAssoConfig_();
  if (!cfg) throw new Error('Le paiement HelloAsso n\'est pas encore configuré par le club.');
  const data = getCompetitions(playerId);
  const c = data.competitions.find(function (x) { return x.id === Number(competitionId); });
  if (!c) throw new Error('Épreuve introuvable.');
  const today = todayStart_();
  if (!inscriptionOpen_(data.competitions, c, today)) {
    throw new Error(singleInscription_(c.nom) ? 'Inscription impossible : le 1er tour de « ' + c.nom + ' » est passé.'
      : 'Inscription impossible : date limite dépassée ou catégorie non concernée.');
  }
  const perEp = perEpreuvePayment_(c.nom);
  const targets = perEp ? [c] : data.competitions.filter(function (x) {
    const d = parseFrDate_(x.date);
    return x.nom === c.nom && x.type === c.type && x.eligible && d && d >= today;
  });
  const prices = data.settings.prices || { serie: {}, epreuve: {}, categorie: {} };
  const euros = priceOf_(c, prices, data.categorie);
  if (!euros) throw new Error('Montant non renseigné pour « ' + c.nom + ' » : demandez au club.');
  const player = readTable_('Joueurs').find(function (p) { return String(p.ID) === String(playerId); }) || {};
  const catPrice = ((prices.categorie || {})[c.nom] || {})[data.categorie];
  const label = ('Inscription ' + c.nom + (catPrice ? ' (' + data.categorie + ')' : '') + (perEp ? ' – ' + (c.epreuve ? c.epreuve + ' – ' : '') + c.date : ' – saison (' + targets.length + ' épreuve' + (targets.length > 1 ? 's' : '') + ')') +
    ' – ' + (player.Prénom || '') + ' ' + (player.Nom || '')).replace(/\s+/g, ' ').trim().slice(0, 250);
  const base = cfg.appUrl + (cfg.appUrl.indexOf('?') >= 0 ? '&' : '?');
  const cents = Math.round(Number(euros) * 100);
  const clean = function (v) { return String(v || '').replace(/[^A-Za-zÀ-ÖØ-öø-ÿ' -]/g, '').trim(); };
  const body = {
    totalAmount: cents, initialAmount: cents, itemName: label,
    backUrl: base + 'helloasso=retour', errorUrl: base + 'helloasso=erreur', returnUrl: base + 'helloasso=ok',
    containsDonation: false,
    metadata: { joueur: String(playerId), competitions: targets.map(function (x) { return x.id; }).join(',') },
  };
  const first = clean(player.Prénom), last = clean(player.Nom);
  if (first && last) body.payer = { firstName: first, lastName: last };
  const intent = helloAssoApi_(cfg, 'post', '/checkout-intents', body);
  if (!intent || !intent.redirectUrl) throw new Error('HelloAsso n\'a pas renvoyé de page de paiement.');
  withScriptLock_(function () {
    ensurePaymentsSheet_().appendRow([String(intent.id), playerId, targets.map(function (x) { return x.id; }).join(','), Number(euros), label, 'En attente', new Date(), '', '']);
  });
  return { redirectUrl: intent.redirectUrl, id: intent.id };
}

/** Interroge HelloAsso pour un paiement en attente ; s'il est payé,
 * inscrit (validé) le joueur aux épreuves concernées. */
function verifyCheckout_(cfg, rowIndex, row, sheet) {
  let intent;
  try { intent = helloAssoApi_(cfg, 'get', '/checkout-intents/' + encodeURIComponent(row[0])); }
  catch (e) { console.warn('Vérification HelloAsso impossible : ' + e.message); return row[5]; }
  const order = intent && intent.order;
  if (!order) {
    // Paiement jamais terminé : on le marque expiré au bout de 2 heures.
    const created = row[6] instanceof Date ? row[6].getTime() : Date.parse(row[6]);
    if (created && Date.now() - created > 2 * 3600 * 1000) {
      sheet.getRange(rowIndex, 6).setValue('Expiré');
      return 'Expiré';
    }
    return 'En attente';
  }
  const now = new Date();
  sheet.getRange(rowIndex, 6, 1, 4).setValues([['Payé', row[6], now, String(order.id || '')]]);
  // Adhésion au club (site d'inscription) : règlement du dossier.
  if (String(row[1]).indexOf('ADH:') === 0) {
    adhMarkPaid_(String(row[1]).slice(4), row[3]);
    return 'Payé';
  }
  const ids = String(row[2]).split(',').map(Number).filter(Boolean);
  withInscriptionsSheet_(function (rows) {
    ids.forEach(function (cid) {
      const r = findInscriptionRow_(rows, cid, row[1]);
      if (r) {
        r[2] = 'Validée';
        while (r.length < 8) r.push('');
        r[7] = now;
      } else {
        rows.push([cid, row[1], 'Validée', now, '', '', '', now]);
      }
    });
  });
  return 'Payé';
}

/** Vérifie les paiements en attente (d'un joueur, ou de tous si playerId
 * est vide). Appelé par l'appli au retour de HelloAsso et par la
 * notification HelloAsso. */
function checkPayments_(playerId) {
  const cfg = helloAssoConfig_();
  if (!cfg) return 0;
  const sheet = ensurePaymentsSheet_();
  const values = sheet.getDataRange().getValues();
  let paid = 0;
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row[5] !== 'En attente') continue;
    if (playerId !== null && playerId !== undefined && String(row[1]) !== String(playerId)) continue;
    if (verifyCheckout_(cfg, i + 1, row, sheet) === 'Payé') paid++;
  }
  return paid;
}

/** L'appli demande où en sont les paiements du joueur. */
function checkMyPayments(playerId) {
  ensureSheets_();
  withScriptLock_(function () { checkPayments_(playerId); });
  return getCompetitions(playerId);
}

/** À programmer dans Apps Script (Déclencheurs › Ajouter › fonction
 * verifierPaiementsEnAttente › déclencheur horaire, toutes les 10 minutes) :
 * valide les paiements HelloAsso même si le joueur ne revient pas dans
 * l'appli. Remplace l'URL de notification HelloAsso, déjà utilisée par un
 * autre script du club. */
function verifierPaiementsEnAttente() {
  const n = withScriptLock_(function () { return checkPayments_(null); });
  if (n) bumpDataVersion_();
  Logger.log(n + ' paiement(s) validé(s).');
}

/** Admin : état de la configuration HelloAsso. */
function getHelloAssoStatus() {
  const cfg = helloAssoConfig_();
  if (!cfg) return { configured: false };
  const st = { configured: true, org: cfg.org, sandbox: cfg.sandbox, ok: true };
  try { helloAssoApi_(cfg, 'get', ''); } catch (e) { st.ok = false; st.error = e.message; }
  return st;
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
    // Date limite de la compétition (réglée par l'admin pour toutes ses
    // épreuves), si l'épreuve n'a pas sa propre date limite.
    if (!deadline && (settings.deadlines || {})[String(r.Nom)]) deadline = settings.deadlines[String(r.Nom)];
    if (!deadline && date) {
      const d = new Date(date.getTime());
      d.setDate(d.getDate() - settings.delaiJours);
      deadline = formatFrDate_(d);
      deadlineAuto = true;
    }
    return {
      id: Number(r.ID), type: String(r.Type), nom: String(r.Nom), epreuve: String(r.Epreuve || ''),
      date: String(r.Date), lieu: String(r.Lieu || ''), dateLimite: deadline, dateLimiteAuto: deadlineAuto, dateLimiteEpreuve: r.DateLimite ? String(r.DateLimite) : '', dateLimiteSerie: String((settings.deadlines || {})[String(r.Nom)] || ''),
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
  'Championnat Jeunes (équipe de 2)': ['CJ1', 'CJ2', 'CJ3', 'CJ4'],
  'Championnat Jeunes': ['CJ1', 'CJ2', 'CJ3', 'CJ4'],
};
function competitionTeams_(nom) {
  return COMPETITION_TEAMS_[nom] || [];
}

/** Admin : équipe (ex. CDR2) d'un inscrit, pour une épreuve. */
function setInscriptionEquipe(competitionId, playerId, equipe) {
  ensureSheets_();
  // L'équipe (CDR / CJ) vaut pour toutes les journées de la compétition :
  // elle est reportée sur toutes les inscriptions du joueur à cette
  // compétition, sauf les journées déjà passées (on ne modifie plus
  // l'équipe d'un match joué).
  const all = readCompetitions_(competitionSettings_());
  const c = all.find(function (x) { return Number(x.id) === Number(competitionId); });
  const today = todayStart_();
  const isPast = function (x) { const d = parseFrDate_(x.date); return d && d < today; };
  if (c && isPast(c)) throw new Error('Match passé : l\'équipe ne peut plus être modifiée.');
  const series = c ? all.filter(function (x) { return x.nom === c.nom && x.type === c.type && !isPast(x); }).map(function (x) { return Number(x.id); })
    : [Number(competitionId)];
  withInscriptionsSheet_(function (rows) {
    const row = rows.find(function (r) { return Number(r[0]) === Number(competitionId) && String(r[1]) === String(playerId); });
    if (!row) throw new Error('Inscription introuvable.');
    rows.forEach(function (r) {
      if (String(r[1]) !== String(playerId) || series.indexOf(Number(r[0])) < 0) return;
      while (r.length < 7) r.push('');
      r[6] = String(equipe || '');
    });
  });
  return getCompetitionsAdmin();
}

/** Admin, compétitions à inscription pour la saison (Coupe de Rouen,
 * Challenge SERANO) : valide / refuse / retire un joueur pour tous les
 * tours à venir (les tours passés ne sont pas modifiés). */
function setSeasonInscription(competitionId, playerId, statut) {
  ensureSheets_();
  if (['Demandée', 'Validée', 'Refusée', ''].indexOf(statut) < 0) throw new Error('Statut invalide.');
  const all = readCompetitions_(competitionSettings_());
  const c = all.find(function (x) { return Number(x.id) === Number(competitionId); });
  if (!c) throw new Error('Compétition introuvable.');
  const today = todayStart_();
  const series = all.filter(function (x) { return x.nom === c.nom && x.type === c.type; });
  const seriesIds = series.map(function (x) { return Number(x.id); });
  const future = series.filter(function (x) { const d = parseFrDate_(x.date); return d && d >= today; });
  withInscriptionsSheet_(function (rows) {
    const team = seriesTeam_(rows, seriesIds, playerId);
    future.forEach(function (x) {
      const idx = rows.findIndex(function (r) { return Number(r[0]) === Number(x.id) && String(r[1]) === String(playerId); });
      if (!statut) { if (idx >= 0) rows.splice(idx, 1); return; }
      if (idx >= 0) rows[idx][2] = statut;
      else rows.push([Number(x.id), playerId, statut, new Date(), '', '', team]);
    });
  });
  if (statut === 'Validée' || statut === 'Refusée') {
    sendPush_([playerId], (statut === 'Validée' ? '✅ Inscription validée' : '❌ Inscription refusée'),
      c.nom + ' : saison (' + future.length + ' tour' + (future.length > 1 ? 's' : '') + ')' + (statut === 'Validée' ? '.' : ''));
  }
  return getCompetitionsAdmin();
}

/** Identifiants des épreuves de la même compétition (même nom, même type). */
function seriesIds_(competitionId) {
  const all = readCompetitions_(competitionSettings_());
  const c = all.find(function (x) { return Number(x.id) === Number(competitionId); });
  if (!c) return [Number(competitionId)];
  return all.filter(function (x) { return x.nom === c.nom && x.type === c.type; }).map(function (x) { return Number(x.id); });
}

/** Équipe déjà attribuée au joueur sur une autre journée de la compétition. */
function seriesTeam_(rows, series, playerId) {
  const r = rows.find(function (x) { return String(x[1]) === String(playerId) && series.indexOf(Number(x[0])) >= 0 && x[6]; });
  return r ? String(r[6]) : '';
}

function readInscriptions_() {
  return readTable_('Inscriptions').map(function (r) {
    return {
      competitionId: Number(r.CompetitionID), playerId: r.JoueurID, statut: String(r.Statut || ''),
      demandeLe: r.DemandeLe instanceof Date ? r.DemandeLe.toISOString() : String(r.DemandeLe || ''),
      participation: String(r.Participation || ''),
      equipe: String(r.Equipe || ''),
      reglement: r.Reglement instanceof Date ? formatFrDate_(r.Reglement) : String(r.Reglement || ''),
    };
  });
}

function isEligible_(compet, categorie, sexe) {
  // Rassemblements féminins : réservés aux joueuses (Sexe = F).
  if (/f[ée]minin/i.test(String(compet.nom || '')) && String(sexe || '').toUpperCase() !== 'F') return false;
  if (!categorie) return true; // catégorie non renseignée : on ne bloque pas
  const allowed = compet.categories.length ? compet.categories
    : compet.type === 'Jeunes' || compet.type === 'Stages' ? YOUTH_CATEGORIES_
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
  const sexe = String(player.Sexe || '').toUpperCase();
  const mine = {};
  readInscriptions_().forEach(function (i) {
    if (String(i.playerId) === String(playerId)) mine[i.competitionId] = i;
  });
  return {
    settings: settings,
    categorie: categorie,
    competitions: readCompetitions_(settings).map(function (c) {
      c.eligible = isEligible_(c, categorie, sexe);
      return c;
    }),
    inscriptions: mine,
    paiementsEnAttente: pendingPaymentsOf_(playerId),
  };
}

/** Épreuves dont le paiement HelloAsso du joueur est en cours (moins de 2 h). */
function pendingPaymentsOf_(playerId) {
  const sh = ss_().getSheetByName('Paiements');
  if (!sh) return [];
  const out = [];
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    const created = r[6] instanceof Date ? r[6].getTime() : Date.parse(r[6]);
    if (String(r[1]) === String(playerId) && r[5] === 'En attente' && created && Date.now() - created < 2 * 3600 * 1000) {
      String(r[2]).split(',').map(Number).filter(Boolean).forEach(function (id) { out.push(id); });
    }
  });
  return out;
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
  // Inscription unique (Coupe de Rouen, SERANO) : la demande vaut pour toutes les
  // épreuves à venir de la compétition.
  const ids = [];
  (competitionIds || []).forEach(function (cid) {
    const c = byId[Number(cid)];
    if (c && singleInscription_(c.nom)) {
      data.competitions.forEach(function (x) {
        const d = parseFrDate_(x.date);
        if (x.nom === c.nom && x.type === c.type && d && d >= today && ids.indexOf(x.id) < 0) ids.push(x.id);
      });
    } else if (ids.indexOf(Number(cid)) < 0) ids.push(Number(cid));
  });
  competitionIds = ids;
  withInscriptionsSheet_(function (rows) {
    (competitionIds || []).forEach(function (cid) {
      const c = byId[Number(cid)];
      if (!c) return;
      if (!inscriptionOpen_(data.competitions, c, today)) { refused.push(c.nom + ' ' + c.epreuve); return; }
      const row = findInscriptionRow_(rows, c.id, playerId);
      if (row) {
        if (row[2] === 'Refusée' || row[2] === 'Annulée') { row[2] = 'Demandée'; row[3] = new Date(); row[4] = ''; row[5] = ''; }
      } else {
        const series = data.competitions.filter(function (x) { return x.nom === c.nom && x.type === c.type; }).map(function (x) { return Number(x.id); });
        rows.push([c.id, playerId, 'Demandée', new Date(), '', '', seriesTeam_(rows, series, playerId)]);
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
      // Lignes vides de la feuille (sans nom ni prénom) ignorées.
      .filter(function (p) { return p.Actif !== false && p.Actif !== 'FAUX' && p.Actif !== 'NON' && String(p.Nom || p.Prénom || '').trim(); })
      .map(function (p) {
        return {
          id: p.ID, nom: p.Nom, prenom: p.Prénom, categorie: String(p.Catégorie || ''), sexe: String(p.Sexe || '').toUpperCase(), phone: String(p.Téléphone || ''),
          licence: String(p.Licence || ''), points: p.Points === '' || p.Points === undefined ? '' : Number(p.Points) || '',
        };
      }),
  };
}

/** Admin : valide / refuse une demande, ou inscrit directement un joueur
 * (statut "Validée"). statut vide = supprime l'inscription. */
function setInscriptionStatus(competitionId, playerId, statut) {
  ensureSheets_();
  if (['Demandée', 'Validée', 'Refusée', ''].indexOf(statut) < 0) throw new Error('Statut invalide.');
  const series = seriesIds_(competitionId);
  withInscriptionsSheet_(function (rows) {
    const idx = rows.findIndex(function (r) { return Number(r[0]) === Number(competitionId) && String(r[1]) === String(playerId); });
    if (!statut) { if (idx >= 0) rows.splice(idx, 1); return; }
    if (idx >= 0) { rows[idx][2] = statut; }
    else { rows.push([Number(competitionId), playerId, statut, new Date(), '', '', seriesTeam_(rows, series, playerId)]); }
  });
  if (statut === 'Validée' || statut === 'Refusée') {
    const c = readCompetitions_(competitionSettings_()).find(function (x) { return Number(x.id) === Number(competitionId); });
    if (c) {
      sendPush_([playerId], (statut === 'Validée' ? '✅ Inscription validée' : '❌ Inscription refusée'),
        c.nom + (c.epreuve ? ' – ' + c.epreuve : '') + ' (' + c.date + ')' + (statut === 'Validée' ? ' : pensez à confirmer votre participation.' : ''));
    }
  }
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

// Date limite d'inscription commune à toutes les épreuves d'une compétition
// (Config « DateLimite:<nom> ») ; une date saisie sur une épreuve reste
// prioritaire. Vide = date par défaut (date de l'épreuve - délai).
function setSeriesDeadline(nom, value) {
  ensureSheets_();
  nom = String(nom || '').trim();
  if (!nom) throw new Error('Compétition inconnue.');
  const raw = String(value || '').trim();
  const d = normalizeDateFr_(raw);
  if (raw && !d) throw new Error('Date invalide : ' + raw + ' (format jj/mm/aaaa).');
  setSetting('DateLimite:' + nom, d);
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
// PROFIL DU JOUEUR (page d'accueil)
// ----------------------------------------------------------------------------------

/** Tout ce qu'il faut pour l'encadré « Mon profil » : identité, licence /
 * points FFTT, affectations (toutes phases), indisponibilités, brûlage par
 * phase et inscriptions aux compétitions individuelles. */
function getProfile(playerId) {
  ensureSheets_();
  const p = readTable_('Joueurs').find(function (r) { return String(r.ID) === String(playerId); });
  if (!p) throw new Error('Joueur introuvable.');
  const assignments = readTable_('Affectations')
    .filter(function (r) { return String(r.JoueurID) === String(playerId) && r.EquipeJouee !== ''; })
    .map(function (r) { return { phase: Number(r.Phase), date: r.Date, teamId: Number(r.EquipeJouee) }; });
  const indispos = readTable_('Disponibilites')
    .filter(function (r) { return String(r.JoueurID) === String(playerId) && r.Statut === 'Indisponible'; })
    .map(function (r) { return { phase: Number(r.Phase), date: r.Date }; });
  const burned = {};
  [1, 2].forEach(function (phase) {
    const counts = {};
    assignments.forEach(function (a) { if (a.phase === phase) counts[a.teamId] = (counts[a.teamId] || 0) + 1; });
    burned[phase] = computeBurnedTeams_(counts);
  });
  const comp = getCompetitions(playerId);
  const byId = {};
  comp.competitions.forEach(function (c) { byId[c.id] = c; });
  const inscriptions = Object.keys(comp.inscriptions).map(function (cid) {
    const i = comp.inscriptions[cid];
    const c = byId[Number(cid)];
    return c ? { nom: c.nom, type: c.type, epreuve: c.epreuve, date: c.date, lieu: c.lieu, statut: i.statut, participation: i.participation, equipe: i.equipe } : null;
  }).filter(Boolean);
  const licence = String(p.Licence || '').trim();
  const points = p.Points === '' || p.Points === undefined ? null : Number(p.Points) || null;
  const template = String(readConfigValue_('LienFicheFFTT', '') || '');
  return {
    nom: p.Nom, prenom: p.Prénom, homeTeam: Number(p.EquipeDomicile) || null,
    categorie: String(p.Catégorie || ''),
    capitaine: p.Capitaine === true || p.Capitaine === 'VRAI' || p.Capitaine === 'TRUE',
    licence: licence,
    points: points,
    // Classement FFTT = points / 100 (arrondi inférieur), 5 au minimum.
    classement: points ? Math.max(5, Math.floor(points / 100)) : null,
    ffttUrl: String(p.LienEspace || '').trim() ||
      (licence && template ? template.replace(/\{licence\}/g, encodeURIComponent(licence)) : ''),
    lienEspace: String(p.LienEspace || '').trim(),
    assignments: assignments,
    indispos: indispos,
    burned: burned,
    inscriptions: inscriptions,
  };
}

/** Admin : modèle de lien vers la fiche FFTT d'un joueur, où {licence} est
 * remplacé par son numéro de licence. Vide = pas de lien. */
function setFfttLink(template) {
  return setSetting('LienFicheFFTT', template);
}

/** Admin : modifie un réglage de la feuille Config (liste limitée). */
function setSetting(key, value) {
  ensureSheets_();
  key = String(key || '');
  if (['LienFicheFFTT', 'SiteClub'].concat(CLUB_INFO_KEYS_).indexOf(key) < 0 && key.indexOf('MailInscription:') !== 0 && key.indexOf('PrixHelloAsso:') !== 0 && key.indexOf('DateLimite:') !== 0 && !/^LienPoule:\d+:\d+$/.test(key)) {
    throw new Error('Réglage non modifiable : ' + key);
  }
  return writeConfig_(key, value);
}

/** Écrit une valeur (texte) dans la feuille Config, en ajoutant la clé si
 * besoin. */
function writeConfig_(key, value) {
  const v = String(value || '').trim();
  const sheet = ss_().getSheetByName('Config');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === key) {
      const cell = sheet.getRange(i + 1, 2);
      cell.setNumberFormat('@');
      cell.setValue(v);
      return { ok: true };
    }
  }
  const row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, 2).setNumberFormat('@').setValues([[key, v]]);
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

function addPlayer(nom, prenom, teamId, pin, capitaine, telephone, categorie, sexe, licence) {
  ensureSheets_();
  assertPinAvailable_(pin);
  const sheet = ss_().getSheetByName('Joueurs');
  const players = readTable_('Joueurs');
  const maxId = players.reduce(function (m, p) { return Math.max(m, Number(p.ID) || 0); }, 0);
  const newId = maxId + 1;
  // teamId 0 / vide = « Sans équipe » (hors championnat par équipes).
  sheet.appendRow([newId, nom, prenom, Number(teamId) || '', pin, true, !!capitaine, String(telephone || ''), String(categorie || ''), String(licence || ''), '', normSexe_(sexe)]);
  return { ok: true, id: newId };
}

function normLink_(v) {
  const s = String(v || '').trim();
  if (!s) return '';
  if (!/^https?:\/\/\S+$/i.test(s)) throw new Error('Lien invalide (il doit commencer par https://) : ' + s);
  return s;
}

/** Le joueur enregistre lui-même son lien (espace FFTT / PingPocket). */
function setPlayerLink(playerId, link) {
  ensureSheets_();
  const v = normLink_(link);
  const sheet = ss_().getSheetByName('Joueurs');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(playerId)) {
      sheet.getRange(i + 1, 13).setValue(v);
      return { ok: true, lienEspace: v };
    }
  }
  throw new Error('Joueur introuvable.');
}

function normSexe_(v) {
  const s = String(v || '').trim().toUpperCase();
  return s === 'F' || s === 'M' ? s : '';
}

function updatePlayer(id, nom, prenom, teamId, pin, actif, capitaine, telephone, categorie, licence, points, sexe, lienEspace) {
  ensureSheets_();
  assertPinAvailable_(pin, id);
  const sheet = ss_().getSheetByName('Joueurs');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sheet.getRange(i + 1, 2, 1, 6).setValues([[nom, prenom, Number(teamId) || '', pin, actif, !!capitaine]]);
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
      if (licence !== undefined) {
        const licCell = sheet.getRange(i + 1, 10);
        licCell.setNumberFormat('@');
        licCell.setValue(String(licence || '').trim());
      }
      if (points !== undefined) {
        sheet.getRange(i + 1, 11).setValue(points === '' || points === null ? '' : Number(points) || '');
      }
      if (sexe !== undefined) {
        sheet.getRange(i + 1, 12).setValue(normSexe_(sexe));
      }
      if (lienEspace !== undefined) {
        sheet.getRange(i + 1, 13).setValue(normLink_(lienEspace));
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


// ----------------------------------------------------------------------------------
// NOTIFICATIONS PUSH (Firebase Cloud Messaging)
// ----------------------------------------------------------------------------------
// Activées seulement si la propriété du script FCM_SERVICE_ACCOUNT contient
// la clé JSON d'un compte de service Firebase (Paramètres du projet ›
// Propriétés du script). Sans elle, tout le reste fonctionne normalement et
// aucune notification n'est envoyée. Voir README.md.

/** L'appli installée enregistre le jeton de notification du téléphone. */
function registerPushToken(playerId, token, platform) {
  ensureSheets_();
  token = String(token || '').trim();
  if (!token) return { ok: false };
  const sheet = ss_().getSheetByName('Appareils');
  const values = sheet.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][1]) === token) sheet.deleteRow(i + 1);
  }
  sheet.appendRow([playerId, token, String(platform || ''), new Date()]);
  return { ok: true };
}

/** Déconnexion : le téléphone ne reçoit plus les notifications de ce joueur. */
function unregisterPushToken(token) {
  ensureSheets_();
  removePushTokens_([String(token || '')]);
  return { ok: true };
}

function removePushTokens_(tokens) {
  const set = {};
  tokens.forEach(function (t) { if (t) set[t] = true; });
  const sheet = ss_().getSheetByName('Appareils');
  if (!sheet) return;
  const values = sheet.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) {
    if (set[String(values[i][1])]) sheet.deleteRow(i + 1);
  }
}

function fcmServiceAccount_() {
  const raw = PropertiesService.getScriptProperties().getProperty('FCM_SERVICE_ACCOUNT');
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

/** Jeton d'accès OAuth du compte de service (gardé 50 min en cache). */
function fcmAccessToken_(sa) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('FCM_ACCESS_TOKEN');
  if (cached) return cached;
  const now = Math.floor(Date.now() / 1000);
  const enc = function (o) { return Utilities.base64EncodeWebSafe(JSON.stringify(o)).replace(/=+$/, ''); };
  const unsigned = enc({ alg: 'RS256', typ: 'JWT' }) + '.' + enc({
    iss: sa.client_email, scope: 'https://www.googleapis.com/auth/firebase.messaging',
    aud: sa.token_uri || 'https://oauth2.googleapis.com/token', iat: now, exp: now + 3600,
  });
  const sig = Utilities.base64EncodeWebSafe(Utilities.computeRsaSha256Signature(unsigned, sa.private_key)).replace(/=+$/, '');
  const res = UrlFetchApp.fetch(sa.token_uri || 'https://oauth2.googleapis.com/token', {
    method: 'post', muteHttpExceptions: true,
    payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: unsigned + '.' + sig },
  });
  const data = JSON.parse(res.getContentText() || '{}');
  if (!data.access_token) throw new Error('Firebase : jeton refusé (' + res.getContentText().slice(0, 200) + ')');
  cache.put('FCM_ACCESS_TOKEN', data.access_token, 3000);
  return data.access_token;
}

/** Envoie une notification aux téléphones des joueurs donnés (null = tous).
 * Ne lève jamais d'erreur : une notification ratée ne doit pas bloquer
 * l'enregistrement (convocation, inscription…). Retourne le nombre envoyé. */
function sendPush_(playerIds, title, body) {
  try {
    const sa = fcmServiceAccount_();
    if (!sa) return 0;
    const wanted = {};
    (playerIds || []).forEach(function (id) { wanted[String(id)] = true; });
    const devices = readTable_('Appareils').filter(function (d) {
      return d.Jeton && (playerIds === null || wanted[String(d.JoueurID)]);
    });
    if (!devices.length) return 0;
    const access = fcmAccessToken_(sa);
    const url = 'https://fcm.googleapis.com/v1/projects/' + sa.project_id + '/messages:send';
    const requests = devices.map(function (d) {
      return {
        url: url, method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { Authorization: 'Bearer ' + access },
        payload: JSON.stringify({ message: {
          token: String(d.Jeton),
          notification: { title: String(title || ''), body: String(body || '') },
          android: { priority: 'high', notification: { sound: 'default' } },
          apns: { payload: { aps: { sound: 'default' } } },
        } }),
      };
    });
    const responses = UrlFetchApp.fetchAll(requests);
    const dead = [];
    let sent = 0;
    responses.forEach(function (r, i) {
      const code = r.getResponseCode();
      if (code === 200) { sent++; return; }
      // Appli désinstallée / jeton expiré : on oublie ce téléphone.
      if (code === 404 || /UNREGISTERED|not a valid FCM registration token/i.test(r.getContentText())) dead.push(String(devices[i].Jeton));
      else console.warn('FCM ' + code + ' : ' + r.getContentText().slice(0, 300));
    });
    if (dead.length) removePushTokens_(dead);
    return sent;
  } catch (err) {
    console.warn('Notification push non envoyée : ' + (err && err.message ? err.message : err));
    return 0;
  }
}

/** Admin : état des notifications push (réglages). */
function getPushStatus() {
  ensureSheets_();
  const sa = fcmServiceAccount_();
  return {
    configured: !!sa,
    projectId: sa ? String(sa.project_id || '') : '',
    devices: readTable_('Appareils').filter(function (d) { return d.Jeton; }).length,
  };
}

/** Admin : notification de test envoyée à tous les téléphones inscrits. */
function sendTestPush() {
  ensureSheets_();
  if (!fcmServiceAccount_()) throw new Error("Notifications push non configurées : ajoutez la propriété du script FCM_SERVICE_ACCOUNT (voir README).");
  const sent = sendPush_(null, '🔔 RebonTT', 'Notification de test : les notifications fonctionnent.');
  return { sent: sent, devices: getPushStatus().devices };
}

/** À lancer UNE FOIS depuis l'éditeur Apps Script (menu Exécuter) après
 * avoir collé cette version du script : Google demande alors l'autorisation
 * « se connecter à un service externe » (nécessaire pour Firebase).
 * Envoie aussi une notification de test si Firebase est configuré. */
function testerNotifications() {
  if (!fcmServiceAccount_()) {
    Logger.log('Autorisation OK. Notifications push pas encore configurées (propriété FCM_SERVICE_ACCOUNT absente) : les rappels locaux fonctionnent quand même.');
    return;
  }
  Logger.log(JSON.stringify(sendTestPush()));
}

// ----------------------------------------------------------------------------------
// ADHÉSIONS AU CLUB — site d'inscription public (www/inscription.html)
// ----------------------------------------------------------------------------------
// Une famille crée son « dossier » (responsable légal + membres + créneaux
// choisis + mode de règlement), puis y revient avec son n° de dossier et son
// téléphone pour le consulter, remplir le questionnaire de santé, accepter le
// règlement intérieur, envoyer des documents ou demander une modification.
// Données dans de nouvelles feuilles (Adhesions, AdhesionMembres,
// AdhesionDocuments, AdhesionMessages) : l'effectif (Joueurs) n'est jamais
// modifié ; l'administrateur peut seulement y AJOUTER un nouveau joueur.
// Questionnaire de santé : les réponses restent sur l'appareil de la famille,
// seule l'attestation (« NON à toutes les questions » ou « certificat
// médical requis ») est enregistrée.

const ADH_SHEETS_ = {
  Adhesions: ['Dossier', 'CreeLe', 'Saison', 'Nom', 'Prénom', 'Adresse', 'CP', 'Ville', 'Téléphone', 'Email', 'Téléphone2', 'Email2',
    'Tarif', 'Total', 'Aides', 'Mode', 'Echeances', 'Statut', 'Réglé', 'Note'],
  AdhesionMembres: ['Dossier', 'N', 'Nom', 'Prénom', 'Naissance', 'Sexe', 'Catégorie', 'Creneau1', 'Creneau2', 'Licence', 'NumLicence',
    'Montant', 'QS', 'QSLe', 'Reglement', 'ReglementLe', 'Statut', 'JoueurID'],
  AdhesionDocuments: ['Dossier', 'N', 'Type', 'Fichier', 'Lien', 'EnvoyeLe'],
  AdhesionMessages: ['Date', 'Dossier', 'Nom', 'Email', 'Téléphone', 'Objet', 'Message', 'Statut'],
  // Chèques remis par la famille (bordereau de remise en banque).
  AdhesionCheques: ['Dossier', 'N', 'Titulaire', 'Banque', 'Numero', 'Montant', 'DeposeLe'],
};
const ADH_DOC_TYPES_ = ['Certificat médical', 'Attestation PPS', 'Justificatif', 'Autre'];
const ADH_STATUTS_ = ['À régler', 'Partiellement réglé', 'Réglé', 'Annulé'];

function adhSheet_(name) {
  let sh = ss_().getSheetByName(name);
  if (!sh) {
    const h = ADH_SHEETS_[name];
    sh = ss_().insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), h.length).setNumberFormat('@'); // téléphones : garder le 0
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function adhRows_(name) { adhSheet_(name); return readTable_(name); }
function adhAppend_(name, row) {
  const sh = adhSheet_(name);
  const r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, row.length).setNumberFormat('@').setValues([row.map(function (v) { return v === null || v === undefined ? '' : String(v); })]);
}
/** Modifie les lignes d'une feuille d'adhésion qui vérifient match(obj) :
 * seules ces lignes sont réécrites. */
function adhUpdate_(name, match, patch) {
  const sh = adhSheet_(name);
  const values = sh.getDataRange().getValues();
  const h = values[0];
  let n = 0;
  for (let i = 1; i < values.length; i++) {
    const obj = {};
    h.forEach(function (k, j) { obj[k] = values[i][j]; });
    if (!match(obj)) continue;
    const p = typeof patch === 'function' ? patch(obj) : patch;
    Object.keys(p || {}).forEach(function (k) {
      const j = h.indexOf(k);
      if (j >= 0) sh.getRange(i + 1, j + 1).setNumberFormat('@').setValue(String(p[k] === null || p[k] === undefined ? '' : p[k]));
    });
    n++;
  }
  return n;
}

function adhNow_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'); }
function adhNum_(v) { return Number(String(v === undefined || v === null ? '' : v).replace(',', '.').replace(/[^\d.-]/g, '')) || 0; }
function adhTxt_(v, max) { return String(v === undefined || v === null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max || 120); }
function adhNorm_(v) { return String(v || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]/g, ''); }
function adhTel_(v) { return String(v || '').replace(/\D/g, '').replace(/^33/, '0').slice(-9); }

function defaultSaison_() {
  const d = new Date();
  const y = d.getMonth() >= 6 ? d.getFullYear() : d.getFullYear() - 1;
  return y + '/' + (y + 1);
}

function configMap_() {
  const m = {};
  readTable_('Config').forEach(function (r) { m[String(r.Clé)] = String(r.Valeur === undefined || r.Valeur === null ? '' : r.Valeur); });
  return m;
}

/** Réglages des inscriptions (clés « Adh… » de la feuille Config). */
function adhSettings_() {
  const m = configMap_();
  const v = function (k, d) { return m[k] !== undefined && m[k] !== '' ? m[k] : (d || ''); };
  const aides = v('AdhAides', "Pass'Sport|50\nCoupons Sport ANCV\nChèques-Vacances ANCV\nAtout Normandie\nPass'Jeunes 76")
    .split('\n').map(function (l) {
      const p = l.split('|');
      return { nom: adhTxt_(p[0], 60), montant: adhNum_(p[1]) };
    }).filter(function (a) { return a.nom; });
  // HelloAsso : paiement intégré (clés API) ou simple lien vers la page de
  // paiement HelloAsso du club.
  const helloassoLien = v('AdhHelloAssoLien', '');
  const helloasso = !!helloAssoConfig_() || !!helloassoLien;
  return {
    helloassoLien: helloassoLien,
    helloassoApi: !!helloAssoConfig_(),
    saison: v('AdhSaison', defaultSaison_()),
    ouvert: !/^(non|0|false|fermé|ferme)$/i.test(v('AdhOuvert', 'oui')),
    commune: v('AdhCommune', v('ClubVille', '')),
    supplement: adhNum_(v('AdhSupplementExterieur', '0')),
    reductionFamille: adhNum_(v('AdhReductionFamille', '0')),
    modes: v('AdhModes', 'Chèque,Espèces,HelloAsso').split(',').map(function (s) { return s.trim(); })
      .filter(function (s) { return s && (s !== 'HelloAsso' || helloasso); }),
    chequesMax: Math.max(1, Math.min(10, adhNum_(v('AdhChequesMax', '3')) || 3)),
    ordreCheque: v('AdhOrdreCheque', v('ClubNom', '')),
    iban: v('AdhIban', ''),
    // Coordonnées bancaires du club (virements, bordereau de remise de chèques).
    banque: { titulaire: v('AdhBanqueTitulaire', v('AdhOrdreCheque', v('ClubNom', ''))), nom: v('AdhBanqueNom', ''),
      agence: v('AdhBanqueAgence', ''), iban: v('AdhIban', ''), bic: v('AdhBic', '') },
    aides: aides,
    reglementUrl: v('AdhReglementUrl', ''),
    reglementTexte: v('AdhReglementTexte', ''),
    contact: v('AdhContactEmail', ''),
    infos: v('AdhInfos', ''),
    // PPS : à faire dans l'espace licencié FFTT du joueur.
    ppsUrl: v('AdhPpsUrl', 'https://malicence.fftt.com/login'),
    // Formulaire FFTT d'inscription en ligne du club (demande de licence).
    licenceUrl: v('AdhLicenceUrl', 'https://inscriptionenligne.fftt.com/club/09760374'),
    licences: adhLicences_(v('AdhLicences', JSON.stringify(ADH_LICENCES_DEFAUT_))),
    clubNom: v('ClubNom', ''),
    clubVille: v('ClubVille', ''),
    clubAdresse: v('ClubAdresse', ''),
    site: v('SiteClub', ''),
  };
}

/** Prix de la licence FFTT du club par catégorie (modifiables dans l'onglet
 * Adhésions). */
const ADH_LICENCES_DEFAUT_ = { Poussin: 37, Benjamin: 37, Minime: 37, Cadet: 52, Junior: 70, Senior: 70,
  V40: 70, V50: 70, V60: 70, V70: 70, V80: 70 };

/** Prix de la licence FFTT par catégorie : {"Poussin": 37, …} (JSON). */
function adhLicences_(raw) {
  let o = {};
  try { o = JSON.parse(String(raw || '{}')) || {}; } catch (e) { o = {}; }
  const out = {};
  ADH_CATS_.forEach(function (c) { const v = adhNum_(o[c[1]]); if (v > 0) out[c[1]] = v; });
  return out;
}
/** Licence d'un membre selon sa catégorie (à défaut : ancien montant saisi
 * sur le créneau). */
function adhLicence_(st, cat, c) {
  return st.licences[cat] > 0 ? st.licences[cat] : (c ? c.licence : 0);
}

/** Catégorie FFTT d'après l'année de naissance et la saison (âge atteint
 * pendant la 1re année civile de la saison). */
const ADH_CATS_ = [[8, 'Poussin'], [10, 'Benjamin'], [12, 'Minime'], [14, 'Cadet'], [18, 'Junior'], [39, 'Senior'],
  [49, 'V40'], [59, 'V50'], [69, 'V60'], [79, 'V70'], [200, 'V80']];
function adhBirth_(s) {
  s = String(s || '').trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return { y: +m[1], m: +m[2], d: +m[3], fr: m[3] + '/' + m[2] + '/' + m[1] };
  m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) return { y: +m[3], m: +m[2], d: +m[1], fr: s };
  return null;
}
function adhCategorie_(naissance, saison) {
  const b = adhBirth_(naissance);
  if (!b) return '';
  const age = Number(String(saison).slice(0, 4)) - b.y;
  for (let i = 0; i < ADH_CATS_.length; i++) if (age <= ADH_CATS_[i][0]) return ADH_CATS_[i][1];
  return '';
}
function adhJeune_(cat) { return YOUTH_CATEGORIES_.indexOf(cat) >= 0; }
function adhEligible_(c, cat) {
  if (!c) return false;
  if (c.public === 'Jeunes') return adhJeune_(cat);
  if (c.public === 'Adultes') return !adhJeune_(cat);
  return true;
}
/** Créneau lié à c (même niveau, 2 séances par semaine), ou null. */
function adhPartner_(c, crs) {
  if (!c || c.seances !== 2) return null;
  return crs.find(function (x) { return x.id !== c.id && x.seances === 2 && adhNorm_(x.niveau) === adhNorm_(c.niveau); }) || null;
}
function adhSlotLabel_(c) {
  return c ? (c.niveau || 'Entraînement') + ' – ' + c.jour + ' ' + c.debut + (c.fin ? '-' + c.fin : '') : '';
}

/** Créneaux avec le nombre d'inscrits (dossiers non annulés). */
function adhCreneaux_() {
  const crs = parseCreneaux_(readConfigValue_('ClubCreneaux', ''));
  const cancelled = {};
  adhRows_('Adhesions').forEach(function (a) { if (a.Statut === 'Annulé') cancelled[a.Dossier] = true; });
  const used = {};
  adhRows_('AdhesionMembres').forEach(function (m) {
    if (cancelled[m.Dossier] || m.Statut === 'Annulé') return;
    [m.Creneau1, m.Creneau2].forEach(function (id) { if (id) used[id] = (used[id] || 0) + 1; });
  });
  return crs.map(function (c) {
    c.inscrits = used[c.id] || 0;
    c.restantes = c.places > 0 ? Math.max(0, c.places - c.inscrits) : null;
    return c;
  });
}

/** Page d'inscription : club, saison, réglages publics et créneaux. */
function getAdhesionInfo() {
  ensureSheets_();
  const st = adhSettings_();
  delete st.contact;
  return { settings: st, creneaux: adhCreneaux_(), helloasso: st.modes.indexOf('HelloAsso') >= 0 };
}

/** « Déjà adhérent ? » : retrouve un dossier (ou, à défaut, un joueur de
 * l'effectif) par téléphone + 3 premières lettres du nom. */
function findAdherent(tel, nom3) {
  ensureSheets_();
  const t = adhTel_(tel), n = adhNorm_(nom3).slice(0, 3);
  if (t.length < 9 || n.length < 3) throw new Error('Saisissez votre numéro de téléphone et les 3 premières lettres du nom.');
  const doss = adhRows_('Adhesions').filter(function (a) {
    return (adhTel_(a.Téléphone) === t || adhTel_(a.Téléphone2) === t) && adhNorm_(a.Nom).indexOf(n) === 0;
  });
  if (doss.length) {
    const a = doss[doss.length - 1];
    return {
      source: 'adhesion',
      responsable: { nom: a.Nom, prenom: a.Prénom, adresse: a.Adresse, cp: a.CP, ville: a.Ville, tel: a.Téléphone, email: a.Email, tel2: a.Téléphone2, email2: a.Email2 },
      membres: adhRows_('AdhesionMembres').filter(function (m) { return m.Dossier === a.Dossier && m.Statut !== 'Annulé'; })
        .map(function (m) { return { nom: m.Nom, prenom: m.Prénom, naissance: m.Naissance, sexe: m.Sexe, numLicence: m.NumLicence }; }),
    };
  }
  const players = readTable_('Joueurs').filter(function (p) { return adhTel_(p.Téléphone) === t && String(p.Nom || '').trim(); });
  const first = players.find(function (p) { return adhNorm_(p.Nom).indexOf(n) === 0; });
  if (!first) throw new Error('Aucune adhésion trouvée avec ce téléphone et ce nom : remplissez le formulaire.');
  return {
    source: 'effectif',
    responsable: { nom: first.Nom, prenom: first.Prénom, tel: first.Téléphone },
    membres: players.map(function (p) { return { nom: p.Nom, prenom: p.Prénom, sexe: normSexe_(p.Sexe), numLicence: String(p.Licence || '') }; }),
  };
}

/** Vérifie et chiffre une demande d'adhésion (même calcul que la page). */
function adhCompute_(d, st, crs) {
  d = d || {};
  const r = d.responsable || {};
  const resp = {
    nom: adhTxt_(r.nom, 60).toUpperCase(), prenom: adhTxt_(r.prenom, 60), adresse: adhTxt_(r.adresse, 150), cp: adhTxt_(r.cp, 10),
    ville: adhTxt_(r.ville, 60).toUpperCase(), tel: adhTxt_(r.tel, 20), email: adhTxt_(r.email, 120).toLowerCase(),
    tel2: adhTxt_(r.tel2, 20), email2: adhTxt_(r.email2, 120).toLowerCase(),
  };
  ['nom', 'prenom', 'adresse', 'cp', 'ville', 'tel', 'email'].forEach(function (k) {
    if (!resp[k]) throw new Error('Responsable légal : champ obligatoire manquant (' + k + ').');
  });
  if (!/^\S+@\S+\.\S+$/.test(resp.email) || (resp.email2 && !/^\S+@\S+\.\S+$/.test(resp.email2))) throw new Error('Adresse e-mail invalide.');
  if (adhTel_(resp.tel).length < 9) throw new Error('Numéro de téléphone invalide.');
  const exterieur = !!st.commune && adhNorm_(resp.ville) !== adhNorm_(st.commune);
  const byId = {};
  crs.forEach(function (c) { byId[c.id] = c; });
  const list = (d.membres || []).filter(function (m) { return m && (adhTxt_(m.nom) || adhTxt_(m.prenom)); });
  if (!list.length) throw new Error('Ajoutez au moins un membre.');
  if (list.length > 10) throw new Error('10 membres au maximum par dossier.');
  let total = 0;
  const membres = list.map(function (m, i) {
    const label = 'Membre ' + (i + 1);
    const b = adhBirth_(m.naissance);
    const out = {
      nom: adhTxt_(m.nom, 60).toUpperCase(), prenom: adhTxt_(m.prenom, 60), naissance: b ? b.fr : '', sexe: normSexe_(m.sexe),
      // Licence FFTT facultative (nécessaire pour la compétition).
      licence: m.licence === 'Renouvellement' || m.licence === 'Nouvelle' ? m.licence : 'Sans licence', numLicence: adhTxt_(m.numLicence, 20),
    };
    if (!out.nom || !out.prenom) throw new Error(label + ' : nom et prénom obligatoires.');
    if (!b || b.y < 1900 || b.y > new Date().getFullYear()) throw new Error(out.prenom + ' : date de naissance invalide.');
    if (!out.sexe) throw new Error(out.prenom + ' : précisez le sexe (licence FFTT).');
    out.categorie = adhCategorie_(out.naissance, st.saison);
    const c1 = byId[m.creneau1];
    if (!c1) throw new Error(out.prenom + ' : choisissez un créneau.');
    if (!adhEligible_(c1, out.categorie)) throw new Error(out.prenom + ' (' + out.categorie + ') : le créneau « ' + adhSlotLabel_(c1) + ' » ne correspond pas à son âge.');
    // Créneaux liés (2 séances par semaine) : l'inscription vaut pour les
    // deux créneaux du même niveau.
    const c2 = adhPartner_(c1, crs);
    out.creneau1 = c1.id;
    out.creneau2 = c2 ? c2.id : '';
    out.montant = Math.max(0, c1.tarif + (out.licence !== 'Sans licence' ? adhLicence_(st, out.categorie, c1) : 0) + (exterieur ? st.supplement : 0) - (i > 0 ? st.reductionFamille : 0));
    total += out.montant;
    return out;
  });
  const aides = st.aides.filter(function (a) { return (d.aides || []).indexOf(a.nom) >= 0; });
  const aidesMontant = Math.min(total, aides.reduce(function (s, a) { return s + a.montant; }, 0));
  const mode = String(d.mode || '');
  if (st.modes.indexOf(mode) < 0) throw new Error('Choisissez un mode de règlement.');
  const echeances = mode === 'Chèque' ? Math.max(1, Math.min(st.chequesMax, Math.floor(adhNum_(d.echeances)) || 1)) : 1;
  // Chèques : titulaire, banque, numéro et montant de chacun ; la somme doit
  // faire le reste à régler.
  let cheques = [];
  if (mode === 'Chèque') {
    const list = (d.cheques || []).slice(0, echeances);
    if (list.length !== echeances) throw new Error('Renseignez les ' + echeances + ' chèque(s).');
    cheques = list.map(function (c, i) {
      const x = { titulaire: adhTxt_(c.titulaire, 80), banque: adhTxt_(c.banque, 60), numero: adhTxt_(c.numero, 20), montant: Math.round(adhNum_(c.montant) * 100) / 100 };
      if (!x.titulaire || !x.banque || !x.numero) throw new Error('Chèque ' + (i + 1) + ' : titulaire, banque et numéro obligatoires.');
      if (!(x.montant > 0)) throw new Error('Chèque ' + (i + 1) + ' : montant invalide.');
      return x;
    });
    const somme = cheques.reduce(function (t, c) { return t + c.montant; }, 0);
    const du = Math.max(0, total - aidesMontant);
    if (Math.abs(somme - du) > 0.01) throw new Error('Le total des chèques (' + somme + ' €) doit être égal au reste à régler (' + du + ' €).');
  }
  return {
    responsable: resp, exterieur: exterieur, membres: membres, total: total,
    aides: aides.map(function (a) { return a.nom; }), aidesMontant: aidesMontant, mode: mode, echeances: echeances, cheques: cheques,
  };
}

function adhPublicUrl_() {
  const p = PropertiesService.getScriptProperties();
  const app = p.getProperty('APP_URL') || 'https://ttfrisno-design.github.io/Brulage/';
  return app.replace(/[?#].*$/, '').replace(/[^/]*$/, '') + 'inscription.html';
}

/** Consignes de virement : coordonnées bancaires du club + libellé. */
function adhVirementTexte_(st, dossier) {
  const b = st.banque || {};
  if (!b.iban) return ' (coordonnées bancaires communiquées par le club)';
  return ' — titulaire : ' + b.titulaire + ', IBAN : ' + b.iban + (b.bic ? ', BIC : ' + b.bic : '') + (b.nom ? ' (' + b.nom + ')' : '') +
    (dossier ? ', libellé : « Adhésion ' + dossier + ' »' : '');
}

function adhMail_(to, subject, body, replyTo) {
  if (!to) return false;
  try {
    const o = { to: to, subject: subject, body: body, name: readConfigValue_('ClubNom', 'RebonTT') || 'RebonTT' };
    if (replyTo) o.replyTo = replyTo; // la famille répond directement au club
    MailApp.sendEmail(o);
    return true;
  } catch (e) {
    console.warn('E-mail non envoyé (' + to + ') : ' + e.message);
    return false;
  }
}

function adhRecap_(a, membres, crs, st) {
  const byId = {};
  crs.forEach(function (c) { byId[c.id] = c; });
  const lines = membres.map(function (m) {
    return '- ' + m.prenom + ' ' + m.nom + ' (' + m.categorie + ', ' + m.licence.toLowerCase() + ') : ' + adhSlotLabel_(byId[m.creneau1]) +
      (m.creneau2 ? ' + ' + adhSlotLabel_(byId[m.creneau2]) : '') + ' — ' + m.montant + ' €';
  });
  let pay = 'Mode de règlement : ' + a.mode;
  if (a.mode === 'Chèque') {
    pay += ' en ' + a.echeances + ' fois, à l\'ordre de « ' + st.ordreCheque + ' »' + (a.cheques || []).map(function (c, i) {
      return '\n  Chèque ' + (i + 1) + ' : ' + c.montant + ' € – n° ' + c.numero + ' – ' + c.banque + ' – ' + c.titulaire;
    }).join('');
  }
  if (a.mode === 'Virement') pay += adhVirementTexte_(st, a.dossier);
  return lines.join('\n') + '\n\nTotal : ' + a.total + ' €' +
    (a.aidesMontant ? '\nAides déclarées (' + a.aides.join(', ') + ') : -' + a.aidesMontant + ' € (justificatif à envoyer)' : '') +
    '\nReste à régler : ' + Math.max(0, a.total - a.aidesMontant) + ' €\n' + pay;
}

/** E-mail de suivi envoyé à la famille (e-mail principal et e-mail 2) à
 * chaque avancement de son dossier : ce qui vient de se passer, l'état de
 * chaque membre et ce qu'il reste à faire. */
function adhFamilyMail_(dossier, sujet, evenement) {
  try {
    const a = adhRows_('Adhesions').find(function (x) { return String(x.Dossier) === String(dossier); });
    if (!a || !a.Email) return;
    const st = adhSettings_(), crs = adhCreneaux_();
    const membres = adhRows_('AdhesionMembres').filter(function (m) { return m.Dossier === a.Dossier && m.Statut !== 'Annulé'; });
    const docs = adhRows_('AdhesionDocuments').filter(function (x) { return x.Dossier === a.Dossier; });
    const hasDoc = function (n, type) { return docs.some(function (x) { return adhNum_(x.N) === n && x.Type === type; }); };
    const todo = [], etat = [];
    if (a.Statut !== 'Annulé') {
      membres.forEach(function (m) {
        const n = adhNum_(m.N), who = m.Prénom, form = adhSanteForm_(m, crs, a.Saison);
        const certOk = hasDoc(n, 'Certificat médical');
        if (!m.QS && !certOk) {
          todo.push(who + ' : ' + ({
            mineur: 'remplir le questionnaire de santé pour mineur (FFTT n° 26-10-2) et signer l\'attestation (joueur et représentant légal)',
            decouverte: 'remplir le questionnaire de santé pour majeur – licence Découverte (FFTT n° 26-10-3) et signer l\'attestation',
            pps: 'faire le PPS dans l\'espace MaLicence (' + st.ppsUrl + ') et envoyer l\'attestation PPS',
            certificat: 'envoyer un certificat médical (entrée en catégorie vétéran V65, V70, V80, V85 ou V90)',
          })[form] + '.');
        } else if (/requis/.test(m.QS || '') && !certOk) {
          todo.push(who + ' : envoyer le certificat médical (réponse OUI au questionnaire de santé).');
        }
        if (!m.Reglement) todo.push(who + ' : accepter le règlement intérieur.');
        etat.push('- ' + m.Prénom + ' ' + m.Nom + ' : santé ' + (m.QS ? '✓ ' + m.QS : (certOk ? '✓ certificat médical reçu' : '✗ à faire')) +
          ' · règlement intérieur ' + (m.Reglement ? '✓' : '✗ à accepter') + (m.JoueurID ? ' · inscrit(e) dans l\'effectif du club' : ''));
      });
      if (a.Aides && !docs.some(function (x) { return x.Type === 'Justificatif'; })) todo.push('Envoyer le justificatif des aides déclarées (' + a.Aides + ').');
      const reste = adhReste_(a);
      if (reste > 0) {
        const ch = adhRows_('AdhesionCheques').filter(function (x) { return x.Dossier === a.Dossier && !x.DeposeLe; });
        todo.push(a.Mode === 'Chèque' && ch.length
          ? 'Remettre au club (si ce n\'est pas déjà fait) vos chèques à l\'ordre de « ' + st.ordreCheque + ' » : ' +
            ch.map(function (x) { return 'chèque ' + x.N + ' n° ' + x.Numero + ' de ' + adhNum_(x.Montant) + ' €'; }).join(', ') + '. Ils sont encaissés un par un.'
          : 'Régler ' + reste + ' € : ' + (a.Mode === 'Chèque' ? 'chèque(s) à l\'ordre de « ' + st.ordreCheque + ' » (' + a.Echeances + ' fois), à remettre au club'
          : a.Mode === 'HelloAsso' ? 'paiement en ligne depuis votre compte (bouton « Payer en ligne »)'
          : a.Mode === 'Virement' ? 'virement' + adhVirementTexte_(st, a.Dossier) : 'en espèces, au club') + '.');
      }
    }
    const lic = membres.some(function (m) { return m.Licence === 'Nouvelle' || m.Licence === 'Renouvellement'; });
    const body = 'Bonjour ' + a.Prénom + ',\n\n' + evenement + '\n\n' +
      'État de votre dossier ' + a.Dossier + ' (saison ' + a.Saison + ') :\n' + etat.join('\n') +
      '\n- Règlement : ' + a.Statut + ' — total ' + adhNum_(a.Total) + ' €' + (adhAidesMontant_(a) ? ', aides ' + adhAidesMontant_(a) + ' €' : '') +
      ', déjà réglé ' + adhNum_(a.Réglé) + ' €, reste ' + adhReste_(a) + ' €\n\n' +
      (a.Statut === 'Annulé' ? 'Ce dossier est annulé.\n'
        : todo.length ? 'Ce qu\'il vous reste à faire :\n' + todo.map(function (t, i) { return (i + 1) + '. ' + t; }).join('\n') + '\n'
        : '✅ Votre dossier est complet. Merci et bonne saison !\n') +
      (lic && a.Statut !== 'Annulé' ? '\nÀ ne pas oublier — licence FFTT : faites la demande en ligne dès que vous êtes sur votre compte licencié (' + st.ppsUrl +
        ') ou avec le formulaire du club : ' + st.licenceUrl + '. Déjà licencié(e) : connectez-vous absolument avec votre numéro de licence ' +
        '(en cas d\'erreur avec l\'ancien compte, un nouveau compte est créé et il faut contacter la FFTT pour demander la correction). ' +
        'Nouveau joueur : créez votre compte licencié FFTT.\n' : '') +
      '\nSuivre votre dossier (n° ' + a.Dossier + ' + votre téléphone) : ' + adhPublicUrl_() +
      '\n\nSportivement,\n' + (st.clubNom || 'Le club');
    [a.Email, a.Email2].filter(function (e, i, arr) { return e && arr.indexOf(e) === i; }).forEach(function (to) {
      adhMail_(to, sujet + ' – dossier ' + a.Dossier, body, st.contact);
    });
  } catch (e) {
    console.warn('E-mail de suivi non envoyé (' + dossier + ') : ' + e.message);
  }
}

/** Crée le dossier d'adhésion d'une famille. */
function createAdhesion(d) {
  ensureSheets_();
  d = d || {};
  if (d.website) throw new Error('Demande refusée.'); // champ piège anti-robots
  const st = adhSettings_();
  if (!st.ouvert) throw new Error('Les inscriptions en ligne sont fermées pour le moment.');
  const crs = adhCreneaux_();
  const a = adhCompute_(d, st, crs);
  // Places restantes (pour l'ensemble des membres du dossier).
  const need = {};
  a.membres.forEach(function (m) { [m.creneau1, m.creneau2].forEach(function (id) { if (id) need[id] = (need[id] || 0) + 1; }); });
  crs.forEach(function (c) {
    if (need[c.id] && c.restantes !== null && need[c.id] > c.restantes) {
      throw new Error('Plus assez de places pour « ' + adhSlotLabel_(c) + ' » (' + c.restantes + ' restante' + (c.restantes > 1 ? 's' : '') + ').');
    }
  });
  const prefix = 'TT' + String(st.saison).slice(2, 4) + '-';
  // N° de dossier : TT + année de saison + 4 caractères tirés au hasard (sans
  // caractères ambigus 0/O, 1/I/L), unique.
  const used = {};
  adhRows_('Adhesions').forEach(function (x) { used[String(x.Dossier).toUpperCase()] = true; });
  const ALPHA = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
  let dossier = '';
  do {
    dossier = prefix;
    for (let k = 0; k < 4; k++) dossier += ALPHA.charAt(Math.floor(Math.random() * ALPHA.length));
  } while (used[dossier]);
  const r = a.responsable;
  adhAppend_('Adhesions', [dossier, adhNow_(), st.saison, r.nom, r.prenom, r.adresse, r.cp, r.ville, r.tel, r.email, r.tel2, r.email2,
    a.exterieur ? 'Extérieur' : 'Commune', a.total, a.aidesMontant ? a.aides.join(', ') + ' (' + a.aidesMontant + ' €)' : '',
    a.mode, a.echeances, 'À régler', 0, '']);
  a.cheques.forEach(function (c, i) {
    adhAppend_('AdhesionCheques', [dossier, i + 1, c.titulaire, c.banque, c.numero, c.montant, '']);
  });
  a.membres.forEach(function (m, i) {
    adhAppend_('AdhesionMembres', [dossier, i + 1, m.nom, m.prenom, m.naissance, m.sexe, m.categorie, m.creneau1, m.creneau2,
      m.licence, m.numLicence, m.montant, '', '', '', '', 'Inscrit', '']);
  });
  a.dossier = dossier;
  const recap = adhRecap_(a, a.membres, crs, st);
  adhFamilyMail_(dossier, 'Inscription enregistrée', 'Votre inscription au tennis de table pour la saison ' + st.saison +
    ' est bien enregistrée sous le n° de dossier ' + dossier + ' : notez-le, il vous sera demandé avec votre téléphone.\n\n' + recap);
  adhMail_(st.contact, 'Nouvelle inscription ' + dossier + ' – ' + r.nom + ' ' + r.prenom,
    r.prenom + ' ' + r.nom + ' – ' + r.tel + ' – ' + r.email + '\n' + r.adresse + ', ' + r.cp + ' ' + r.ville + '\n\n' + recap);
  return { dossier: dossier, total: a.total, aidesMontant: a.aidesMontant, reste: Math.max(0, a.total - a.aidesMontant), mode: a.mode, email: r.email };
}

/** Accès à un dossier : n° de dossier + téléphone du responsable. */
function adhAuth_(dossier, tel) {
  const d = String(dossier || '').trim().toUpperCase();
  const t = adhTel_(tel);
  const a = adhRows_('Adhesions').find(function (x) { return String(x.Dossier).toUpperCase() === d; });
  if (!a || t.length < 9 || (adhTel_(a.Téléphone) !== t && adhTel_(a.Téléphone2) !== t)) {
    throw new Error('Dossier introuvable ou téléphone incorrect.');
  }
  return a;
}
function adhAidesMontant_(a) { const m = String(a.Aides || '').match(/\(([\d.,]+) €\)\s*$/); return m ? adhNum_(m[1]) : 0; }
// « Réglé » posé par le club = dossier soldé, quel que soit le montant saisi.
function adhReste_(a) { return a.Statut === 'Annulé' || a.Statut === 'Réglé' ? 0 : Math.max(0, adhNum_(a.Total) - adhAidesMontant_(a) - adhNum_(a.Réglé)); }

/** Consultation d'un dossier par la famille. */
function getAdhesion(dossier, tel) {
  ensureSheets_();
  let a = adhAuth_(dossier, tel);
  // Paiement HelloAsso en attente pour ce dossier : on le revérifie.
  if (helloAssoConfig_() && readTable_Safe_('Paiements').some(function (p) { return String(p.JoueurID) === 'ADH:' + a.Dossier && p.Statut === 'En attente'; })) {
    withScriptLock_(function () { checkPayments_('ADH:' + a.Dossier); });
    a = adhAuth_(dossier, tel);
  }
  const crs = adhCreneaux_();
  const byId = {};
  crs.forEach(function (c) { byId[c.id] = c; });
  const docs = adhRows_('AdhesionDocuments').filter(function (x) { return x.Dossier === a.Dossier; });
  return {
    dossier: a.Dossier, saison: a.Saison, creeLe: a.CreeLe,
    responsable: { nom: a.Nom, prenom: a.Prénom, adresse: a.Adresse, cp: a.CP, ville: a.Ville, tel: a.Téléphone, email: a.Email, tel2: a.Téléphone2, email2: a.Email2 },
    tarif: a.Tarif, total: adhNum_(a.Total), aides: a.Aides, aidesMontant: adhAidesMontant_(a), mode: a.Mode, echeances: adhNum_(a.Echeances) || 1,
    statut: a.Statut, regle: adhNum_(a.Réglé), reste: adhReste_(a),
    membres: adhRows_('AdhesionMembres').filter(function (m) { return m.Dossier === a.Dossier; }).map(function (m) {
      return {
        n: adhNum_(m.N), nom: m.Nom, prenom: m.Prénom, naissance: m.Naissance, sexe: m.Sexe, categorie: m.Catégorie,
        creneau1: adhSlotLabel_(byId[m.Creneau1]) || m.Creneau1, creneau2: m.Creneau2 ? (adhSlotLabel_(byId[m.Creneau2]) || m.Creneau2) : '',
        licence: m.Licence, numLicence: m.NumLicence, montant: adhNum_(m.Montant), statut: m.Statut,
        qs: m.QS, qsLe: m.QSLe, reglement: m.Reglement, reglementLe: m.ReglementLe,
        mineur: (adhAge_(m.Naissance) || 0) < 18,
        sante: adhSanteForm_(m, crs, a.Saison),
      };
    }),
    documents: docs.map(function (x) { return { n: adhNum_(x.N), type: x.Type, fichier: x.Fichier, date: x.EnvoyeLe }; }),
    cheques: adhRows_('AdhesionCheques').filter(function (x) { return x.Dossier === a.Dossier; }).map(function (x) {
      return { n: adhNum_(x.N), titulaire: x.Titulaire, banque: x.Banque, numero: x.Numero, montant: adhNum_(x.Montant), deposeLe: x.DeposeLe };
    }),
    helloasso: adhSettings_().modes.indexOf('HelloAsso') >= 0,
  };
}
function readTable_Safe_(name) { return ss_().getSheetByName(name) ? readTable_(name) : []; }

function adhMembre_(a, n) {
  const m = adhRows_('AdhesionMembres').find(function (x) { return x.Dossier === a.Dossier && adhNum_(x.N) === Number(n); });
  if (!m) throw new Error('Membre introuvable dans ce dossier.');
  return m;
}

/** Âge exact aujourd'hui. */
function adhAge_(naissance) {
  const b = adhBirth_(naissance);
  if (!b) return null;
  const d = new Date();
  return d.getFullYear() - b.y - ((d.getMonth() + 1 < b.m || (d.getMonth() + 1 === b.m && d.getDate() < b.d)) ? 1 : 0);
}
/** Certification médicale demandée par la FFTT (saison 2026/2027) :
 * - mineur : questionnaire de santé pour mineur n° 26-10-2 ;
 * - majeur en loisir : questionnaire pour majeur « licence Découverte »
 *   n° 26-10-3 ;
 * - majeur en compétition : PPS en ligne chaque année ; à 65 ans et plus,
 *   certificat médical à l'entrée dans chaque catégorie vétéran (V65, V70,
 *   V80, V85, V90).
 * Un certificat médical en cours de validité remplace toujours le PPS ou le
 * questionnaire. */
function adhSanteForm_(m, crs, saison) {
  const age = adhAge_(m.Naissance);
  if (age !== null && age < 18) return 'mineur';
  const byId = {};
  (crs || []).forEach(function (c) { byId[c.id] = c; });
  const c1 = byId[m.Creneau1], c2 = byId[m.Creneau2];
  if (!((c1 && c1.competition) || (c2 && c2.competition))) return 'decouverte';
  const b = adhBirth_(m.Naissance);
  const catAge = b ? Number(String(saison || defaultSaison_()).slice(0, 4)) - b.y : 0;
  return [65, 70, 80, 85, 90].indexOf(catAge) >= 0 ? 'certificat' : 'pps';
}
const ADH_FORMS_ = {
  mineur: { num: '26-10-2', titre: 'Questionnaire de Santé pour Mineur' },
  decouverte: { num: '26-10-3', titre: 'Questionnaire de Santé pour Majeur – Licence Découverte' },
};
function adhSigOk_(dataUrl) { return /^data:image\/png;base64,[A-Za-z0-9+/=]{200,}$/.test(String(dataUrl || '')) && String(dataUrl).length < 400000; }

/** Questionnaire de santé FFTT : les réponses restent sur l'appareil. Si
 * toutes sont NON, l'attestation (coupon du formulaire FFTT) est signée en
 * ligne et enregistrée en PDF dans le Drive du club ; sinon un certificat
 * médical est demandé. */
function submitQS(dossier, tel, n, toutNon, att) {
  ensureSheets_();
  const a = adhAuth_(dossier, tel);
  const m = adhMembre_(a, n);
  const form = adhSanteForm_(m, adhCreneaux_(), a.Saison);
  if (!ADH_FORMS_[form]) throw new Error('Pour ce membre, la FFTT demande ' + (form === 'pps' ? 'l\'attestation PPS' : 'un certificat médical') + ' (pas de questionnaire).');
  att = att || {};
  const f = ADH_FORMS_[form];
  const where = function (x) { return x.Dossier === a.Dossier && adhNum_(x.N) === adhNum_(m.N); };
  if (!toutNon) {
    adhUpdate_('AdhesionMembres', where, { QS: 'Certificat médical requis (questionnaire ' + f.num + ')', QSLe: adhNow_() });
    adhFamilyMail_(a.Dossier, 'Questionnaire de santé', 'Questionnaire de santé de ' + m.Prénom + ' : au moins une réponse OUI, un certificat médical est à fournir.');
    return getAdhesion(dossier, tel);
  }
  const nom = adhTxt_(att.nom || m.Nom, 60).toUpperCase(), prenom = adhTxt_(att.prenom || m.Prénom, 60);
  const lic = adhTxt_(att.licence, 20), lieu = adhTxt_(att.lieu, 60);
  const sigs = form === 'mineur' ? [['Signature du joueur mineur', att.sigJoueur], ['Signature du représentant légal', att.sigParent]]
    : [['Signature du titulaire ou du représentant légal', att.sigJoueur]];
  sigs.forEach(function (x) { if (!adhSigOk_(x[1])) throw new Error('Signature manquante : ' + x[0].toLowerCase() + '.'); });
  if (form === 'mineur' && !adhTxt_(att.parent, 80)) throw new Error('Indiquez le nom du représentant légal.');
  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const esc = function (v) { return String(v || '').replace(/&/g, '&amp;').replace(/</g, '&lt;'); };
  const decl = form === 'mineur'
    ? 'déclare avoir répondu NON à toutes les questions du questionnaire de santé.'
    : 'atteste avoir répondu NON à toutes les questions du questionnaire de santé nécessaire à ma prise de licence Découverte (uniquement)';
  const html = '<html><body style="font-family:Arial,sans-serif;font-size:12pt;">' +
    '<p style="font-size:10pt;color:#555;">' + esc(readConfigValue_('ClubNom', '')) + ' – dossier ' + esc(a.Dossier) + ' – saison ' + esc(a.Saison) + '</p>' +
    '<h2 style="text-align:center;margin-bottom:0;">N° ' + f.num + ' — ATTESTATION</h2>' +
    '<p style="text-align:center;margin-top:4px;">' + esc(f.titre) + '</p>' +
    '<p>Je soussigné :</p><p>NOM : <b>' + esc(nom) + '</b> &nbsp; PRÉNOM : <b>' + esc(prenom) + '</b></p>' +
    '<p>N° de licence : ' + esc(lic || '—') + (form === 'decouverte' ? ' &nbsp; Date de naissance : ' + esc(m.Naissance) : '') + '</p>' +
    '<p>' + decl + '</p>' +
    '<p>Je sais qu\'à travers cette attestation, et comme cela est précisé dans la circulaire administrative de la FFTT, j\'engage ma propre ' +
    'responsabilité et qu\'en aucun cas celle de la FFTT ne pourra être recherchée.</p>' +
    '<p>Fait' + (lieu ? ' à ' + esc(lieu) : '') + ' le ' + date + ' (signature électronique sur le site d\'inscription du club).</p>' +
    '<table style="width:100%;margin-top:10px;"><tr>' + sigs.map(function (x, i) {
      return '<td style="width:50%;vertical-align:top;"><p>' + esc(x[0]) + (form === 'mineur' && i === 1 ? ' : ' + esc(att.parent) : '') + '<br>Date : ' + date + '</p>' +
        '<img src="' + x[1] + '" style="width:260px;border:1px solid #999;"></td>';
    }).join('') + '</tr></table></body></html>';
  const fname = a.Dossier + ' – ' + prenom + ' ' + nom + ' – Attestation QS ' + f.num;
  let file;
  try {
    const pdf = HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf').setName(fname + '.pdf');
    file = adhFolder_().createFile(pdf);
  } catch (e) {
    file = adhFolder_().createFile(Utilities.newBlob(html, 'text/html', fname + '.html'));
  }
  adhAppend_('AdhesionDocuments', [a.Dossier, adhNum_(m.N), 'Attestation QS ' + f.num, fname, file.getUrl(), adhNow_()]);
  if (lic && !m.NumLicence) adhUpdate_('AdhesionMembres', where, { NumLicence: lic });
  adhUpdate_('AdhesionMembres', where, { QS: 'Attestation ' + f.num + ' signée (NON à toutes les questions)', QSLe: adhNow_() });
  adhFamilyMail_(a.Dossier, 'Questionnaire de santé', 'Nous avons bien reçu l\'attestation de santé signée (FFTT n° ' + f.num + ') de ' + m.Prénom + '.');
  return getAdhesion(dossier, tel);
}

/** Règlement intérieur accepté pour un membre. */
function acceptReglement(dossier, tel, n, signataire) {
  ensureSheets_();
  const a = adhAuth_(dossier, tel);
  const m = adhMembre_(a, n);
  const sig = adhTxt_(signataire, 80);
  if (!sig) throw new Error('Indiquez le nom du signataire.');
  adhUpdate_('AdhesionMembres', function (x) { return x.Dossier === a.Dossier && adhNum_(x.N) === adhNum_(m.N); },
    { Reglement: 'Accepté (' + sig + ')', ReglementLe: adhNow_() });
  adhFamilyMail_(a.Dossier, 'Règlement intérieur', 'Le règlement intérieur est accepté pour ' + m.Prénom + '.');
  return getAdhesion(dossier, tel);
}

function adhFolder_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('ADH_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* dossier supprimé : on le recrée */ } }
  const f = DriveApp.createFolder('RebonTT – Documents des adhésions');
  p.setProperty('ADH_FOLDER_ID', f.getId());
  return f;
}

/** Document envoyé par la famille (certificat médical, justificatif d'aide…),
 * enregistré dans un dossier PRIVÉ du Drive du club (non partagé). */
function uploadAdhesionDocument(dossier, tel, n, type, filename, mime, base64) {
  ensureSheets_();
  const a = adhAuth_(dossier, tel);
  let who = 'Famille';
  if (Number(n) > 0) { const m = adhMembre_(a, n); who = m.Prénom + ' ' + m.Nom; }
  if (ADH_DOC_TYPES_.indexOf(type) < 0) throw new Error('Type de document invalide.');
  if (['image/jpeg', 'image/png', 'application/pdf'].indexOf(mime) < 0) throw new Error('Format accepté : photo (JPEG, PNG) ou PDF.');
  const data = String(base64 || '').replace(/^data:[^,]+,/, '');
  if (!data) throw new Error('Fichier vide.');
  if (data.length > 7 * 1024 * 1024) throw new Error('Fichier trop volumineux (5 Mo maximum).');
  const name = adhTxt_(filename, 80).replace(/[\\/:*?"<>|]/g, '_') || 'document';
  const blob = Utilities.newBlob(Utilities.base64Decode(data), mime, a.Dossier + ' – ' + who + ' – ' + type + ' – ' + name);
  const file = adhFolder_().createFile(blob);
  adhAppend_('AdhesionDocuments', [a.Dossier, Number(n) || 0, type, name, file.getUrl(), adhNow_()]);
  // Un certificat médical (toujours accepté) ou l'attestation PPS valent
  // certification médicale pour le membre.
  if (Number(n) > 0 && (type === 'Certificat médical' || type === 'Attestation PPS')) {
    adhUpdate_('AdhesionMembres', function (x) { return x.Dossier === a.Dossier && adhNum_(x.N) === Number(n); },
      { QS: type === 'Certificat médical' ? 'Certificat médical fourni' : 'Attestation PPS fournie', QSLe: adhNow_() });
  }
  adhMail_(adhSettings_().contact, 'Document reçu – dossier ' + a.Dossier, type + ' pour ' + who + ' : ' + file.getUrl());
  adhFamilyMail_(a.Dossier, 'Document reçu', 'Nous avons bien reçu le document « ' + type + ' » (' + name + ') pour ' + who + '.');
  return getAdhesion(dossier, tel);
}

/** Demande de modification des activités (traitée par le club). */
function requestAdhesionChange(dossier, tel, message) {
  ensureSheets_();
  const a = adhAuth_(dossier, tel);
  const msg = String(message || '').trim().slice(0, 2000);
  if (!msg) throw new Error('Décrivez la modification souhaitée.');
  adhAppend_('AdhesionMessages', [adhNow_(), a.Dossier, a.Prénom + ' ' + a.Nom, a.Email, a.Téléphone, 'Modification des activités', msg, 'Nouveau']);
  adhFamilyMail_(a.Dossier, 'Demande de modification reçue', 'Nous avons bien reçu votre demande de modification des activités :\n« ' + msg + ' »\nLe club vous recontacte rapidement.');
  adhMail_(adhSettings_().contact, 'Demande de modification – dossier ' + a.Dossier,
    a.Prénom + ' ' + a.Nom + ' (' + a.Téléphone + ', ' + a.Email + ') :\n\n' + msg);
  return { ok: true };
}

/** Message libre envoyé au club depuis la page d'inscription. */
function sendAdhesionMessage(d) {
  ensureSheets_();
  d = d || {};
  if (d.website) throw new Error('Demande refusée.');
  const nom = adhTxt_(d.nom, 80), email = adhTxt_(d.email, 120), tel = adhTxt_(d.tel, 20);
  const objet = adhTxt_(d.objet, 120) || 'Question', msg = String(d.message || '').trim().slice(0, 3000);
  if (!nom || !msg) throw new Error('Indiquez votre nom et votre message.');
  if (!/^\S+@\S+\.\S+$/.test(email) && adhTel_(tel).length < 9) throw new Error('Indiquez un e-mail ou un téléphone pour la réponse.');
  adhAppend_('AdhesionMessages', [adhNow_(), adhTxt_(d.dossier, 20).toUpperCase(), nom, email, tel, objet, msg, 'Nouveau']);
  adhMail_(adhSettings_().contact, 'Message (inscriptions) : ' + objet, nom + ' (' + [tel, email].filter(Boolean).join(', ') + ') :\n\n' + msg);
  return { ok: true };
}

/** Paiement en ligne du reste à régler d'un dossier (HelloAsso). */
function createAdhesionCheckout(dossier, tel) {
  ensureSheets_();
  const cfg = helloAssoConfig_();
  const a = adhAuth_(dossier, tel);
  const reste = adhReste_(a);
  if (!reste) throw new Error('Rien à régler pour ce dossier.');
  // Sans clés API : page de paiement HelloAsso du club (montant à saisir par
  // la famille ; le club valide le règlement à réception).
  if (!cfg) {
    const lien = adhSettings_().helloassoLien;
    if (!lien) throw new Error('Le paiement en ligne n\'est pas configuré par le club.');
    return { redirectUrl: lien, lien: true, montant: reste };
  }
  const label = ('Adhésion tennis de table ' + a.Saison + ' – dossier ' + a.Dossier + ' – ' + a.Prénom + ' ' + a.Nom).slice(0, 250);
  const page = adhPublicUrl_() + '?dossier=' + encodeURIComponent(a.Dossier) + '&helloasso=';
  const cents = Math.round(reste * 100);
  const clean = function (v) { return String(v || '').replace(/[^A-Za-zÀ-ÖØ-öø-ÿ' -]/g, '').trim(); };
  const body = {
    totalAmount: cents, initialAmount: cents, itemName: label,
    backUrl: page + 'retour', errorUrl: page + 'erreur', returnUrl: page + 'ok',
    containsDonation: false, metadata: { adhesion: a.Dossier },
  };
  const first = clean(a.Prénom), last = clean(a.Nom);
  if (first && last) body.payer = { firstName: first, lastName: last, email: a.Email };
  const intent = helloAssoApi_(cfg, 'post', '/checkout-intents', body);
  if (!intent || !intent.redirectUrl) throw new Error('HelloAsso n\'a pas renvoyé de page de paiement.');
  withScriptLock_(function () {
    ensurePaymentsSheet_().appendRow([String(intent.id), 'ADH:' + a.Dossier, '', reste, label, 'En attente', new Date(), '', '']);
  });
  return { redirectUrl: intent.redirectUrl };
}

/** Paiement HelloAsso confirmé pour un dossier d'adhésion. */
function adhMarkPaid_(dossier, montant) {
  adhUpdate_('Adhesions', function (x) { return String(x.Dossier) === String(dossier); }, function (x) {
    const regle = adhNum_(x.Réglé) + adhNum_(montant);
    const du = adhNum_(x.Total) - adhAidesMontant_(x);
    return { Réglé: regle, Statut: regle >= du ? 'Réglé' : 'Partiellement réglé', Note: (x.Note ? x.Note + ' ; ' : '') + 'HelloAsso ' + montant + ' € le ' + adhNow_() };
  });
  adhFamilyMail_(dossier, 'Paiement reçu', 'Votre paiement en ligne de ' + adhNum_(montant) + ' € (HelloAsso) est bien reçu. Merci !');
}

// --- Administration (jeton remis à la connexion administrateur) ---

function adminToken_() {
  const p = PropertiesService.getScriptProperties();
  let t = p.getProperty('ADMIN_TOKEN');
  if (!t) { t = Utilities.getUuid(); p.setProperty('ADMIN_TOKEN', t); }
  return t;
}
function requireAdmin_(token) {
  const t = PropertiesService.getScriptProperties().getProperty('ADMIN_TOKEN');
  if (!t || String(token || '') !== t) throw new Error('Session administrateur expirée : déconnectez-vous puis reconnectez-vous.');
}

/** Admin : réglages, créneaux (places), dossiers et messages. */
function getAdhesionsAdmin(token) {
  ensureSheets_();
  requireAdmin_(token);
  const m = configMap_();
  const raw = {};
  ['AdhSaison', 'AdhOuvert', 'AdhCommune', 'AdhSupplementExterieur', 'AdhReductionFamille', 'AdhModes', 'AdhChequesMax', 'AdhOrdreCheque',
    'AdhIban', 'AdhAides', 'AdhReglementUrl', 'AdhReglementTexte', 'AdhContactEmail', 'AdhInfos', 'AdhLicences', 'AdhPpsUrl', 'AdhLicenceUrl', 'AdhHelloAssoLien', 'AdhBanqueTitulaire', 'AdhBanqueNom', 'AdhBanqueAgence', 'AdhBic'].forEach(function (k) { raw[k] = m[k] || ''; });
  const crs = adhCreneaux_();
  const byId = {};
  crs.forEach(function (c) { byId[c.id] = c; });
  const membres = adhRows_('AdhesionMembres');
  const docs = adhRows_('AdhesionDocuments');
  const cheques = adhRows_('AdhesionCheques');
  const msgSheet = adhSheet_('AdhesionMessages').getDataRange().getValues();
  const msgs = [];
  for (let i = 1; i < msgSheet.length; i++) {
    const r = msgSheet[i];
    if (!r.join('')) continue;
    msgs.push({ row: i + 1, date: String(r[0]), dossier: r[1], nom: r[2], email: r[3], tel: r[4], objet: r[5], message: r[6], statut: r[7] });
  }
  return {
    raw: raw, settings: adhSettings_(), helloasso: !!helloAssoConfig_(), url: adhPublicUrl_(),
    creneaux: crs,
    dossiers: adhRows_('Adhesions').map(function (a) {
      return {
        dossier: a.Dossier, creeLe: a.CreeLe, saison: a.Saison, nom: a.Nom, prenom: a.Prénom, adresse: a.Adresse, cp: a.CP, ville: a.Ville,
        tel: a.Téléphone, email: a.Email, tel2: a.Téléphone2, email2: a.Email2, tarif: a.Tarif, total: adhNum_(a.Total), aides: a.Aides,
        mode: a.Mode, echeances: a.Echeances, statut: a.Statut, regle: adhNum_(a.Réglé), reste: adhReste_(a), note: a.Note,
        membres: membres.filter(function (x) { return x.Dossier === a.Dossier; }).map(function (x) {
          return {
            n: adhNum_(x.N), nom: x.Nom, prenom: x.Prénom, naissance: x.Naissance, sexe: x.Sexe, categorie: x.Catégorie,
            creneau1: adhSlotLabel_(byId[x.Creneau1]) || x.Creneau1, creneau2: x.Creneau2 ? (adhSlotLabel_(byId[x.Creneau2]) || x.Creneau2) : '',
            licence: x.Licence, numLicence: x.NumLicence, montant: adhNum_(x.Montant), qs: x.QS, reglement: x.Reglement,
            sante: adhSanteForm_(x, crs, a.Saison),
            statut: x.Statut, joueurId: x.JoueurID,
          };
        }),
        documents: docs.filter(function (x) { return x.Dossier === a.Dossier; })
          .map(function (x) { return { n: adhNum_(x.N), type: x.Type, fichier: x.Fichier, lien: x.Lien, date: x.EnvoyeLe }; }),
        cheques: cheques.filter(function (x) { return x.Dossier === a.Dossier; }).map(function (x) {
          return { n: adhNum_(x.N), titulaire: x.Titulaire, banque: x.Banque, numero: x.Numero, montant: adhNum_(x.Montant), deposeLe: x.DeposeLe };
        }),
      };
    }).reverse(),
    messages: msgs.reverse(),
  };
}

/** Admin : réglages des inscriptions. */
function saveAdhesionSettings(token, s) {
  requireAdmin_(token);
  Object.keys(s || {}).forEach(function (k) {
    if (!/^Adh[A-Za-z]+$/.test(k)) throw new Error('Réglage inconnu : ' + k);
    const v = String(s[k] === undefined || s[k] === null ? '' : s[k]);
    if (k === 'AdhLicences') { writeConfig_(k, JSON.stringify(adhLicences_(v))); return; }
    if (k === 'AdhIban' && v.trim() && !/^[A-Z]{2}\d{2}[A-Z0-9 ]{10,40}$/i.test(v.trim())) throw new Error('IBAN invalide (ex. FR76 1234 5678 9012 3456 7890 123).');
    if (k === 'AdhBic' && v.trim() && !/^[A-Z]{6}[A-Z0-9]{2}([A-Z0-9]{3})?$/i.test(v.trim().replace(/\s/g, ''))) throw new Error('BIC invalide (8 ou 11 caractères, ex. AGRIFRPP866).');
    if (k === 'AdhHelloAssoLien' && v.trim() && !/^https:\/\/(www\.)?helloasso\.com\/\S+$/i.test(v.trim())) throw new Error('Lien HelloAsso invalide (https://www.helloasso.com/…).');
    if (k === 'AdhLicenceUrl' && v.trim() && !/^https?:\/\/\S+$/i.test(v.trim())) throw new Error('Lien du formulaire de licence invalide (https://…).');
    if (k === 'AdhPpsUrl' && v.trim() && !/^https?:\/\/\S+$/i.test(v.trim())) throw new Error('Lien PPS invalide (https://…).');
    if (k === 'AdhReglementUrl' && v.trim() && !/^https?:\/\/\S+$/i.test(v.trim())) throw new Error('Lien du règlement invalide (https://…).');
    writeConfig_(k, v);
  });
  return getAdhesionsAdmin(token);
}

/** Admin : statut de règlement, montant réglé et note d'un dossier. */
function setAdhesionAdmin(token, dossier, patch) {
  requireAdmin_(token);
  patch = patch || {};
  const p = {};
  if (patch.statut !== undefined) {
    if (ADH_STATUTS_.indexOf(patch.statut) < 0) throw new Error('Statut invalide.');
    p.Statut = patch.statut;
  }
  if (patch.regle !== undefined) p.Réglé = adhNum_(patch.regle);
  if (patch.note !== undefined) p.Note = String(patch.note || '').slice(0, 500);
  const before = adhRows_('Adhesions').find(function (x) { return String(x.Dossier) === String(dossier); });
  if (!adhUpdate_('Adhesions', function (x) { return String(x.Dossier) === String(dossier); }, p)) throw new Error('Dossier introuvable.');
  // Avancement du règlement : la famille est prévenue.
  const statutChange = p.Statut !== undefined && p.Statut !== before.Statut;
  const regleChange = p.Réglé !== undefined && p.Réglé !== adhNum_(before.Réglé);
  if (statutChange || regleChange) {
    adhFamilyMail_(dossier, p.Statut === 'Annulé' ? 'Inscription annulée' : 'Règlement',
      p.Statut === 'Annulé' ? 'Votre inscription a été annulée par le club. Pour toute question, répondez à ce message ou contactez le club.'
        : 'Le club a mis à jour le règlement de votre dossier : ' + (p.Statut || before.Statut) + (regleChange ? ', montant réglé ' + p.Réglé + ' €' : '') + '.');
  }
  return getAdhesionsAdmin(token);
}

/** Admin : message traité / à traiter. */
function setAdhesionMessageStatut(token, row, statut) {
  requireAdmin_(token);
  const sh = adhSheet_('AdhesionMessages');
  if (Number(row) < 2 || Number(row) > sh.getLastRow()) throw new Error('Message introuvable.');
  sh.getRange(Number(row), 8).setValue(statut === 'Traité' ? 'Traité' : 'Nouveau');
  return getAdhesionsAdmin(token);
}

/** Admin : ajoute un membre d'un dossier à l'effectif (nouvelle ligne de
 * Joueurs, sans équipe, avec un code personnel tiré au hasard). */
function createPlayerFromAdhesion(token, dossier, n) {
  requireAdmin_(token);
  const a = adhRows_('Adhesions').find(function (x) { return String(x.Dossier) === String(dossier); });
  if (!a) throw new Error('Dossier introuvable.');
  const m = adhMembre_(a, n);
  if (m.JoueurID) throw new Error('Ce membre est déjà dans l\'effectif.');
  const used = {};
  readTable_('Joueurs').forEach(function (p) { used[String(p.PIN)] = true; });
  used[String(readConfigValue_('AdminPIN', ''))] = true;
  let pin = '';
  do { pin = String(Math.floor(100000 + Math.random() * 900000)); } while (used[pin]);
  const res = addPlayer(m.Nom, m.Prénom, 0, pin, false, a.Téléphone, m.Catégorie, m.Sexe, m.NumLicence);
  adhUpdate_('AdhesionMembres', function (x) { return x.Dossier === a.Dossier && adhNum_(x.N) === adhNum_(m.N); }, { JoueurID: res.id });
  adhFamilyMail_(a.Dossier, 'Bienvenue au club', m.Prénom + ' est inscrit(e) dans l\'effectif du club. Son code personnel pour l\'application RebonTT ' +
    '(disponibilités, convocations, compétitions) est : ' + pin + '. Il est personnel : ne le communiquez pas.');
  return { id: res.id, pin: pin, admin: getAdhesionsAdmin(token) };
}

/** À lancer UNE FOIS depuis l'éditeur Apps Script (menu Exécuter) : Google
 * demande l'autorisation d'envoyer des e-mails (confirmations d'inscription)
 * et d'enregistrer les documents dans le Drive du club. */
function autoriserInscriptions() {
  adhFolder_();
  Logger.log('Autorisations OK. Quota d\'e-mails restant aujourd\'hui : ' + MailApp.getRemainingDailyQuota());
}

// ----------------------------------------------------------------------------------
// COMITÉ DÉPARTEMENTAL (CD76TT) — inscriptions validées de tous les clubs
// ----------------------------------------------------------------------------------
// Le comité (www/comite.html, ou son propre script backend/Comite.gs) interroge
// le script de chaque club avec un code secret. Ce code est rangé dans les
// propriétés du script du club (COMITE_CODE, 12 caractères minimum) : il
// n'apparaît ni dans le code, ni dans l'application. Après 10 codes erronés,
// l'accès est bloqué 15 minutes.

function comiteCheck_(code) {
  const expected = String(PropertiesService.getScriptProperties().getProperty('COMITE_CODE') || '');
  if (expected.length < 12) throw new Error('Accès comité non configuré pour ce club.');
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('COMITE_FAILS') || 0);
  if (fails >= 10) throw new Error('Trop de tentatives : réessayez dans 15 minutes.');
  const given = String(code || '');
  let diff = given.length ^ expected.length;
  for (let i = 0; i < expected.length; i++) diff |= expected.charCodeAt(i) ^ given.charCodeAt(i % Math.max(given.length, 1));
  if (diff !== 0) {
    cache.put('COMITE_FAILS', String(fails + 1), 900);
    throw new Error('Code comité incorrect.');
  }
}

/** Comité : inscriptions VALIDÉES par le club aux compétitions individuelles
 * (jeunes, adultes, stages…), hors participations annulées. */
function getComiteExport(code) {
  ensureSheets_();
  comiteCheck_(code);
  const settings = competitionSettings_();
  const comps = {};
  readCompetitions_(settings).forEach(function (c) { comps[c.id] = c; });
  const players = {};
  readTable_('Joueurs').forEach(function (p) { players[String(p.ID)] = p; });
  const rows = readInscriptions_().filter(function (i) {
    return i.statut === 'Validée' && i.participation !== 'Annulée' && comps[i.competitionId];
  }).map(function (i) {
    const c = comps[i.competitionId], p = players[String(i.playerId)] || {};
    return {
      type: c.type, competition: c.nom, epreuve: c.epreuve, date: c.date, lieu: c.lieu,
      nom: String(p.Nom || ''), prenom: String(p.Prénom || ''), categorie: String(p.Catégorie || ''),
      sexe: String(p.Sexe || '').toUpperCase(), licence: String(p.Licence || ''),
      points: p.Points === '' || p.Points === undefined ? '' : Number(p.Points) || '',
      equipe: i.equipe, participation: i.participation, reglement: i.reglement,
    };
  });
  return {
    club: { nom: String(readConfigValue_('ClubNom', '') || ''), ville: String(readConfigValue_('ClubVille', '') || ''), version: REBONTT_VERSION },
    generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
    inscriptions: rows,
    // Compétitions du calendrier du club (pour les fiches documents du comité).
    competitions: Object.keys(comps).reduce(function (acc, id) {
      const c = comps[id], k = c.type + '|' + c.nom;
      if (!acc.seen[k]) { acc.seen[k] = true; acc.list.push({ type: c.type, nom: c.nom }); }
      return acc;
    }, { seen: {}, list: [] }).list,
  };
}

// ----------------------------------------------------------------------------------
// CLUB D'ESSAI — données fictives pour tester (nouveau classeur uniquement)
// ----------------------------------------------------------------------------------
/** À lancer UNE FOIS depuis l'éditeur Apps Script d'un classeur NEUF (copie
 * vide pour les essais) : crée un club fictif (équipes, joueurs, codes,
 * présentation, créneaux). Refuse de s'exécuter si l'effectif contient déjà
 * des joueurs : ne jamais lancer sur le classeur d'un vrai club. */
function initialiserClubEssai() {
  // Chaque étape est notée dans le journal : en cas d'erreur Google
  // (« erreur inconnue »…), on voit où elle s'est arrêtée, et on peut relancer.
  const etape = function (n, nom, fn) {
    Logger.log(n + '/6 ' + nom + '…');
    try { return fn(); } catch (e) { throw new Error('Étape ' + n + '/6 (' + nom + ') : ' + (e && e.message || e)); }
  };
  etape(1, 'Feuilles du classeur', function () { forceSheetsCheck(); });
  // Relance possible : seuls des joueurs d'essai (licence ESSAI…) déjà créés
  // par cette fonction sont acceptés ; le classeur d'un vrai club est refusé.
  const players = readTable_('Joueurs').filter(function (p) { return String(p.Nom || p.Prénom || '').trim(); });
  const reels = players.filter(function (p) { return String(p.Licence || '').indexOf('ESSAI') !== 0; });
  if (reels.length) throw new Error('Ce classeur contient déjà ' + reels.length + ' joueur(s) réel(s) : initialisation d\'essai refusée.');
  etape(2, 'Équipes', function () {
    const eq = ss_().getSheetByName('Equipes');
    const teams = [[1, 'ESSAI1', 'Régional 3', 'Dimanche', '', 'A'], [2, 'ESSAI2', 'Départemental 1', 'Vendredi', '20:00', 'B'],
      [3, 'ESSAI3', 'Départemental 2', 'Vendredi', '20:00', 'B'], [4, 'ESSAI4', 'Départemental 3', 'Vendredi', '20:00', 'C']];
    if (eq.getLastRow() > 1) eq.getRange(2, 1, eq.getLastRow() - 1, 6).clearContent();
    eq.getRange(2, 1, teams.length, 6).setNumberFormat('@').setValues(teams.map(function (r) { return r.map(String); }));
    teamsLoaded_ = false; loadTeams_();
  });
  // Joueurs fictifs (codes 6 chiffres affichés dans le journal d'exécution)
  const noms = [['MARTIN', 'Lucas', 'Senior', 'M', 1], ['BERNARD', 'Emma', 'Senior', 'F', 1], ['DUBOIS', 'Hugo', 'V40', 'M', 1], ['THOMAS', 'Léa', 'Senior', 'F', 1],
    ['ROBERT', 'Louis', 'V50', 'M', 2], ['RICHARD', 'Chloé', 'Junior', 'F', 2], ['PETIT', 'Jules', 'Senior', 'M', 2], ['DURAND', 'Inès', 'V40', 'F', 3],
    ['LEROY', 'Nathan', 'Cadet', 'M', 3], ['MOREAU', 'Manon', 'Senior', 'F', 3], ['SIMON', 'Tom', 'V60', 'M', 4], ['LAURENT', 'Zoé', 'Minime', 'F', 4],
    ['LEFEBVRE', 'Théo', 'Benjamin', 'M', 0], ['MICHEL', 'Jade', 'Poussin', 'F', 0]];
  const used = {};
  players.forEach(function (p) { used[String(p.PIN)] = true; });
  const codes = [];
  etape(3, 'Joueurs fictifs', function () {
    noms.forEach(function (n, i) {
      const deja = players.find(function (p) { return String(p.Nom) === n[0] && String(p.Prénom) === n[1]; });
      if (deja) { codes.push(n[1] + ' ' + n[0] + ' : ' + deja.PIN + ' (déjà créé)'); return; }
      let pin;
      do { pin = String(Math.floor(100000 + Math.random() * 900000)); } while (used[pin]);
      used[pin] = true;
      addPlayer(n[0], n[1], n[4], pin, i % 4 === 0, '06000000' + ('0' + i).slice(-2), n[2], n[3], 'ESSAI' + (1000 + i));
      codes.push(n[1] + ' ' + n[0] + ' : ' + pin + (i % 4 === 0 ? ' (capitaine)' : ''));
    });
  });
  // Dates de phase 1 autour d'aujourd'hui (une journée tous les 14 jours,
  // vendredi pour les groupes B et C, dimanche pour le groupe A).
  etape(4, 'Dates', function () {
  const ds = ss_().getSheetByName('Dates');
  const dv = ds.getDataRange().getValues();
  const start = new Date(); start.setHours(12, 0, 0, 0); start.setDate(start.getDate() - 28 - start.getDay() + 5); // un vendredi
  for (let i = 1; i < dv.length; i++) {
    if (Number(dv[i][0]) !== 1 || ['A', 'B', 'C'].indexOf(String(dv[i][1])) < 0) continue;
    const k = Number(dv[i][2]) - 1, d = new Date(start.getTime());
    d.setDate(d.getDate() + 14 * k + (String(dv[i][1]) === 'A' ? 2 : 0));
    ds.getRange(i + 1, 4).setNumberFormat('@').setValue(formatFrDate_(d));
    ds.getRange(i + 1, 6).setValue('J' + (k + 1));
  }
  });
  let admin;
  do { admin = String(Math.floor(100000 + Math.random() * 900000)); } while (used[admin]);
  etape(5, 'Code administrateur', function () { setAdminPin(admin); });
  etape(6, 'Présentation du club', function () {
  setClubInfo({
    nom: 'Club d\'essai RebonTT', ville: 'Rouen', adresse: 'Salle fictive, 76000 Rouen',
    entrainements: 'Club fictif pour tester RebonTT : les données ne sont pas réelles.', tarifs: 'Jeunes : 130 €\nAdultes : 150 €',
    creneaux: [
      { jour: 'Lundi', debut: '18:00', fin: '19:30', niveau: 'Jeunes débutants', public: 'Jeunes', tarif: 130, places: 12 },
      { jour: 'Mercredi', debut: '17:30', fin: '19:00', niveau: 'Jeunes compétiteurs', public: 'Jeunes', tarif: 150, seances: 2, places: 10, competition: 'oui' },
      { jour: 'Vendredi', debut: '17:30', fin: '19:00', niveau: 'Jeunes compétiteurs', public: 'Jeunes', tarif: 150, seances: 2, places: 10, competition: 'oui' },
      { jour: 'Mardi', debut: '20:00', fin: '22:00', niveau: 'Loisirs', public: 'Adultes', tarif: 120, places: 20 },
      { jour: 'Jeudi', debut: '20:00', fin: '22:00', niveau: 'Adultes compétiteurs', public: 'Adultes', tarif: 150, places: 16, competition: 'oui' },
    ],
  });
  });
  bumpDataVersion_();
  const msg = 'Club d\'essai prêt.\nCode administrateur : ' + admin + '\nCodes joueurs :\n' + codes.join('\n');
  Logger.log(msg);
  return msg;
}

/** Admin : chèques remis en banque (bordereau). refs = [{dossier, n}] ;
 * chaque chèque reçoit la date de dépôt, son montant s'ajoute au « réglé »
 * du dossier et la famille est prévenue. */
function setChequesDeposes(token, refs, date) {
  requireAdmin_(token);
  const d = adhTxt_(date, 10) || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  if (!/^\d{2}\/\d{2}\/\d{4}$/.test(d)) throw new Error('Date de dépôt invalide (JJ/MM/AAAA).');
  const parDossier = {};
  (refs || []).forEach(function (r) {
    let montant = 0;
    adhUpdate_('AdhesionCheques', function (x) { return String(x.Dossier) === String(r.dossier) && adhNum_(x.N) === Number(r.n) && !x.DeposeLe; }, function (x) {
      montant = adhNum_(x.Montant);
      return { DeposeLe: d };
    });
    if (montant) (parDossier[r.dossier] = parDossier[r.dossier] || []).push({ n: Number(r.n), montant: montant });
  });
  Object.keys(parDossier).forEach(function (dos) {
    const list = parDossier[dos], somme = list.reduce(function (t, c) { return t + c.montant; }, 0);
    adhUpdate_('Adhesions', function (x) { return String(x.Dossier) === dos; }, function (x) {
      const regle = adhNum_(x.Réglé) + somme, du = adhNum_(x.Total) - adhAidesMontant_(x);
      return { Réglé: regle, Statut: x.Statut === 'Annulé' ? 'Annulé' : (regle >= du ? 'Réglé' : 'Partiellement réglé'),
        Note: (x.Note ? x.Note + ' ; ' : '') + 'chèque(s) ' + list.map(function (c) { return c.n; }).join(', ') + ' déposé(s) le ' + d };
    });
    adhFamilyMail_(dos, 'Chèque encaissé', list.map(function (c) { return 'Le chèque n° ' + c.n + ' de ' + c.montant + ' €'; }).join(', ') +
      (list.length > 1 ? ' ont été remis' : ' a été remis') + ' en banque le ' + d + '.');
  });
  return getAdhesionsAdmin(token);
}
