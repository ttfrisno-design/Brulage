// Configuration de l'application mobile RebonTT.
// SERVER_URL : URL du déploiement "Application Web" du script Google Apps
// Script (backend/Code.gs), se terminant par /exec. Voir README.md.
// CLUBS : annuaire des clubs (www/clubs.js). Chaque club a son propre
// script (serverUrl) ; sans serverUrl, c'est SERVER_URL qui est utilisé.
window.APP_CONFIG = {
  SERVER_URL: 'https://script.google.com/macros/s/AKfycbzcCd1pZkHyZkY7KH704kZ3pzr2Hl5blMRL4SUdzAQeGgzS9DeKI1UR1vt459gBVw4Jog/exec',
  // Espace comité départemental (comite.html, non affiché dans l'appli) :
  // nom et lien du Google Sheet du comité (backend/Comite.gs). Aucun secret
  // ici : le code comité est saisi à la connexion.
  // apiUrl : adresse /exec du script du comité déployé en Application Web
  // (documents des compétitions affichés dans l'appli).
  COMITE: { nom: 'CD76TT', sheetUrl: 'https://docs.google.com/spreadsheets/d/11Q5sF76it4w4TWKHgh2wfHrnfriDhzo_NIt_kGMEdzQ/edit', apiUrl: 'https://script.google.com/macros/s/AKfycbyrqCf5FpPG6MQaKSnk1c0sIYLGAO6-aLAfF-cfeSoqUsg7d1EGS2vD76V6h-w7g76gIA/exec' },
  // Annuaire des clubs : www/clubs.js (chargé avant ce fichier).
  CLUBS: window.REBONTT_CLUBS,
  CLUBS_URL: 'https://ttfrisno-design.github.io/Brulage/clubs.js',
};
