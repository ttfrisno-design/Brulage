// Configuration de l'application mobile RebonTT.
// SERVER_URL : URL du déploiement "Application Web" du script Google Apps
// Script (backend/Code.gs), se terminant par /exec. Voir README.md.
// CLUBS : annuaire des clubs proposés sur la page d'accueil (choix du club
// par son nom ou sa ville). Chaque club a son propre script (serverUrl) ;
// sans serverUrl, c'est SERVER_URL qui est utilisé.
window.APP_CONFIG = {
  SERVER_URL: 'https://script.google.com/macros/s/AKfycbzcCd1pZkHyZkY7KH704kZ3pzr2Hl5blMRL4SUdzAQeGgzS9DeKI1UR1vt459gBVw4Jog/exec',
  // Espace comité départemental (comite.html, non affiché dans l'appli) :
  // nom et lien du Google Sheet du comité (backend/Comite.gs). Aucun secret
  // ici : le code comité est saisi à la connexion.
  COMITE: { nom: 'CD76TT', sheetUrl: 'https://docs.google.com/spreadsheets/d/11Q5sF76it4w4TWKHgh2wfHrnfriDhzo_NIt_kGMEdzQ/edit' },
  CLUBS: [
    { id: 'fri', nom: 'FR Isneauville Tennis de Table', ville: 'Isneauville', cp: '76230', logo: 'assets/logo-fri.png' },
  ],
};
