// Configuration de l'application mobile RebonTT.
// SERVER_URL : URL du déploiement "Application Web" du script Google Apps
// Script (backend/Code.gs), se terminant par /exec. Voir README.md.
// CLUBS : annuaire des clubs proposés sur la page d'accueil (choix du club
// par son nom ou sa ville). Chaque club a son propre script (serverUrl) ;
// sans serverUrl, c'est SERVER_URL qui est utilisé.
window.APP_CONFIG = {
  SERVER_URL: 'https://script.google.com/macros/s/AKfycbzcCd1pZkHyZkY7KH704kZ3pzr2Hl5blMRL4SUdzAQeGgzS9DeKI1UR1vt459gBVw4Jog/exec',
  CLUBS: [
    { id: 'fri', nom: 'FR Isneauville Tennis de Table', ville: 'Isneauville', cp: '76230', logo: 'assets/logo-fri.png' },
  ],
};
