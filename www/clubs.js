// Annuaire des clubs RebonTT (une ligne par club). Utilisé par l'appli
// (choix du club), le site d'inscription, l'espace comité et le script du
// comité (backend/Comite.gs). Le tableau doit rester au format JSON
// (guillemets doubles, pas de virgule finale) : l'appli installée le
// relit en ligne pour connaître les nouveaux clubs sans mise à jour.
//   id        : identifiant court, unique (minuscules, sans espace)
//   nom, ville, cp
//   serverUrl : adresse /exec du script Apps Script du club (vide = SERVER_URL)
//   logo      : chemin ou adresse du logo (facultatif)
//   essai     : true = club de test, masqué (visible avec ?essai=1)
window.REBONTT_CLUBS_MAJ = '2026-10-22b';
window.REBONTT_CLUBS = [
  { "id": "fri", "nom": "FR Isneauville Tennis de Table", "ville": "Isneauville", "cp": "76230", "logo": "assets/logo-fri.png",
    "serverUrl": "https://script.google.com/macros/s/AKfycbzcCd1pZkHyZkY7KH704kZ3pzr2Hl5blMRL4SUdzAQeGgzS9DeKI1UR1vt459gBVw4Jog/exec" },
  { "id": "essai", "nom": "Club d'essai RebonTT", "ville": "Rouen", "cp": "76000", "serverUrl": "https://script.google.com/macros/s/AKfycbxOEatd9JSW11O7m5X20E-i80ul1uZ3GVCtp7umu2XftNDa1-uvIfqTVyxkTfzFW0gR/exec", "essai": true }
];
