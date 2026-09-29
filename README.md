# Brûlages FRI – Application mobile (Android & iOS)

Application de gestion des brûlages du FRI Tennis de Table (disponibilités des
joueurs, feuilles de match des capitaines, tableau de brûlage de
l'administrateur), en version **application mobile Android et iPhone**.

## Principe

```
 Application mobile (Capacitor)          Google Apps Script           Google Sheets
 ┌──────────────────────────┐   fetch   ┌──────────────────┐        ┌──────────────┐
 │ www/index.html (interface)│ ───────▶ │ backend/Code.gs  │ ─────▶ │ Config,      │
 │ www/config.js (URL /exec) │  POST    │ doPost (API JSON)│        │ Joueurs, ... │
 └──────────────────────────┘           └──────────────────┘        └──────────────┘
```

- **Le serveur ne change pas** : `backend/Code.gs` est le script Apps Script
  existant. Il expose déjà une API JSON (`doPost`) que l'application appelle.
  La version web (`doGet` + `backend/Index.html`) continue de fonctionner.
- **L'interface** (`www/index.html`) reprend `Index.html`, adaptée au mobile :
  - URL du serveur lue dans `www/config.js` (au lieu de `<?!= scriptUrl ?>`) ;
  - logos en fichiers (`www/assets/`) au lieu du base64 intégré ;
  - prise en compte de l'encoche iPhone / barre d'état Android ;
  - session conservée entre deux lancements (plus besoin de retaper son code
    à chaque ouverture) ; la configuration (dates) est rafraîchie en
    arrière-plan au démarrage.
- **Capacitor** emballe cette interface dans une vraie application native
  (`android/`, `ios/`), publiable sur le Play Store et l'App Store.

## 1. Configurer l'URL du serveur (obligatoire)

1. Dans l'éditeur Apps Script : **Déployer › Gérer les déploiements** (ou
   *Nouveau déploiement › Application Web*).
   - *Exécuter en tant que* : **Moi**
   - *Qui a accès* : **Tout le monde** (indispensable : l'application appelle
     le script sans compte Google).
2. Copier l'URL qui se termine par `/exec`.
3. La coller dans `www/config.js` :
   ```js
   window.APP_CONFIG = {
     SERVER_URL: 'https://script.google.com/macros/s/AKfycb.../exec',
   };
   ```

> Après chaque modification de `Code.gs`, mettre à jour le **même**
> déploiement (*Gérer les déploiements › Modifier › Nouvelle version*) pour
> garder la même URL.

## 2. Obtenir un APK Android sans rien installer

Chaque `push` sur GitHub lance le workflow **Applications mobiles**
(`.github/workflows/mobile.yml`) qui compile un APK de test.
Onglet **Actions** › dernier run › artefact `brulages-fri-android-debug` :
le télécharger, le transférer sur le téléphone et l'installer (autoriser les
« sources inconnues »). Idéal pour tester avant publication.

## 3. Développement local

Prérequis : Node.js 22+, et
- Android : [Android Studio](https://developer.android.com/studio) (JDK 21 inclus) ;
- iOS : un Mac avec Xcode 16+ (obligatoire pour iOS, imposé par Apple).

```bash
npm install
npx cap sync            # copie www/ dans les projets natifs
npm run open:android    # ouvre Android Studio  -> bouton ▶ Run
npm run open:ios        # ouvre Xcode           -> bouton ▶ Run
```

À chaque modification de `www/`, relancer `npx cap sync`.

Icônes et écran de démarrage : générés depuis `assets/` (icône : logo TTFRI `assets/logo-ttfri.gif`) avec
`npm run icons`.

## 4. Publication dans les stores

### Google Play (Android)
1. Créer un compte [Google Play Console](https://play.google.com/console)
   (25 $ une fois).
2. Android Studio › **Build › Generate Signed App Bundle** › créer une clé de
   signature (la conserver précieusement !) › fichier `.aab`.
3. Play Console › créer l'application › importer le `.aab`, remplir la fiche
   (description, captures, politique de confidentialité), publier
   (d'abord en « test interne » pour les joueurs du club, c'est le plus simple).

### App Store (iPhone)
1. Compte [Apple Developer](https://developer.apple.com/programs/) (99 €/an).
2. Xcode › cible *App* › **Signing & Capabilities** › choisir l'équipe.
3. **Product › Archive** › *Distribute App* › App Store Connect.
4. Distribuer aux joueurs via **TestFlight** (le plus simple pour un club),
   ou soumettre à la validation Apple pour une publication publique.

Identifiant de l'application : `fr.fri.brulages` (modifiable dans
`capacitor.config.json` **avant** la première publication).

## Arborescence

| Chemin | Rôle |
|---|---|
| `backend/Code.gs` | Script Google Apps Script (serveur, inchangé) |
| `backend/Index.html` | Version web d'origine (servie par `doGet`) |
| `www/` | Interface de l'application mobile |
| `www/config.js` | URL du déploiement Apps Script |
| `assets/` | Sources des icônes / écran de démarrage |
| `android/`, `ios/` | Projets natifs générés par Capacitor |
| `capacitor.config.json` | Nom, identifiant, couleurs de l'application |
