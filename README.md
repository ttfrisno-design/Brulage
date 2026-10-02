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

## 3. Version PC (application web installable)

Le même code (`www/`) est publié sur **GitHub Pages** par le workflow
**Version web (PC)** (`.github/workflows/pages.yml`) à chaque mise à jour de
la branche `main` :

> https://ttfrisno-design.github.io/Brulage/

Sur PC, ouvrir ce lien dans **Chrome ou Edge**, puis cliquer sur l'icône
**« Installer »** dans la barre d'adresse (ou menu ⋮ › *Installer
l'application*) : l'application s'ajoute au menu Démarrer / bureau avec
l'icône TTFRI et s'ouvre dans sa propre fenêtre. Fonctionne aussi sur Mac,
et dans n'importe quel navigateur sans installation.

Réglages GitHub à faire une seule fois :
1. **Settings › Branches** : branche par défaut = `main`.
2. **Settings › Pages** : *Source* = **GitHub Actions**.
3. Onglet **Actions** › *Version web (PC)* › **Run workflow** (ou toute
   fusion dans `main`).

## 4. Développement local

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

## 5. Publication dans les stores

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

## Convocations

- **Effectif** (admin) : renseigner le **Téléphone** de chaque joueur (colonne
  H de la feuille *Joueurs*).
- **Envoi** : carte « 📣 Envoyer les convocations » sous le tableau de brûlage
  (admin : toutes les équipes) ou sous la feuille de match (capitaine : son
  équipe). Choisir le match, relire le message, puis **Envoyer** :
  - la convocation est enregistrée (feuille *Convocations*) et s'affiche en
    haut de l'appli des joueurs sélectionnés, avec les boutons **Présent /
    Absent** ;
  - sur téléphone, l'appli SMS s'ouvre avec les numéros et le message
    pré-remplis (il reste à appuyer sur Envoyer). Sur PC, les numéros sont
    affichés.
- Le tableau de la carte indique pour chaque joueur : non envoyée, envoyée
  sans réponse, ✓ Présent ou ✕ Absent.

## Coupe de Rouen (CDR1-CDR4) et Championnat Jeunes (CJ1-CJ4)

Ces deux compétitions par équipes se gèrent dans les **compétitions
individuelles** (et non dans le championnat FRI1-FRI8, ni dans le tableau
de brûlage) :

- **Coupe de Rouen** : compétitions individuelles *adultes* ;
  **Championnat Jeunes (équipe de 2)** : compétitions individuelles
  *jeunes* (réservé aux catégories jeunes).
- Le joueur demande son inscription à chaque tour / journée ;
  l'administrateur la valide et choisit l'équipe (**CDR1 à CDR4** ou **CJ1
  à CJ4**) dans la colonne « Équipe » de l'onglet Compétitions
  individuelles (feuille *Inscriptions*, colonne G) ; le joueur voit
  « Inscription validée — CJ3 » puis confirme sa participation.

## Compétitions individuelles (jeunes / adultes)

- Après la saisie du code, le joueur choisit : **Championnat par équipes**,
  **Compétitions individuelles jeunes** ou **Compétitions individuelles
  adultes**. L'administrateur arrive directement sur l'administration.
- Le calendrier 2026-2027 des plaquettes du club est créé automatiquement
  dans la feuille *Competitions* (une ligne par épreuve). L'onglet admin
  **Compétitions individuelles** permet d'ajouter, supprimer une épreuve et
  de modifier sa **date limite** (sinon : date − « délai par défaut »).
- **Effectif** : colonne **Catégorie** (Poussin, Benjamin, Minime, Cadet,
  Junior, Senior, V40…V80), qui détermine les compétitions proposées. Seul
  l'en-tête est ajouté à la feuille *Joueurs* : les joueurs existants ne
  sont pas modifiés.
- Parcours : le joueur **demande son inscription** avant la date limite → le
  club la **valide** (ou refuse) dans l'onglet admin → le joueur
  **confirme sa participation** (ou se désiste) à l'approche de l'épreuve.
  Suivi dans la feuille *Inscriptions*.
- **Rappels** : à l'ouverture de l'appli, un encadré « 🔔 À faire » (et un
  badge sur le bouton) signale les inscriptions dont la date limite approche
  et les participations à confirmer (délai réglable, 10 jours par défaut).
  L'admin peut aussi relancer par **SMS** les joueurs concernés non inscrits
  ou non confirmés.

## Limite de 4 joueurs par journée

Dans le tableau de brûlage (administrateur) et la feuille de match
(capitaine), une équipe ne peut avoir que **4 joueurs sélectionnés** pour
une même date : le choix d'un 5e joueur est refusé avec un message
d'avertissement listant les 4 joueurs déjà sélectionnés (retirer d'abord
l'un d'eux).

## Joueurs « Sans équipe »

Onglet **Effectif** : la colonne « Équipe » propose **Sans équipe** pour les
joueurs qui ne participent pas aux championnats par équipes (compétitions
individuelles uniquement). Ils sont regroupés dans le cadre « Sans équipe »,
n'apparaissent pas dans le **tableau de brûlage**, et le bouton
« Championnat par équipes » n'est pas proposé sur leur page d'accueil
(colonne *EquipeDomicile* vide dans la feuille *Joueurs*). Choisir une équipe
FRI pour les remettre dans le championnat.

## Mon profil (page d'accueil)

Après la saisie du code, le joueur voit un encadré **Mon profil** : équipe,
catégorie, matchs joués dans la saison (par équipe), prochain match et
sélections à venir (adversaire, lieu, réponse à la convocation), prochaine
compétition individuelle, indisponibilités à venir et équipes pour
lesquelles il est brûlé.

**FFTT** : l'onglet Effectif comporte les colonnes **Licence** et **Points**
(feuille *Joueurs*, colonnes J et K) ; le classement est calculé (points /
100). Dans *Dates & Réglages*, « Fiche joueur FFTT » permet de saisir
l'adresse de la fiche d'un joueur, où `{licence}` est remplacé par son
numéro de licence : un lien « Ma fiche FFTT » apparaît alors dans son
profil. La récupération automatique des points nécessite l'API officielle
de la FFTT (identifiants à demander par le club).

## Inscriptions par mail aux organisateurs

Onglet admin **Compétitions individuelles** : les épreuves sont présentées
**une colonne par compétition**. Pour une épreuve ayant des inscrits
validés, le bouton **« ✉️ Mail d'inscription à l'organisateur »** prépare le
message (liste des joueurs avec licence, catégorie, points et équipe
CDR / CJ) et l'ouvre dans la messagerie ; le texte est aussi affiché pour
être copié. Les adresses se règlent dans la carte « Mails d'inscription aux
organisateurs » (par défaut : Challenge SERANO → girard76@wanadoo.fr,
Coupe de Rouen → coupederouen@gmail.com, Championnat Jeunes →
lefebvrejmichel@hotmail.com ; feuille *Config*, clés `MailInscription:…`).

Le profil du joueur contient un lien vers le **site du club** (modifiable
dans *Dates & Réglages*).

## Arborescence

| Chemin | Rôle |
|---|---|
| `backend/Code.gs` | Script Google Apps Script (serveur, inchangé) |
| `backend/Index.html` | Version web d'origine (servie par `doGet`) |
| `www/` | Interface de l'application mobile |
| `www/config.js` | URL du déploiement Apps Script |
| `www/manifest.json`, `www/sw.js` | Version web installable (PC) |
| `assets/` | Sources des icônes / écran de démarrage |
| `android/`, `ios/` | Projets natifs générés par Capacitor |
| `capacitor.config.json` | Nom, identifiant, couleurs de l'application |
