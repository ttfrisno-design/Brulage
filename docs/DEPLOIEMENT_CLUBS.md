# RebonTT — déployer sur une centaine de clubs (Seine-Maritime)

## 1. Principe retenu : un classeur et un script par club

Chaque club garde **ses données dans son propre Google Sheet**, avec **son propre
script Apps Script** (copie de `backend/Code.gs`) déployé en « Application Web ».

- **Données séparées** : chaque club est responsable de ses données (RGPD) ; une
  erreur ou une panne chez un club ne touche pas les autres.
- **Quotas Google par club** : e-mails, exécutions, Drive, déclencheurs… sont
  ceux du compte Google du club. Un script unique pour 100 clubs atteindrait vite
  les limites d'un seul compte (30 exécutions simultanées, 100 à 1 500 e-mails
  par jour).
- **Une seule application** (GitHub Pages et appli Android) pour tous les clubs :
  le joueur choisit son club, et l'appli appelle le script de ce club.
- **L'annuaire `www/clubs.js`** liste les clubs (nom, ville, code postal, adresse
  /exec, logo). L'appli installée le relit en ligne : un nouveau club apparaît
  **sans mise à jour de l'appli**.
- **Le comité** (`comite.html` et `backend/Comite.gs`) interroge le script de
  chaque club avec son code secret ; sans liste `COMITE_CLUBS`, il prend
  l'annuaire.

## 2. Le club d'essai

Il sert à tester chaque nouvelle version avant de la proposer aux clubs.

1. Créer un **Google Sheet vide** (par exemple « RebonTT – club d'essai ») sur
   un compte Google de test, puis *Extensions › Apps Script* et coller
   `backend/Code.gs`.
2. Exécuter **`initialiserClubEssai`** (menu Exécuter). Cela crée :
   - 4 équipes ESSAI1 à ESSAI4 ;
   - 14 joueurs fictifs, avec des capitaines ;
   - des dates autour d'aujourd'hui ;
   - les créneaux et la présentation du club.

   Le **code administrateur et les codes joueurs** s'affichent dans le journal.
   La fonction refuse de s'exécuter si le classeur contient déjà des joueurs.
3. *Déployer › Nouveau déploiement › Application Web* (exécuter en tant que
   moi, accès : tout le monde), puis copier l'adresse /exec dans `clubs.js`,
   ligne `"id": "essai"`, champ `serverUrl`.
4. Ouvrir l'appli avec **`?essai=1`** (par exemple
   `https://ttfrisno-design.github.io/Brulage/?essai=1`) : le club d'essai
   apparaît dans la liste. Il reste masqué pour tout le monde sinon, et `?essai=0`
   le masque à nouveau. Côté comité, `comite.html?essai=1` l'inclut.

## 3. Installer un nouveau club (environ 15 minutes)

| Étape | Qui | Action |
|---|---|---|
| 1 | Référent RebonTT | Envoyer le lien « Faire une copie » du **classeur modèle**. C'est un classeur vide qui contient déjà le script : une copie de Google Sheet emporte son script. |
| 2 | Club | Faire la copie avec le compte Google du club (de préférence un compte « club », pas personnel). |
| 3 | Club | *Extensions › Apps Script* : exécuter `autoriserInscriptions` (autorisations), puis *Déployer › Application Web* et copier l'adresse /exec. |
| 4 | Club | Remplir l'onglet **Equipes** (nom, division, jour, heure, groupe de dates A/B/C) et l'onglet **Joueurs**, ou importer l'effectif. |
| 5 | Club | Dans l'appli : code admin par défaut **0000**, à changer aussitôt. Puis renseigner les dates, la présentation du club, les créneaux, les adhésions et HelloAsso. |
| 6 | Club | Propriétés du script : `COMITE_CODE` (le code du comité), et le cas échéant les clés HelloAsso. |
| 7 | Référent | Ajouter une ligne dans `www/clubs.js`, mettre à jour `REBONTT_CLUBS_MAJ`, puis fusionner. Le club apparaît dans l'appli et chez le comité. |

Pour aller vite à 100 clubs : une **fiche d'installation** d'une page, une
**visio collective** par vague de 10 clubs, et un **formulaire** où le club colle
son adresse /exec (le référent l'ajoute à l'annuaire).

## 4. Mises à jour du code sur 100 clubs

- **Numéro de version** : `REBONTT_VERSION` dans `Code.gs`. La page comité
  l'affiche au survol de chaque club, ce qui permet de voir qui n'est pas à jour.
- **Aujourd'hui** : chaque club recolle `Code.gs` et redéploie une nouvelle
  version. C'est acceptable pour quelques clubs, mais pas pour 100.
- **Étape suivante, à prévoir avant 20 clubs : une bibliothèque Apps Script
  « RebonTT »**.
  - Le script du club ne contient plus que quelques lignes, qui appellent la
    bibliothèque : `doGet`, `doPost`, les déclencheurs et les fonctions de
    configuration.
  - Une mise à jour = une nouvelle version de la bibliothèque. Les clubs la
    reçoivent sans rien recoller : il suffit de choisir le mode « dernière
    version », ou que le référent fasse passer les clubs d'une version à l'autre.
  - Travail nécessaire : passer au code de la bibliothèque les propriétés et le
    classeur du club (`PropertiesService` et `SpreadsheetApp` côté club), car une
    bibliothèque lit sinon ses propres propriétés.
- **Recette** : chaque version est d'abord installée sur le **club d'essai**,
  puis sur 2 ou 3 clubs pilotes, avant d'être diffusée.

## 5. Ce qui est propre à chaque club

| Élément | Où |
|---|---|
| Équipes (noms, divisions, jours, groupes de dates) | onglet **Equipes** |
| Dates de championnat (phases, journées) | appli › Dates & Réglages |
| Joueurs, capitaines, codes | onglet Joueurs / appli › Effectif |
| Présentation, créneaux, tarifs, adhésions | appli › Dates & Réglages, Adhésions |
| Compétitions individuelles, tarifs, dates limites | appli › Compétitions individuelles |
| HelloAsso, code comité, notifications | propriétés du script |

À mutualiser ensuite : le **calendrier départemental** des compétitions
individuelles et des stages est le même pour tous les clubs. Le comité pourrait
le publier une fois, et chaque club l'importerait.

## 6. Montée en charge par étapes

1. **Pilote (1 à 3 mois)** : FRI, club d'essai et 2 ou 3 clubs volontaires.
   Fiche d'installation, retours, corrections.
2. **Vague 1 (20 clubs)** : bibliothèque Apps Script en place, classeur modèle,
   visios collectives, formulaire d'inscription des clubs.
3. **Généralisation (100 clubs)** : support par le comité (FAQ, référent),
   suivi des versions depuis l'espace comité, annuaire alimenté par le
   formulaire.

## 7. Points de vigilance

- **Compte Google du club** : privilégier une adresse « club » partagée, pour
  ne pas perdre l'accès quand un bénévole part.
- **Codes** : code admin à changer dès l'installation ; `COMITE_CODE` long
  (12 caractères ou plus), identique chez tous les clubs et transmis hors
  e-mail.
- **RGPD** : chaque club est responsable de ses données, et le comité ne voit
  que les inscriptions validées. Une mention d'information est à prévoir dans
  chaque club.
- **Comité** : avec 100 clubs, la page interroge 100 scripts. Il faut compter
  de 30 secondes à 1 minute, et le Google Sheet du comité (mis à jour toutes les
  heures) reste la vue de référence.
