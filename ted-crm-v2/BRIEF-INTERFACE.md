# TED CRM — Brief interface

Document de reprise. Il fixe les **règles de travail**, les **décisions déjà prises**
et les **valeurs exactes** en vigueur, pour que toute évolution de l'interface reste
cohérente avec l'existant.

Tout ce qui suit est **déjà implémenté** dans `src/App.js`, sauf la section finale
« Chantiers ouverts ».

---

## 1. Contraintes de travail

| Règle | Détail |
|---|---|
| **Base de données** | Ne jamais modifier le schéma, les policies ou les données Supabase sans demande explicite. Les évolutions d'interface doivent s'appuyer sur le schéma existant. |
| **Périmètre** | Ne toucher qu'à ce qui est demandé. Le reste du CRM (Réservations, Clients, Menu, Communications) n'est pas à refactorer au passage. |
| **Données de test** | Préfixe `ZZDemo` / `ZZStat`, systématiquement supprimées après vérification. |
| **Avant commit** | `npx react-scripts build` doit passer, et le rendu être vérifié à l'écran (pas seulement compilé). |
| **Réglages de production** | Si un réglage est modifié pour tester (ex. `acceptation_auto`), le remettre à sa valeur d'origine. |

**Piège rencontré plusieurs fois :** un script d'édition qui échoue à mi-parcours
n'écrit rien mais laisse croire au succès. Toujours revérifier par mesure après coup.

**Piège plus vicieux, qui a déjà vidé ce fichier :** en Python,
`open(p,'w').write(open(p).read().replace(...))` **tronque le fichier avant** de
le lire — le contenu part, et une chaîne vide est écrite. Toujours lire d'abord
dans une variable, écrire ensuite.

---

## 2. Cible d'usage

L'écran de référence est une **tablette Android en paysage**, tenue en main pendant
le service. Hauteur réellement utile ≈ **620 px** (la barre d'adresse en mange ~150).

Conséquences :

- Tout se juge à **1024 × 620**, pas à 1024 × 768.
- Priorité à la **lecture d'un coup d'œil** et aux **cibles tactiles**.
- Le commerçant n'a pas le temps de chercher : l'information utile doit être en haut,
  sans défilement, et les actions doivent être atteignables sans viser.

---

## 3. Paliers responsive

```js
useIsMobile()      // < 900 px  → mise en page mobile (une colonne, nav basse)
useEcranEtroit()   // < 1180 px → mise en page bureau compactée (tablette paysage)
// ≥ 1180 px       → bureau, inchangé
```

La variable locale s'appelle `etroit` par convention. **Toute réduction de taille
destinée à la tablette doit être conditionnée par `etroit`**, jamais appliquée
globalement : le bureau ne doit pas être dégradé.

---

## 4. Règle « plein écran », et son exception

**Principe :** une page occupe la hauteur de l'écran, l'en-tête et les filtres restent
visibles, et seule la zone de contenu défile en interne. Objectif : ne pas obliger à
faire défiler pour accéder aux informations.

Appliqué à **Réservations**, **Clients**, **Communications**, **Menu**.

**Exception assumée — la page Click and Collect défile entièrement**, sur demande :
la liste des commandes en préparation se lit mieux d'une traite. Ne pas « corriger »
ce point en croyant à un oubli.

Le bandeau *Nouvelles commandes à traiter* est en `position: sticky; top: 10` : il se
fige en haut quand on descend, et reprend sa place quand on remonte.

---

## 5. Statuts et couleurs

| Statut | Fond | Texte |
|---|---|---|
| Nouvelle | `#dc2626` rouge | blanc |
| En préparation | `#E8C547` jaune de marque | `#111` |
| Prête | `#16a34a` vert | blanc |
| Récupérée | `#111` noir | blanc |
| Annulée | `#f5f5f5` gris | `#999` |

**Bouton d'action des cartes** — il porte le sens de l'étape suivante :

- statut *en préparation* → bouton **`#E8C547`** (jaune de la marque, celui de
  « Nouvelle commande ») libellé **« Marquer prête »**
- statut *prête* → bouton **`#16a34a`** vert libellé **« À récupérer »**, qui termine
  la commande

Jaune de marque : `#E8C547`. Doré des données : `#b8860b`. Bleu des données : `#2563eb`.
Le jaune de marque **échoue** les tests de contraste sur du texte : ne jamais l'utiliser
comme couleur de texte ni comme série de graphique.

---

## 6. Click and Collect — parcours

### Filtres

Trois filtres, **compteur toujours affiché même à zéro** (le libellé ne doit pas
changer de forme quand on navigue) :

`En préparation (N)` · `À récupérer (N)` · `Terminées (N)`

Le repère n'est pas la date du calendrier mais la **journée de service**, qui
**bascule à 6 h du matin et non à minuit** : le service du soir déborde après minuit,
si bien qu'à 2 h du matin on travaille encore le service de la veille. Une commande
en préparation à 23 h 59 est toujours là à 2 h — c'est le même service.

```js
HEURE_FIN_DE_NUIT = 6
service en cours   = 6h–14h59 → midi   |   15h–5h59 → soir
jour de service    = date du jour, sauf avant 6h → la veille
```

- « En préparation » et « À récupérer » montrent le **service en cours**, et rien
  du service qui n'a pas commencé (le soir reste caché pendant le midi).
- Elles remontent en revanche tout ce qui **traîne d'avant** — midi non clôturé
  pendant le soir, jours précédents jamais terminés — sinon une commande oubliée
  devient invisible.
- « Terminées » couvre la **journée de service entière**, midi et soir réunis,
  comme les deux tuiles de tête.
- Une commande dont l'heure de retrait est avant 6 h (ex. 00 h 30) relève du
  **service du soir de la veille**.

### Tuiles de tête

Une seule ligne, quatre colonnes : `Commandes du jour` · `Total du jour` ·
`Statistiques` · `Calendrier`. Elles suivent le jour et le service affichés.
Hauteur 42 px en tablette : c'est un rappel, pas le sujet de la page.

### Calendrier des commandes

Parcours en trois temps, à respecter :

1. je clique un jour → il se sélectionne, **la modale reste ouverte**
2. je choisis Midi ou Soir dans la colonne de droite
3. je valide avec le bouton du bas, qui rappelle le choix :
   « Valider — vendredi 21 août · Midi »

La croix ferme sans rien appliquer.

- **Glissement horizontal** pour changer de mois, repris à l'identique du calendrier
  des réservations : deux grilles superposées, seuil à 28 % de la largeur.
  Les cases réagissent à `onClick` (pas `onPointerDown`) et le glissement appelle
  `preventDefault()` — sans quoi un swipe démarré sur un chiffre le sélectionne.
- Colonne de droite : **Service, Midi, Soir**, rien d'autre. Pas de date en toutes
  lettres au-dessus, elle repoussait le bouton Soir hors de l'écran.
- Chiffre **centré** dans la case ; compteurs midi/soir en position absolue en bas,
  8,5 px, affichés seulement s'il y a des commandes.

### Modale « + Nouvelle commande »

Six sections, dans cet ordre :

1. **Téléphone du client** — pilote tout le reste
2. **Date de retrait** — calendrier identique à « Nouvelle réservation », mêmes
   animations (`date-flash` 200 ms puis `cal-fermeture` 300 ms)
3. **Service** — Midi / Soir
4. **Heure** — 5 créneaux proposés (les prochains du service), bouton
   « Plus d'horaires » qui **complète la même grille** (surtout pas un second bloc)
5. **Articles**
6. **Notes**

Règles :

- Il n'y a **pas de champ « Nom du client »**. Le nom vient du client : repris de la
  fiche si le numéro est reconnu, saisi dans le bloc de création sinon.
  **Le téléphone est donc obligatoire.**
- Numéro inconnu → bloc genre / prénom / nom / email comme en réservation, et la fiche
  client est créée avec la commande (`client_id` renseigné). **L'email est optionnel**,
  contrairement aux réservations : bloquer une commande téléphone sur un email serait
  intenable en service.
- En **modification**, la fiche client n'est pas redemandée.
- Fermer avec des modifications non enregistrées demande confirmation.
- Le contenu défile jusqu'aux notes (`minHeight: 0` sur le conteneur flex — sans lui
  rien ne défile).

### Modale d'ajout d'articles

- **Pas de titre** : la croix est dans la barre de recherche, en haut à droite.
- **Recherche et filtres de catégorie figés en haut**, seule la liste défile.
- Le **nom du produit est cliquable** pour l'ajouter, en plus du bouton `+` :
  zone de frappe large.
- Bouton du bas = **Enregistrer**. La croix demande confirmation et **annule
  réellement** les articles ajoutés pendant cette ouverture (retour au panier d'avant).

### Cartes de commande

- **Ligne 1** : nom du client et badges à gauche, **« Prête dans X min » à droite**
- **Ligne 2** : numéro et téléphone à gauche, **jour et heure de retrait à droite**
- Le nom affiché est **« Prénom Nom »** via une jointure sur la fiche client
  (`select('*, client:clients(prenom,nom,entreprise)')`). Les commandes en ligne
  n'ayant pas de `client_id`, la fiche est retrouvée **par le téléphone**
  (`cleTel` : chiffres seuls, 9 derniers, pour ignorer le 0 ou le +33).
  Repli sur le nom saisi si aucune fiche ne correspond.
- Une commande récupérée affiche « Récupérée aujourd'hui / hier / le … à HH:MM »
  (horodatage `traited_at`, posé automatiquement au passage du statut).

### Tri des commandes

1. les commandes en cours avant les terminées
2. les terminées entre elles par **heure de récupération décroissante**
3. les commandes en cours groupées par jour de retrait, puis par ordre de réception

### Suppression

Bouton discret en bas du détail, confirmation en deux temps annonçant que l'action
est **irréversible** et que la commande disparaît des listes, du calendrier, des
statistiques et des exports.

### Statistiques des commandes

Trois axes seulement : **la carte, le chiffre d'affaires, les clients**. Tout
indicateur qui ne sert aucune de ces trois décisions n'a pas sa place ici.

- Périodes : `Aujourd'hui · 7 jours · Ce mois · Cette année · Dates au choix`.
- **Chaque chiffre est comparé à la période précédente** de même durée, placée
  juste avant. Un chiffre sans comparaison ne se regarde qu'une fois ; sans
  passé comparable on affiche « pas de comparable », jamais un faux 0 %.
- Quatre tuiles : CA, commandes, panier moyen (avec leur évolution), clients
  (nouveaux / déjà venus).
- Courbe : CA par jour tant que la fenêtre tient en 31 jours, sinon par mois.
  Hauteur des barres réduite à 116 px en `etroit` — sinon tout le reste passe
  sous le pli sur les 620 px de la tablette.
- « Ce qui se vend » : top 8, **basculable quantité / chiffre** — le plat le plus
  vendu n'est pas celui qui rapporte — avec la part du CA de chacun.
- « Personne n'en a pris » : les produits de la carte jamais commandés sur la
  période. En dessous de 28 jours, un avertissement rappelle que le chiffre ne
  veut rien dire à cette échelle.
- « Vos clients » : nouveaux, déjà venus, part de ceux qui reviennent, top 5 par CA.
- Les deux listes de produits gardent un aperçu court (8 lignes / 14 étiquettes) et
  un bouton **« Voir les N »** qui ouvre une modale dédiée : liste entière,
  **recherche** en tête et bascule quantité / chiffre conservée.
  La carte est **dédoublonnée par nom** — elle contient des homonymes qui, sans
  cela, faussaient le compte et cassaient le filtrage (clés React en double).

---

### Fiche client — trois cartes d'une même famille

L'en-tête **Coordonnées** suit la même grammaire que les volets : avatar 38 px +
titre 15/800, actions à droite (Appeler, SMS, et **Modifier en jaune de marque**),
puis trois tuiles grises — Téléphone, E-mail, Client depuis. **Le commentaire du
client vit dans cette carte**, en tuile grise s'il existe, en bouton pointillé
sinon.

La modale **Classement clients** (tuile « Top client ») reprend la même
grammaire : fond gris, une carte blanche, en-tête à badge, trois tuiles
(meilleur client, clients classés, réservations) puis le classement à barres
`#E8C547`. Une ligne ouvre la fiche du client.

### Deux volets bâtis pareil

`BlocReservations` et `BlocClickCollect` partagent exactement la même grammaire,
et toute évolution de l'un doit se répercuter sur l'autre :

- carte blanche `borderRadius:16`, `padding:'20px 24px'` ;
- en-tête : icône 18 + titre 15/800, et à droite le badge de service préféré ;
- **quatre tuiles** grises `#f9f9f9` (libellé 9,5 px majuscules, valeur 16/900,
  précision 11 px) ;
- des **colonnes qui se replient seules** (`repeat(auto-fit, minmax(210px,1fr))`) :
  un historique daté, puis les classements à barres `#E8C547` — **top 3** partout.
  Réservations en compte deux, Click and Collect trois (articles, puis jours) ;
- version `compact` pour la fiche mobile : tuiles sur deux colonnes, listes empilées.

Les tuiles à icône colorée et la grille « Historique / Jours favoris » d'origine
ont été remplacées par ce volet.

### Volet Click and Collect de la fiche client

Présent dans la fiche bureau (entre le commentaire et l'historique des réservations)
et dans la fiche mobile, en version `compact` (tuiles sur deux colonnes, listes
empilées). Les commandes sont rattachées au client **par `client_id` ou par
téléphone** (`cleTel`), les commandes en ligne n'ayant pas de `client_id`.

- **Commandes / Chiffre d'affaires / Panier moyen** : calculés sur les seules
  commandes **récupérées** — l'argent réellement encaissé. Les commandes encore
  en cours sont signalées à part (« + N en cours »).
- **Dernière commande**, **Articles préférés** (top 5, barres relatives au plus
  commandé) et **service préféré** : sur tout l'historique, annulées exclues.
- **Dernières commandes** : les 5 dernières, avec date, service, heure, montant
  et statut.

---

## 7. Dimensions de référence en tablette (`etroit`)

| Élément | Bureau | Tablette |
|---|---|---|
| Bouton d'action de carte | 150 × 150 | **96 × 96** |
| Colonne d'action | 200 px | 118 px |
| Nom du client sur la carte | 17 px | 13,5 px |
| Lignes d'articles | 15,5 px | 12,5 px |
| Total de la carte | 19 px | 15,5 px |
| Filtres | 36 px / 13 px | 31 px / 11,5 px |
| Tuiles de tête | 52 px | 42 px |
| Cases du calendrier Réservations | 48 px | 33 px |

Boutons `−` / `+` du catalogue : 40 px — **sous les 44 px recommandés**, compensé par
le nom du produit cliquable. À remonter si la saisie tactile pose problème.

---

## 8. Barre latérale

- Onglet **« Click and Collect »**, avec **pastille rouge clignotante** dès qu'il y a
  des commandes à traiter, identique à celle des réservations.
  Le comptage est tenu par le shell avec son propre abonnement temps réel, pour rester
  juste quel que soit l'onglet ouvert.
- Onglets **Jeux** et **Système** masqués (commentés, pas supprimés).
- Pastille d'état système masquée.

---

## 8 bis. Paramètres

Onglet de la barre latérale, **au-dessus de Déconnexion**. Réglages **durables**,
par opposition au « Statut du jour » du Click and Collect qui ne vaut qu'une journée.

**Sommaire groupé** dans une colonne de gauche : trois titres de groupe, chacun
suivi de ses entrées cliquables. Tout est visible d'un coup d'œil, sans repli
ni navigation à plusieurs niveaux.

```
PARAMÈTRES DU COMPTE      ÉTABLISSEMENT       APPLICATION
  Compte et accès           Identité            Toutes les applications
  Sauvegardes               Dates / Horaires    …puis une entrée par
                                                  application ACTIVE
```

### Applications

`APPLICATIONS` décrit chaque application : résumé, texte détaillé, points forts,
trois vignettes d'aperçu, teinte. « Toutes les applications » les range en
**Mes applications actives** puis **À découvrir** ; un clic ouvre une modale large
avec la description complète et un bouton **Activer l'application**.

**Deux états distincts, à ne pas confondre :**

- `applications_installees` — l'application fait partie de la maison. Elle reste
  au sommaire même éteinte. Une application non installée s'ajoute par
  **« Installer l'application »**.
- `applications_actives` — allumée ou éteinte. Une fois installée, on ne fait
  plus que basculer cet état ; l'application ne quitte jamais la liste.

`APPS_TOUJOURS_ACTIVES` (Réservations, Fichier clients) ne peuvent pas être coupées.

`APPS_MASQUEES` retire des applications du sommaire **et** du catalogue des
Paramètres sans toucher à leur code. Actuellement **vide**.

Ne pas confondre avec les **onglets de la barre latérale**, masqués séparément
en commentant leur ligne dans le tableau de navigation — actuellement
**Communications, Menu et Jeux**. Leurs pages et leurs réglages restent
entièrement accessibles ; seul l'onglet disparaît.

La page Paramètres n'a **pas de sous-titre** sous son titre, et le **sommaire de
gauche est fixe** (`position: sticky`) : seule la colonne de droite défile.

### Messages automatiques

`MESSAGES_AUTO` décrit chaque moment du parcours client : confirmation et rappel
de réservation, refus, commande acceptée / prête / récupérée / refusée. Pour
chacun : les **canaux** (SMS, e-mail, ou les deux), un **délai** pour le rappel,
et le **texte**, avec des variables `{prenom}`, `{date}`, `{numero}`… insérables
d'un appui. Rangé dans `messages_auto`, une entrée par message.

Un message sans canal ne peut pas être activé. Seuls les messages des
applications installées sont proposés.

⚠️ Les messages sont **réglés mais pas encore envoyés** : il reste à les brancher
sur les changements de statut et à câbler un service d'envoi SMS.

### Créneaux — réservation et retrait

Les créneaux **suivent toujours les horaires d'ouverture**. Il n'y a pas de choix
de mode : changer un horaire dans « Dates / Horaires » suffit, les créneaux
s'adaptent. On ne règle que deux choses par service :

- **le rythme** — un créneau toutes les 15 / 30 / 45 / 60 minutes ;
- **jusqu'où aller** :
  - réservation → **la durée du service** (30 min à 2 h, réglée séparément midi
    et soir). Le dernier créneau est celui qui laisse le temps de finir :
    `dernier = fermeture − durée`. Un service de 45 min sur une ouverture
    12:00–14:30 s'arrête à **13:45** ; réserver 30 min avant la fermeture est
    impossible, le repas déborderait.
  - retrait → **la marge avant fermeture** (0 / 15 / 30 / 45 min) : il n'y a pas
    de table à libérer, juste une commande à remettre.

`creneauxService()` fait le calcul et tient compte des dates particulières.
Un encadré explique le résultat en toutes lettres, avec un exemple daté.

### Couleur des boutons de sélection

Dans Paramètres, un élément sélectionné est **fond jaune `#E8C547`, texte noir**
— jamais l'inverse. Vaut pour le sommaire, les rythmes de créneaux et les délais.

Chaque application affiche en tête un **interrupteur** activée / inactive ;
il est grisé pour les applications essentielles.

Le titre de groupe est une étiquette grise en capitales avec son icône — il ne
se clique pas. L'entrée active prend le fond noir et le jaune de marque, comme
la barre latérale. Le panneau de droite rappelle le nom de l'entrée et sa ligne
d'explication : les blocs qui répéteraient ce titre n'en portent pas.

La **semaine type** liste les sept jours. Un jour dont tous les services sont
fermés est un jour de fermeture — il n'y a pas de frise séparée.

Chaque journée porte une **liste de services** (jusqu'à `MAX_SERVICES` = 4) :
`[{id, nom, ouvert, debut, fin}]`. Les deux premiers, `midi` et `soir`
(`SERVICES_SOCLE`), ne sont ni renommables ni supprimables — **tout le reste de
l'application s'appuie sur ces deux identifiants**. Les suivants (`service3`,
`service4`) portent un nom libre et se retirent.

`servicesDuJour()` accepte l'ancienne forme `{midi:{…}, soir:{…}}` comme la
nouvelle : ne pas casser cette compatibilité.

Chaque jour a un bouton **« Appliquer à toute la semaine »** qui recopie ses
services sur les sept jours.

Le champ *nom du service* garde sa saisie en **state local** et n'enregistre
qu'au `blur` : écrire à chaque frappe déclenchait une écriture par caractère,
qui entrait en course avec la recopie et perdait le nom.

⚠️ Les services ajoutés sont **configurés mais pas encore exploités** :
réservations, commandes, calendriers et statistiques raisonnent toujours en
midi / soir.

Le **logo** ouvre le bloc Identité, avant le nom de l'établissement.

**Compte et accès** contient le changement de mot de passe
(`supabase.auth.updateUser`, 8 caractères minimum, saisie confirmée, œil
d'affichage) et la **gestion de l'équipe**.

### Comptes et rôles

Table `collaborateurs` adossée à `auth.users` (`on delete cascade`) : `email`,
`nom`, `role`, `actif`. Quatre rôles — `service`, `cuisine`, `manager`,
`proprietaire` — contrôlés par une contrainte.

```sql
public.mon_role()   -- SECURITY DEFINER, réservée à `authenticated`
                    -- évite que la policy de lecture ne se rappelle elle-même
```

Politiques : chacun lit sa fiche, le propriétaire lit toute l'équipe ; seul le
propriétaire crée, modifie ou retire.

La création d'un accès passe par la fonction **`gerer-collaborateurs`**
(`verify_jwt: true`) : elle vérifie que l'appelant est propriétaire **actif**
avant d'utiliser la clé d'administration. Actions : `creer`, `supprimer`,
`motDePasse`. Si l'insertion de la fiche échoue, le compte auth tout juste créé
est supprimé — pas d'orphelin. Personne ne peut supprimer son propre compte.

**La clé `service_role` ne doit jamais apparaître dans le frontend.**

⚠️ Les rôles sont **définis et attribués, mais pas encore appliqués** : l'interface
ne restreint rien selon le rôle, et les politiques des autres tables ne le
consultent pas. C'est le chantier suivant.

Les horaires se choisissent **dans une liste déroulante au quart d'heure**
(`ChoixHeure`, 96 options de 00:00 à 23:45), jamais au clavier : le CRM se tient
sur une tablette en plein service. Ne pas revenir à `<input type="time">`.

**Le lien public et le QR code appartiennent au module** qui les concerne : celui
de la réservation dans Réservations, celui de la commande dans Click and Collect.
Les adresses sont en consultation seule.

Tout passe par un objet unique :

```js
REGLAGES_DEFAUT   // les valeurs qui étaient en dur, servent de repli
REGLAGES          // objet vivant, lu au rendu par toute l'application
chargerReglages() // hydraté depuis commandes_config AVANT le premier rendu du CRM
appliquerReglage()// effet immédiat quand on enregistre depuis la page
```

**Le chargement bloque le premier rendu** : remonter le CRM après coup (via une
`key`) coupait les abonnements temps réel — « Maximum call stack size exceeded ».

Les valeurs par défaut ne doivent **jamais** se référencer elles-mêmes : une
substitution globale trop large a déjà produit `lien_reservation: REGLAGES.lien_reservation`,
ce qui plantait le module au chargement (page blanche, aucune erreur console).

Réglages **effectivement câblés** : créneaux de réservation et de retrait, bascule
midi/soir, fin de nuit, liens publics, nom/adresse/téléphone de la fiche imprimée,
acceptation automatique, délai, horizon.
Réglages **enregistrés mais pas encore appliqués** : jours de fermeture, capacités
par service, e-mail de contact (les pages publiques `reserver.html` et
`commander.html` ne lisent pas encore la configuration).

---

## 8 quater. Plan de salle

Éditeur de plan de salle dans la page Réservations. Le plan n'est pas un dessin :
c'est l'outil du service. Bascule **Plan de salle ⇄ Vue calendrier** dans la
ligne du titre ; le plan prend la colonne de gauche, le bloc « Réservations du »
ne bouge jamais (377 px, calé à droite).

**Aucune table en base.** Le plan vit dans le réglage `plan_salle`
(`commandes_config`), les placements dans `reservations.table_plan`. Les deux
existaient déjà : la fonctionnalité complète n'a demandé aucune migration.

### Le modèle

Coordonnées en **centimètres**, jamais en pixels ni en pourcentage — le plan est
indépendant de la taille de l'écran et une table ronde est vraiment ronde.

```
{ v:2, zones:[ { id, nom, type:'salle|etage|terrasse|bar', largeur, hauteur,
                 tables:[ {id, nom, forme, places, x, y, l, h, rot, statut, notes} ],
                 decor: [ {id, nom, type, x, y, l, h, rot} ] } ] }
```

`normaliserPlan()` accepte l'ancien format en pourcentage et le convertit à la
volée — rien à migrer. `PLAN_SALLE_DEFAUT` est le plan livré (3 espaces,
18 tables, 7 décors), utilisé tant que rien n'a été enregistré.

Les **décors** (bar, entrée, escalier, cuisine, WC, bloc) sont séparés des
tables : pas de couverts, jamais comptés, jamais assignables.

### Le canvas (`PlanCanvas`)

Un SVG dont on pilote le cadrage par `translate/scale` — un seul transform, net
à tous les zooms et léger au doigt.

- zoom molette (écouteur natif **non passif**, sinon `preventDefault` est ignoré),
  pincement à deux doigts, déplacement sur le vide, bouton « ajuster à l'écran » ;
- `etendueZone()` calcule les bornes **depuis le contenu** : approcher une table
  du bord agrandit le plan tout seul, il n'y a jamais de mur ;
- épaisseurs et poignées dimensionnées en `1/k` : taille constante à l'écran
  quel que soit le zoom ; cibles ≥ 44 px.

Trois pièges déjà payés, à ne pas défaire :

- **le cadrage ne se recalcule qu'au changement d'espace** (`zoneRef`), sinon le
  plan saute à chaque table déplacée ;
- **l'objet manipulé vit dans une ref** (`provRef`) autant que dans l'état : le
  relâchement peut suivre le déplacement dans la même frame et lirait sinon un
  état périmé — le déplacement serait perdu ;
- **la table visée est retenue dès l'appui** (`geste.current.table`) : la capture
  du pointeur redirige le `pointerup` vers le SVG, on ne saurait plus sur quoi le
  doigt s'est posé et un simple tap n'ouvrirait jamais la fiche de table.

### Mode édition

Par défaut on ne voit que le plan et la palette. Les options d'une table
n'apparaissent **qu'à la sélection**, dans un panneau latéral.

- palette à glisser-déposer (ronde, ovale, carrée, rectangle, banquette + décors),
  posée par des écouteurs `window` et `apiRef.versPlan()` ;
- déplacer, tourner (pas de 15°), redimensionner aux 4 coins, dupliquer, supprimer ;
- grille de 5 cm + **guides magnétiques** sur les bords et centres des voisins
  (9 cm d'accrochage) ; chevauchement signalé en orange, jamais bloqué ;
- espaces : ajouter, renommer, changer de type, réordonner, supprimer ;
- undo/redo (pile de 80), autosave débouncé 600 ms avec témoin « Enregistré » ;
- raccourcis : `Suppr`, `⌘D`, flèches (5 cm, 25 avec `Maj`), `⌘Z` / `⌘⇧Z`.

### Mode service

Date + service repris de la page (pas de second sélecteur à maintenir).

- code couleur **jamais seul** : la table porte le prénom, les couverts et l'heure ;
- glisser une réservation de la liste sur une table (le glisser-déposer de la page
  est inchangé, les tables portent `data-table`) ; ou clic table → choisir la résa ;
- **combinaison** : `table_plan` accepte plusieurs identifiants séparés par une
  virgule (`tablesDeResa`). La capacité d'un groupe est la **somme** des tables,
  sinon 9 personnes sur 6+8 s'afficheraient en dépassement. Contour violet ;
  « Séparer » revient à une seule table ;
- compteurs : chaque espace annonce `Salle (13/32)`, le pied donne les couverts
  placés et les tables libres.

### En-tête

Une seule ligne dès qu'il y a la place : date en bouton (calendrier en surcouche),
Midi/Soir, espaces chiffrés, puis le crayon d'édition. Sur écran étroit la date
perd son jour de semaine et les services passent en icônes seules ; en dessous de
~1180 px la ligne se replie plutôt que de cacher un espace derrière un
défilement — c'est l'occupation qu'on vient y lire.

⚠️ La piste de grille `1fr` a pour largeur minimale son min-content : elle refuse
de rétrécir et fait déborder la page entière. C'est `minmax(0,1fr)` qu'il faut.

Le plan n'existe qu'au-dessus de 900 px (`useIsMobile`) : en dessous, la page
Réservations garde sa vue mobile en liste, sans plan.


## 8 ter. Approbations

Onglet de la barre latérale, avec **pastille rouge** du nombre de demandes en
attente, tenue par le shell comme celles des réservations et des commandes.

Table `approbations` : `type` (campagne / site / menu / tarifs / reservation /
autre), `titre`, `description`, `contenu` (jsonb — ce que l'assistant propose),
`statut` (en_attente / approuvee / refusee), `motif_refus`, `revision_de`
(pointe vers la demande corrigée), `decide_le`, `decide_par`.

Politiques : **toute l'équipe lit**, seuls `proprietaire` et `manager`
décident. Un rôle sans droit voit un bandeau qui le dit, et aucun bouton.

Point d'entrée de l'assistant : fonction **`approbations`**, `verify_jwt: false`
mais protégée par la clé secrète `CLE_ASSISTANT` (en-tête `x-cle-assistant`).
Actions : `creer`, `lire`, `lister`. La clé d'administration ne sort jamais du
serveur. **Sans la variable d'environnement, la fonction refuse tout** — c'est
volontaire.

**Lisibilité — la règle du domaine d'abord.** Les demandes sont **regroupées par
domaine** (Communication, Site internet, La carte, Tarifs, Réservations, Fichier
clients, Divers), chacun avec son en-tête, son icône et **sa couleur**, reprise
en liseré de 5 px sur le bord gauche de chaque carte. On sait de quoi on parle
avant de lire une ligne. Ne jamais revenir à une liste à plat.

**Le détail s'affiche selon ce dont il s'agit**, jamais en vidage clé/valeur
générique :

- un **message** (campagne) se lit comme un message — bulle blanche à liseré,
  puis trois pastilles : canal, destinataires, date d'envoi ;
- une **modification** se juge en comparant — deux colonnes *Aujourd'hui*
  (barré, gris) et *Après la modification* (vert) ;
- le reste en lignes libellé / valeur espacées.

**On n'approuve pas un tableau, on approuve ce que le client verra.** Le bouton
principal est **« Voir un aperçu »** ; la décision se prend dans la modale, au
bout de l'aperçu. Chaque aperçu imite le support réel :

- **SMS** → un téléphone (châssis noir, barre d'état, en-tête de conversation,
  bulle reçue grise à coin cassé) ;
- **E-mail** → une boîte mail (fenêtre à trois pastilles, objet, expéditeur, corps) ;
- **Site** → un navigateur avec barre d'URL, la page esquissée et **le bloc
  modifié encadré en vert**, ancien texte barré au-dessus du nouveau ;
- **Carte / tarifs** → une ligne de carte, plat barré ou ancien prix barré ;
- sinon, le détail en lignes libellé / valeur.

**Un refus doit être motivé** : le bouton reste inactif tant que le motif est
vide, avec quatre motifs fréquents proposés. C'est ce motif que l'assistant
relit pour proposer une version corrigée ; la révision affiche alors un bandeau
rappelant la demande d'origine et la raison du refus.

**Exemples de démonstration.** Tant que l'assistant n'est pas branché, l'écran
vide propose **« Voir des exemples »** : six demandes couvrant les six domaines,
marquées `demandeur = 'exemple'`. Un bandeau le dit et offre de les retirer.
Deux politiques étroites autorisent **uniquement** l'insertion et la suppression
de lignes `demandeur = 'exemple'`, et seulement pour propriétaire et manager :
une vraie demande ne peut jamais être fabriquée depuis le navigateur.

La table est ajoutée à la publication `supabase_realtime` — sans quoi une
demande n'apparaît qu'au rechargement.

---

## 9. Chantiers ouverts

À traiter uniquement sur demande, mais à connaître :

1. **Aucune notification client.** Ni à l'acceptation, ni au refus, ni en cas de
   modification de commande. Le commerçant doit rappeler lui-même.
2. **Notifications push cassées** pour les réservations client : le trigger appelle
   `/send-push-onesignal`, qui renvoie 403 depuis le durcissement JWT.
3. **Table `clients` lisible publiquement** via la clé anon — faille identifiée,
   jamais corrigée faute de demande.
4. **Commandes annulées inatteignables** : aucun des trois filtres ne les montre,
   elles ne peuvent donc plus être consultées ni supprimées depuis l'interface.
5. **Commandes en ligne non rattachées** à une fiche client (`client_id` vide) :
   elles affichent le nom saisi par le client, pas « Prénom Nom ».
6. **Créneaux de retrait écrits en dur** : midi 11:30–14:15, soir 18:00–21:30.
   À rendre paramétrables si les horaires changent.
7. Bascule **midi / soir figée à 15 h** (et fin de nuit à 6 h), alignée sur les
   réservations. À rendre paramétrable si les horaires changent.
