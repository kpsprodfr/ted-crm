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
  (`select('*, client:clients(prenom,nom,entreprise)')`), avec repli sur le nom saisi.
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
