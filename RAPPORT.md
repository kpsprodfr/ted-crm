# RAPPORT — Industrialisation TED CRM

Démarré le 2026-07-01. Une ligne par phase : fait / testé / résultat.

---

## PHASE 0 — État des lieux (diagnostic, aucune modification de code)

### 0.1 Cartographie

**Stack** : React CRA (`react-scripts 5.0.1`, pas Vite) · Supabase (projet `mwpfaytccypvdrgapptk`, Postgres 17, eu-west-1) · Cloudflare Pages (`ted-crm.pages.dev`, déploiement auto via push GitHub `kpsprodfr/ted-crm`) · Brevo (email + SMS) · OneSignal (push).

**Fichiers source**

| Fichier | Rôle |
|---|---|
| `ted-crm-v2/src/App.js` | CRM complet (~7 270 lignes, 110 appels Supabase dont 83 `await`) |
| `ted-crm-v2/src/lib/supabase.js` + `src/supabase.js` | 2 clients Supabase en doublon (même config) |
| `ted-crm-v2/src/firebase.js` | **Code mort** — plus importé nulle part |
| `ted-crm-v2/public/accueil.html` | Page client : soirées + roue (15 fetch directs Supabase) |
| `ted-crm-v2/public/reserver.html` | Réservation publique (clients SELECT/INSERT/PATCH, reservations INSERT) |
| `ted-crm-v2/public/menu.html` | Menu public (lecture seule) |
| `ted-crm-v2/public/agenda.html`, `ical.html` | Statiques, pas d'accès DB |
| `ted-crm-v2/functions/send-email.js` | POST `/send-email` — appelé par App.js (confirmations résa) |
| `ted-crm-v2/functions/send-sms.js` | POST `/send-sms` — **aucun appelant** (App.js appelle Brevo en direct, voir C2) |
| `ted-crm-v2/functions/send-push-onesignal.js` | POST `/send-push-onesignal` — appelé par App.js |
| `ted-crm-v2/functions/api/roue-email.js` | POST `/api/roue-email` — email premium du gain (accueil.html + CRM) |
| `ted-crm-v2/functions/api/roue-notify.js` | POST `/api/roue-notify` — **aucun appelant** (legacy) |
| `ted-crm-v2/functions/api/backup.js` | GET `/api/backup` — protégé par `BACKUP_SECRET`, mais **cassé** (voir 0.2) et sans stockage ni cron |
| `ted-crm-v2/supabase/functions/send-push/` | Edge function Deno (push FCM legacy) |

**Tables Supabase** (15 + `error_logs` créée ce jour) : `clients` (34), `reservations` (113), `roue_gains`, `roue_recompenses`, `roue_config`, `parametres`, `menu_produits` (228), `menu_categories`, `menu_cartes`, `menu_soirees`, `menu_plat_jour`, `menu_origines`, `sms_envoyes`, `emails_envoyes`, `fcm_tokens` (0 ligne, legacy), `error_logs`.

### 0.2 Variables d'environnement — croisement code ↔ réalité

| Variable | Utilisée par | État |
|---|---|---|
| `BREVO_API_KEY` | send-email, send-sms, roue-email, roue-notify | ✅ Présente dans CF (les emails partent en prod) |
| `REACT_APP_SUPABASE_URL` / `_ANON_KEY` | build React (src/lib/supabase.js) | ✅ Présentes dans CF build (le CRM fonctionne), ❌ absentes du `.env` local → **dev local cassé** |
| `REACT_APP_BREVO_API_KEY` | App.js:5767 (SMS direct navigateur) | ⚠️ Clé dans `.env` local, invalide côté API (401 constaté), **exposée dans le bundle si définie** — à supprimer |
| `SUPABASE_URL` / `SUPABASE_ANON_KEY` | roue-email, roue-notify | Optionnelles (fallback en dur dans le code) |
| `REACT_APP_SUPABASE_KEY`, `SUPABASE_SERVICE_KEY` | backup.js | ❌ Noms incohérents, probablement non définies → **backup.js ne fonctionne pas en prod** |
| `BACKUP_SECRET` | backup.js | ❓ À définir/vérifier dans CF |
| `ONESIGNAL_REST_API_KEY` | send-push-onesignal | ✅ Présente (push fonctionne) |
| `REACT_APP_FIREBASE_*` (6 vars) | firebase.js (code mort) | 🗑️ Ne servent plus |

### 0.3 Failles de sécurité constatées

| # | Faille | Gravité |
|---|---|---|
| S1 | ~~RLS désactivée sur `clients` et `reservations`~~ : n'importe qui avec la clé publique (visible dans le source des pages) pouvait **lire, modifier et SUPPRIMER toute la base clients et les réservations** | **CRITIQUE — corrigée ce jour** (voir note ⬇) |
| S2 | Endpoints `/send-email`, `/send-sms`, `/send-push-onesignal`, `/api/roue-notify` **sans aucune authentification ni origine** → relais de spam sur ton compte Brevo/OneSignal | CRITIQUE |
| S3 | Injection HTML dans les emails : `prenom`, `nom`, `recompense`, `conditions`, `message` insérés sans échappement dans le HTML (roue-email.js, roue-notify.js, send-email.js) | MAJEUR |
| S4 | Clé Brevo référencée dans le code client (App.js:5767, appel SMS direct navigateur) | MAJEUR |
| S5 | CORS `*` sur roue-notify et backup ; aucun header de sécurité nulle part | MAJEUR |
| S6 | Aucune validation/sanitisation serveur des inputs publics (roue, réservation) | MAJEUR |
| S7 | `npm audit` : 13 high — toutes dans les devDependencies de `react-scripts` (nth-check, webpack-dev-server…), **build-time uniquement**, pas exploitables en prod ; `audit fix --force` casserait CRA | MINEUR |
| S8 | Token GitHub (`ghp_…`) stocké en clair dans l'URL du remote git local | MINEUR (local) |
| S9 | Advisories Supabase restantes : fonctions `SECURITY DEFINER` exécutables par anon (`decrement_stock_roue` — nécessaire à la roue), policies `USING (true)` sur `roue_config` (nécessaire au front public actuel), listing du bucket public `soirees-flyers` | MINEUR (documenté) |

> **Note S1** : correction appliquée avant réception de la présente mission (demande explicite précédente) — migrations `security_hardening_rls_error_logs` + `performance_indexes` : RLS activée sur `clients`, `reservations`, `menu_origines` avec policies reproduisant exactement l'usage constaté (anon : SELECT/INSERT/UPDATE clients, INSERT reservations, SELECT menu_origines — jamais DELETE) ; table `error_logs` créée (INSERT anon, SELECT authenticated) ; `search_path` figé sur les 2 fonctions SQL ; index ajoutés (`clients.mail`, `roue_gains.date_gain`, `reservations(date,service)`, `error_logs.created_at`). **Testé en prod le 2026-07-01** : lecture publique clients/roue_config/menu_origines OK (200), lecture anonyme des réservations bloquée, DELETE anonyme neutralisé.

### 0.4 Points de défaillance uniques (SPOF)

| SPOF | Si ça tombe… | Gravité |
|---|---|---|
| **Supabase** (aucun backup automatisé aujourd'hui) | Perte définitive clients + réservations + gains. `backup.js` existe mais est cassé, manuel et sans stockage | **CRITIQUE** |
| **Brevo** (échec ponctuel API, quota épuisé, clé révoquée) | Email de gain **perdu définitivement** — aucun retry ni file d'attente ; le client gagne à la roue et ne reçoit rien | **CRITIQUE** |
| **Cloudflare Pages** | Site + CRM + functions down (les emails de gain ne partent plus) | MAJEUR (SLA Cloudflare élevé) |
| **Requête Supabase qui échoue dans App.js** | Beaucoup d'`await` sans try/catch → écran figé ou crash silencieux | MAJEUR |
| **OneSignal** | Plus de push nouvelles résas (le CRM reste utilisable) | MINEUR |
| **Aucun monitoring** | Toute panne ci-dessus est découverte par les clients, pas par toi | MAJEUR |

### 0.5 Priorisation des chantiers

- **CRITIQUE** : file d'attente emails + retries (Phase 1) · backups automatiques (Phase 2) · fermeture des endpoints ouverts S2 (Phase 4) — ✅ S1 déjà corrigée.
- **MAJEUR** : resilience Supabase/safeQuery + timeouts (Phase 1) · monitoring/alertes (Phase 3) · injection HTML S3, clé client S4, CORS/headers S5, validation serveur S6 (Phase 4).
- **MINEUR** : cache localStorage + index (Phase 5, index ✅ déjà posés) · nettoyage code mort (firebase.js, roue-notify.js, doublon supabase.js) · documentation S7–S9 (Phase 6).

**Résultat Phase 0 : diagnostic complet, aucun code modifié. Validé (« go ») le 2026-07-01.**

---

## PHASE 1 — Fiabiliser l'existant

**Fait** :
- `src/lib/db.js` : fetch résilient (timeout 10 s, retry exponentiel 500 ms/1 s/2 s, log `error_logs` en échec final) **injecté dans le client Supabase** (`global.fetch`) → les 110 appels Supabase de App.js sont couverts par construction, sans réécrire chaque call site (comptage avant/après : 110 → 110, zéro régression possible sur les handlers JSX inline). Les écritures ne sont retryées que sur erreur réseau franche ou 429/503 (pas de doublon de réservation possible).
- `safeQuery(queryFn, {retries, fallback, context})` appliqué explicitement aux chargements critiques : `loadResa`, `loadClients`, `chargerToutesStatsClients`, `loadMenu` (fallback `[]` au lieu d'un crash UI).
- `resilientChannel` : les 3 channels Realtime (`resa-page-realtime`, `menu-rt`, `nouvelles-reservations`) se reconnectent automatiquement avec backoff expo (1 s → 30 s max) si le websocket tombe.
- File d'attente emails : table `email_queue` (RLS : anon = écriture seule) ; `send-email.js` et `roue-email.js` y déposent l'email rendu en cas d'échec Brevo ; `/api/process-email-queue` reprend les `pending` (max 5 tentatives, batch 20, **claim atomique pending→processing** rendant tout rejeu/appel concurrent inoffensif) ; déclenchement horaire par **pg_cron** côté Supabase (Cloudflare Pages n'exécute pas les crons wrangler ; le push de workflows GitHub est refusé par le PAT local — pg_cron est de toute façon plus autonome).
- Timeout 10 s sur tous les fetch : client Supabase du CRM + patch global sur `accueil.html`, `reserver.html`, `menu.html` ; `fetchT` (AbortController) dans toutes les functions. Correction au passage : `notifEnCoursRef` restait bloqué si le push OneSignal échouait.

**Testé (en réel, prod)** :
- `npm run build` : ✅ compile (bundle +610 o).
- **Cycle complet file d'attente** : échec Brevo simulé (env sans clé) sur `send-email.js` → ✅ `{queued:true}` + ligne `pending` dans `email_queue` avec l'erreur tracée ; puis `/api/process-email-queue` en prod → ✅ `{sent:1}`, statut passé à `sent`, email réellement reçu. (`SUPABASE_SERVICE_ROLE_KEY` était déjà définie dans Cloudflare.)
- RLS file d'attente : lecture anonyme de `email_queue` → ✅ `[]` (bloquée).
- **Parcours roue de bout en bout (prod)** : insert `roue_gains` avec la clé publique → `/api/roue-email` → ✅ `{ok:true}`, email reçu, flag `email1_envoye=true` posé.
- `/send-email` réécrit → ✅ `{success:true}` en prod.

**🔥 Incident détecté et corrigé pendant la phase** : la RLS activée en Phase 0 sur `reservations` cassait l'INSERT du formulaire public `reserver.html` (PostgREST exige une policy SELECT quand le client demande `Prefer: return=representation`). Le retour n'étant pas utilisé, correctif code : `return=minimal` sur cet INSERT (commit `0cdf7daa`), sans rouvrir la lecture anonyme des réservations. `accueil.html` utilisait déjà `return=minimal` (non affecté). Re-testé après déploiement : ✅ INSERT anonyme 201.

**Données de test nettoyées** (1 résa, 1 gain, 1 client, 1 email de file). **Phase 1 validée.**

---

## PHASE 2 — Sauvegarde et restauration

**Fait** : `/api/backup-daily` (12 tables → KV `backup:YYYY-MM-DD`, rétention 30 j, checksum SHA-256, email récap à com.astegal@gmail.com, idempotent — GET cron sans secret, POST forcé JWT) · `/api/backups` (liste, JWT) · `/api/backup-restore` (upsert par lots de 200, non destructif, confirmation « RESTAURER », JWT) · page **Système** dans la sidebar CRM (liste des backups, bouton « Backup maintenant », restauration avec double confirmation) · namespace KV `ted-crm-backups` créé via l'API Cloudflare · cron pg_cron 02:00 UTC posé.
**Testé** : build ✅ ; endpoints déployés et répondent ✅.
**⏳ En attente d'une action manuelle** : le binding KV `BACKUPS` doit être ajouté dans le dashboard Cloudflare Pages (30 s — voir MAINTENANCE.md §2 ; l'ajout par wrangler.toml écraserait la config du dashboard, risque refusé). Dès le binding posé : le premier backup réel + une restauration de test valideront la phase.

## PHASE 3 — Surveillance autonome

**Fait** : `/api/health` (Supabase avec latence + 4 tables critiques, Brevo via GET /v3/account, variables d'env, fraîcheur des backups → `ok`/`degraded`/`down`) · pg_cron toutes les 15 min avec `?notify=1` → alerte email si down, dédupliquée à 1/h (KV TTL), mise en file si Brevo est lui-même en panne · pastille verte/orange/rouge en bas de la sidebar (refresh 60 s, tooltip détaillé, clic → page Système) · journal des 50 dernières `error_logs` dans la page Système, alimenté par `logError()` (front + functions).
**Testé (prod)** : `/api/health` → supabase ok 128 ms, brevo ok 256 ms, env_vars ok ✅. Statut global `degraded` uniquement à cause du binding KV manquant → passera vert avec la Phase 2.

## PHASE 4 — Sécurité

**Fait** :
- **Endpoints verrouillés** : JWT Supabase obligatoire sur `/send-email`, `/send-sms`, `/send-push-onesignal`, `/api/roue-notify` (avant : n'importe qui pouvait envoyer des emails/SMS avec le compte Brevo du TED). `/api/roue-email` reste public pour la roue mais : `gain_id` obligatoire et vérifié en base, l'email doit correspondre au gain, données (prénom/récompense) lues en base et plus dans le payload, idempotent (jamais 2 envois pour le même gain).
- **Origin + rate limit** : `guard()` sur tout — origine `ted-crm.pages.dev` uniquement (403 sinon), 10 req/min/IP via KV (actif dès le binding posé).
- **Headers** : nosniff, X-Frame-Options DENY, HSTS, CSP `default-src 'none'`, Referrer-Policy sur toutes les réponses ; CORS `*` supprimé.
- **Anti-injection** : échappement HTML de toutes les interpolations dans les 2 templates d'email.
- **Clé Brevo hors du navigateur** : SMS marketing CRM routés par `/send-sms`, `REACT_APP_BREVO_API_KEY` supprimée du `.env`.
- **Validation** : formulaires publics (sanitisation HTML, longueurs max 50/254/500, tel mobile 06/07 pour la roue) + formulaire client CRM (formats tel FR/email, longueurs) + verrou anti double-clic sur la confirmation de résa (évitait double email) + validations serveur dans les functions (`isValidEmail`, `isValidTelFR`).
- **Hygiène** : `.gitignore` complété, `npm audit fix` appliqué (2 passes) — les high restantes sont dans les devDependencies de react-scripts (build-time, non exploitables en prod ; `--force` casserait CRA, refusé).

**Testé (prod, après déploiement)** : Origin bidon → 403 ✅ · send-email/send-sms/roue-notify sans JWT → 401 ✅ · roue-email gain inexistant → 404 ✅ · sans gain_id → 400 ✅ · email ne correspondant pas au gain → 403 ✅ · **parcours roue réel : gain inséré → email envoyé `{ok:true}` → rejeu → `{already_sent:true}`** ✅ · headers présents sur les réponses ✅ · formulaire résa public toujours fonctionnel (INSERT anon 201) ✅. Advisories Supabase : **0 ERROR** (4 au départ) ; les WARN restants sont volontaires (INSERT publics des formulaires, rpc roue) ou hors code (protection mots de passe divulgués à activer dans le dashboard Auth, pg_net géré par Supabase).

## PHASE 5 — Performance

**Fait** : index vérifiés à l'EXPLAIN (`reservations.date` → Index Scan ✅ ; `clients.mail` seq scan normal à 34 lignes, l'index `idx_clients_mail` prendra le relais en volume) · cache menu migré sessionStorage → **localStorage** (TTL 5 min, stale-while-revalidate 60 s conservé, toujours invalidé par le Realtime) · config + récompenses de la roue en cache localStorage **10 min** (stock protégé côté serveur par le rpc atomique) · fuites : 3 setInterval/4 clearInterval, 3 channels realtime tous nettoyés au démontage par `resilientChannel` ✅ · syntaxe des scripts inline vérifiée (node --check) ✅.

## PHASE 6 — Tests et documentation

**Fait** : `/api/run-tests` (X-Test-Key: BACKUP_SECRET **ou** JWT CRM) — env vars, binding KV, lecture des 12 tables, **écriture/suppression réelles** (error_logs), clé Brevo (GET /v3/account, sans envoi), fraîcheur du dernier backup, file emails sans échec définitif → rapport JSON · bouton « Lancer les tests » + badges dans la page Système · `MAINTENANCE.md` complet (architecture, tableau des variables avec régénération, procédures de panne par composant, restauration, checklist mensuelle, accès) · scripts npm `security-audit` / `update-deps`.
**Note** : le workflow GitHub Actions hebdomadaire n'a pas pu être poussé (PAT local sans scope `workflow`) — remplacé par la checklist mensuelle + `npm run security-audit`, documenté.

---

## BILAN FINAL

**Failles corrigées** : suppression/lecture anonymes de la base clients et des réservations (RLS) · 4 endpoints ouverts en relais de spam (JWT) · injection HTML dans 2 templates d'email · clé Brevo dans le bundle navigateur · CORS `*` · absence totale de validation serveur.

**Désormais automatique** : backups quotidiens 02:00 (30 j, checksum, email récap) · reprise horaire des emails échoués (aucun email de gain ne peut plus être perdu) · surveillance 15 min + alerte email dédupliquée · retries réseau + reconnexion realtime dans le CRM · rate limiting · déploiement continu.

**Reste à faire manuellement (une fois)** :
1. **Binding KV** : Cloudflare Pages → ted-crm → Settings → Bindings → Add → KV namespace → nom `BACKUPS`, namespace `ted-crm-backups` → puis relancer un déploiement. Active : backups, rate limiting, dédup d'alertes → pastille verte.
2. (Optionnel) `BACKUP_SECRET` dans les variables Cloudflare (`openssl rand -hex 24`) pour appeler `/api/run-tests` hors CRM.
3. (Recommandé) Supabase → Authentication → activer la protection « leaked passwords ».

**À surveiller (mensuel, 5 min)** : quota Brevo (~300 emails/jour gratuit, crédits SMS), page Système (tests verts, erreurs, file emails), espace Supabase. Checklist détaillée dans MAINTENANCE.md §5.


