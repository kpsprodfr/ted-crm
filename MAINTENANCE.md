# MAINTENANCE — TED CRM

Guide de maintenance pour faire tourner le système sans intervention, et savoir quoi faire quand quelque chose casse. Mis à jour le 2026-07-02.

---

## 1. Architecture

```
ted-crm.pages.dev (Cloudflare Pages, deploy auto sur push GitHub kpsprodfr/ted-crm)
│
├── CRM (React CRA, ted-crm-v2/src/App.js) — accès via login Supabase Auth
├── Pages publiques : accueil.html (soirées + roue), reserver.html, menu.html
│
├── Cloudflare Functions (ted-crm-v2/functions/)
│   ├── /send-email            POST, JWT CRM — emails (file d'attente si échec)
│   ├── /send-sms              POST, JWT CRM — SMS (clé Brevo jamais côté client)
│   ├── /send-push-onesignal   POST, JWT CRM — push nouvelles résas
│   ├── /api/roue-email        POST — email de gain (public : gain_id vérifié en base, idempotent)
│   ├── /api/roue-notify       POST, JWT CRM — legacy (aucun appelant)
│   ├── /api/health            GET — santé (pastille CRM + cron alertes)
│   ├── /api/process-email-queue GET — reprise file emails (cron, claim atomique)
│   ├── /api/backup-daily      GET cron / POST JWT — backup quotidien vers KV
│   ├── /api/backups           GET, JWT — liste des backups
│   ├── /api/backup-restore    POST, JWT — restauration (upsert non destructif)
│   └── /api/run-tests         GET, X-Test-Key ou JWT — suite de tests
│
├── Supabase (projet mwpfaytccypvdrgapptk, eu-west-1, Postgres 17)
│   ├── Tables : clients, reservations, roue_gains, roue_recompenses, roue_config,
│   │            parametres, menu_* (6 tables), sms_envoyes, emails_envoyes,
│   │            email_queue, error_logs, fcm_tokens (legacy)
│   ├── RLS activée partout ; anon = usage minimal des pages publiques
│   │   (jamais de DELETE, jamais de lecture des réservations/file emails)
│   └── pg_cron : process-email-queue (xx:05), backup-daily (02:00 UTC),
│                 health?notify=1 (toutes les 15 min)
│
├── Cloudflare KV : namespace ted-crm-backups (backups + rate limiting + dédup alertes)
├── Brevo : emails transactionnels/campagnes + SMS (expéditeur com.astegal@gmail.com)
└── OneSignal : push CRM (app_id 87b29550-ffb0-412a-9682-05fdace514fc)
```

**Toute la sécurité des endpoints** est dans `functions/_utils.js` : origine autorisée (`ted-crm.pages.dev`, previews, localhost), rate limit 10 req/min/IP (via KV), headers (nosniff, DENY, HSTS, CSP), échappement HTML, validations email/tel.

---

## 2. Variables d'environnement

### Cloudflare Pages (dashboard → Workers & Pages → ted-crm → Settings → Variables and secrets)

| Variable | Rôle | Si elle manque | Comment la régénérer |
|---|---|---|---|
| `BREVO_API_KEY` | Envoi emails + SMS | Emails en file d'attente, SMS KO, alerte santé | Brevo → Settings → SMTP & API → API Keys |
| `SUPABASE_SERVICE_ROLE_KEY` | File emails, backups, restauration, tests | File non traitée, backups KO (santé « down ») | Supabase → Project Settings → API keys → service_role (⚠️ ne jamais l'exposer côté client) |
| `REACT_APP_SUPABASE_URL` | Build React | Build échoue | `https://mwpfaytccypvdrgapptk.supabase.co` |
| `REACT_APP_SUPABASE_ANON_KEY` | Build React | Build échoue | Supabase → API keys → clé publishable |
| `ONESIGNAL_REST_API_KEY` | Push résas | Plus de push (non bloquant) | OneSignal → Settings → Keys & IDs |
| `BACKUP_SECRET` | Header X-Test-Key de /api/run-tests | Tests accessibles uniquement via CRM connecté | `openssl rand -hex 24` |

### Bindings Cloudflare Pages (Settings → Bindings)

| Binding | Type | Valeur | Si absent |
|---|---|---|---|
| `BACKUPS` | KV namespace | `ted-crm-backups` (id `b9d3c42b261b4ababdb9f8eab8fa9c71`) | Backups inactifs, rate limiting inactif, alertes non dédupliquées — santé « degraded » |

### Local (`ted-crm-v2/.env`, jamais commité)

`REACT_APP_SUPABASE_URL` et `REACT_APP_SUPABASE_ANON_KEY` sont nécessaires pour `npm start`/`npm run build` en local (sinon le build jette une erreur explicite).

---

## 3. Procédures de panne par composant

**Réflexe n°1 : CRM → Système** (pastille en bas de la sidebar : vert/orange/rouge). Un email d'alerte automatique part vers com.astegal@gmail.com si un composant est down (max 1/heure).

| Panne | Symptôme | Procédure |
|---|---|---|
| **Brevo down / quota épuisé** | Alerte santé « brevo down » ; emails de gain non reçus | Rien à faire dans l'immédiat : les emails partent en file (`email_queue`) et sont repris toutes les heures (5 tentatives). Vérifier le compte Brevo (quota gratuit ≈ 300 emails/jour, crédits SMS prépayés). Après rétablissement, la file se vide seule. |
| **Clé Brevo révoquée** | `brevo: down — clé invalide` | Régénérer la clé (tableau §2), la remplacer dans Cloudflare, redéployer (Deployments → Retry). |
| **Supabase injoignable** | Pastille rouge, CRM vide | Vérifier https://status.supabase.com. Le front retry automatiquement (3×, backoff). Si la panne dure : rien à faire côté code, tout revient seul. |
| **Données corrompues / supprimées par erreur** | — | CRM → Système → Restaurer (voir §4). |
| **Cloudflare Pages down** | Site injoignable | Vérifier https://www.cloudflarestatus.com. Deploy auto au prochain push sinon Deployments → Retry deployment. |
| **Push OneSignal KO** | Plus de notifications résas | Non bloquant. Vérifier ONESIGNAL_REST_API_KEY et le dashboard OneSignal. |
| **Emails « failed » dans la file** | Test `email_queue` en échec | Table `email_queue`, lignes `status=failed` : lire `error_message`, corriger la cause, puis repasser les lignes en `pending` (elles seront reprises). |

---

## 4. Restauration depuis un backup

Les backups tournent tous les jours à 02:00 UTC (12 tables → Cloudflare KV, rétention 30 jours, checksum SHA-256, email récap à com.astegal@gmail.com).

1. CRM → **Système** → tableau des backups.
2. Bouton **Restaurer** sur la date voulue → choisir la table → taper `RESTAURER` → confirmer.
3. La restauration est un **upsert** : les lignes du backup sont recréées/écrasées par identifiant ; les lignes créées depuis le backup ne sont **pas** supprimées.
4. Vérifier le résultat dans la page concernée (Clients, Réservations…).

Récupération brute sans CRM : Cloudflare dashboard → Workers & Pages → KV → `ted-crm-backups` → clé `backup:YYYY-MM-DD` (JSON complet téléchargeable).

---

## 5. Checklist mensuelle (5 minutes)

- [ ] CRM → Système : pastille **verte**, « Lancer les tests » → tout OK.
- [ ] Un email « ✅ Backup TED CRM » reçu ce matin sur com.astegal@gmail.com.
- [ ] Quota Brevo : dashboard Brevo → consommation emails (≈300/jour gratuit) et crédits SMS restants.
- [ ] Table `error_logs` (page Système) : pas d'erreur récurrente inexpliquée.
- [ ] Table `email_queue` : aucun `status=failed`.
- [ ] Espace Supabase : Project → Reports (plan gratuit : 500 Mo DB).
- [ ] `npm run security-audit` en local de temps en temps (vulnérabilités restantes connues : devDependencies de react-scripts, build-time uniquement, ne pas lancer `npm audit fix --force` qui casserait CRA).

---

## 6. Accès et comptes

| Service | Accès | Sert à |
|---|---|---|
| GitHub `kpsprodfr/ted-crm` | compte kpsprodfr | Code + déclencheur de déploiement |
| Cloudflare Pages `ted-crm` | compte Cloudflare du projet | Hébergement, functions, KV, variables |
| Supabase `ted-crm` (mwpfaytccypvdrgapptk) | compte Supabase du projet | Base de données, Auth CRM, pg_cron |
| Brevo | com.astegal@gmail.com | Emails + SMS |
| OneSignal | compte du projet | Push CRM |

⚠️ Le PAT GitHub embarqué dans le remote git local n'a pas le scope `workflow` : les fichiers `.github/workflows/*` ne peuvent pas être poussés depuis cette machine (pas bloquant : les crons passent par pg_cron côté Supabase).

---

## 7. Ce qui est automatique (aucune intervention)

- **Backups** quotidiens 02:00 UTC + rétention 30 j + email récap.
- **Reprise des emails échoués** toutes les heures (max 5 tentatives par email).
- **Surveillance** toutes les 15 min + alerte email si panne (dédupliquée 1/h).
- **Retries réseau** : tous les appels Supabase du CRM (timeout 10 s, 3 retries backoff) ; reconnexion automatique du temps réel.
- **Déploiement** : chaque push sur `main` déploie automatiquement.
