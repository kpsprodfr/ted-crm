# TED — Fichier Clients v2

CRM avec login Supabase, base de données partagée, accessible depuis n'importe quel appareil.

## Variables d'environnement

Créez un fichier `.env` à la racine du projet :

```
REACT_APP_SUPABASE_URL=https://VOTRE_ID.supabase.co
REACT_APP_SUPABASE_ANON_KEY=VOTRE_CLE_PUBLISHABLE
```

## Lancer en local

```
npm install
npm start
```

## Déployer sur Netlify

Dans Netlify → Site configuration → Environment variables, ajoutez :
- REACT_APP_SUPABASE_URL
- REACT_APP_SUPABASE_ANON_KEY
