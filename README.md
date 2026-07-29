# Retail Tracker

A React + Vite app for tracking retail sales across a solo shop or a team of
shops with a boss dashboard. Data is stored in Appwrite.

This build was verified to install and compile cleanly (`npm install && npm run build`)
before being handed to you.

## What changed from the Claude artifact version

The only thing that changed is the storage layer: `window.storage` (which only
exists inside Claude.ai) was swapped for Appwrite. All the login, dashboard,
and hardening logic (hashed boss/shop/PIN codes) is untouched.

Storage model: each data bucket (`shops`, `users`, `sales`, `solo_businesses`,
`boss_code_hash`) is stored as one Appwrite document, using that name as the
document's own ID, with the JSON payload stringified into a single `value`
text column in the `kv_store` table.

---

## 1. Appwrite setup (already done, for reference)

You've already:
1. Created an Appwrite project.
2. Created a database.
3. Created a `kv_store` table inside it with one column, `value` (Text, Required).
4. Set permissions on the table: role **Any** with Create, Read, Update checked.

You'll need these four values (you already have all of them):

| Value | Where to find it |
|---|---|
| API Endpoint | Project → Overview |
| Project ID | Project → Overview |
| Database ID | Databases → your database → shown near the top / in URL |
| Table ID | Inside `kv_store` → shown near the top / in URL |

## 2. Push this project to GitHub

From this project folder:

```bash
git init
git add .
git commit -m "Initial commit: retail tracker"
gh repo create retail-tracker --private --source=. --push
```

(No GitHub CLI? Create an empty repo on github.com instead, then:
`git remote add origin <your-repo-url> && git branch -M main && git push -u origin main`)

## 3. Deploy to Vercel

1. Go to [vercel.com](https://vercel.com) → **Add New → Project** → import the
   GitHub repo you just pushed.
2. Vercel auto-detects Vite — leave the build settings as-is
   (`npm run build`, output directory `dist`).
3. Before deploying, open **Environment Variables** and add all four:
   - `VITE_APPWRITE_ENDPOINT`
   - `VITE_APPWRITE_PROJECT_ID`
   - `VITE_APPWRITE_DATABASE_ID`
   - `VITE_APPWRITE_TABLE_ID`
4. Click **Deploy**.

You'll get a live `*.vercel.app` URL. Every push to `main` redeploys
automatically.

## Local development

```bash
npm install
cp .env.example .env.local   # then fill in your four Appwrite values
npm run dev
```

## Notes

- The boss master code, shop portal codes, and all PINs are SHA-256 hashed
  client-side before being written to Appwrite — the same hardening from the
  Claude artifact version.
- The `kv_store` table is intentionally simple (a JSON blob per document) so
  this ports 1:1 from the original data model. A future improvement would be
  proper structured tables (`shops`, `users`, `sales`) with real columns and
  relationships — useful once you're past the "get it live" stage.
- Appwrite's "Any" role permission means anyone with your Project ID can
  read/write this one table via the public SDK — same trust model as the
  original shared storage. The app's own PIN and portal-code screens are what
  actually gate who can do what inside the UI.
