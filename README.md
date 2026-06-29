# Balicky App (SK/CZ)

Webova aplikace pro upravy balicku a exporty do Agendy.

## Lokalni spusteni

```bash
npm install
npm start
```

Aplikace bezi na `http://localhost:8787`.

## Nasazeni na Google Cloud Run

### 1) Predpoklady

- Google Cloud projekt s povolenym billingem.
- Nainstalovany Google Cloud CLI (`gcloud`).
- Prihlaseni v CLI:

```bash
gcloud auth login
gcloud config set project TVUJ_PROJECT_ID
```

### 2) Zapni potrebne sluzby

```bash
gcloud services enable run.googleapis.com cloudbuild.googleapis.com artifactregistry.googleapis.com firestore.googleapis.com
```

### 3) Inicializuj Firestore (Native mode)

V Google Cloud konzoli otevri Firestore a vytvor databazi v **Native mode**.

### 4) Nasad aplikaci na Cloud Run

Spust z rootu repozitare:

```bash
gcloud run deploy balicky-app ^
  --source . ^
  --region europe-central2 ^
  --allow-unauthenticated ^
  --set-env-vars SESSION_BACKEND=firestore
```

Poznamka: Na Linux/macOS pouzij `\` misto `^` na konci radku.

### 5) Otestuj produkcni URL

Po deploy vypise Cloud Run URL, napr.:

`https://balicky-app-xxxxx-ew.a.run.app`

Zkontroluj:

- `GET /api/health`
- otevreni root URL v prohlizeci

## Session backendy

Server umi 2 rezimy:

- `SESSION_BACKEND=firestore` -> produkce na Google (doporuceno)
- bez promenne -> lokalni `data/sessions.json`

Cloud Run nema perzistentni disk pro vice instanci, proto je Firestore pro sdilenou rozpracovanou praci nutny.

## Uzivatele (login + PIN)

Aplikace pouziva prihlaseni:

- login ve formatu `prijmeni.jmeno`
- `PIN` presne 6 cislic

Backend drzi uzivatele v:

- Firestore kolekci `users` (produkce)
- nebo lokalne `data/users.json` (lokalni beh)

Pro lokalni test je pripraven uzivatel:

- login: `panak.vladimir`
- pin: `123456`

Poznamka: pro produkci doporucujeme piny neukladat v plain textu, ale hashovat.

## Data soubory

- SK: `balicky_sk.xlsx` + `katalog.js`
- CZ: pridej `balicky_cz.xlsx` (nebo `balicky_cz.js`) + `katalog_cz.js`

## API

- `GET /api/health`
- `GET /api/countries`
- `POST /api/auth/login`
- `GET /api/data/:country`
- `GET /api/session?country=sk&user=...&pm=...`
- `PUT /api/session`
- `DELETE /api/session?country=sk&user=...&pm=...`
