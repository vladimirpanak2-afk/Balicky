# Balicky App (SK/CZ)

Webova aplikace pro upravy balicku a exporty do Agendy.

## Lokalni spusteni

```bash
npm install
npm start
```

Aplikace bezi na `http://localhost:8787`.

## Produkcni nasazeni (Render)

Repo obsahuje `render.yaml`, takze Render si nastaveni nacte automaticky.

1. Otevri [Render Dashboard](https://dashboard.render.com/)
2. **New +** -> **Blueprint**
3. Vyber repo `vladimirpanak2-afk/Balicky`
4. Potvrd nasazeni
5. Pockej na build a ziskej verejny URL odkaz

Poznamky:
- `DATA_DIR` je nastaven na `/var/data`, kde se drzi `sessions.json` (rozpracovana prace).
- SK data se nactou z `balicky_sk.xlsx` + `katalog.js`.
- Pro CZ staci do repa pridat `balicky_cz.xlsx` (pripadne `balicky_cz.js`) a `katalog_cz.js`.

## API

- `GET /api/health`
- `GET /api/countries`
- `GET /api/data/:country`
- `GET /api/session?country=sk&user=...&pm=...`
- `PUT /api/session`
- `DELETE /api/session?country=sk&user=...&pm=...`
