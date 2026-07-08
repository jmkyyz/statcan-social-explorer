# StatCan Social Explorer

Explore non-economic Statistics Canada data — health, education, Indigenous
statistics, justice & safety, demographics, and income support — with charting,
frequency conversion, % change transforms, indexing, and XLSX export.

This is a structural sibling of **StatCan Econ Explorer** (`statcan-explorer`):
same Flask proxy to the StatCan Web Data Service (WDS) API, same single-file
frontend architecture, same admin Wizard for catalog management — but a
completely independent codebase, catalog, and (eventually) deployment.

## Files

| File | What it does |
|------|-------------|
| `proxy.py` | Flask server: proxies the StatCan WDS API, serves the frontend, and hosts the catalog/Wizard APIs. Listens on **port 5004** by default (Econ Explorer uses 5001, so both run side by side). |
| `social-explorer.html` | Single-file frontend (vanilla JS + Chart.js + SheetJS), served at `/`. |
| `wizard.html` | Admin tool at `/wizard` for adding/editing/deleting catalog series. |
| `Vectors.xlsx` | The series catalog (sheet `series`). Ships **empty** — add verified series via the Wizard. |

## Run locally

```bash
pip install -r requirements.txt
python proxy.py
# → http://localhost:5004        (the app)
# → http://localhost:5004/wizard (catalog admin)
```

## Catalog

The category picker is driven entirely by the `category` column in
`Vectors.xlsx`. Six default buckets are seeded in the UI even while the catalog
is empty: Health, Education, Indigenous Statistics, Justice & Safety,
Demographics & Population, Social Assistance & Income Support.

Columns (same schema as Econ Explorer): `category, freq, series_id,
series_name, table_id, dim1–5 name/value, vector, full_label, short_label,
dim1_group` (+ optional `dim1–5_level`).

## Wizard auth

Same mechanism as Econ Explorer:

- Set `ADMIN_KEY` in the environment → wizard endpoints require the
  `X-Admin-Key` header (the wizard UI prompts for the key and caches it in
  localStorage).
- No `ADMIN_KEY` locally → open (local dev convenience).
- On Render (`RENDER` env var present) auth is **always** enforced, even
  before `ADMIN_KEY` is configured.

With `GITHUB_TOKEN` set, every catalog write is committed back to the GitHub
repo (`GITHUB_REPO`, default `jmkyyz/statcan-social-explorer`) so Render
redeploys and the change survives ephemeral-disk restarts.

## Deployment

Live on Render (service `statcan-social-explorer`, 512MB Starter) at
`social.jasonkirby.ca` (Cloudflare CNAME → Render); see `render.yaml`.
The buildCommand runs `build_catalog_cache.py` to pre-build the gzipped
catalog cache (`catalog_cache/`, gitignored) — without it, the first
`/api/catalog*` request after a deploy streams the whole 140k-row
Vectors.xlsx (~minutes on 0.5 vCPU) and times out behind Cloudflare.
The server validates the cache against Vectors.xlsx by content hash and
falls back to streaming if it's missing or stale.
