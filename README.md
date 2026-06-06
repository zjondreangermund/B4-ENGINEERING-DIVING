# B4 Nautilus Operations Suite

Railway-ready prototype for the B4 Engineering & Diving operations platform.

## What is included

- `public/index.html` - current clickable prototype
- `database/schema.sql` - first PostgreSQL schema draft
- `api/openapi.yaml` - first API route draft
- `docs/` - workbook analysis, product blueprint, screen specs and import mapping
- `imports/` - dummy data mapping summaries

## Run locally

```bash
npm install
npm start
```

Open:

```text
http://localhost:3000
```

Health check:

```text
http://localhost:3000/health
```

## Deploy on Railway from GitHub

1. Create a new GitHub repo, for example `b4-nautilus-operations`.
2. Upload/push all files in this folder to the repo root.
3. Go to Railway.
4. Click **New Project**.
5. Choose **Deploy from GitHub repo**.
6. Select the repo.
7. Railway should detect Node/Nixpacks automatically.
8. Start command is already set in `railway.json` as:

```bash
npm start
```

9. After deployment, open the Railway-generated domain.

## Future build path

This prototype is currently static. The next version should add:

- PostgreSQL on Railway
- Authentication
- Job register CRUD
- Excel import/export service
- Dashboard calculations from database
- User roles and permissions
