# PPTX Design Review Web UI

Mobile-first static SPA for filling `REV-NNN` review decisions.

## Development

```bash
cd web
npm install
npm run dev
```

Open `/` for the review hub. Open
`/review/?deck=260329-seminar-curriculum-proposal&rev=017` for the REV-017
screen.

Local development reads TSV files from `../doc/reviews/` through the Vite dev
server. Static builds can still show mock deck metadata when unauthenticated.

## GitHub Read Mode

Set these environment variables before running the dev server or build:

```bash
VITE_GITHUB_CLIENT_ID=...
VITE_GITHUB_OWNER=...
VITE_GITHUB_REPO=...
```

The app uses GitHub OAuth Device Flow and stores the access token in
`localStorage` with the `pptx-review:` prefix. Authenticated reads use the
GitHub Contents API. TSV write-back is intentionally not implemented in this
phase; the save action downloads a TSV file.

## Build

```bash
cd web
npm run build
```

The static output is written to `web/dist/`.
