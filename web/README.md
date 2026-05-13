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

## GitHub Contents API Mode

Set these environment variables before running the dev server or build:

```bash
VITE_GITHUB_CLIENT_ID=...
VITE_GITHUB_OWNER=...
VITE_GITHUB_REPO=...
```

The app uses GitHub OAuth Device Flow with the `repo` scope and stores the
access token in `localStorage` with the `pptx-review:` prefix. Authenticated
reads use the GitHub Contents API. Write-back requires a token that can update
repository contents, either the Device Flow token or a fine-grained PAT with
Contents: write.

On the review screen, `Download TSV` remains available as the unauthenticated
fallback. `Commit to repo` is enabled only when the TSV was loaded from GitHub
and the user is authenticated. It sends a Contents API `PUT` with the last read
file SHA, creating a commit on the repository default branch. There is no branch
or pull request flow in this phase.

If GitHub returns a SHA conflict, the page shows a reload action. Reloading
fetches the latest TSV from GitHub and discards in-progress edits. Unsaved
`rationale` and `finding_dispositions` edits are otherwise kept in
`sessionStorage` using the `pptx-review:` prefix, and are cleared after a
successful commit.

## Build

```bash
cd web
npm run build
```

The static output is written to `web/dist/`.

## Review Artifact Snapshots

Rendered slide PNGs and lint JSON live under `tmp/review/`, which is ignored.
To make evidence available on the deployed static site, publish an explicit
snapshot and commit it:

```bash
python3 scripts/publish_review_snapshot.py --deck 260329-seminar-curriculum-proposal --rev 017
git add tmp/review-snapshot/260329-seminar-curriculum-proposal/rev-017
git commit
git push
```

The snapshot output is `tmp/review-snapshot/<deck>/rev-<NNN>/` with:

- `images/slide-XX.png`
- `lint.json`
- `priorities.json` when a matching priorities file exists

The script only copies and re-encodes existing artifacts. It does not run the
review orchestrator, render slides, or regenerate lint.

## Deploy

GitHub Actions builds `web/`, copies `doc/reviews/` and the optional
`tmp/review-snapshot/` tree into `web/dist/`, and deploys the result to GitHub
Pages. The deploy workflow runs after pushing relevant web, review, snapshot,
or workflow changes to the repository default branch, and can also be started
with `workflow_dispatch`.
