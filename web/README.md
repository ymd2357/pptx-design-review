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
Open
`/visual/?deck=260329-seminar-curriculum-proposal&rev=017&observation=P0-3`
for the slide PNG review surface scoped to one observation.

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

On the review screen, `Download files` remains available as the unauthenticated
fallback and downloads both `rev-NNN-decisions.tsv` and
`rev-NNN-finding-judgements.json`. `Commit to repo` is enabled only when the TSV
was loaded from GitHub and the user is authenticated. It sends Contents API
`PUT` requests for both files using the last read file SHA where present,
creating commits on the repository default branch. There is no branch or pull
request flow in this phase.

If GitHub returns a SHA conflict, the page shows a reload action. Reloading
fetches the latest TSV from GitHub and discards in-progress edits. Unsaved
`rationale` and `finding_dispositions` edits are otherwise kept in
`sessionStorage` using the `pptx-review:` prefix, and are cleared after a
successful commit.

## Visual Review

The visual route loads `tmp/review-snapshot/<deck>/rev-<NNN>/lint.json` and
`images/slide-XX.png`. It renders one slide at a time, supports left/right
keyboard navigation plus pointer swipes, and draws SVG rectangles for finding
`bbox_pt` coordinates on top of the PNG. Selecting a rectangle opens a finding
drawer with the lint details and per-finding judgement controls.

Each finding judgement is stored immediately in `localStorage` with the
`pptx-review:finding-judgements:<deck>:<rev>:<group_key>` prefix. The review
page merges those drafts with the persisted JSON artifact and computes
`finding_dispositions` from individual findings whenever it renders or saves.
The persisted artifact lives next to the TSV:

```text
doc/reviews/<deck>/rev-NNN-finding-judgements.json
```

If the JSON file is missing, the client initializes every lint finding as
`review_status=unreviewed` with `judgement_reason=null`; the first review-page
save creates the file.

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
