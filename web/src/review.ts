import "./styles.css";
import { requireAuth } from "./auth/auth-gate";
import {
  parseDecisionTsv,
  serializeDecisionTsv,
  validateDecisionRow,
  type DecisionRow,
} from "./data/decisions-tsv";
import { submitFeedback } from "./data/feedback-client";
import {
  aggregateFindingDispositions,
  clearFindingJudgementDrafts,
  findingJudgementsPath,
  initializeFindingJudgements,
  mergeLocalJudgementDrafts,
  parseFindingJudgementsJson,
  serializeFindingJudgementsJson,
  type FindingJudgementsFile,
} from "./data/finding-judgements";
import { parseLintJson, type LintFinding } from "./data/lint-json";
import {
  fetchDecisionTsv,
  fetchJsonFile,
  fetchReviewSnapshot,
  type ReviewSnapshot,
} from "./github/contents";
import { sitePath } from "./site-path";
import { renderObservationCard } from "./ui/observation-card";

const appElement = document.querySelector<HTMLDivElement>("#app");
if (!appElement) throw new Error("Missing #app");
const app = appElement;

const params = new URLSearchParams(location.search);
const deck = params.get("deck") ?? "260329-seminar-curriculum-proposal";
const rev = params.get("rev") ?? "017";

let rows: DecisionRow[] = [];
let sourceSha: string | undefined;
let snapshot: ReviewSnapshot | undefined;
let lintFindings: LintFinding[] = [];
let findingJudgements: FindingJudgementsFile = { deck, rev, judgements: {} };
let findingJudgementsFilePath = findingJudgementsPath(deck, rev);

void (async () => {
  await requireAuth(app);
  await renderReview();
})();

async function renderReview(): Promise<void> {
  app.replaceChildren(shell("Loading decisions..."));
  try {
    const file = await fetchDecisionTsv(deck, rev);
    sourceSha = file.sha;
    const parsed = parseDecisionTsv(file.text);
    rows = cloneRows(parsed.rows);
    applyDrafts(rows);
    snapshot = await loadSnapshot();
    await loadFindingJudgements();
    applyFindingJudgementAggregates(rows);
    renderLoaded(file.source, parsed.errors);
  } catch (error) {
    app.replaceChildren(
      shell(error instanceof Error ? error.message : "Failed to load decisions."),
    );
  }
}

function renderLoaded(source: "github" | "local", parseErrors: string[]): void {
  const root = shell("");

  const summary = document.createElement("section");
  summary.className = "review-summary";
  const invalidCount = rows.filter((row) => validateDecisionRow(row).length > 0).length;
  summary.innerHTML = `
    <div>
      <p class="eyebrow">${source}${sourceSha ? ` / ${sourceSha.slice(0, 7)}` : ""}</p>
      <h2>${deck} / REV-${rev}</h2>
    </div>
    <p>${rows.length} observations / ${invalidCount} validation issues</p>
  `;

  const messages = document.createElement("div");
  messages.className = "message-list";
  for (const error of parseErrors) {
    const item = document.createElement("p");
    item.className = "inline-error";
    item.textContent = error;
    messages.append(item);
  }

  const cardList = document.createElement("section");
  cardList.className = "card-list";
  const rerenderSummary = () => {
    const count = rows.filter((row) => validateDecisionRow(row).length > 0).length;
    summary.querySelector("p:last-child")!.textContent =
      `${rows.length} observations / ${count} validation issues`;
    persistDrafts();
  };
  rows.forEach((row) => {
    cardList.append(renderObservationCard(row, rerenderSummary, visualReviewHref(row)));
  });

  const actions = document.createElement("div");
  actions.className = "sticky-actions";
  const back = document.createElement("a");
  back.className = "secondary-link";
  back.href = sitePath("");
  back.textContent = "Hub";
  const download = document.createElement("button");
  download.type = "button";
  download.className = "primary-button";
  download.textContent = "Download files";
  download.addEventListener("click", downloadReviewFiles);
  const submit = document.createElement("button");
  submit.type = "button";
  submit.className = "primary-button";
  submit.textContent = "Submit to KV";
  submit.title = "Encrypt with age public key and POST to the shared KV.";
  submit.addEventListener("click", () => {
    submit.disabled = true;
    void handleSubmit().finally(() => {
      submit.disabled = false;
    });
  });
  actions.append(back, download, submit);

  root.append(summary, messages, renderSnapshotPanel(), cardList, actions);
  app.replaceChildren(root);
}

async function loadSnapshot(): Promise<ReviewSnapshot | undefined> {
  try {
    return await fetchReviewSnapshot(deck, rev);
  } catch (error) {
    console.warn("Review snapshot fetch failed.", error);
    return undefined;
  }
}

async function loadFindingJudgements(): Promise<void> {
  lintFindings = snapshot?.lint ? parseLintJson(snapshot.lint).findings : [];
  findingJudgementsFilePath = findingJudgementsPath(deck, rev);
  const file = await fetchJsonFile<unknown>(findingJudgementsFilePath);
  findingJudgements = mergeLocalJudgementDrafts(
    initializeFindingJudgements(
      deck,
      rev,
      lintFindings,
      file ? parseFindingJudgementsJson(file.data, deck, rev) : undefined,
    ),
    deck,
    rev,
    lintFindings,
  );
}

function renderSnapshotPanel(): HTMLElement {
  const panel = document.createElement("section");
  panel.className = "artifact-panel";

  const title = document.createElement("div");
  const eyebrow = document.createElement("p");
  eyebrow.className = "eyebrow";
  eyebrow.textContent = "snapshot";
  const heading = document.createElement("h2");
  heading.textContent = snapshot ? "Published review artifacts" : "No published artifacts";
  title.append(eyebrow, heading);

  const meta = document.createElement("p");
  if (!snapshot) {
    meta.textContent = "Run the snapshot publisher and commit tmp/review-snapshot to show evidence.";
    panel.append(title, meta);
    return panel;
  }

  meta.textContent =
    `${snapshot.imageUrls.length} slide images / ${lintCount(snapshot.lint)} lint findings` +
    (snapshot.priorities ? " / priorities.json" : "");

  const imageGrid = document.createElement("div");
  imageGrid.className = "artifact-grid";
  for (const imageUrl of snapshot.imageUrls.slice(0, 4)) {
    const image = document.createElement("img");
    image.src = imageUrl;
    image.alt = "Published slide snapshot";
    image.loading = "lazy";
    imageGrid.append(image);
  }

  const lintLink = document.createElement("a");
  lintLink.className = "text-link";
  lintLink.href = `${snapshot.basePath}/lint.json`;
  lintLink.target = "_blank";
  lintLink.rel = "noreferrer";
  lintLink.textContent = "lint.json";

  panel.append(title, meta);
  if (imageGrid.childElementCount > 0) panel.append(imageGrid);
  panel.append(lintLink);
  return panel;
}

function shell(statusText: string): HTMLElement {
  const root = document.createElement("main");
  root.className = "app-shell";
  const header = document.createElement("header");
  header.className = "app-header";
  header.innerHTML = `
    <div>
      <p class="eyebrow">PPTX Design Review</p>
      <h1>REV-${rev}</h1>
    </div>
  `;
  root.append(header);
  if (statusText) {
    const status = document.createElement("p");
    status.className = "status-text";
    status.textContent = statusText;
    root.append(status);
  }
  return root;
}

function downloadReviewFiles(): void {
  applyFindingJudgementAggregates(rows);
  const tsv = serializeDecisionTsv(rows);
  downloadBlob(tsv, `rev-${rev}-decisions.tsv`, "text/tab-separated-values;charset=utf-8");
  const json = serializeFindingJudgementsJson(prepareFindingJudgementsForSave());
  downloadBlob(json, `rev-${rev}-finding-judgements.json`, "application/json;charset=utf-8");
}

function downloadBlob(content: string, filename: string, type: string): void {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

async function handleSubmit(): Promise<void> {
  applyFindingJudgementAggregates(rows);

  const validationErrors = rows.flatMap((row) =>
    validateDecisionRow(row).map((error) => `${row.review_no} / ${row.check_id}: ${error}`),
  );
  if (validationErrors.length > 0) {
    window.alert(`Resolve validation issues before submitting.\n\n${validationErrors.join("\n")}`);
    return;
  }

  try {
    const judgements = prepareFindingJudgementsForSave();
    const { key } = await submitFeedback({
      deck,
      rev,
      decisions: rows,
      findingJudgements: judgements,
    });
    clearDrafts();
    clearFindingJudgementDrafts(deck, rev);
    showSubmitSuccessBanner(key);
  } catch (error) {
    window.alert(error instanceof Error ? error.message : "Failed to submit review.");
  }
}

function showSubmitSuccessBanner(key: string): void {
  removeBanners();
  const banner = document.createElement("section");
  banner.className = "commit-banner success";
  const note = document.createElement("span");
  note.textContent = `Submitted (${key}). Run scripts/fetch-reviews.py on your PC to apply.`;
  banner.append(note);
  app.querySelector(".app-shell")?.prepend(banner);
  window.setTimeout(() => banner.remove(), 8000);
}

function removeBanners(): void {
  app.querySelectorAll(".commit-banner").forEach((banner) => banner.remove());
}

function applyDrafts(targetRows: DecisionRow[]): void {
  for (const row of targetRows) {
    const rationale = sessionStorage.getItem(draftKey(row, "rationale"));
    const findingDispositions = sessionStorage.getItem(draftKey(row, "finding_dispositions"));
    if (rationale !== null) row.rationale = rationale;
    if (findingDispositions !== null) row.finding_dispositions = findingDispositions;
  }
}

function persistDrafts(): void {
  for (const row of rows) {
    sessionStorage.setItem(draftKey(row, "rationale"), row.rationale);
    sessionStorage.setItem(draftKey(row, "finding_dispositions"), row.finding_dispositions);
  }
}

function clearDrafts(): void {
  const prefix = draftPrefix();
  for (let index = sessionStorage.length - 1; index >= 0; index -= 1) {
    const key = sessionStorage.key(index);
    if (key?.startsWith(prefix)) sessionStorage.removeItem(key);
  }
}

function draftKey(row: DecisionRow, field: "rationale" | "finding_dispositions"): string {
  return `${draftPrefix()}${encodeURIComponent(row.review_no)}:${encodeURIComponent(row.check_id)}:${field}`;
}

function draftPrefix(): string {
  return `pptx-review:draft:${deck}:${rev}:`;
}

function cloneRows(sourceRows: DecisionRow[]): DecisionRow[] {
  return sourceRows.map((row) => ({ ...row }));
}

function lintCount(lint: unknown): number {
  return Array.isArray(lint) ? lint.length : 0;
}

function visualReviewHref(row: DecisionRow): string {
  const query = new URLSearchParams({
    deck,
    rev,
    observation: row.review_no,
  });
  return sitePath(`visual/?${query.toString()}`);
}

function applyFindingJudgementAggregates(targetRows: DecisionRow[]): void {
  if (lintFindings.length === 0) return;
  for (const row of targetRows) {
    const hasFindings = lintFindings.some((finding) => finding.check === row.check_id);
    if (!hasFindings) continue;
    row.finding_dispositions = aggregateFindingDispositions(
      lintFindings,
      findingJudgements,
      row.check_id,
    );
  }
}

function prepareFindingJudgementsForSave(): FindingJudgementsFile {
  findingJudgements = mergeLocalJudgementDrafts(
    initializeFindingJudgements(deck, rev, lintFindings, findingJudgements),
    deck,
    rev,
    lintFindings,
  );
  return findingJudgements;
}

