import "./styles.css";
import {
  clearStoredToken,
  getStoredToken,
  startDeviceFlow,
  type DeviceFlowState,
} from "./auth/device-flow";
import {
  countChangedDecisionRows,
  parseDecisionTsv,
  serializeDecisionTsv,
  validateDecisionRow,
  type DecisionRow,
} from "./data/decisions-tsv";
import {
  fetchDecisionTsv,
  fetchReviewSnapshot,
  GitHubContentError,
  putFile,
  type PutFileResult,
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
let originalRows: DecisionRow[] = [];
let sourceSha: string | undefined;
let filePath = `doc/reviews/${deck}/rev-${rev}-decisions.tsv`;
let fileSource: "github" | "local" = "local";
let snapshot: ReviewSnapshot | undefined;

void renderReview();

async function renderReview(): Promise<void> {
  app.replaceChildren(shell("Loading decisions..."));
  try {
    const file = await fetchDecisionTsv(deck, rev);
    sourceSha = file.sha;
    filePath = file.path;
    fileSource = file.source;
    const parsed = parseDecisionTsv(file.text);
    originalRows = cloneRows(parsed.rows);
    rows = cloneRows(parsed.rows);
    applyDrafts(rows);
    snapshot = await loadSnapshot();
    renderLoaded(file.source, parsed.errors);
  } catch (error) {
    app.replaceChildren(
      shell(error instanceof Error ? error.message : "Failed to load decisions."),
    );
  }
}

function renderLoaded(source: "github" | "local", parseErrors: string[]): void {
  const root = shell("");
  root.append(renderAuthPanel(() => void renderReview()));

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
    cardList.append(renderObservationCard(row, rerenderSummary));
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
  download.textContent = "Download TSV";
  download.addEventListener("click", downloadTsv);
  const commit = document.createElement("button");
  commit.type = "button";
  commit.className = "primary-button";
  commit.textContent = "Commit to repo";
  commit.disabled = !getStoredToken() || source !== "github" || !sourceSha;
  commit.title = commit.disabled
    ? "Sign in and load the TSV from GitHub before committing."
    : "Commit edited decisions.tsv to GitHub.";
  commit.addEventListener("click", () => {
    commit.disabled = true;
    void handleCommit().finally(() => {
      commit.disabled = !getStoredToken() || fileSource !== "github" || !sourceSha;
    });
  });
  actions.append(back, download, commit);

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

function renderAuthPanel(onAuthChange: () => void): HTMLElement {
  const panel = document.createElement("section");
  panel.className = "auth-panel compact";
  const token = getStoredToken();
  const status = document.createElement("p");
  status.textContent = token ? "GitHub Contents API read mode." : "Local TSV read mode.";
  const actions = document.createElement("div");
  actions.className = "button-row";

  if (token) {
    const signOut = document.createElement("button");
    signOut.type = "button";
    signOut.className = "secondary-button";
    signOut.textContent = "Sign out";
    signOut.addEventListener("click", () => {
      clearStoredToken();
      onAuthChange();
    });
    actions.append(signOut);
  } else {
    const signIn = document.createElement("button");
    signIn.type = "button";
    signIn.className = "secondary-button";
    signIn.textContent = "Sign in with GitHub";
    signIn.addEventListener("click", () => {
      signIn.disabled = true;
      void startDeviceFlow((state) => renderDeviceState(panel, state, onAuthChange)).catch(
        (error: unknown) =>
          renderDeviceState(
            panel,
            {
              status: "error",
              message: error instanceof Error ? error.message : String(error),
            },
            onAuthChange,
          ),
      );
    });
    actions.append(signIn);
  }

  panel.append(status, actions);
  return panel;
}

function renderDeviceState(
  panel: HTMLElement,
  state: DeviceFlowState,
  onAuthChange: () => void,
): void {
  let device = panel.querySelector<HTMLElement>(".device-state");
  if (!device) {
    device = document.createElement("div");
    device.className = "device-state";
    panel.append(device);
  }
  if (state.status === "code") {
    device.innerHTML = `
      <strong class="user-code">${state.code.user_code}</strong>
      <a class="text-link" href="${state.code.verification_uri}" target="_blank" rel="noreferrer">
        Open GitHub verification
      </a>
    `;
  } else if (state.status === "authenticated") {
    onAuthChange();
  } else {
    device.textContent = state.status === "pending" ? state.message : state.message;
  }
}

function downloadTsv(): void {
  const tsv = serializeDecisionTsv(rows);
  const blob = new Blob([tsv], { type: "text/tab-separated-values;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `rev-${rev}-decisions.tsv`;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

async function handleCommit(): Promise<void> {
  if (!sourceSha || fileSource !== "github") return;

  const validationErrors = rows.flatMap((row) =>
    validateDecisionRow(row).map((error) => `${row.review_no} / ${row.check_id}: ${error}`),
  );
  if (validationErrors.length > 0) {
    window.alert(`Resolve validation issues before committing.\n\n${validationErrors.join("\n")}`);
    return;
  }

  const changedCount = countChangedDecisionRows(originalRows, rows);
  const defaultMessage =
    `docs(reviews): REV-${rev} update from web UI (${changedCount} observations changed)`;
  const message = window.prompt("Commit message", defaultMessage)?.trim();
  if (!message) return;
  if (!isValidCommitMessage(message)) {
    window.alert("Commit message must be a subject line plus optional body after a blank line.");
    return;
  }

  const tsv = serializeDecisionTsv(rows);
  await commitWithRetry(message, tsv, true);
}

async function commitWithRetry(
  message: string,
  tsv: string,
  retryOnAuthError: boolean,
): Promise<void> {
  try {
    if (!sourceSha) return;
    const result = await putFile({
      path: filePath,
      message,
      content: tsv,
      sha: sourceSha,
    });
    handleCommitSuccess(result);
  } catch (error) {
    if (error instanceof GitHubContentError && error.kind === "conflict") {
      window.alert("Conflict: file changed on the server");
      showConflictBanner();
      return;
    }
    if (error instanceof GitHubContentError && error.kind === "auth" && retryOnAuthError) {
      await reauthenticateForRetry();
      await commitWithRetry(message, tsv, false);
      return;
    }
    window.alert(error instanceof Error ? error.message : "Failed to commit TSV.");
  }
}

function handleCommitSuccess(result: PutFileResult): void {
  sourceSha = result.contentSha;
  originalRows = cloneRows(rows);
  clearDrafts();
  const sourceLabel = app.querySelector<HTMLElement>(".review-summary .eyebrow");
  if (sourceLabel) sourceLabel.textContent = `${fileSource} / ${sourceSha.slice(0, 7)}`;
  showCommitSuccessBanner(result);
}

function showCommitSuccessBanner(result: PutFileResult): void {
  removeBanners();
  const banner = document.createElement("section");
  banner.className = "commit-banner success";
  const link = document.createElement("a");
  link.className = "text-link";
  link.href = result.commitUrl;
  link.target = "_blank";
  link.rel = "noreferrer";
  link.textContent = result.commitSha.slice(0, 7);
  banner.append("Committed ", link);
  app.querySelector(".app-shell")?.prepend(banner);
  window.setTimeout(() => banner.remove(), 8000);
}

function showConflictBanner(): void {
  removeBanners();
  const banner = document.createElement("section");
  banner.className = "commit-banner error";
  const message = document.createElement("p");
  message.textContent = "Conflict: file changed on the server.";
  const reload = document.createElement("button");
  reload.type = "button";
  reload.className = "secondary-button";
  reload.textContent = "Reload";
  reload.addEventListener("click", () => {
    if (!window.confirm("Reload latest TSV from GitHub? In-progress edits will be lost.")) return;
    clearDrafts();
    void renderReview();
  });
  banner.append(message, reload);
  app.querySelector(".app-shell")?.prepend(banner);
}

async function reauthenticateForRetry(): Promise<void> {
  const panel = document.createElement("section");
  panel.className = "auth-panel compact";
  const status = document.createElement("p");
  status.textContent = "GitHub token expired. Re-authenticate to retry commit.";
  panel.append(status);
  app.querySelector(".app-shell")?.prepend(panel);
  await startDeviceFlow((state) => renderDeviceState(panel, state, () => undefined));
}

function removeBanners(): void {
  app.querySelectorAll(".commit-banner").forEach((banner) => banner.remove());
}

function isValidCommitMessage(message: string): boolean {
  const lines = message.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
  return Boolean(lines[0]?.trim()) && (lines.length === 1 || lines[1] === "");
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
