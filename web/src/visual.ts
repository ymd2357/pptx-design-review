import "./styles.css";
import { requireAuth } from "./auth/auth-gate";
import {
  findingJudgementsPath,
  initializeFindingJudgements,
  judgedFindingCount,
  mergeLocalJudgementDrafts,
  parseFindingJudgementsJson,
  storeFindingJudgementDraft,
  type FindingJudgement,
  type FindingJudgementsFile,
} from "./data/finding-judgements";
import { parseDecisionTsv, type DecisionRow } from "./data/decisions-tsv";
import {
  findingsForObservation,
  loadSlideSizePt,
  loadSnapshotLint,
  type LintFinding,
  type SlideSizePt,
} from "./data/lint-json";
import {
  fetchDecisionTsv,
  fetchJsonFile,
  fetchReviewSnapshot,
  getAuthenticatedUserLogin,
} from "./github/contents";
import { sitePath } from "./site-path";
import { renderFindingDrawer } from "./ui/finding-drawer";
import { renderSlideGallery } from "./ui/slide-gallery";

const appElement = document.querySelector<HTMLDivElement>("#app");
if (!appElement) throw new Error("Missing #app");
const app = appElement;

const params = new URLSearchParams(location.search);
const deck = params.get("deck") ?? "260329-seminar-curriculum-proposal";
const rev = params.get("rev") ?? "017";
const observationParam = params.get("observation") ?? "P0-3";

let observation: DecisionRow | undefined;
let findings: LintFinding[] = [];
let judgements: FindingJudgementsFile = { deck, rev, judgements: {} };
let slideSize: SlideSizePt = { w: 1440, h: 810 };
let imageUrls: string[] = [];
let updatedBy = "local";
let progressText: HTMLElement | undefined;

void (async () => {
  await requireAuth(app);
  await renderVisualReview();
})();

async function renderVisualReview(): Promise<void> {
  app.replaceChildren(shell("Loading visual review..."));
  try {
    const [decisionFile, lintData, snapshot, loadedSlideSize, judgementFile] = await Promise.all([
      fetchDecisionTsv(deck, rev),
      loadSnapshotLint(deck, rev),
      fetchReviewSnapshot(deck, rev),
      loadSlideSizePt(),
      fetchJsonFile<unknown>(findingJudgementsPath(deck, rev)),
    ]);
    const decisions = parseDecisionTsv(decisionFile.text).rows;
    observation = findObservation(decisions, observationParam);
    if (!observation) {
      throw new Error(`Observation not found: ${observationParam}`);
    }
    findings = findingsForObservation(lintData.findings, observation.check_id);
    slideSize = loadedSlideSize;
    imageUrls = snapshot?.imageUrls ?? [];
    judgements = mergeLocalJudgementDrafts(
      initializeFindingJudgements(
        deck,
        rev,
        lintData.findings,
        judgementFile
          ? parseFindingJudgementsJson(judgementFile.data, deck, rev)
          : undefined,
      ),
      deck,
      rev,
      lintData.findings,
    );
    void getAuthenticatedUserLogin()
      .then((login) => {
        if (login) updatedBy = login;
      })
      .catch(() => undefined);
    renderLoaded();
  } catch (error) {
    app.replaceChildren(
      shell(error instanceof Error ? error.message : "Failed to load visual review."),
    );
  }
}

function renderLoaded(): void {
  if (!observation) return;
  const root = shell("");

  const summary = document.createElement("section");
  summary.className = "visual-summary";
  const back = sitePath(
    `review/?deck=${encodeURIComponent(deck)}&rev=${encodeURIComponent(rev)}`,
  );
  summary.innerHTML = `
    <div>
      <p class="eyebrow">${escapeHtml(deck)} / REV-${escapeHtml(rev)}</p>
      <h2>${escapeHtml(observation.review_no)} ${escapeHtml(observation.check_id)}</h2>
    </div>
    <a class="secondary-link" href="${back}">Back</a>
  `;

  progressText = document.createElement("p");
  progressText.className = "visual-progress";
  updateProgress();

  const gallery =
    imageUrls.length > 0
      ? renderSlideGallery({
          imageUrls,
          findings,
          judgements,
          slideSizePt: slideSize,
          initialSlideNo: findings[0]?.slideNo,
          onSelectFinding: openFinding,
        })
      : renderEmptyState("No slide PNGs found in the review snapshot.");

  root.append(summary, progressText, gallery);
  app.replaceChildren(root);
}

function openFinding(finding: LintFinding): void {
  const existing = document.querySelector(".drawer-backdrop");
  existing?.remove();
  const judgement = judgements.judgements[finding.key] ?? {
    review_status: "unreviewed",
    judgement_reason: null,
  };
  document.body.append(
    renderFindingDrawer({
      finding,
      judgement,
      onClose: () => document.querySelector(".drawer-backdrop")?.remove(),
      onChange: (next) => updateJudgement(finding, next),
    }),
  );
}

function updateJudgement(finding: LintFinding, next: FindingJudgement): void {
  const judgement: FindingJudgement = {
    ...next,
    updated_at: new Date().toISOString(),
    updated_by: updatedBy,
  };
  judgements.judgements[finding.key] = judgement;
  storeFindingJudgementDraft(deck, rev, finding.key, judgement);
  updateProgress();
}

function updateProgress(): void {
  if (!progressText) return;
  progressText.textContent = `${judgedFindingCount(findings, judgements)} of ${findings.length} findings judged`;
}

function findObservation(rows: readonly DecisionRow[], value: string): DecisionRow | undefined {
  return rows.find((row) => row.review_no === value || row.check_id === value);
}

function shell(statusText: string): HTMLElement {
  const root = document.createElement("main");
  root.className = "app-shell visual-shell";
  const header = document.createElement("header");
  header.className = "app-header";
  header.innerHTML = `
    <div>
      <p class="eyebrow">PPTX Design Review</p>
      <h1>Visual Review</h1>
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

function renderEmptyState(message: string): HTMLElement {
  const element = document.createElement("section");
  element.className = "artifact-panel";
  element.textContent = message;
  return element;
}

function escapeHtml(value: string): string {
  return value.replace(
    /[&<>"']/g,
    (char) =>
      ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#039;",
      })[char] ?? char,
  );
}
