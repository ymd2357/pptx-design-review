import "./styles.css";
import { requireAuth } from "./auth/auth-gate";
import {
  aggregateFindingDispositions,
  clearFindingJudgementDrafts,
  findingJudgementsPath,
  initializeFindingJudgements,
  judgedFindingCount,
  mergeLocalJudgementDrafts,
  parseFindingJudgementsJson,
  serializeFindingJudgementsJson,
  storeFindingJudgementDraft,
  type FindingJudgement,
  type FindingJudgementsFile,
} from "./data/finding-judgements";
import {
  parseDecisionTsv,
  serializeDecisionTsv,
  type DecisionRow,
} from "./data/decisions-tsv";
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
} from "./github/contents";
import { submitFeedback } from "./data/feedback-client";
import { sitePath } from "./site-path";
import {
  renderInlineFindingReview,
  type InlineFindingGroup,
  type InlineFindingReviewHandle,
} from "./ui/inline-finding-review";
import { renderSlideGallery, type SlideGalleryHandle } from "./ui/slide-gallery";

const appElement = document.querySelector<HTMLDivElement>("#app");
if (!appElement) throw new Error("Missing #app");
const app = appElement;

const params = new URLSearchParams(location.search);
const deck = params.get("deck") ?? "260329-seminar-curriculum-proposal";
const rev = params.get("rev") ?? "017";

let decisionRows: DecisionRow[] = [];
let groups: InlineFindingGroup[] = [];
let allFindings: LintFinding[] = [];
let judgements: FindingJudgementsFile = { deck, rev, judgements: {} };
let slideSize: SlideSizePt = { w: 1440, h: 810 };
let imageUrls: string[] = [];
const updatedBy = "local";

let progressText: HTMLElement | undefined;
let galleryHandle: SlideGalleryHandle | undefined;
let reviewHandle: InlineFindingReviewHandle | undefined;

void (async () => {
  await requireAuth(app);
  await renderVisualReview();
})();

async function renderVisualReview(): Promise<void> {
  app.replaceChildren(shell("視覚レビューを読み込み中..."));
  try {
    const [decisionFile, lintData, snapshot, loadedSlideSize, judgementFile] = await Promise.all([
      fetchDecisionTsv(deck, rev),
      loadSnapshotLint(deck, rev),
      fetchReviewSnapshot(deck, rev),
      loadSlideSizePt(),
      fetchJsonFile<unknown>(findingJudgementsPath(deck, rev)),
    ]);

    decisionRows = parseDecisionTsv(decisionFile.text).rows;
    allFindings = lintData.findings;
    slideSize = loadedSlideSize;
    imageUrls = snapshot?.imageUrls ?? [];
    judgements = mergeLocalJudgementDrafts(
      initializeFindingJudgements(
        deck,
        rev,
        lintData.findings,
        judgementFile ? parseFindingJudgementsJson(judgementFile.data, deck, rev) : undefined,
      ),
      deck,
      rev,
      lintData.findings,
    );

    groups = buildGroups(decisionRows, allFindings);
    if (groups.length === 0) {
      throw new Error("lint.json に finding がありません。");
    }
    renderLoaded();
  } catch (error) {
    app.replaceChildren(
      shell(error instanceof Error ? error.message : "視覚レビューの読み込みに失敗しました。"),
    );
  }
}

/**
 * decisions.tsv の観点順 (P0→P1→P2) に従って check ごとに finding を束ねる。
 * decisions に無いが lint に出た check は末尾に補完する。
 */
function buildGroups(rows: readonly DecisionRow[], findings: readonly LintFinding[]): InlineFindingGroup[] {
  const result: InlineFindingGroup[] = [];
  const seen = new Set<string>();
  for (const row of rows) {
    if (!row.check_id || seen.has(row.check_id)) continue;
    const matched = findingsForObservation(findings, row.check_id);
    if (matched.length === 0) continue;
    seen.add(row.check_id);
    result.push({ observation: row, findings: matched });
  }
  const leftovers = new Map<string, LintFinding[]>();
  for (const finding of findings) {
    if (seen.has(finding.check)) continue;
    const list = leftovers.get(finding.check) ?? [];
    list.push(finding);
    leftovers.set(finding.check, list);
  }
  for (const [check, matched] of leftovers) {
    result.push({ observation: syntheticRow(check, matched.length), findings: matched });
  }
  return result;
}

function syntheticRow(checkId: string, count: number): DecisionRow {
  return {
    review_no: "—",
    check_id: checkId,
    priority: "P3",
    latest_lint_count: String(count),
    observation_decision: "remaining",
    finding_dispositions: "",
    rationale: "",
    related_artifacts: "",
  } as DecisionRow;
}

function renderLoaded(): void {
  const root = shell("");

  const summary = document.createElement("section");
  summary.className = "visual-summary";
  const back = sitePath(`review/?deck=${encodeURIComponent(deck)}&rev=${encodeURIComponent(rev)}`);
  summary.innerHTML = `
    <div>
      <p class="eyebrow">${escapeHtml(deck)} / REV-${escapeHtml(rev)}</p>
      <h2>視覚レビュー (finding 単位)</h2>
    </div>
    <a class="secondary-link" href="${back}">戻る</a>
  `;
  root.append(summary);

  // --- sticky 上部: スライド画像 + 進捗 ---
  const sticky = document.createElement("div");
  sticky.className = "visual-sticky";

  progressText = document.createElement("p");
  progressText.className = "visual-progress";

  if (imageUrls.length > 0) {
    galleryHandle = renderSlideGallery({
      imageUrls,
      findings: allFindings,
      judgements,
      slideSizePt: slideSize,
      initialSlideNo: allFindings[0]?.slideNo,
      onSelectFinding: focusFinding,
    });
    sticky.append(galleryHandle.element, progressText);
    const firstWithBox = allFindings.find((f) => f.bboxPt);
    if (firstWithBox) galleryHandle.focus(firstWithBox);
  } else {
    galleryHandle = undefined;
    sticky.append(
      renderEmptyState("レビュースナップショットにスライド PNG が見つかりません。"),
      progressText,
    );
  }
  root.append(sticky);
  updateProgress();

  // --- スクロール領域: 全 check の inline 判定 ---
  reviewHandle = renderInlineFindingReview({
    groups,
    judgements,
    onChange: updateJudgement,
    onFocusFinding: focusFinding,
  });
  root.append(reviewHandle.element);

  // --- 固定下部: 送信 / ダウンロード ---
  root.append(renderActions());

  app.replaceChildren(root);
}

function focusFinding(finding: LintFinding): void {
  galleryHandle?.focus(finding);
  reviewHandle?.focusFinding(finding.key);
}

function updateJudgement(key: string, next: FindingJudgement): void {
  const judgement: FindingJudgement = {
    ...next,
    updated_at: new Date().toISOString(),
    updated_by: updatedBy,
  };
  judgements.judgements[key] = judgement;
  storeFindingJudgementDraft(deck, rev, key, judgement);
  updateProgress();
  galleryHandle?.refresh();
  reviewHandle?.refresh();
}

function updateProgress(): void {
  if (!progressText) return;
  const judged = judgedFindingCount(allFindings, judgements);
  progressText.textContent = `判定済 ${judged} / ${allFindings.length} finding`;
}

function renderActions(): HTMLElement {
  const actions = document.createElement("div");
  actions.className = "sticky-actions";

  const download = document.createElement("button");
  download.type = "button";
  download.className = "secondary-button";
  download.textContent = "ダウンロード";
  download.title = "判定を TSV / JSON でローカル保存します。";
  download.addEventListener("click", downloadReviewFiles);

  const submit = document.createElement("button");
  submit.type = "button";
  submit.className = "primary-button";
  submit.textContent = "送信";
  submit.title = "判定を age 公開鍵で暗号化して共有 KV に送信します。";
  submit.addEventListener("click", () => {
    submit.disabled = true;
    void handleSubmit().finally(() => {
      submit.disabled = false;
    });
  });

  actions.append(download, submit);
  return actions;
}

function applyAggregates(): void {
  for (const row of decisionRows) {
    if (!findingsForObservation(allFindings, row.check_id).length) continue;
    row.finding_dispositions = aggregateFindingDispositions(allFindings, judgements, row.check_id);
  }
}

async function handleSubmit(): Promise<void> {
  applyAggregates();
  try {
    const { key } = await submitFeedback({
      deck,
      rev,
      decisions: decisionRows,
      findingJudgements: judgements,
    });
    clearFindingJudgementDrafts(deck, rev);
    showBanner(`送信完了 (${key})。PC で scripts/fetch-reviews.py --apply を実行すると判定が取り込まれます。`, "success");
  } catch (error) {
    showBanner(error instanceof Error ? error.message : "判定の送信に失敗しました。", "error");
  }
}

function downloadReviewFiles(): void {
  applyAggregates();
  triggerDownload(`rev-${rev}-decisions.tsv`, serializeDecisionTsv(decisionRows), "text/tab-separated-values");
  triggerDownload(
    `rev-${rev}-finding-judgements.json`,
    serializeFindingJudgementsJson(judgements),
    "application/json",
  );
}

function triggerDownload(filename: string, content: string, mime: string): void {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function showBanner(text: string, kind: "success" | "error"): void {
  app.querySelectorAll(".commit-banner").forEach((b) => b.remove());
  const banner = document.createElement("section");
  banner.className = `commit-banner ${kind}`;
  banner.textContent = text;
  app.querySelector(".app-shell")?.prepend(banner);
  window.setTimeout(() => banner.remove(), 8000);
}

function shell(statusText: string): HTMLElement {
  const root = document.createElement("main");
  root.className = "app-shell visual-shell";
  const header = document.createElement("header");
  header.className = "app-header";
  header.innerHTML = `
    <div>
      <p class="eyebrow">PPTX デザインレビュー</p>
      <h1>視覚レビュー</h1>
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
