import "./styles.css";
import { requireAuth } from "./auth/auth-gate";
import { submitFeedback } from "./data/feedback-client";
import { sitePath } from "./site-path";

const appElement = document.querySelector<HTMLDivElement>("#app");
if (!appElement) throw new Error("Missing #app");
const app = appElement;

const params = new URLSearchParams(location.search);
const deck = params.get("deck") ?? "260329-seminar-curriculum-proposal";
const rev = params.get("rev") ?? "017";

type SlideDecision = {
  decision: "adopt" | "reject" | null;
  memo: string;
};

type SlideEntry = {
  slideNo: number;
  beforeUrl: string;
  afterUrl: string;
  diffUrl: string;
  changed: boolean;
};

type CompareMeta = {
  changed_slides?: number[];
  slide_diffs?: Array<{ slide_no?: number }>;
};

const draftKey = `pptx-review:compare:${deck}:${rev}`;
let slides: SlideEntry[] = [];
const decisions = new Map<number, SlideDecision>();
let progressText: HTMLElement | undefined;
let showUnchanged = false;

void (async () => {
  await requireAuth(app);
  await renderCompare();
})();

async function renderCompare(): Promise<void> {
  app.replaceChildren(shell("比較画像を読み込み中..."));
  try {
    slides = await discoverSlides(deck, rev);
    if (slides.length === 0) {
      app.replaceChildren(shell("before/after PNG が見つかりません。"));
      return;
    }
    loadDraft();
    renderLoaded();
  } catch (error) {
    app.replaceChildren(shell(error instanceof Error ? error.message : "比較画像の読み込みに失敗しました。"));
  }
}

function renderLoaded(): void {
  const root = shell("");

  const summary = document.createElement("section");
  summary.className = "visual-summary";
  summary.innerHTML = `
    <div>
      <p class="eyebrow">${escapeHtml(deck)} / REV-${escapeHtml(rev)}</p>
      <h2>修正後をスライドごとに採用 / 不採用</h2>
    </div>
    <a class="secondary-link" href="${sitePath("")}">一覧へ</a>
  `;

  progressText = document.createElement("p");
  progressText.className = "visual-progress";
  updateProgress();

  const filterRow = document.createElement("div");
  filterRow.className = "compare-filter-row";
  const filterLabel = document.createElement("label");
  filterLabel.className = "compare-filter-label";
  const filterCheckbox = document.createElement("input");
  filterCheckbox.type = "checkbox";
  filterCheckbox.checked = showUnchanged;
  filterCheckbox.addEventListener("change", () => {
    showUnchanged = filterCheckbox.checked;
    renderList();
  });
  filterLabel.append(filterCheckbox, document.createTextNode(" 変更なしスライドも表示"));
  filterRow.append(filterLabel);

  const list = document.createElement("section");
  list.className = "compare-list";

  function renderList(): void {
    list.replaceChildren();
    const visible = slides.filter((s) => showUnchanged || s.changed);
    if (visible.length === 0) {
      const empty = document.createElement("p");
      empty.className = "status-text";
      empty.textContent = "変更ありのスライドはありません。";
      list.append(empty);
      return;
    }
    for (const slide of visible) {
      list.append(renderSlideCard(slide));
    }
  }
  renderList();

  const actions = document.createElement("div");
  actions.className = "sticky-actions";
  const submit = document.createElement("button");
  submit.type = "button";
  submit.className = "primary-button";
  submit.textContent = "送信";
  submit.title = "採用 / 不採用の判定を共有 KV に送ります。";
  submit.addEventListener("click", () => {
    submit.disabled = true;
    void handleSubmit().finally(() => {
      submit.disabled = false;
    });
  });
  actions.append(submit);

  root.append(summary, progressText, filterRow, list, actions);
  app.replaceChildren(root);
}

function renderSlideCard(slide: SlideEntry): HTMLElement {
  const card = document.createElement("article");
  card.className = "compare-card";
  if (!slide.changed) card.classList.add("unchanged");

  const header = document.createElement("header");
  header.className = "compare-card-header";
  const changedLabel = slide.changed ? "" : ' <span class="compare-unchanged-tag">変更なし</span>';
  header.innerHTML = `<h3>スライド ${slide.slideNo}${changedLabel}</h3>`;

  const view = document.createElement("div");
  view.className = "compare-view";
  const img = document.createElement("img");
  img.alt = `スライド ${slide.slideNo}`;
  img.loading = "lazy";
  const missingNote = document.createElement("p");
  missingNote.className = "compare-missing-image";
  missingNote.hidden = true;
  view.append(img, missingNote);

  // Unchanged slides only render the `before` image (= after / diff are
  // intentionally not produced for them; the snapshot omits those files).
  if (!slide.changed) {
    img.src = slide.beforeUrl;
    img.onerror = () => {
      img.onerror = null;
      img.hidden = true;
      missingNote.hidden = false;
      missingNote.textContent = "画像が見つかりません。";
    };
    card.append(header, view);
    return card;
  }

  const tabs = document.createElement("div");
  tabs.className = "compare-tabs";

  let currentMode: "before" | "after" | "diff" = "after";
  const buttons: Record<string, HTMLButtonElement> = {};
  for (const mode of ["before", "after", "diff"] as const) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "compare-tab";
    btn.textContent = mode === "before" ? "修正前" : mode === "after" ? "修正後" : "差分";
    btn.addEventListener("click", () => {
      currentMode = mode;
      updateImage();
      updateTabs();
    });
    buttons[mode] = btn;
    tabs.append(btn);
  }

  function updateImage(): void {
    const url =
      currentMode === "before"
        ? slide.beforeUrl
        : currentMode === "after"
          ? slide.afterUrl
          : slide.diffUrl;
    img.hidden = false;
    missingNote.hidden = true;
    img.onerror = () => {
      img.onerror = null;
      img.hidden = true;
      missingNote.hidden = false;
      missingNote.textContent =
        currentMode === "diff"
          ? "差分画像は生成されていません。"
          : currentMode === "after"
            ? "修正後画像が見つかりません。"
            : "修正前画像が見つかりません。";
    };
    img.src = url;
  }
  function updateTabs(): void {
    for (const mode of Object.keys(buttons)) {
      buttons[mode]!.classList.toggle("active", mode === currentMode);
    }
  }
  updateImage();
  updateTabs();

  const form = document.createElement("div");
  form.className = "compare-form";
  const current: SlideDecision = decisions.get(slide.slideNo) ?? { decision: null, memo: "" };

  const adoptBtn = document.createElement("button");
  adoptBtn.type = "button";
  adoptBtn.className = "compare-decision-button adopt";
  adoptBtn.textContent = "採用";

  const rejectBtn = document.createElement("button");
  rejectBtn.type = "button";
  rejectBtn.className = "compare-decision-button reject";
  rejectBtn.textContent = "不採用";

  const memo = document.createElement("textarea");
  memo.rows = 2;
  memo.className = "compare-memo";
  memo.placeholder = "メモ (任意): 不採用の理由などを書く";
  memo.value = current.memo;
  memo.addEventListener("input", () => {
    const d = decisions.get(slide.slideNo) ?? { decision: null, memo: "" };
    d.memo = memo.value;
    decisions.set(slide.slideNo, d);
    persistDraft();
  });

  function updateDecisionUi(): void {
    const d = decisions.get(slide.slideNo);
    adoptBtn.classList.toggle("active", d?.decision === "adopt");
    rejectBtn.classList.toggle("active", d?.decision === "reject");
    card.classList.toggle("adopted", d?.decision === "adopt");
    card.classList.toggle("rejected", d?.decision === "reject");
  }

  adoptBtn.addEventListener("click", () => {
    decisions.set(slide.slideNo, { decision: "adopt", memo: memo.value });
    persistDraft();
    updateDecisionUi();
    updateProgress();
  });
  rejectBtn.addEventListener("click", () => {
    decisions.set(slide.slideNo, { decision: "reject", memo: memo.value });
    persistDraft();
    updateDecisionUi();
    updateProgress();
  });

  const buttonRow = document.createElement("div");
  buttonRow.className = "compare-button-row";
  buttonRow.append(adoptBtn, rejectBtn);

  form.append(buttonRow, memo);
  card.append(header, tabs, view, form);
  if (current.decision) updateDecisionUi();
  return card;
}

function updateProgress(): void {
  if (!progressText) return;
  const targets = slides.filter((s) => s.changed);
  let adopt = 0;
  let reject = 0;
  for (const slide of targets) {
    const d = decisions.get(slide.slideNo);
    if (d?.decision === "adopt") adopt += 1;
    if (d?.decision === "reject") reject += 1;
  }
  const remaining = targets.length - adopt - reject;
  progressText.textContent = `変更あり ${targets.length} 枚 / 判定済 ${adopt + reject} (採用 ${adopt} / 不採用 ${reject} / 未判定 ${remaining})`;
}

async function handleSubmit(): Promise<void> {
  const undecided: number[] = [];
  for (const slide of slides) {
    if (!slide.changed) continue;
    if (!decisions.get(slide.slideNo)?.decision) undecided.push(slide.slideNo);
  }
  if (undecided.length > 0) {
    const ok = window.confirm(
      `変更ありスライドのうち未判定が ${undecided.length} 件あります (スライド ${undecided.join(", ")})。\nこのまま送信しますか？`,
    );
    if (!ok) return;
  }
  try {
    const payload = {
      mode: "compare" as const,
      slides: slides
        .filter((slide) => slide.changed)
        .map((slide) => ({
          slideNo: slide.slideNo,
          decision: decisions.get(slide.slideNo)?.decision ?? null,
          memo: decisions.get(slide.slideNo)?.memo ?? "",
        })),
    };
    const { key } = await submitFeedback({
      deck,
      rev,
      decisions: payload,
      findingJudgements: {},
    });
    sessionStorage.removeItem(draftKey);
    const banner = document.createElement("section");
    banner.className = "commit-banner success";
    banner.textContent = `送信完了 (${key})。PC で scripts/fetch-reviews.py --apply を実行してください。`;
    app.querySelector(".app-shell")?.prepend(banner);
    window.setTimeout(() => banner.remove(), 8000);
  } catch (error) {
    window.alert(error instanceof Error ? error.message : "送信に失敗しました。");
  }
}

async function discoverSlides(deckId: string, revId: string): Promise<SlideEntry[]> {
  const revBase = sitePath(
    `tmp/review-snapshot/${encodeURIComponent(deckId)}/rev-${encodeURIComponent(revId)}`,
  );
  const basePath = `${revBase}/images`;
  const meta = await fetchMeta(`${revBase}/meta.json`);
  const changedSet = new Set(meta?.changed_slides ?? []);
  const metaSlideNos = meta?.slide_diffs
    ?.map((slide) => slide.slide_no)
    .filter(
      (slideNo): slideNo is number =>
        typeof slideNo === "number" && Number.isInteger(slideNo) && slideNo > 0,
    )
    .sort((a, b) => a - b);
  if (metaSlideNos?.length) {
    return metaSlideNos.map((slideNo) => buildSlideEntry(basePath, slideNo, changedSet.has(slideNo)));
  }

  const results: SlideEntry[] = [];
  for (let i = 1; i <= 99; i += 1) {
    const beforeUrl = slideImageUrl(basePath, "before", i);
    const beforeExists = await assetExists(beforeUrl);
    if (!beforeExists) {
      if (results.length > 0) break;
      continue;
    }
    // Always reference the real URL; if the file is absent the slide card's
    // <img> handler shows an explicit "image not generated" notice rather
    // than silently falling back to the before image.
    results.push({
      slideNo: i,
      beforeUrl,
      afterUrl: slideImageUrl(basePath, "after", i),
      diffUrl: slideImageUrl(basePath, "diff", i),
      changed: meta ? changedSet.has(i) : true,
    });
  }
  return results;
}

function buildSlideEntry(basePath: string, slideNo: number, changed: boolean): SlideEntry {
  return {
    slideNo,
    beforeUrl: slideImageUrl(basePath, "before", slideNo),
    afterUrl: slideImageUrl(basePath, "after", slideNo),
    diffUrl: slideImageUrl(basePath, "diff", slideNo),
    changed,
  };
}

function slideImageUrl(basePath: string, kind: "before" | "after" | "diff", slideNo: number): string {
  const num = String(slideNo).padStart(2, "0");
  return `${basePath}/${kind}/slide-${num}.png`;
}

async function fetchMeta(url: string): Promise<CompareMeta | undefined> {
  try {
    const res = await fetch(url);
    if (!res.ok) return undefined;
    return (await res.json()) as CompareMeta;
  } catch {
    return undefined;
  }
}

async function assetExists(url: string): Promise<boolean> {
  try {
    const res = await fetch(url, { method: "HEAD" });
    return res.ok;
  } catch {
    return false;
  }
}

function loadDraft(): void {
  const raw = sessionStorage.getItem(draftKey);
  if (!raw) return;
  try {
    const parsed = JSON.parse(raw) as Record<string, SlideDecision>;
    for (const [k, v] of Object.entries(parsed)) {
      decisions.set(Number.parseInt(k, 10), v);
    }
  } catch {
    sessionStorage.removeItem(draftKey);
  }
}

function persistDraft(): void {
  const obj: Record<string, SlideDecision> = {};
  for (const [k, v] of decisions.entries()) obj[String(k)] = v;
  sessionStorage.setItem(draftKey, JSON.stringify(obj));
}

function shell(statusText: string): HTMLElement {
  const root = document.createElement("main");
  root.className = "app-shell";
  const header = document.createElement("header");
  header.className = "app-header";
  header.innerHTML = `
    <div>
      <p class="eyebrow">PPTX デザインレビュー</p>
      <h1>修正比較</h1>
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
