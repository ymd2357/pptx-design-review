import { labelForReviewStatus } from "../data/enums";
import type { FindingJudgementsFile } from "../data/finding-judgements";
import type { LintFinding } from "../data/lint-json";

export type FindingListOptions = {
  findings: readonly LintFinding[];
  judgements: FindingJudgementsFile;
  onSelect: (finding: LintFinding) => void;
};

export type FindingListHandle = {
  element: HTMLElement;
  refresh(): void;
};

export function renderFindingList(options: FindingListOptions): FindingListHandle {
  const section = document.createElement("section");
  section.className = "finding-list";

  const header = document.createElement("header");
  header.className = "finding-list-header";
  const heading = document.createElement("h3");
  heading.textContent = "全 finding 一覧";
  const note = document.createElement("p");
  note.className = "status-text inline";
  note.textContent = "タップで判定 drawer を開きます。bbox が無い finding もここから到達できます。";
  header.append(heading, note);
  section.append(header);

  const list = document.createElement("ol");
  list.className = "finding-row-list";
  section.append(list);

  render();
  return { element: section, refresh: render };

  function render(): void {
    list.replaceChildren();
    for (const finding of options.findings) {
      list.append(renderRow(finding));
    }
  }

  function renderRow(finding: LintFinding): HTMLElement {
    const judgement = options.judgements.judgements[finding.key];
    const judged = isJudged(judgement);

    const item = document.createElement("li");
    item.className = "finding-row";
    if (judged) item.classList.add("judged");

    const button = document.createElement("button");
    button.type = "button";
    button.className = "finding-row-button";
    button.addEventListener("click", () => options.onSelect(finding));

    const head = document.createElement("div");
    head.className = "finding-row-head";

    const left = document.createElement("div");
    left.className = "finding-row-left";
    const slide = document.createElement("span");
    slide.className = "finding-row-slide";
    slide.textContent = `スライド ${finding.slideNo}`;
    const shape = document.createElement("span");
    shape.className = "finding-row-shape";
    shape.textContent = finding.shapeName || finding.key;
    left.append(slide, shape);

    const badge = document.createElement("span");
    badge.className = "finding-row-badge";
    if (judged && judgement) {
      badge.classList.add("judged");
      badge.textContent = labelForReviewStatus(judgement.review_status);
    } else {
      badge.textContent = "未判定";
    }
    head.append(left, badge);

    const message = document.createElement("p");
    message.className = "finding-row-message";
    message.textContent = finding.message;

    button.append(head, message);
    item.append(button);
    return item;
  }
}

function isJudged(judgement: { review_status?: string; judgement_reason?: string | null } | undefined): boolean {
  if (!judgement) return false;
  if (!judgement.review_status || judgement.review_status === "unreviewed") return false;
  return Boolean(judgement.judgement_reason);
}
