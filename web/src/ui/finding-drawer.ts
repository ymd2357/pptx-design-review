import type { FindingJudgement, FindingJudgementsFile } from "../data/finding-judgements";
import type { LintFinding } from "../data/lint-json";
import {
  labelForJudgementReason,
  labelForReviewStatus,
  reasonsForStatus,
  REVIEW_STATUSES,
  type ReviewStatus,
} from "../data/enums";

export type FindingDrawerOptions = {
  findings: readonly LintFinding[];
  initialKey: string;
  judgements: FindingJudgementsFile;
  onChange: (key: string, judgement: FindingJudgement) => void;
  onClose: () => void;
  onComplete?: () => void;
};

export type FindingDrawerHandle = {
  element: HTMLElement;
};

export function renderFindingDrawer(options: FindingDrawerOptions): HTMLElement {
  const backdrop = document.createElement("div");
  backdrop.className = "drawer-backdrop";

  const drawer = document.createElement("aside");
  drawer.className = "finding-drawer";
  drawer.setAttribute("aria-label", "finding 判定");

  let currentIndex = Math.max(
    0,
    options.findings.findIndex((f) => f.key === options.initialKey),
  );
  if (options.findings.length === 0) {
    options.onClose();
    return backdrop;
  }

  const close = document.createElement("button");
  close.type = "button";
  close.className = "icon-button drawer-close";
  close.textContent = "x";
  close.title = "閉じる";
  close.setAttribute("aria-label", "finding 詳細を閉じる");
  close.addEventListener("click", options.onClose);

  const counter = document.createElement("p");
  counter.className = "drawer-counter eyebrow";

  const title = document.createElement("div");
  title.className = "drawer-title";

  const message = document.createElement("p");
  message.className = "finding-message";

  const details = document.createElement("dl");
  details.className = "finding-details";

  const statusField = document.createElement("label");
  statusField.className = "field";
  statusField.innerHTML = "<span>レビュー状態</span>";
  const status = document.createElement("select");
  for (const value of REVIEW_STATUSES) {
    status.append(new Option(labelForReviewStatus(value), value));
  }
  statusField.append(status);

  const reasonField = document.createElement("label");
  reasonField.className = "field";
  reasonField.innerHTML = "<span>判定理由</span>";
  const reason = document.createElement("select");
  reasonField.append(reason);

  const rationaleField = document.createElement("label");
  rationaleField.className = "field";
  rationaleField.innerHTML = "<span>補足コメント</span>";
  const rationale = document.createElement("textarea");
  rationale.rows = 3;
  rationaleField.append(rationale);

  const footer = document.createElement("div");
  footer.className = "drawer-footer";
  const prevBtn = document.createElement("button");
  prevBtn.type = "button";
  prevBtn.className = "secondary-button";
  prevBtn.textContent = "← 前へ";
  const nextBtn = document.createElement("button");
  nextBtn.type = "button";
  nextBtn.className = "secondary-button";
  nextBtn.textContent = "次へ →";
  const nextUnreviewedBtn = document.createElement("button");
  nextUnreviewedBtn.type = "button";
  nextUnreviewedBtn.className = "primary-button";
  nextUnreviewedBtn.textContent = "✓ 保存して次の未判定へ";
  footer.append(prevBtn, nextBtn, nextUnreviewedBtn);

  drawer.append(close, counter, title, message, details, statusField, reasonField, rationaleField, footer);

  prevBtn.addEventListener("click", () => navigateTo(currentIndex - 1));
  nextBtn.addEventListener("click", () => navigateTo(currentIndex + 1));
  nextUnreviewedBtn.addEventListener("click", () => {
    const nextIdx = findNextUnreviewedIndex(currentIndex);
    if (nextIdx === -1) {
      options.onComplete?.();
      options.onClose();
    } else {
      navigateTo(nextIdx);
    }
  });

  status.addEventListener("change", () => {
    const newStatus = status.value as ReviewStatus;
    const reasons = reasonsForStatus(newStatus);
    const judgement: FindingJudgement = {
      review_status: newStatus,
      judgement_reason: reasons[0] ?? null,
      rationale: rationale.value || undefined,
    };
    updateJudgement(judgement);
    updateReasonOptions(judgement);
  });
  reason.addEventListener("change", () => {
    const judgement: FindingJudgement = {
      review_status: status.value as ReviewStatus,
      judgement_reason: reason.value || null,
      rationale: rationale.value || undefined,
    };
    updateJudgement(judgement);
  });
  rationale.addEventListener("input", () => {
    const judgement: FindingJudgement = {
      review_status: status.value as ReviewStatus,
      judgement_reason: reason.value || null,
      rationale: rationale.value || undefined,
    };
    updateJudgement(judgement);
  });

  backdrop.append(drawer);
  backdrop.addEventListener("click", (event) => {
    if (event.target === backdrop) options.onClose();
  });

  applyCurrent();
  return backdrop;

  function navigateTo(targetIndex: number): void {
    if (targetIndex < 0 || targetIndex >= options.findings.length) return;
    currentIndex = targetIndex;
    applyCurrent();
  }

  function applyCurrent(): void {
    const finding = options.findings[currentIndex]!;
    const judgement = options.judgements.judgements[finding.key] ?? {
      review_status: "unreviewed" as ReviewStatus,
      judgement_reason: null,
    };

    counter.textContent = `${currentIndex + 1} / ${options.findings.length}`;
    title.innerHTML = `
      <p class="eyebrow">${escapeHtml(finding.check)} / スライド ${finding.slideNo}</p>
      <h2>${escapeHtml(finding.shapeName || finding.key)}</h2>
    `;
    message.textContent = finding.message;

    details.replaceChildren();
    appendDetail(details, "severity", finding.severity);
    appendDetail(details, "shape_id", String(finding.shapeId ?? ""));
    appendDetail(details, "measured_value", stringifyValue(finding.measuredValue));
    appendDetail(details, "bbox_pt", finding.bboxPt?.join(", ") ?? "");
    appendDetail(details, "group_key", finding.key);

    status.value = judgement.review_status;
    rationale.value = judgement.rationale ?? "";
    updateReasonOptions(judgement);

    prevBtn.disabled = currentIndex === 0;
    nextBtn.disabled = currentIndex === options.findings.length - 1;
  }

  function updateReasonOptions(judgement: { review_status: string; judgement_reason?: string | null }): void {
    reason.replaceChildren();
    const reasons = reasonsForStatus(judgement.review_status);
    reason.disabled = reasons.length === 0;
    if (reasons.length === 0) {
      reason.append(new Option("(なし)", "", true, true));
      return;
    }
    for (const value of reasons) {
      reason.append(new Option(labelForJudgementReason(value), value, false, judgement.judgement_reason === value));
    }
    if (!judgement.judgement_reason || !reasons.includes(judgement.judgement_reason)) {
      reason.value = reasons[0] ?? "";
    }
  }

  function updateJudgement(judgement: FindingJudgement): void {
    const finding = options.findings[currentIndex]!;
    options.judgements.judgements[finding.key] = judgement;
    options.onChange(finding.key, judgement);
  }

  function findNextUnreviewedIndex(fromIndex: number): number {
    for (let i = fromIndex + 1; i < options.findings.length; i += 1) {
      const finding = options.findings[i]!;
      if (!isJudged(options.judgements.judgements[finding.key])) return i;
    }
    for (let i = 0; i <= fromIndex; i += 1) {
      const finding = options.findings[i]!;
      if (!isJudged(options.judgements.judgements[finding.key])) return i;
    }
    return -1;
  }
}

function isJudged(judgement: { review_status?: string; judgement_reason?: string | null } | undefined): boolean {
  if (!judgement) return false;
  if (!judgement.review_status || judgement.review_status === "unreviewed") return false;
  return Boolean(judgement.judgement_reason);
}

function appendDetail(target: HTMLDListElement, label: string, value: string): void {
  if (!value) return;
  const dt = document.createElement("dt");
  dt.textContent = label;
  const dd = document.createElement("dd");
  dd.textContent = value;
  target.append(dt, dd);
}

function stringifyValue(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  return JSON.stringify(value);
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
