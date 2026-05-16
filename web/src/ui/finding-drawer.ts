import type { FindingJudgement } from "../data/finding-judgements";
import type { LintFinding } from "../data/lint-json";
import { reasonsForStatus, REVIEW_STATUSES, type ReviewStatus } from "../data/enums";

export type FindingDrawerOptions = {
  finding: LintFinding;
  judgement: FindingJudgement;
  onClose: () => void;
  onChange: (judgement: FindingJudgement) => void;
};

export function renderFindingDrawer(options: FindingDrawerOptions): HTMLElement {
  const backdrop = document.createElement("div");
  backdrop.className = "drawer-backdrop";

  const drawer = document.createElement("aside");
  drawer.className = "finding-drawer";
  drawer.setAttribute("aria-label", "finding 判定");

  const close = document.createElement("button");
  close.type = "button";
  close.className = "icon-button drawer-close";
  close.textContent = "x";
  close.title = "閉じる";
  close.setAttribute("aria-label", "finding 詳細を閉じる");
  close.addEventListener("click", options.onClose);

  const title = document.createElement("div");
  title.className = "drawer-title";
  title.innerHTML = `
    <p class="eyebrow">${escapeHtml(options.finding.check)} / slide ${options.finding.slideNo}</p>
    <h2>${escapeHtml(options.finding.shapeName || options.finding.key)}</h2>
  `;

  const message = document.createElement("p");
  message.className = "finding-message";
  message.textContent = options.finding.message;

  const details = document.createElement("dl");
  details.className = "finding-details";
  appendDetail(details, "severity", options.finding.severity);
  appendDetail(details, "shape_id", String(options.finding.shapeId ?? ""));
  appendDetail(details, "measured_value", stringifyValue(options.finding.measuredValue));
  appendDetail(details, "bbox_pt", options.finding.bboxPt?.join(", ") ?? "");
  appendDetail(details, "group_key", options.finding.key);

  const statusField = document.createElement("label");
  statusField.className = "field";
  statusField.innerHTML = "<span>レビュー状態 (review_status)</span>";
  const status = document.createElement("select");
  for (const value of REVIEW_STATUSES) {
    status.append(
      new Option(value, value, false, options.judgement.review_status === value),
    );
  }
  statusField.append(status);

  const reasonField = document.createElement("label");
  reasonField.className = "field";
  reasonField.innerHTML = "<span>判定理由 (judgement_reason)</span>";
  const reason = document.createElement("select");
  reasonField.append(reason);

  const rationaleField = document.createElement("label");
  rationaleField.className = "field";
  rationaleField.innerHTML = "<span>補足コメント (rationale)</span>";
  const rationale = document.createElement("textarea");
  rationale.rows = 4;
  rationale.value = options.judgement.rationale ?? "";
  rationaleField.append(rationale);

  const current: FindingJudgement = { ...options.judgement };

  status.addEventListener("change", () => {
    current.review_status = status.value as ReviewStatus;
    current.judgement_reason = reasonsForStatus(current.review_status)[0] ?? null;
    updateReasonOptions();
    emitChange();
  });
  reason.addEventListener("change", () => {
    current.judgement_reason = reason.value || null;
    emitChange();
  });
  rationale.addEventListener("input", () => {
    current.rationale = rationale.value;
    emitChange();
  });

  updateReasonOptions();
  drawer.append(close, title, message, details, statusField, reasonField, rationaleField);
  backdrop.append(drawer);
  backdrop.addEventListener("click", (event) => {
    if (event.target === backdrop) options.onClose();
  });
  return backdrop;

  function updateReasonOptions(): void {
    reason.replaceChildren();
    const reasons = reasonsForStatus(current.review_status);
    reason.disabled = reasons.length === 0;
    if (reasons.length === 0) {
      reason.append(new Option("none", "", true, true));
      current.judgement_reason = null;
      return;
    }
    for (const value of reasons) {
      reason.append(new Option(value, value, false, current.judgement_reason === value));
    }
    if (!current.judgement_reason || !reasons.includes(current.judgement_reason)) {
      current.judgement_reason = reasons[0] ?? null;
      reason.value = current.judgement_reason ?? "";
    }
  }

  function emitChange(): void {
    options.onChange({ ...current });
  }
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
