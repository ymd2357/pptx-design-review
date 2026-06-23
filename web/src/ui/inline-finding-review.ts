import {
  labelForJudgementReason,
  labelForReviewStatus,
  reasonsForStatus,
  REVIEW_STATUSES,
  type ReviewStatus,
} from "../data/enums";
import {
  judgedFindingCount,
  type FindingJudgement,
  type FindingJudgementsFile,
} from "../data/finding-judgements";
import type { DecisionRow } from "../data/decisions-tsv";
import type { LintFinding } from "../data/lint-json";

export type InlineFindingGroup = {
  observation: DecisionRow;
  findings: readonly LintFinding[];
};

export type InlineFindingReviewOptions = {
  groups: readonly InlineFindingGroup[];
  judgements: FindingJudgementsFile;
  onChange: (key: string, judgement: FindingJudgement) => void;
  onFocusFinding: (finding: LintFinding) => void;
};

export type InlineFindingReviewHandle = {
  element: HTMLElement;
  refresh(): void;
  focusFinding(key: string): void;
};

type RowHandle = {
  finding: LintFinding;
  element: HTMLElement;
  syncJudged(): void;
};

type GroupHandle = {
  count: HTMLElement;
  findings: readonly LintFinding[];
  rows: RowHandle[];
};

export function renderInlineFindingReview(
  options: InlineFindingReviewOptions,
): InlineFindingReviewHandle {
  const section = document.createElement("section");
  section.className = "vfr";

  const groupHandles: GroupHandle[] = [];
  const rowByKey = new Map<string, RowHandle>();

  for (const group of options.groups) {
    const { element, handle } = renderGroup(group);
    section.append(element);
    groupHandles.push(handle);
    for (const row of handle.rows) rowByKey.set(row.finding.key, row);
  }

  refresh();
  return { element: section, refresh, focusFinding };

  function refresh(): void {
    for (const group of groupHandles) {
      for (const row of group.rows) row.syncJudged();
      group.count.textContent = groupCountLabel(group);
    }
  }

  function focusFinding(key: string): void {
    const row = rowByKey.get(key);
    if (!row) return;
    row.element.classList.add("vfr-row--active");
    window.setTimeout(() => row.element.classList.remove("vfr-row--active"), 1600);
  }

  function groupCountLabel(group: GroupHandle): string {
    const judged = judgedFindingCount(group.findings, options.judgements);
    return `判定 ${judged} / ${group.findings.length}`;
  }

  function renderGroup(group: InlineFindingGroup): { element: HTMLElement; handle: GroupHandle } {
    const wrapper = document.createElement("section");
    wrapper.className = "vfr-group";

    const header = document.createElement("div");
    header.className = "vfr-group-head";

    const title = document.createElement("div");
    title.className = "vfr-group-title";
    const priority = document.createElement("span");
    priority.className = `vfr-prio vfr-prio--${(group.observation.priority || "").toLowerCase()}`;
    priority.textContent = group.observation.priority || "";
    const name = document.createElement("span");
    name.className = "vfr-group-name";
    name.textContent = `${group.observation.review_no} ${group.observation.check_id}`;
    title.append(priority, name);

    const count = document.createElement("span");
    count.className = "vfr-group-count";

    header.append(title, count);
    wrapper.append(header);

    const rows: RowHandle[] = [];
    for (const finding of group.findings) {
      const row = renderRow(finding);
      rows.push(row);
      wrapper.append(row.element);
    }

    return { element: wrapper, handle: { count, findings: group.findings, rows } };
  }

  function renderRow(finding: LintFinding): RowHandle {
    const item = document.createElement("div");
    item.className = "vfr-row";

    // --- left: finding identity (click to sync gallery) ---
    const info = document.createElement("button");
    info.type = "button";
    info.className = "vfr-row-info";
    info.addEventListener("click", () => options.onFocusFinding(finding));

    const meta = document.createElement("span");
    meta.className = "vfr-row-meta";
    const slide = document.createElement("span");
    slide.className = "vfr-row-slide";
    slide.textContent = `S${finding.slideNo}`;
    const shape = document.createElement("span");
    shape.className = "vfr-row-shape";
    shape.textContent = finding.shapeName || finding.key;
    meta.append(slide, shape);

    const message = document.createElement("span");
    message.className = "vfr-row-message";
    message.textContent = finding.message;

    info.append(meta, message);

    // --- right: inline judgement controls ---
    const controls = document.createElement("div");
    controls.className = "vfr-row-controls";

    const status = document.createElement("select");
    status.className = "vfr-select vfr-status";
    status.setAttribute("aria-label", "レビュー状態");
    for (const value of REVIEW_STATUSES) {
      status.append(new Option(labelForReviewStatus(value), value));
    }

    const reason = document.createElement("select");
    reason.className = "vfr-select vfr-reason";
    reason.setAttribute("aria-label", "判定理由");

    controls.append(status, reason);
    item.append(info, controls);

    const current = options.judgements.judgements[finding.key] ?? {
      review_status: "unreviewed" as ReviewStatus,
      judgement_reason: null,
    };
    status.value = current.review_status;
    fillReasonOptions(reason, current);

    status.addEventListener("change", () => {
      const newStatus = status.value as ReviewStatus;
      const reasons = reasonsForStatus(newStatus);
      const judgement: FindingJudgement = {
        review_status: newStatus,
        judgement_reason: reasons[0] ?? null,
      };
      fillReasonOptions(reason, judgement);
      commit(judgement);
    });
    reason.addEventListener("change", () => {
      commit({
        review_status: status.value as ReviewStatus,
        judgement_reason: reason.value || null,
      });
    });

    const handle: RowHandle = {
      finding,
      element: item,
      syncJudged(): void {
        item.classList.toggle("vfr-row--judged", isJudged(options.judgements.judgements[finding.key]));
      },
    };

    function commit(judgement: FindingJudgement): void {
      options.judgements.judgements[finding.key] = judgement;
      handle.syncJudged();
      options.onChange(finding.key, judgement);
    }

    return handle;
  }
}

function fillReasonOptions(
  select: HTMLSelectElement,
  judgement: { review_status: string; judgement_reason?: string | null },
): void {
  select.replaceChildren();
  const reasons = reasonsForStatus(judgement.review_status);
  select.disabled = reasons.length === 0;
  if (reasons.length === 0) {
    select.append(new Option("(理由なし)", "", true, true));
    return;
  }
  for (const value of reasons) {
    select.append(
      new Option(labelForJudgementReason(value), value, false, judgement.judgement_reason === value),
    );
  }
  if (!judgement.judgement_reason || !reasons.includes(judgement.judgement_reason)) {
    select.value = reasons[0] ?? "";
  }
}

function isJudged(
  judgement: { review_status?: string; judgement_reason?: string | null } | undefined,
): boolean {
  if (!judgement) return false;
  if (!judgement.review_status || judgement.review_status === "unreviewed") return false;
  return Boolean(judgement.judgement_reason);
}
