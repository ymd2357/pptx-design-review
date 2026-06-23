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
  /** ユーザーが行を選んだ (= アクティブを切替えた)。スライド側を同期するのに使う。 */
  onActivate: (finding: LintFinding) => void;
};

export type InlineFindingReviewHandle = {
  element: HTMLElement;
  refresh(): void;
  /** 外部 (スライドの青枠タップ等) からアクティブ行を指定。展開してスクロール表示する。 */
  activate(key: string, options?: { scroll?: boolean }): void;
};

type RowHandle = {
  finding: LintFinding;
  element: HTMLElement;
  setOpen(open: boolean): void;
  syncBadge(): void;
};

type GroupHandle = {
  count: HTMLElement;
  findings: readonly LintFinding[];
};

export function renderInlineFindingReview(
  options: InlineFindingReviewOptions,
): InlineFindingReviewHandle {
  const section = document.createElement("section");
  section.className = "vfr";

  const groupHandles: GroupHandle[] = [];
  const rowByKey = new Map<string, RowHandle>();
  let activeKey: string | null = null;

  for (const group of options.groups) {
    const { element, handle, rows } = renderGroup(group);
    section.append(element);
    groupHandles.push(handle);
    for (const row of rows) rowByKey.set(row.finding.key, row);
  }

  refresh();
  return { element: section, refresh, activate };

  function refresh(): void {
    for (const row of rowByKey.values()) row.syncBadge();
    for (const group of groupHandles) group.count.textContent = groupCountLabel(group);
  }

  /** 行クリック由来: アクティブ切替 + スライド同期 (onActivate)。 */
  function selectFromRow(finding: LintFinding): void {
    setActive(finding.key);
    options.onActivate(finding);
  }

  /** 外部由来 (スライドの枠タップ): アクティブ切替 + スクロール表示のみ。 */
  function activate(key: string, opts?: { scroll?: boolean }): void {
    setActive(key);
    if (opts?.scroll !== false) {
      rowByKey.get(key)?.element.scrollIntoView({ block: "nearest", behavior: "smooth" });
    }
  }

  function setActive(key: string): void {
    if (activeKey === key) return;
    if (activeKey) rowByKey.get(activeKey)?.setOpen(false);
    activeKey = key;
    rowByKey.get(key)?.setOpen(true);
  }

  function groupCountLabel(group: GroupHandle): string {
    const judged = judgedFindingCount(group.findings, options.judgements);
    return `判定 ${judged} / ${group.findings.length}`;
  }

  function renderGroup(
    group: InlineFindingGroup,
  ): { element: HTMLElement; handle: GroupHandle; rows: RowHandle[] } {
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
    return { element: wrapper, handle: { count, findings: group.findings }, rows };
  }

  function renderRow(finding: LintFinding): RowHandle {
    const item = document.createElement("div");
    item.className = "vfr-row";

    // --- 常時表示: コンパクトな見出し行 (タップで展開) ---
    const head = document.createElement("button");
    head.type = "button";
    head.className = "vfr-row-head";
    head.addEventListener("click", () => selectFromRow(finding));

    const slide = document.createElement("span");
    slide.className = "vfr-row-slide";
    slide.textContent = `S${finding.slideNo}`;
    const shape = document.createElement("span");
    shape.className = "vfr-row-shape";
    shape.textContent = finding.shapeName || finding.key;
    const badge = document.createElement("span");
    badge.className = "vfr-row-badge";
    head.append(slide, shape, badge);

    // --- 展開時のみ表示: メッセージ + 判定コントロール ---
    const body = document.createElement("div");
    body.className = "vfr-row-body";

    const message = document.createElement("p");
    message.className = "vfr-row-message";
    message.textContent = finding.message;

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

    body.append(message, controls);
    item.append(head, body);

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
      setOpen(open: boolean): void {
        item.classList.toggle("vfr-row--open", open);
      },
      syncBadge(): void {
        const judgement = options.judgements.judgements[finding.key];
        const judged = isJudged(judgement);
        item.classList.toggle("vfr-row--judged", judged);
        badge.classList.toggle("judged", judged);
        badge.textContent = judged && judgement ? labelForReviewStatus(judgement.review_status) : "未判定";
      },
    };

    function commit(judgement: FindingJudgement): void {
      options.judgements.judgements[finding.key] = judgement;
      handle.syncBadge();
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
