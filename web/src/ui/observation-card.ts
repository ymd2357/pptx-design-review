import {
  parseFindingDispositions,
  serializeFindingDispositions,
  validateDecisionRow,
  type DecisionRow,
  type FindingDisposition,
} from "../data/decisions-tsv";
import {
  JUDGEMENT_REASONS,
  OBSERVATION_DECISIONS,
  reasonsForStatus,
  type DispositionStatus,
  type ObservationDecision,
} from "../data/enums";

export function renderObservationCard(
  row: DecisionRow,
  onChange: () => void,
  visualHref?: string,
): HTMLElement {
  const card = document.createElement("article");
  card.className = "observation-card";

  const header = document.createElement("div");
  header.className = "card-header";
  header.innerHTML = `
    <div>
      <p class="eyebrow">${escapeHtml(row.review_no)} / ${escapeHtml(row.priority)}</p>
      <h2>${escapeHtml(row.check_id)}</h2>
    </div>
    <div class="count-badge">${escapeHtml(row.latest_lint_count)}</div>
  `;

  const meta = document.createElement("div");
  meta.className = "card-meta";
  meta.innerHTML = `
    <span>判定</span>
    <strong>${escapeHtml(row.observation_decision || "not_recorded")}</strong>
  `;

  const decisionLabel = document.createElement("label");
  decisionLabel.className = "field";
  decisionLabel.innerHTML = `<span>観点判定 (observation_decision)</span>`;
  const decisionSelect = document.createElement("select");
  for (const decision of OBSERVATION_DECISIONS) {
    decisionSelect.append(new Option(decision, decision, false, row.observation_decision === decision));
  }
  decisionSelect.addEventListener("change", () => {
    row.observation_decision = decisionSelect.value as ObservationDecision;
    renderDynamicArea();
    onChange();
  });
  decisionLabel.append(decisionSelect);

  const dynamicArea = document.createElement("div");
  dynamicArea.className = "dynamic-area";

  const errorBox = document.createElement("p");
  errorBox.className = "inline-error";

  card.append(header, meta);
  if (visualHref) {
    const visualLink = document.createElement("a");
    visualLink.className = "secondary-link visual-review-link";
    visualLink.href = visualHref;
    visualLink.textContent = "視覚レビューへ →";
    card.append(visualLink);
  }
  card.append(decisionLabel, dynamicArea, errorBox);

  function renderDynamicArea(): void {
    dynamicArea.replaceChildren();
    if (row.observation_decision !== "remaining") {
      errorBox.textContent = "";
      return;
    }

    const dispositions = parseFindingDispositions(row.finding_dispositions);
    const list = document.createElement("div");
    list.className = "disposition-list";
    if (dispositions.length === 0) {
      dispositions.push({
        review_status: "fix_required",
        judgement_reason: JUDGEMENT_REASONS.fix_required[0],
        count: Number.parseInt(row.latest_lint_count, 10) || 1,
      });
      row.finding_dispositions = serializeFindingDispositions(dispositions);
    }

    dispositions.forEach((disposition, index) => {
      list.append(renderDispositionRow(disposition, index, dispositions, commitDispositions));
    });

    const addButton = document.createElement("button");
    addButton.type = "button";
    addButton.className = "secondary-button";
    addButton.textContent = "判定内訳を追加";
    addButton.addEventListener("click", () => {
      dispositions.push({
        review_status: "fix_required",
        judgement_reason: JUDGEMENT_REASONS.fix_required[0],
        count: 1,
      });
      commitDispositions();
      renderDynamicArea();
    });

    const rationale = document.createElement("label");
    rationale.className = "field";
    rationale.innerHTML = `<span>補足コメント (rationale)</span>`;
    const textarea = document.createElement("textarea");
    textarea.rows = 4;
    textarea.value = row.rationale;
    textarea.addEventListener("input", () => {
      row.rationale = textarea.value;
      onChange();
    });
    rationale.append(textarea);

    dynamicArea.append(list, addButton, rationale);
    updateErrors();

    function commitDispositions(): void {
      row.finding_dispositions = serializeFindingDispositions(dispositions);
      updateErrors();
      onChange();
    }
  }

  function updateErrors(): void {
    errorBox.textContent = validateDecisionRow(row).join(" ");
  }

  renderDynamicArea();
  updateErrors();
  return card;
}

function renderDispositionRow(
  disposition: FindingDisposition,
  index: number,
  dispositions: FindingDisposition[],
  onCommit: () => void,
): HTMLElement {
  const row = document.createElement("div");
  row.className = "disposition-row";

  const status = document.createElement("select");
  for (const value of Object.keys(JUDGEMENT_REASONS)) {
    status.append(new Option(value, value, false, disposition.review_status === value));
  }

  const reason = document.createElement("select");
  const renderReasons = () => {
    reason.replaceChildren();
    for (const value of reasonsForStatus(disposition.review_status)) {
      reason.append(new Option(value, value, false, disposition.judgement_reason === value));
    }
    if (!reasonsForStatus(disposition.review_status).includes(disposition.judgement_reason)) {
      disposition.judgement_reason = reason.value;
    }
  };

  const count = document.createElement("input");
  count.type = "number";
  count.min = "0";
  count.step = "1";
  count.inputMode = "numeric";
  count.value = String(disposition.count);

  const remove = document.createElement("button");
  remove.type = "button";
  remove.className = "icon-button";
  remove.textContent = "x";
  remove.title = "判定内訳を削除";
  remove.setAttribute("aria-label", "判定内訳を削除");

  status.addEventListener("change", () => {
    disposition.review_status = status.value as DispositionStatus;
    disposition.judgement_reason = reasonsForStatus(disposition.review_status)[0] ?? "";
    renderReasons();
    onCommit();
  });
  reason.addEventListener("change", () => {
    disposition.judgement_reason = reason.value;
    onCommit();
  });
  count.addEventListener("input", () => {
    disposition.count = Number.parseInt(count.value, 10) || 0;
    onCommit();
  });
  remove.addEventListener("click", () => {
    dispositions.splice(index, 1);
    onCommit();
    row.remove();
  });

  renderReasons();
  row.append(status, reason, count, remove);
  return row;
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
