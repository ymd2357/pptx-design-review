import {
  parseFindingDispositions,
  validateDecisionRow,
  type DecisionRow,
} from "../data/decisions-tsv";
import {
  labelForJudgementReason,
  labelForObservationDecision,
  labelForReviewStatus,
  OBSERVATION_DECISIONS,
  type ObservationDecision,
} from "../data/enums";

export function renderObservationCard(
  row: DecisionRow,
  onChange: () => void,
  visualHref?: string,
): HTMLElement {
  const card = document.createElement("article");
  card.className = "observation-card";
  card.id = `obs-${row.review_no}`;

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
    <strong>${escapeHtml(labelForObservationDecision(row.observation_decision || "not_recorded"))}</strong>
  `;

  const decisionLabel = document.createElement("label");
  decisionLabel.className = "field";
  decisionLabel.innerHTML = `<span>観点判定</span>`;
  const decisionSelect = document.createElement("select");
  for (const decision of OBSERVATION_DECISIONS) {
    decisionSelect.append(new Option(labelForObservationDecision(decision), decision, false, row.observation_decision === decision));
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
    visualLink.addEventListener("click", () => {
      sessionStorage.setItem("pptx-review:scroll-anchor", row.review_no);
    });
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
    const summary = document.createElement("div");
    summary.className = "disposition-summary";
    if (dispositions.length === 0) {
      const empty = document.createElement("p");
      empty.className = "status-text inline";
      empty.textContent = "視覚レビューで finding 単位の判定を行うと、ここに集計が表示されます。";
      summary.append(empty);
    } else {
      const list = document.createElement("ul");
      list.className = "disposition-readonly-list";
      for (const disposition of dispositions) {
        const item = document.createElement("li");
        const statusLabel = labelForReviewStatus(disposition.review_status);
        const reasonLabel = labelForJudgementReason(disposition.judgement_reason);
        item.innerHTML = `
          <span class="disposition-pair">${escapeHtml(statusLabel)} / ${escapeHtml(reasonLabel)}</span>
          <strong class="disposition-count">${disposition.count}</strong>
        `;
        list.append(item);
      }
      summary.append(list);
    }

    const rationale = document.createElement("label");
    rationale.className = "field";
    rationale.innerHTML = `<span>補足コメント</span>`;
    const textarea = document.createElement("textarea");
    textarea.rows = 3;
    textarea.value = row.rationale;
    textarea.addEventListener("input", () => {
      row.rationale = textarea.value;
      onChange();
    });
    rationale.append(textarea);

    dynamicArea.append(summary, rationale);
    updateErrors();
  }

  function updateErrors(): void {
    errorBox.textContent = validateDecisionRow(row).join(" ");
  }

  renderDynamicArea();
  updateErrors();
  return card;
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
