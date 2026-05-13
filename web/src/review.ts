import "./styles.css";
import {
  clearStoredToken,
  getStoredToken,
  startDeviceFlow,
  type DeviceFlowState,
} from "./auth/device-flow";
import {
  parseDecisionTsv,
  serializeDecisionTsv,
  validateDecisionRow,
  type DecisionRow,
} from "./data/decisions-tsv";
import { fetchDecisionTsv } from "./github/contents";
import { renderObservationCard } from "./ui/observation-card";

const appElement = document.querySelector<HTMLDivElement>("#app");
if (!appElement) throw new Error("Missing #app");
const app = appElement;

const params = new URLSearchParams(location.search);
const deck = params.get("deck") ?? "260329-seminar-curriculum-proposal";
const rev = params.get("rev") ?? "017";

let rows: DecisionRow[] = [];
let sourceSha: string | undefined;

void renderReview();

async function renderReview(): Promise<void> {
  app.replaceChildren(shell("Loading decisions..."));
  try {
    const file = await fetchDecisionTsv(deck, rev);
    sourceSha = file.sha;
    const parsed = parseDecisionTsv(file.text);
    rows = parsed.rows;
    renderLoaded(file.source, parsed.errors);
  } catch (error) {
    app.replaceChildren(
      shell(error instanceof Error ? error.message : "Failed to load decisions."),
    );
  }
}

function renderLoaded(source: "github" | "local", parseErrors: string[]): void {
  const root = shell("");
  root.append(renderAuthPanel(() => void renderReview()));

  const summary = document.createElement("section");
  summary.className = "review-summary";
  const invalidCount = rows.filter((row) => validateDecisionRow(row).length > 0).length;
  summary.innerHTML = `
    <div>
      <p class="eyebrow">${source}${sourceSha ? ` / ${sourceSha.slice(0, 7)}` : ""}</p>
      <h2>${deck} / REV-${rev}</h2>
    </div>
    <p>${rows.length} observations / ${invalidCount} validation issues</p>
  `;

  const messages = document.createElement("div");
  messages.className = "message-list";
  for (const error of parseErrors) {
    const item = document.createElement("p");
    item.className = "inline-error";
    item.textContent = error;
    messages.append(item);
  }

  const cardList = document.createElement("section");
  cardList.className = "card-list";
  const rerenderSummary = () => {
    const count = rows.filter((row) => validateDecisionRow(row).length > 0).length;
    summary.querySelector("p:last-child")!.textContent =
      `${rows.length} observations / ${count} validation issues`;
  };
  rows.forEach((row) => {
    cardList.append(renderObservationCard(row, rerenderSummary));
  });

  const actions = document.createElement("div");
  actions.className = "sticky-actions";
  const back = document.createElement("a");
  back.className = "secondary-link";
  back.href = "/";
  back.textContent = "Hub";
  const download = document.createElement("button");
  download.type = "button";
  download.className = "primary-button";
  download.textContent = "Download TSV";
  download.addEventListener("click", downloadTsv);
  actions.append(back, download);

  root.append(summary, messages, cardList, actions);
  app.replaceChildren(root);
}

function shell(statusText: string): HTMLElement {
  const root = document.createElement("main");
  root.className = "app-shell";
  const header = document.createElement("header");
  header.className = "app-header";
  header.innerHTML = `
    <div>
      <p class="eyebrow">PPTX Design Review</p>
      <h1>REV-${rev}</h1>
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

function renderAuthPanel(onAuthChange: () => void): HTMLElement {
  const panel = document.createElement("section");
  panel.className = "auth-panel compact";
  const token = getStoredToken();
  const status = document.createElement("p");
  status.textContent = token ? "GitHub Contents API read mode." : "Local TSV read mode.";
  const actions = document.createElement("div");
  actions.className = "button-row";

  if (token) {
    const signOut = document.createElement("button");
    signOut.type = "button";
    signOut.className = "secondary-button";
    signOut.textContent = "Sign out";
    signOut.addEventListener("click", () => {
      clearStoredToken();
      onAuthChange();
    });
    actions.append(signOut);
  } else {
    const signIn = document.createElement("button");
    signIn.type = "button";
    signIn.className = "secondary-button";
    signIn.textContent = "Sign in with GitHub";
    signIn.addEventListener("click", () => {
      signIn.disabled = true;
      void startDeviceFlow((state) => renderDeviceState(panel, state, onAuthChange)).catch(
        (error: unknown) =>
          renderDeviceState(
            panel,
            {
              status: "error",
              message: error instanceof Error ? error.message : String(error),
            },
            onAuthChange,
          ),
      );
    });
    actions.append(signIn);
  }

  panel.append(status, actions);
  return panel;
}

function renderDeviceState(
  panel: HTMLElement,
  state: DeviceFlowState,
  onAuthChange: () => void,
): void {
  let device = panel.querySelector<HTMLElement>(".device-state");
  if (!device) {
    device = document.createElement("div");
    device.className = "device-state";
    panel.append(device);
  }
  if (state.status === "code") {
    device.innerHTML = `
      <strong class="user-code">${state.code.user_code}</strong>
      <a class="text-link" href="${state.code.verification_uri}" target="_blank" rel="noreferrer">
        Open GitHub verification
      </a>
    `;
  } else if (state.status === "authenticated") {
    onAuthChange();
  } else {
    device.textContent = state.status === "pending" ? state.message : state.message;
  }
}

function downloadTsv(): void {
  const tsv = serializeDecisionTsv(rows);
  const blob = new Blob([tsv], { type: "text/tab-separated-values;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `rev-${rev}-decisions.tsv`;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}
