import "./styles.css";
import {
  clearStoredToken,
  getStoredToken,
  startDeviceFlow,
  type DeviceFlowState,
} from "./auth/device-flow";
import { listReviewDecks, type ReviewDeck } from "./github/contents";

const appElement = document.querySelector<HTMLDivElement>("#app");
if (!appElement) throw new Error("Missing #app");
const app = appElement;

void renderHub();

async function renderHub(): Promise<void> {
  app.replaceChildren(shell("Loading reviews..."));
  const decks = await listReviewDecks();
  const root = shell("");
  root.append(renderAuthPanel(() => void renderHub()));
  root.append(renderDeckGrid(decks));
  app.replaceChildren(root);
}

function shell(title: string): HTMLElement {
  const root = document.createElement("main");
  root.className = "app-shell";
  const header = document.createElement("header");
  header.className = "app-header";
  header.innerHTML = `
    <div>
      <p class="eyebrow">PPTX Design Review</p>
      <h1>Review Hub</h1>
    </div>
  `;
  if (title) {
    const status = document.createElement("p");
    status.className = "status-text";
    status.textContent = title;
    root.append(header, status);
    return root;
  }
  root.append(header);
  return root;
}

function renderDeckGrid(decks: ReviewDeck[]): HTMLElement {
  const grid = document.createElement("section");
  grid.className = "deck-grid";
  for (const deck of decks) {
    const card = document.createElement("a");
    card.className = "deck-card";
    const rev = deck.revs.at(-1) ?? "017";
    card.href = `/review/?deck=${encodeURIComponent(deck.deck)}&rev=${encodeURIComponent(rev)}`;
    card.innerHTML = `
      <div class="thumb-placeholder"></div>
      <div>
        <p class="eyebrow">${deck.source}</p>
        <h2>${deck.deck}</h2>
        <p>REV-${deck.revs.join(", REV-")}</p>
      </div>
    `;
    grid.append(card);
  }
  return grid;
}

function renderAuthPanel(onAuthChange: () => void): HTMLElement {
  const panel = document.createElement("section");
  panel.className = "auth-panel";
  const token = getStoredToken();
  const status = document.createElement("p");
  status.textContent = token
    ? "GitHub authenticated. Contents API reads are enabled."
    : "Local/mock review data is shown until GitHub sign-in.";

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
    signIn.className = "primary-button";
    signIn.textContent = "Sign in with GitHub";
    signIn.addEventListener("click", () => {
      signIn.disabled = true;
      void startDeviceFlow((state) => {
        renderDeviceState(panel, state, onAuthChange);
      }).catch((error: unknown) => {
        renderDeviceState(panel, {
          status: "error",
          message: error instanceof Error ? error.message : String(error),
        }, onAuthChange);
      });
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
      <p>Enter this code on GitHub:</p>
      <strong class="user-code">${state.code.user_code}</strong>
      <a class="text-link" href="${state.code.verification_uri}" target="_blank" rel="noreferrer">
        Open GitHub verification
      </a>
    `;
  } else if (state.status === "pending") {
    device.textContent = state.message;
  } else if (state.status === "authenticated") {
    onAuthChange();
  } else {
    device.textContent = state.message;
  }
}
