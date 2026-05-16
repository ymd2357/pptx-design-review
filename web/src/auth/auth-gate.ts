import {
  getStoredToken,
  startDeviceFlow,
  type DeviceFlowState,
} from "./device-flow";

export async function requireAuth(app: HTMLElement): Promise<string> {
  const existing = getStoredToken();
  if (existing) return existing;

  return new Promise<string>((resolve, reject) => {
    const card = renderGateCard((onUpdate) => {
      void startDeviceFlow((state) => {
        onUpdate(state);
        if (state.status === "authenticated") {
          const token = getStoredToken();
          if (token) {
            resolve(token);
          } else {
            reject(new Error("Authenticated but token missing"));
          }
        }
      }).catch((error: unknown) => {
        onUpdate({
          status: "error",
          message: error instanceof Error ? error.message : String(error),
        });
      });
    });
    app.replaceChildren(card);
  });
}

function renderGateCard(
  onStart: (onUpdate: (state: DeviceFlowState) => void) => void,
): HTMLElement {
  const root = document.createElement("main");
  root.className = "app-shell auth-gate";

  const header = document.createElement("header");
  header.className = "app-header";
  header.innerHTML = `
    <div>
      <p class="eyebrow">PPTX Design Review</p>
      <h1>Sign in required</h1>
    </div>
  `;

  const note = document.createElement("p");
  note.className = "status-text";
  note.textContent =
    "このページは GitHub の認証が必要です。Sign in を押して表示されたコードを GitHub で入力してください。";

  const button = document.createElement("button");
  button.type = "button";
  button.className = "primary-button";
  button.textContent = "Sign in with GitHub";

  const state = document.createElement("div");
  state.className = "device-state";

  button.addEventListener("click", () => {
    button.disabled = true;
    onStart((s) => renderState(state, s));
  });

  const row = document.createElement("div");
  row.className = "button-row";
  row.append(button);

  root.append(header, note, row, state);
  return root;
}

function renderState(host: HTMLElement, state: DeviceFlowState): void {
  if (state.status === "code") {
    host.innerHTML = `
      <p>Enter this code on GitHub:</p>
      <strong class="user-code">${state.code.user_code}</strong>
      <a class="text-link" href="${state.code.verification_uri}" target="_blank" rel="noreferrer">
        Open GitHub verification
      </a>
    `;
  } else if (state.status === "pending") {
    host.textContent = state.message;
  } else if (state.status === "authenticated") {
    host.textContent = "Signed in. Loading...";
  } else {
    host.textContent = state.message;
  }
}
