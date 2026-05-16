import {
  getStoredToken,
  setStoredToken,
  verifyToken,
} from "./token-store";

export async function requireAuth(app: HTMLElement): Promise<string> {
  const existing = getStoredToken();
  if (existing) return existing;

  return new Promise<string>((resolve) => {
    const card = renderGateCard(async (token, onResult) => {
      try {
        const login = await verifyToken(token);
        setStoredToken(token);
        onResult({ ok: true, message: `Signed in as ${login}. Loading…` });
        resolve(token);
      } catch (error) {
        onResult({
          ok: false,
          message:
            error instanceof Error ? error.message : String(error),
        });
      }
    });
    app.replaceChildren(card);
  });
}

function renderGateCard(
  onSubmit: (
    token: string,
    onResult: (result: { ok: boolean; message: string }) => void,
  ) => void,
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
    "GitHub Personal Access Token を貼り付けてください。fine-grained / classic どちらでも可。必要 scope: Contents (Read and write)。";

  const help = document.createElement("p");
  help.className = "status-text";
  help.innerHTML =
    'Token 発行: <a class="text-link" href="https://github.com/settings/personal-access-tokens/new" target="_blank" rel="noreferrer">fine-grained</a> ' +
    'または <a class="text-link" href="https://github.com/settings/tokens/new?scopes=repo&description=pptx-design-review" target="_blank" rel="noreferrer">classic (repo)</a>';

  const form = document.createElement("form");
  form.className = "auth-form";

  const input = document.createElement("input");
  input.type = "password";
  input.name = "token";
  input.autocomplete = "off";
  input.placeholder = "ghp_… / github_pat_…";
  input.className = "token-input";
  input.required = true;

  const submit = document.createElement("button");
  submit.type = "submit";
  submit.className = "primary-button";
  submit.textContent = "Sign in";

  form.append(input, submit);

  const state = document.createElement("div");
  state.className = "device-state";

  form.addEventListener("submit", (event) => {
    event.preventDefault();
    const token = input.value.trim();
    if (!token) return;
    submit.disabled = true;
    state.textContent = "Verifying token…";
    onSubmit(token, ({ ok, message }) => {
      state.textContent = message;
      if (!ok) submit.disabled = false;
    });
  });

  root.append(header, note, help, form, state);
  return root;
}
