const PIN_SHA256_HEX =
  "255338aaf00799eeaea7a6bcc85984ea0b15a01d13831564628eb0d4d2d3f13c";
const storageKey = "pptx-review:pin-passed";

export async function requireAuth(app: HTMLElement): Promise<void> {
  if (sessionStorage.getItem(storageKey) === "1") return;

  return new Promise<void>((resolve) => {
    const card = renderGateCard(async (pin, onResult) => {
      const ok = await verifyPin(pin);
      if (ok) {
        sessionStorage.setItem(storageKey, "1");
        onResult({ ok: true, message: "Unlocked. Loading…" });
        resolve();
      } else {
        onResult({ ok: false, message: "PIN が違います。" });
      }
    });
    app.replaceChildren(card);
  });
}

async function verifyPin(pin: string): Promise<boolean> {
  const encoder = new TextEncoder();
  const data = encoder.encode(pin.trim());
  const digest = await crypto.subtle.digest("SHA-256", data);
  const hex = Array.from(new Uint8Array(digest))
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");
  return hex === PIN_SHA256_HEX;
}

function renderGateCard(
  onSubmit: (
    pin: string,
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
      <h1>PIN required</h1>
    </div>
  `;

  const note = document.createElement("p");
  note.className = "status-text";
  note.textContent = "PIN を入力してください。";

  const form = document.createElement("form");
  form.className = "auth-form";

  const input = document.createElement("input");
  input.type = "password";
  input.name = "pin";
  input.inputMode = "numeric";
  input.autocomplete = "off";
  input.pattern = "[0-9]*";
  input.placeholder = "PIN";
  input.className = "token-input";
  input.required = true;

  const submit = document.createElement("button");
  submit.type = "submit";
  submit.className = "primary-button";
  submit.textContent = "Unlock";

  form.append(input, submit);

  const state = document.createElement("div");
  state.className = "device-state";

  form.addEventListener("submit", (event) => {
    event.preventDefault();
    const pin = input.value.trim();
    if (!pin) return;
    submit.disabled = true;
    state.textContent = "Checking…";
    onSubmit(pin, ({ ok, message }) => {
      state.textContent = message;
      if (!ok) {
        submit.disabled = false;
        input.value = "";
        input.focus();
      }
    });
  });

  root.append(header, note, form, state);
  return root;
}
