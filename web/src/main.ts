import "./styles.css";
import { requireAuth } from "./auth/auth-gate";
import { listReviewDecks, type ReviewDeck } from "./github/contents";
import { sitePath } from "./site-path";

const appElement = document.querySelector<HTMLDivElement>("#app");
if (!appElement) throw new Error("Missing #app");
const app = appElement;

void (async () => {
  await requireAuth(app);
  await renderHub();
})();

async function renderHub(): Promise<void> {
  app.replaceChildren(shell("レビューを読み込み中..."));
  const decks = await listReviewDecks();
  const root = shell("");
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
      <p class="eyebrow">PPTX デザインレビュー</p>
      <h1>レビュー一覧</h1>
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
    card.href =
      `${sitePath("review/")}?deck=${encodeURIComponent(deck.deck)}&rev=${encodeURIComponent(rev)}`;
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
