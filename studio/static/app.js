// PPTX Lint/Fix Studio — vanilla front-end (no build step).

const state = {
  session: null, // upload response
  // selection: Set of "slide_index::check" that are ON
  selection: new Set(),
};

const $ = (sel) => document.querySelector(sel);

// ---- upload ---------------------------------------------------------------

const dropzone = $("#dropzone");
const fileInput = $("#file-input");

$("#browse-btn").addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", () => {
  if (fileInput.files.length) uploadFile(fileInput.files[0]);
});

["dragenter", "dragover"].forEach((ev) =>
  dropzone.addEventListener(ev, (e) => {
    e.preventDefault();
    dropzone.classList.add("drag");
  })
);
["dragleave", "drop"].forEach((ev) =>
  dropzone.addEventListener(ev, (e) => {
    e.preventDefault();
    dropzone.classList.remove("drag");
  })
);
dropzone.addEventListener("drop", (e) => {
  const f = e.dataTransfer.files[0];
  if (f) uploadFile(f);
});

async function uploadFile(file) {
  if (!file.name.toLowerCase().endsWith(".pptx")) {
    setUploadStatus("PPTX ファイルを選んでください", true);
    return;
  }
  setUploadStatus(`アップロード中… (${(file.size / 1e6).toFixed(1)} MB) → 正規化 → lint`);
  try {
    const res = await fetch("/api/upload", {
      method: "POST",
      headers: {
        "Content-Type": "application/octet-stream",
        "X-Filename": encodeURIComponent(file.name),
      },
      body: file,
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "upload failed");
    onSession(data);
  } catch (err) {
    setUploadStatus("失敗: " + err.message, true);
  }
}

function setUploadStatus(msg, isError) {
  const el = $("#upload-status");
  el.textContent = msg;
  el.classList.toggle("error", !!isError);
}

// ---- render findings ------------------------------------------------------

function onSession(data) {
  state.session = data;
  state.selection = new Set();

  // default: turn on every fixable check
  for (const slide of data.slides) {
    for (const c of slide.checks) {
      if (c.default_on) state.selection.add(key(slide.slide_index, c.check));
    }
  }

  setUploadStatus(
    `lint 完了: ${data.finding_count} findings / ${data.slide_count} slides` +
      (data.normalize.occurrences
        ? ` ・ font 正規化 ${data.normalize.occurrences} 箇所`
        : "")
  );

  $("#deck-name").textContent = data.filename;
  $("#deck-stats").textContent = `${data.slide_count} slides ・ ${data.finding_count} findings`;
  $("#toolbar").classList.remove("hidden");
  $("#apply-summary").classList.add("hidden");
  disableDownload(true);
  renderSlides();
}

function key(slide, check) {
  return `${slide}::${check}`;
}

function renderSlides() {
  const container = $("#slides");
  container.innerHTML = "";
  for (const slide of state.session.slides) {
    const card = document.createElement("div");
    card.className = "slide-card";

    const head = document.createElement("div");
    head.className = "slide-head";
    head.innerHTML = `<span class="sidx">スライド ${slide.slide_index}</span>`;
    const renderBtn = document.createElement("button");
    renderBtn.className = "btn ghost render-btn";
    renderBtn.textContent = "描画して確認";
    renderBtn.addEventListener("click", () => openRender(slide.slide_index));
    head.appendChild(renderBtn);
    card.appendChild(head);

    const chips = document.createElement("div");
    chips.className = "chips";
    if (!slide.checks.length) {
      const e = document.createElement("span");
      e.className = "slide-empty";
      e.textContent = "問題なし";
      chips.appendChild(e);
    }
    for (const c of slide.checks) {
      chips.appendChild(makeChip(slide.slide_index, c));
    }
    card.appendChild(chips);
    container.appendChild(card);
  }
}

function makeChip(slide, c) {
  const chip = document.createElement("span");
  const k = key(slide, c.check);
  const on = state.selection.has(k);
  chip.className = "chip" + (on ? " on" : "") + (c.severity === "error" ? " sev-error" : "");
  if (!c.fixable) chip.classList.add("locked");

  const title = c.messages && c.messages.length ? c.messages.join("\n") : c.check;
  chip.title =
    `${c.check}\nrule: ${c.rule}\nmode: ${c.apply_mode || "—"}` +
    (c.fixable ? "" : "\n(機械修正の実装なし — 表示のみ)") +
    (title !== c.check ? "\n\n" + title : "");

  chip.innerHTML =
    `<span class="dot"></span>` +
    `<span class="label">${c.check}</span>` +
    `<span class="count">${c.count}</span>` +
    (c.apply_mode ? `<span class="mode">${shortMode(c.apply_mode)}</span>` : "");

  if (c.fixable) {
    chip.addEventListener("click", () => {
      if (state.selection.has(k)) state.selection.delete(k);
      else state.selection.add(k);
      chip.classList.toggle("on");
    });
  }
  return chip;
}

function shortMode(mode) {
  return { auto_fix: "auto", judgement_fix: "judge", no_fix: "info" }[mode] || mode;
}

// ---- toolbar actions ------------------------------------------------------

$("#select-all").addEventListener("click", () => {
  for (const slide of state.session.slides)
    for (const c of slide.checks)
      if (c.fixable) state.selection.add(key(slide.slide_index, c.check));
  renderSlides();
});
$("#select-none").addEventListener("click", () => {
  state.selection.clear();
  renderSlides();
});

$("#apply-btn").addEventListener("click", applySelection);

async function applySelection() {
  if (!state.session) return;
  const selections = [...state.selection].map((k) => {
    const [slide, check] = k.split("::");
    return { slide_index: Number(slide), check };
  });
  const btn = $("#apply-btn");
  btn.disabled = true;
  btn.textContent = "適用中…";
  try {
    const res = await fetch("/api/apply", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ session_id: state.session.session_id, selections }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "apply failed");
    showApplySummary(data);
    disableDownload(false);
    // fixed.pptx changed -> invalidate fixed render so preview re-renders
    state._fixedRendered = false;
  } catch (err) {
    showApplyError(err.message);
  } finally {
    btn.disabled = false;
    btn.textContent = "選択を適用";
  }
}

function showApplySummary(data) {
  const el = $("#apply-summary");
  el.classList.remove("hidden");
  const bySlide = {};
  for (const a of data.applied) {
    (bySlide[a.slide_index] = bySlide[a.slide_index] || []).push(a.rule);
  }
  const lines = Object.keys(bySlide)
    .sort((a, b) => a - b)
    .map((s) => `スライド ${s}: ${[...new Set(bySlide[s])].join(", ")}`)
    .join("<br>");
  el.innerHTML =
    `<h3>適用結果 <span class="pill">${data.applied.length} 適用</span>` +
    (data.skipped.length ? ` ・ ${data.skipped.length} スキップ` : "") +
    ` ・ ${data.selected} 選択</h3>` +
    (lines || '<span class="muted">適用対象なし</span>');
}

function showApplyError(msg) {
  const el = $("#apply-summary");
  el.classList.remove("hidden");
  el.innerHTML = `<h3 style="color:var(--err)">適用失敗</h3>${msg}`;
}

function disableDownload(disabled) {
  const a = $("#download-btn");
  if (disabled) {
    a.classList.add("disabled");
    a.removeAttribute("href");
  } else {
    a.classList.remove("disabled");
    a.href = `/api/download?session_id=${state.session.session_id}`;
  }
}

// ---- render preview modal -------------------------------------------------

const modal = $("#modal");
$("#modal-close").addEventListener("click", () => modal.classList.add("hidden"));
modal.addEventListener("click", (e) => {
  if (e.target === modal) modal.classList.add("hidden");
});

async function openRender(slideIndex) {
  modal.classList.remove("hidden");
  $("#modal-title").textContent = `スライド ${slideIndex} のプレビュー`;
  const body = $("#modal-body");
  const hasFixed = !$("#download-btn").classList.contains("disabled");
  body.innerHTML = `<p class="spinner">描画中… (vscode-pptx-viewer + Playwright, 数秒/枚)</p>`;

  try {
    const source = await renderDeck("source");
    let fixed = null;
    if (hasFixed) fixed = await renderDeck("fixed");

    const fname = `slide-${String(slideIndex).padStart(2, "0")}.png`;
    const srcSlide = source.slides.find((s) => s.name === fname);
    const fixSlide = fixed && fixed.slides.find((s) => s.name === fname);

    if (!srcSlide) {
      body.innerHTML = `<p class="spinner">このスライドの描画が見つかりません (${fname})</p>`;
      return;
    }
    body.innerHTML = "";
    const pair = document.createElement("div");
    pair.className = "render-pair";
    pair.appendChild(figure("before (source)", srcSlide.url));
    if (fixSlide) pair.appendChild(figure("after (fixed)", fixSlide.url));
    body.appendChild(pair);
  } catch (err) {
    body.innerHTML = `<p class="spinner" style="color:var(--err)">描画失敗: ${err.message}</p>`;
  }
}

function figure(caption, url) {
  const fig = document.createElement("figure");
  fig.innerHTML = `<figcaption>${caption}</figcaption><img src="${url}" alt="${caption}" />`;
  return fig;
}

const renderCache = {};
async function renderDeck(which) {
  if (renderCache[which] && (which === "source" || state._fixedRendered))
    return renderCache[which];
  const res = await fetch("/api/render", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ session_id: state.session.session_id, which }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "render failed");
  renderCache[which] = data;
  if (which === "fixed") state._fixedRendered = true;
  return data;
}
