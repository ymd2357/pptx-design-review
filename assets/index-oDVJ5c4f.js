import{r as l,s as o}from"./site-path-BDZEjMdB.js";import{l as m}from"./contents-C1TvvOkQ.js";const d=document.querySelector("#app");if(!d)throw new Error("Missing #app");const r=d;(async()=>(await l(r),await u()))();async function u(){r.replaceChildren(s("レビューを読み込み中..."));const t=await m(),e=s("");e.append(h(t)),r.replaceChildren(e)}function s(t){const e=document.createElement("main");e.className="app-shell";const n=document.createElement("header");if(n.className="app-header",n.innerHTML=`
    <div>
      <p class="eyebrow">PPTX デザインレビュー</p>
      <h1>レビュー一覧</h1>
    </div>
  `,t){const c=document.createElement("p");return c.className="status-text",c.textContent=t,e.append(n,c),e}return e.append(n),e}function h(t){const e=document.createElement("section");e.className="deck-grid";for(const n of t){const c=document.createElement("div");c.className="deck-card";const a=n.revs.at(-1)??"017",i=`${o("compare/")}?deck=${encodeURIComponent(n.deck)}&rev=${encodeURIComponent(a)}`,p=`${o("review/")}?deck=${encodeURIComponent(n.deck)}&rev=${encodeURIComponent(a)}`;c.innerHTML=`
      <div class="thumb-placeholder"></div>
      <div>
        <p class="eyebrow">${n.source}</p>
        <h2>${n.deck}</h2>
        <p>REV-${n.revs.join(", REV-")}</p>
        <div class="deck-card-actions">
          <a class="primary-button" href="${i}">修正比較を開く</a>
          <a class="secondary-link" href="${p}">観点別レビュー</a>
        </div>
      </div>
    `,e.append(c)}return e}
