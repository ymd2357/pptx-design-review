import type { FindingJudgementsFile } from "../data/finding-judgements";
import type { LintFinding, SlideSizePt } from "../data/lint-json";

export type SlideGalleryOptions = {
  imageUrls: readonly string[];
  findings: readonly LintFinding[];
  judgements: FindingJudgementsFile;
  slideSizePt: SlideSizePt;
  initialSlideNo?: number;
  onSelectFinding: (finding: LintFinding) => void;
};

export function renderSlideGallery(options: SlideGalleryOptions): HTMLElement {
  let currentIndex = initialSlideIndex(options.imageUrls, options.initialSlideNo);
  let pointerStartX: number | null = null;

  const root = document.createElement("section");
  root.className = "slide-gallery";
  root.tabIndex = 0;

  const toolbar = document.createElement("div");
  toolbar.className = "slide-toolbar";

  const previous = document.createElement("button");
  previous.type = "button";
  previous.className = "secondary-button";
  previous.textContent = "Prev";

  const position = document.createElement("p");
  position.className = "slide-position";

  const next = document.createElement("button");
  next.type = "button";
  next.className = "secondary-button";
  next.textContent = "Next";
  toolbar.append(previous, position, next);

  const stage = document.createElement("div");
  stage.className = "slide-stage";

  previous.addEventListener("click", () => move(-1));
  next.addEventListener("click", () => move(1));
  root.addEventListener("keydown", (event) => {
    if (event.key === "ArrowLeft") move(-1);
    if (event.key === "ArrowRight") move(1);
  });
  stage.addEventListener("pointerdown", (event) => {
    pointerStartX = event.clientX;
  });
  stage.addEventListener("pointerup", (event) => {
    if (pointerStartX === null) return;
    const deltaX = event.clientX - pointerStartX;
    pointerStartX = null;
    if (Math.abs(deltaX) < 44) return;
    move(deltaX < 0 ? 1 : -1);
  });

  root.append(toolbar, stage);
  render();
  return root;

  function move(delta: number): void {
    const nextIndex = Math.min(Math.max(currentIndex + delta, 0), options.imageUrls.length - 1);
    if (nextIndex === currentIndex) return;
    currentIndex = nextIndex;
    render();
  }

  function render(): void {
    stage.replaceChildren();
    const imageUrl = options.imageUrls[currentIndex];
    const slideNo = slideNumberFromUrl(imageUrl) ?? currentIndex + 1;
    const slideFindings = options.findings.filter((finding) => finding.slideNo === slideNo);
    position.textContent = `Slide ${slideNo} / ${options.imageUrls.length}`;
    previous.disabled = currentIndex === 0;
    next.disabled = currentIndex === options.imageUrls.length - 1;

    const frame = document.createElement("div");
    frame.className = "slide-frame";

    const image = document.createElement("img");
    image.alt = `Slide ${slideNo}`;
    image.loading = "eager";
    image.decoding = "async";
    image.src = imageUrl;
    preloadNeighbor(currentIndex - 1);
    preloadNeighbor(currentIndex + 1);

    const overlay = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    overlay.setAttribute("class", "finding-overlay");
    overlay.setAttribute("viewBox", `0 0 ${options.slideSizePt.w} ${options.slideSizePt.h}`);
    overlay.setAttribute("preserveAspectRatio", "none");
    overlay.setAttribute("aria-label", "Finding overlays");
    for (const finding of slideFindings) {
      if (!finding.bboxPt) continue;
      const [x, y, width, height] = finding.bboxPt;
      const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
      rect.setAttribute("x", String(x));
      rect.setAttribute("y", String(y));
      rect.setAttribute("width", String(Math.max(width, 6)));
      rect.setAttribute("height", String(Math.max(height, 6)));
      rect.setAttribute("class", judgementClass(options.judgements, finding.key));
      rect.setAttribute("tabindex", "0");
      rect.setAttribute("role", "button");
      rect.setAttribute("aria-label", `${finding.check}: ${finding.message}`);
      rect.addEventListener("click", () => options.onSelectFinding(finding));
      rect.addEventListener("keydown", (event) => {
        if (event.key === "Enter" || event.key === " ") {
          event.preventDefault();
          options.onSelectFinding(finding);
        }
      });
      overlay.append(rect);
    }

    const empty = document.createElement("p");
    empty.className = "slide-empty";
    empty.textContent = slideFindings.length === 0 ? "No findings on this slide." : "";

    frame.append(image, overlay);
    stage.append(frame);
    if (empty.textContent) stage.append(empty);
  }

  function preloadNeighbor(index: number): void {
    if (index < 0 || index >= options.imageUrls.length) return;
    const link = document.createElement("link");
    link.rel = "preload";
    link.as = "image";
    link.href = options.imageUrls[index] ?? "";
    document.head.append(link);
    window.setTimeout(() => link.remove(), 5000);
  }
}

function judgementClass(judgements: FindingJudgementsFile, groupKey: string): string {
  const judgement = judgements.judgements[groupKey];
  return judgement?.review_status && judgement.review_status !== "unreviewed"
    ? "finding-box judged"
    : "finding-box";
}

function slideNumberFromUrl(url: string): number | null {
  const match = url.match(/slide-(\d+)\.png(?:$|\?)/);
  return match ? Number.parseInt(match[1], 10) : null;
}

function initialSlideIndex(imageUrls: readonly string[], slideNo: number | undefined): number {
  if (!slideNo) return 0;
  const index = imageUrls.findIndex((url) => slideNumberFromUrl(url) === slideNo);
  return index >= 0 ? index : 0;
}
