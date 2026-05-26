#!/usr/bin/env node
/* eslint-disable no-console */
/*
 * Developer-only text width oracle for rendered vscode-pptx-viewer output.
 *
 * Normal lint uses the fast Python estimator in pptx_lint.py. This script is
 * intentionally kept out of the user-facing lint CLI; it launches Chromium,
 * reads actual DOM character rects, and asks CDP which platform fonts rendered
 * each text element. Use it to validate estimator changes against the viewer.
 *
 * Usage:
 *   node text_width_oracle.js VIEWER_DIR [--slide N] [--out OUT.json]
 */

const fs = require("fs");
const path = require("path");
const os = require("os");
const { createRequire } = require("module");

function usage() {
  console.error("usage: text_width_oracle.js VIEWER_DIR [--slide N] [--out OUT.json]");
  process.exit(2);
}

function loadPlaywright() {
  const explicit = process.env.PPTX_VIEWER_PLAYWRIGHT_ROOT;
  if (explicit) {
    return createRequire(explicit)("playwright");
  }
  try {
    return require("playwright");
  } catch (_) {
    // fall through
  }
  const repoRoot = path.resolve(__dirname, "..", "..", "..");
  const candidates = [
    path.join(repoRoot, "package.json"),
    path.join(os.homedir(), "workspace", "GitHub", "vscode-pptx-viewer", "package.json"),
    path.join(os.homedir(), "Documents", "GitHub", "vscode-pptx-viewer", "package.json"),
  ];
  for (const pkg of candidates) {
    if (!fs.existsSync(pkg)) continue;
    try {
      return createRequire(pkg)("playwright");
    } catch (_) {
      // try next
    }
  }
  throw new Error(
    "Could not locate the `playwright` package. Install it in this repo " +
      "or set PPTX_VIEWER_PLAYWRIGHT_ROOT to a package.json that can resolve it."
  );
}

function parseArgs(argv) {
  const args = { viewerDir: argv[2], slide: null, out: null };
  if (!args.viewerDir) usage();
  for (let i = 3; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === "--slide") {
      args.slide = Number(argv[++i]);
    } else if (arg === "--out") {
      args.out = argv[++i];
    } else {
      usage();
    }
  }
  if (args.slide !== null && (!Number.isInteger(args.slide) || args.slide < 1)) {
    usage();
  }
  return args;
}

const { chromium } = loadPlaywright();
const args = parseArgs(process.argv);

async function platformFontsForOracleIds(page, ids) {
  const session = await page.context().newCDPSession(page);
  await session.send("DOM.enable");
  await session.send("CSS.enable");
  const doc = await session.send("DOM.getDocument", { depth: -1, pierce: true });
  const out = {};
  for (const id of ids) {
    const selector = `[data-text-width-oracle-id="${id}"]`;
    const result = await session.send("DOM.querySelector", {
      nodeId: doc.root.nodeId,
      selector,
    });
    if (!result.nodeId) {
      out[id] = [];
      continue;
    }
    try {
      const fonts = await session.send("CSS.getPlatformFontsForNode", {
        nodeId: result.nodeId,
      });
      out[id] = fonts.fonts || [];
    } catch (_) {
      out[id] = [];
    }
  }
  await session.detach();
  return out;
}

async function collectSlide(page, slideIndex) {
  await page.keyboard.press("Home");
  for (let step = 1; step < slideIndex; step++) {
    await page.keyboard.press("ArrowRight");
  }
  await page.waitForTimeout(150);
  const dom = await page.evaluate(() => {
    const canvas = document.querySelector("#slide-canvas");
    if (!canvas) throw new Error("#slide-canvas not found");
    const canvasRect = canvas.getBoundingClientRect();
    const lines = [];
    const ids = [];
    let nextId = 1;
    const walker = document.createTreeWalker(
      canvas,
      NodeFilter.SHOW_TEXT,
      {
        acceptNode(node) {
          return node.nodeValue && node.nodeValue.trim()
            ? NodeFilter.FILTER_ACCEPT
            : NodeFilter.FILTER_REJECT;
        },
      }
    );

    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      const parent = node.parentElement;
      if (!parent) continue;
      const oracleId = String(nextId++);
      parent.dataset.textWidthOracleId = oracleId;
      ids.push(oracleId);
      const byTop = new Map();
      for (let i = 0; i < node.nodeValue.length; i++) {
        const range = document.createRange();
        range.setStart(node, i);
        range.setEnd(node, i + 1);
        const rects = Array.from(range.getClientRects()).filter(
          (rect) => rect.width > 0 || rect.height > 0
        );
        range.detach();
        if (!rects.length) continue;
        const rect = rects[rects.length - 1];
        const topKey = String(Math.round((rect.top - canvasRect.top) * 2) / 2);
        const prev = byTop.get(topKey) || {
          top_px: rect.top - canvasRect.top,
          left_px: rect.left - canvasRect.left,
          right_px: rect.right - canvasRect.left,
          bottom_px: rect.bottom - canvasRect.top,
          text: "",
          oracle_id: oracleId,
        };
        prev.left_px = Math.min(prev.left_px, rect.left - canvasRect.left);
        prev.right_px = Math.max(prev.right_px, rect.right - canvasRect.left);
        prev.top_px = Math.min(prev.top_px, rect.top - canvasRect.top);
        prev.bottom_px = Math.max(prev.bottom_px, rect.bottom - canvasRect.top);
        prev.text += node.nodeValue[i];
        byTop.set(topKey, prev);
      }
      for (const line of byTop.values()) {
        line.width_px = line.right_px - line.left_px;
        line.parent_tag = parent.tagName.toLowerCase();
        line.parent_class = parent.getAttribute("class") || "";
        lines.push(line);
      }
    }
    lines.sort((a, b) => a.top_px - b.top_px || a.left_px - b.left_px);
    return {
      canvas_px: {
        width: canvasRect.width,
        height: canvasRect.height,
      },
      ids,
      lines,
    };
  });
  const platformFonts = await platformFontsForOracleIds(page, dom.ids);
  for (const line of dom.lines) {
    line.platform_fonts = platformFonts[line.oracle_id] || [];
  }
  return dom;
}

async function main() {
  const slidesJson = JSON.parse(fs.readFileSync(path.join(args.viewerDir, "slides.json"), "utf8"));
  const slideCount = slidesJson.meta.slideCount;
  const slideIndexes = args.slide === null
    ? Array.from({ length: slideCount }, (_, idx) => idx + 1)
    : [args.slide];
  const url = "file://" + path.resolve(args.viewerDir, "index.html");
  const browser = await chromium.launch({
    headless: true,
    args: [
      "--disable-features=MachPortRendezvous",
      "--disable-gpu",
      "--single-process",
    ],
  });
  const page = await browser.newPage({ viewport: { width: 1200, height: 675 }, deviceScaleFactor: 2 });
  await page.goto(url);
  await page.waitForFunction(() => window.__viewerReady === true, null, { timeout: 30000 });

  const slides = [];
  for (const slideIndex of slideIndexes) {
    const dom = await collectSlide(page, slideIndex);
    const slideWidthPt = slidesJson.meta.slideWidth / 12700;
    const ptPerPx = slideWidthPt / dom.canvas_px.width;
    for (const line of dom.lines) {
      line.width_pt = line.width_px * ptPerPx;
      line.left_pt = line.left_px * ptPerPx;
      line.right_pt = line.right_px * ptPerPx;
      line.top_pt = line.top_px * ptPerPx;
      line.bottom_pt = line.bottom_px * ptPerPx;
    }
    slides.push({ slide_index: slideIndex, ...dom });
  }
  await browser.close();

  const payload = {
    source: "vscode-pptx-viewer-dom-cdp",
    viewer_dir: path.resolve(args.viewerDir),
    slide_count: slideCount,
    slides,
  };
  const json = JSON.stringify(payload, null, 2);
  if (args.out) {
    fs.writeFileSync(args.out, json, "utf8");
  } else {
    console.log(json);
  }
}

main().catch((error) => {
  console.error(error && error.stack ? error.stack : String(error));
  process.exit(1);
});
