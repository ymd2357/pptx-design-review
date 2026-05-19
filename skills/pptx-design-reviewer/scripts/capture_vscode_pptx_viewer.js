#!/usr/bin/env node
/* eslint-disable no-console */
/*
 * Screenshot every slide in a pre-rendered vscode-pptx-viewer directory by
 * driving Playwright Chromium against the embedded index.html. Pair with
 * `render_with_vscode_pptx_viewer.js`.
 *
 * Usage:
 *   node capture_vscode_pptx_viewer.js VIEWER_DIR OUTDIR
 *
 * Environment overrides (all optional):
 *   PPTX_VIEWER_PLAYWRIGHT_ROOT  absolute path to a package.json from which to
 *                                resolve the `playwright` package. Useful when
 *                                playwright is installed in a sibling repo.
 *
 * Without an override the script tries these in order:
 *   1. require.resolve("playwright") relative to this script
 *   2. ./node_modules/playwright in the repo root
 *   3. ~/workspace/GitHub/vscode-pptx-viewer/node_modules/playwright
 *   4. ~/Documents/GitHub/vscode-pptx-viewer/node_modules/playwright
 *
 * The first one that resolves wins.
 */

const fs = require("fs");
const path = require("path");
const os = require("os");
const { createRequire } = require("module");

function usage() {
  console.error("usage: capture_vscode_pptx_viewer.js VIEWER_DIR OUTDIR");
  process.exit(2);
}

function loadPlaywright() {
  const explicit = process.env.PPTX_VIEWER_PLAYWRIGHT_ROOT;
  if (explicit) {
    return createRequire(explicit)("playwright");
  }
  // 1. Try requiring directly (works if this script's package has playwright).
  try {
    return require("playwright");
  } catch (_) {
    // fall through
  }
  // 2..4: try conventional locations.
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
      "(`npm install playwright`) or set PPTX_VIEWER_PLAYWRIGHT_ROOT to a " +
      "package.json that can resolve it."
  );
}

const { chromium } = loadPlaywright();

const viewerDir = process.argv[2];
const outdir = process.argv[3];
if (!viewerDir || !outdir) usage();

async function main() {
  fs.mkdirSync(outdir, { recursive: true });
  const data = JSON.parse(fs.readFileSync(path.join(viewerDir, "slides.json"), "utf8"));
  const url = "file://" + path.resolve(viewerDir, "index.html");
  const browser = await chromium.launch({
    headless: true,
    args: [
      "--disable-features=MachPortRendezvous",
      "--disable-gpu",
      "--single-process",
    ],
  });
  const page = await browser.newPage({ viewport: { width: 1200, height: 675 }, deviceScaleFactor: 2 });
  page.on("pageerror", (error) => console.error("[pageerror]", error.message));
  page.on("console", (message) => {
    if (message.type() === "error") console.error("[console]", message.text());
  });
  await page.goto(url);
  await page.waitForFunction(() => window.__viewerReady === true, null, { timeout: 30000 });

  for (let i = 0; i < data.meta.slideCount; i++) {
    await page.keyboard.press("Home");
    for (let step = 0; step < i; step++) {
      await page.keyboard.press("ArrowRight");
    }
    await page.waitForTimeout(150);
    const canvas = page.locator("#slide-canvas");
    await canvas.screenshot({ path: path.join(outdir, `slide-${String(i + 1).padStart(2, "0")}.png`) });
  }
  await browser.close();
  console.log(JSON.stringify({ slides: data.meta.slideCount, outdir }, null, 2));
}

main().catch((error) => {
  console.error(error && error.stack ? error.stack : String(error));
  process.exit(1);
});
