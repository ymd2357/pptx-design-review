#!/usr/bin/env node
/* eslint-disable no-console */
/*
 * Render a PPTX into a self-contained HTML viewer directory by re-using the
 * vscode-pptx-viewer VSCode extension's bundled parser + webview assets.
 *
 * Pair with `capture_vscode_pptx_viewer.js` to screenshot each slide via
 * Playwright Chromium. The two scripts together replace the legacy
 * PDF-based capture pipeline (see SKILL.md "Visual export" section).
 *
 * Usage:
 *   node render_with_vscode_pptx_viewer.js INPUT.pptx OUTDIR
 *
 * Environment overrides (all optional):
 *   PPTX_VIEWER_EXT_DIR        absolute path to the unpacked VSCode extension
 *                              directory (the one containing dist/extension.js,
 *                              dist/webview.js, dist/webview.css).
 *
 * Without an override the script searches a small list of conventional
 * locations and uses the highest-versioned match.
 */

const fs = require("fs");
const path = require("path");
const os = require("os");
const vm = require("vm");
const { createRequire } = require("module");

function usage() {
  console.error("usage: render_with_vscode_pptx_viewer.js INPUT.pptx OUTDIR");
  process.exit(2);
}

function locateExtensionDir() {
  const explicit = process.env.PPTX_VIEWER_EXT_DIR;
  if (explicit) {
    if (!fs.existsSync(path.join(explicit, "dist", "extension.js"))) {
      throw new Error(
        `PPTX_VIEWER_EXT_DIR is set to ${explicit} but dist/extension.js is missing.`
      );
    }
    return explicit;
  }
  const candidates = [
    path.join(os.homedir(), ".vscode", "extensions"),
    path.join(os.homedir(), ".cursor", "extensions"),
    path.join(os.homedir(), ".vscode-server", "extensions"),
  ];
  const found = [];
  for (const dir of candidates) {
    if (!fs.existsSync(dir)) continue;
    for (const entry of fs.readdirSync(dir)) {
      if (!/^astx-jp\.vscode-pptx-viewer-/i.test(entry)) continue;
      const full = path.join(dir, entry);
      if (fs.existsSync(path.join(full, "dist", "extension.js"))) {
        found.push(full);
      }
    }
  }
  if (!found.length) {
    throw new Error(
      "Could not locate vscode-pptx-viewer extension. Install astx-jp.vscode-pptx-viewer " +
        "into VSCode/Cursor, or set PPTX_VIEWER_EXT_DIR to its unpacked dist parent."
    );
  }
  // Highest semver wins. Compare numeric parts segment-by-segment so 0.0.10 > 0.0.8.
  function parseVersion(dir) {
    const m = path.basename(dir).match(/-([0-9]+(?:\.[0-9]+)*)$/);
    return m ? m[1].split(".").map((n) => parseInt(n, 10)) : [];
  }
  found.sort((a, b) => {
    const va = parseVersion(a);
    const vb = parseVersion(b);
    const len = Math.max(va.length, vb.length);
    for (let i = 0; i < len; i++) {
      const ai = va[i] || 0;
      const bi = vb[i] || 0;
      if (ai !== bi) return ai - bi;
    }
    return 0;
  });
  return found[found.length - 1];
}

const input = process.argv[2];
const outdir = process.argv[3];
if (!input || !outdir) usage();

const EXT_DIR = locateExtensionDir();
const DIST_DIR = path.join(EXT_DIR, "dist");
const EXTENSION_JS = path.join(DIST_DIR, "extension.js");
const WEBVIEW_JS = path.join(DIST_DIR, "webview.js");
const WEBVIEW_CSS = path.join(DIST_DIR, "webview.css");

function loadParser() {
  const source =
    fs.readFileSync(EXTENSION_JS, "utf8") +
    "\n;module.exports.__parsePptx = Ad; module.exports.__loadSlide = Nd;\n";
  const module = { exports: {} };
  const extRequire = createRequire(EXTENSION_JS);
  const sandbox = {
    module,
    exports: module.exports,
    require(id) {
      if (id === "vscode") {
        return {
          window: {
            registerCustomEditorProvider() {
              return { dispose() {} };
            },
            showErrorMessage(message) {
              throw new Error(message);
            },
          },
          workspace: { fs: { readFile: fs.promises.readFile } },
          Uri: {
            joinPath(...parts) {
              return path.join(...parts.map((p) => (typeof p === "string" ? p : p.fsPath || String(p))));
            },
            parse(value) {
              return value;
            },
          },
          env: { openExternal() {} },
          extensions: { getExtension() { return null; } },
          version: "standalone",
        };
      }
      return extRequire(id);
    },
    console,
    process,
    Buffer,
    __dirname: DIST_DIR,
    __filename: EXTENSION_JS,
    setTimeout,
    clearTimeout,
    setImmediate,
    clearImmediate,
    URLSearchParams,
  };
  vm.runInNewContext(source, sandbox, { filename: EXTENSION_JS });
  return module.exports;
}

async function main() {
  fs.mkdirSync(outdir, { recursive: true });
  const parser = loadParser();
  const buffer = fs.readFileSync(input);
  const { meta, handle } = await parser.__parsePptx(buffer);
  const slides = [];
  for (let i = 0; i < meta.slideCount; i++) {
    slides.push(await parser.__loadSlide(handle, i));
  }

  fs.copyFileSync(WEBVIEW_JS, path.join(outdir, "webview.js"));
  fs.copyFileSync(WEBVIEW_CSS, path.join(outdir, "webview.css"));
  fs.writeFileSync(path.join(outdir, "slides.json"), JSON.stringify({ meta, slides }), "utf8");

  const embeddedData = JSON.stringify({ meta, slides }).replace(/</g, "\\u003c");
  const html = `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <link href="webview.css" rel="stylesheet">
  <style>
    body { margin: 0; background: #fff; }
    #thumbnail-sidebar, #navigation-bar { display: none !important; }
    #app, #main-area, #slide-container { width: 100vw; height: 100vh; }
    #slide-container { overflow: hidden !important; display: flex; align-items: center; justify-content: center; }
  </style>
</head>
<body>
  <div id="app">
    <div id="thumbnail-sidebar"></div>
    <div id="main-area">
      <div id="slide-container">
        <div id="slide-canvas"></div>
      </div>
      <div id="navigation-bar">
        <button id="prev-btn"></button><span id="slide-counter"></span><button id="next-btn"></button>
        <button id="zoom-out-btn"></button><span id="zoom-label"></span><button id="zoom-in-btn"></button><button id="zoom-reset-btn"></button>
        <span id="debug-badge" hidden></span><span id="unsupported-badge"></span><button id="report-issue-btn"></button>
        <div id="unsupported-popover" hidden><ul id="unsupported-popover-list"></ul><a id="unsupported-popover-report" href="#"></a></div>
      </div>
    </div>
  </div>
  <script>
    window.__viewerReady = false;
    window.__viewerDataPromise = Promise.resolve(${embeddedData});
    window.acquireVsCodeApi = function() {
      let state = {};
      return {
        getState() { return state; },
        setState(next) { state = next || {}; },
        postMessage(message) {
          if (message && message.type === "ready") {
            window.__viewerDataPromise.then(({ meta, slides }) => {
              window.postMessage({ type: "init", slideWidth: meta.slideWidth, slideHeight: meta.slideHeight, slideCount: meta.slideCount, firstSlideNum: meta.firstSlideNum, presentationWarnings: meta.presentationWarnings }, "*");
              slides.forEach((slide, index) => window.postMessage({ type: "loadSlide", index, slide }, "*"));
              window.__viewerReady = true;
            });
          }
        }
      };
    };
  </script>
  <script src="webview.js"></script>
</body>
</html>`;
  fs.writeFileSync(path.join(outdir, "index.html"), html, "utf8");
  console.log(JSON.stringify({ html: path.join(outdir, "index.html"), slideCount: meta.slideCount, extensionDir: EXT_DIR }, null, 2));
}

main().catch((error) => {
  console.error(error && error.stack ? error.stack : String(error));
  process.exit(1);
});
