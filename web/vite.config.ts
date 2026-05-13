import fs from "node:fs";
import type { ServerResponse } from "node:http";
import path from "node:path";
import { defineConfig, type Plugin } from "vite";

const repoRoot = path.resolve(__dirname, "..");
const reviewsRoot = path.join(repoRoot, "doc", "reviews");
const snapshotRoot = path.join(repoRoot, "tmp", "review-snapshot");
const repositoryName = process.env.GITHUB_REPOSITORY?.split("/")[1];
const base =
  process.env.GITHUB_ACTIONS === "true" && repositoryName ? `/${repositoryName}/` : "/";

function localReviewsPlugin(): Plugin {
  return {
    name: "local-reviews",
    configureServer(server) {
      server.middlewares.use("/api/local-reviews", (_req, res) => {
        const decks = fs.existsSync(reviewsRoot)
          ? fs
              .readdirSync(reviewsRoot, { withFileTypes: true })
              .filter((entry) => entry.isDirectory())
              .map((entry) => {
                const deck = entry.name;
                const files = fs
                  .readdirSync(path.join(reviewsRoot, deck), {
                    withFileTypes: true,
                  })
                  .filter((file) => file.isFile() && /^rev-\d+-decisions\.tsv$/.test(file.name))
                  .map((file) => file.name);
                return { deck, files };
              })
          : [];
        res.setHeader("Content-Type", "application/json; charset=utf-8");
        res.end(JSON.stringify({ decks }));
      });

      server.middlewares.use("/doc/reviews", (req, res) => {
        serveRepoFile(req.url ?? "/", reviewsRoot, res);
      });

      server.middlewares.use("/tmp/review-snapshot", (req, res) => {
        serveRepoFile(req.url ?? "/", snapshotRoot, res);
      });
    },
  };
}

function serveRepoFile(requestUrl: string, root: string, res: ServerResponse): void {
  const requestPath = decodeURIComponent(requestUrl);
  const normalized = path.normalize(requestPath).replace(/^(\.\.[/\\])+/, "");
  const filePath = path.join(root, normalized);
  if (!filePath.startsWith(root) || !fs.existsSync(filePath)) {
    res.statusCode = 404;
    res.end("Not found");
    return;
  }
  res.setHeader("Content-Type", contentType(filePath));
  fs.createReadStream(filePath).pipe(res);
}

function contentType(filePath: string): string {
  if (filePath.endsWith(".json")) return "application/json; charset=utf-8";
  if (filePath.endsWith(".png")) return "image/png";
  if (filePath.endsWith(".tsv")) return "text/tab-separated-values; charset=utf-8";
  return "application/octet-stream";
}

export default defineConfig({
  root: ".",
  base,
  plugins: [localReviewsPlugin()],
  build: {
    outDir: "dist",
    emptyOutDir: true,
    rollupOptions: {
      input: {
        index: path.resolve(__dirname, "index.html"),
        review: path.resolve(__dirname, "review", "index.html"),
      },
    },
  },
});
