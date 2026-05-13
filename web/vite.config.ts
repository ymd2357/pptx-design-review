import fs from "node:fs";
import path from "node:path";
import { defineConfig, type Plugin } from "vite";

const repoRoot = path.resolve(__dirname, "..");
const reviewsRoot = path.join(repoRoot, "doc", "reviews");

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
        const requestPath = decodeURIComponent(req.url ?? "/");
        const normalized = path.normalize(requestPath).replace(/^(\.\.[/\\])+/, "");
        const filePath = path.join(reviewsRoot, normalized);
        if (!filePath.startsWith(reviewsRoot) || !fs.existsSync(filePath)) {
          res.statusCode = 404;
          res.end("Not found");
          return;
        }
        res.setHeader("Content-Type", "text/tab-separated-values; charset=utf-8");
        fs.createReadStream(filePath).pipe(res);
      });
    },
  };
}

export default defineConfig({
  root: ".",
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
