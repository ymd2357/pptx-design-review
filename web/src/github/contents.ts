import { getStoredToken } from "../auth/device-flow";

const apiRoot = "https://api.github.com/repos";

export type ReviewDeck = {
  deck: string;
  revs: string[];
  source: "github" | "local" | "mock";
};

export type ContentFile = {
  text: string;
  sha?: string;
  source: "github" | "local";
};

type GitHubContent = {
  name: string;
  path: string;
  type: "file" | "dir";
  sha: string;
  content?: string;
  encoding?: string;
};

const mockDecks: ReviewDeck[] = [
  {
    deck: "260329-seminar-curriculum-proposal",
    revs: ["017"],
    source: "mock",
  },
];

export async function listReviewDecks(): Promise<ReviewDeck[]> {
  const token = getStoredToken();
  if (token && hasRepoCoordinates()) {
    try {
      return await listReviewDecksFromGitHub(token);
    } catch (error) {
      console.warn("GitHub deck listing failed; falling back to local data.", error);
    }
  }

  try {
    return await listReviewDecksFromLocalDev();
  } catch {
    return mockDecks;
  }
}

export async function fetchDecisionTsv(deck: string, rev: string): Promise<ContentFile> {
  const token = getStoredToken();
  const filePath = `doc/reviews/${deck}/rev-${rev}-decisions.tsv`;
  if (token && hasRepoCoordinates()) {
    try {
      const file = await githubJson<GitHubContent>(contentsUrl(filePath), token);
      return {
        text: decodeGitHubContent(file),
        sha: file.sha,
        source: "github",
      };
    } catch (error) {
      console.warn("GitHub TSV fetch failed; falling back to local data.", error);
    }
  }

  const response = await fetch(`/${filePath}`);
  if (!response.ok) {
    throw new Error(`Failed to load ${filePath}: ${response.status}`);
  }
  return { text: await response.text(), source: "local" };
}

async function listReviewDecksFromGitHub(token: string): Promise<ReviewDeck[]> {
  const dirs = await githubJson<GitHubContent[]>(contentsUrl("doc/reviews"), token);
  const decks = dirs.filter((item) => item.type === "dir");
  const withRevs = await Promise.all(
    decks.map(async (deck) => {
      const files = await githubJson<GitHubContent[]>(contentsUrl(deck.path), token);
      return {
        deck: deck.name,
        revs: extractRevs(files.map((file) => file.name)),
        source: "github" as const,
      };
    }),
  );
  return withRevs.filter((deck) => deck.revs.length > 0);
}

async function listReviewDecksFromLocalDev(): Promise<ReviewDeck[]> {
  const response = await fetch("/api/local-reviews");
  if (!response.ok) {
    throw new Error("Local review listing unavailable.");
  }
  const data = (await response.json()) as {
    decks: Array<{ deck: string; files: string[] }>;
  };
  return data.decks
    .map((entry) => ({
      deck: entry.deck,
      revs: extractRevs(entry.files),
      source: "local" as const,
    }))
    .filter((deck) => deck.revs.length > 0);
}

async function githubJson<T>(url: string, token: string): Promise<T> {
  const response = await fetch(url, {
    headers: {
      Accept: "application/vnd.github+json",
      Authorization: `Bearer ${token}`,
      "X-GitHub-Api-Version": "2022-11-28",
    },
  });
  if (!response.ok) {
    throw new Error(`GitHub API ${response.status}`);
  }
  return (await response.json()) as T;
}

function contentsUrl(pathname: string): string {
  const owner = import.meta.env.VITE_GITHUB_OWNER;
  const repo = import.meta.env.VITE_GITHUB_REPO;
  return `${apiRoot}/${owner}/${repo}/contents/${pathname}`;
}

function hasRepoCoordinates(): boolean {
  return Boolean(import.meta.env.VITE_GITHUB_OWNER && import.meta.env.VITE_GITHUB_REPO);
}

function extractRevs(files: string[]): string[] {
  return files
    .map((file) => file.match(/^rev-(\d+)-decisions\.tsv$/)?.[1])
    .filter((rev): rev is string => Boolean(rev))
    .sort();
}

function decodeGitHubContent(file: GitHubContent): string {
  if (file.encoding !== "base64" || !file.content) {
    throw new Error("Unsupported GitHub content encoding.");
  }
  const binary = atob(file.content.replace(/\s/g, ""));
  const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
  return new TextDecoder().decode(bytes);
}
