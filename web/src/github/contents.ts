import { clearStoredToken, getStoredToken } from "../auth/device-flow";
import { sitePath } from "../site-path";

const apiRoot = "https://api.github.com/repos";

export type ReviewDeck = {
  deck: string;
  revs: string[];
  source: "github" | "local" | "mock";
};

export type ContentFile = {
  text: string;
  sha: string | undefined;
  source: "github" | "local";
  path: string;
};

export type PutFileOptions = {
  path: string;
  message: string;
  content: string;
  sha?: string;
  branch?: string;
};

export type PutFileResult = {
  contentSha: string;
  commitSha: string;
  commitUrl: string;
};

export type ReviewSnapshot = {
  basePath: string;
  imageUrls: string[];
  lint: unknown | undefined;
  priorities: unknown | undefined;
};

export type JsonContentFile<T> = {
  data: T;
  text: string;
  sha: string | undefined;
  source: "github" | "local";
  path: string;
};

type GitHubContent = {
  name: string;
  path: string;
  type: "file" | "dir";
  sha: string;
  content?: string;
  encoding?: string;
};

type GitHubPutContentResponse = {
  content?: {
    sha: string;
  };
  commit: {
    sha: string;
    html_url: string;
  };
};

export class GitHubContentError extends Error {
  constructor(
    public readonly kind: "auth" | "conflict" | "api",
    message: string,
    public readonly status?: number,
  ) {
    super(message);
    this.name = "GitHubContentError";
  }
}

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
      const file = await getFile(filePath, token);
      return {
        text: decodeGitHubContent(file),
        sha: file.sha,
        source: "github",
        path: filePath,
      };
    } catch (error) {
      console.warn("GitHub TSV fetch failed; falling back to local data.", error);
    }
  }

  const response = await fetch(sitePath(filePath));
  if (!response.ok) {
    throw new Error(`Failed to load ${filePath}: ${response.status}`);
  }
  return { text: await response.text(), sha: undefined, source: "local", path: filePath };
}

export async function fetchReviewSnapshot(
  deck: string,
  rev: string,
): Promise<ReviewSnapshot | undefined> {
  const basePath = sitePath(
    `tmp/review-snapshot/${encodeURIComponent(deck)}/rev-${encodeURIComponent(rev)}`,
  );
  const [lint, priorities] = await Promise.all([
    fetchOptionalJson(`${basePath}/lint.json`),
    fetchOptionalJson(`${basePath}/priorities.json`),
  ]);
  const imageUrls = await discoverSnapshotImages(basePath, lint);
  if (!lint && !priorities && imageUrls.length === 0) return undefined;
  return { basePath, imageUrls, lint, priorities };
}

export async function getFile(path: string, token = getStoredToken()): Promise<GitHubContent> {
  if (!token) {
    throw new GitHubContentError("auth", "GitHub token is not available.");
  }
  return githubJson<GitHubContent>(contentsUrl(path), token);
}

export async function fetchJsonFile<T>(path: string): Promise<JsonContentFile<T> | undefined> {
  const token = getStoredToken();
  if (token && hasRepoCoordinates()) {
    try {
      const file = await getFile(path, token);
      const text = decodeGitHubContent(file);
      return {
        data: JSON.parse(text) as T,
        text,
        sha: file.sha,
        source: "github",
        path,
      };
    } catch (error) {
      if (error instanceof GitHubContentError && error.status === 404) return undefined;
      console.warn("GitHub JSON fetch failed; falling back to local data.", error);
    }
  }

  const response = await fetch(sitePath(path));
  if (response.status === 404) return undefined;
  if (!response.ok) {
    throw new Error(`Failed to load ${path}: ${response.status}`);
  }
  const text = await response.text();
  return {
    data: JSON.parse(text) as T,
    text,
    sha: undefined,
    source: "local",
    path,
  };
}

export async function putFile(opts: PutFileOptions): Promise<PutFileResult> {
  const token = getStoredToken();
  if (!token) {
    throw new GitHubContentError("auth", "GitHub token is not available.");
  }

  const body: Record<string, string> = {
    message: opts.message,
    content: encodeBase64(opts.content),
  };
  if (opts.sha) body.sha = opts.sha;
  if (opts.branch) body.branch = opts.branch;

  const response = await fetch(contentsUrl(opts.path), {
    method: "PUT",
    headers: {
      Accept: "application/vnd.github+json",
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      "X-GitHub-Api-Version": "2022-11-28",
    },
    body: JSON.stringify(body),
  });

  if (response.status === 409) {
    throw new GitHubContentError("conflict", "Conflict: file changed on the server.", 409);
  }
  if (response.status === 401 || response.status === 403) {
    clearStoredToken();
    throw new GitHubContentError("auth", `GitHub API ${response.status}`, response.status);
  }
  if (!response.ok) {
    throw new GitHubContentError("api", `GitHub API ${response.status}`, response.status);
  }

  const data = (await response.json()) as GitHubPutContentResponse;
  return {
    contentSha: data.content?.sha ?? opts.sha ?? "",
    commitSha: data.commit.sha,
    commitUrl: data.commit.html_url,
  };
}

export async function putJsonFile<T>(opts: {
  path: string;
  message: string;
  data: T;
  sha?: string;
  branch?: string;
}): Promise<PutFileResult> {
  return putFile({
    path: opts.path,
    message: opts.message,
    content: `${JSON.stringify(opts.data, null, 2)}\n`,
    sha: opts.sha,
    branch: opts.branch,
  });
}

export async function getAuthenticatedUserLogin(): Promise<string | undefined> {
  const token = getStoredToken();
  if (!token) return undefined;
  const user = await githubJson<{ login?: string }>("https://api.github.com/user", token);
  return user.login;
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
  const response = await fetch(sitePath("api/local-reviews"));
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

async function fetchOptionalJson(pathname: string): Promise<unknown | undefined> {
  const response = await fetch(pathname);
  if (response.status === 404) return undefined;
  if (!response.ok) {
    throw new Error(`Failed to load ${pathname}: ${response.status}`);
  }
  return response.json() as Promise<unknown>;
}

async function discoverSnapshotImages(basePath: string, lint: unknown): Promise<string[]> {
  const slideNumbers = extractSlideNumbers(lint);
  const candidates =
    slideNumbers.length > 0 ? slideNumbers : Array.from({ length: 40 }, (_value, index) => index + 1);
  const checks = await Promise.all(
    candidates.map(async (slideNo) => {
      const url = `${basePath}/images/slide-${String(slideNo).padStart(2, "0")}.png`;
      return (await assetExists(url)) ? url : undefined;
    }),
  );
  return checks.filter((url): url is string => Boolean(url));
}

async function assetExists(pathname: string): Promise<boolean> {
  const response = await fetch(pathname, { method: "HEAD" });
  return response.ok;
}

function extractSlideNumbers(value: unknown): number[] {
  const numbers = new Set<number>();
  const stack: unknown[] = [value];
  while (stack.length > 0) {
    const current = stack.pop();
    if (Array.isArray(current)) {
      stack.push(...current);
      continue;
    }
    if (!current || typeof current !== "object") continue;
    const record = current as Record<string, unknown>;
    addSlideNumber(numbers, record.slide_index, 1);
    addSlideNumber(numbers, record.example_slide_index, 1);
    addSlideNumber(numbers, record.slide, 0);
    addSlideArray(numbers, record.affected_slides);
    stack.push(...Object.values(record));
  }
  return Array.from(numbers).sort((a, b) => a - b);
}

function addSlideNumber(target: Set<number>, value: unknown, offset: number): void {
  if (typeof value !== "number" || !Number.isInteger(value)) return;
  const slideNo = value + offset;
  if (slideNo > 0 && slideNo < 1000) target.add(slideNo);
}

function addSlideArray(target: Set<number>, value: unknown): void {
  if (!Array.isArray(value)) return;
  for (const item of value) {
    addSlideNumber(target, item, 0);
  }
}

async function githubJson<T>(url: string, token: string): Promise<T> {
  const response = await fetch(url, {
    headers: {
      Accept: "application/vnd.github+json",
      Authorization: `Bearer ${token}`,
      "X-GitHub-Api-Version": "2022-11-28",
    },
  });
  if (response.status === 401 || response.status === 403) {
    clearStoredToken();
    throw new GitHubContentError("auth", `GitHub API ${response.status}`, response.status);
  }
  if (!response.ok) {
    throw new GitHubContentError("api", `GitHub API ${response.status}`, response.status);
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

function encodeBase64(value: string): string {
  const bytes = new TextEncoder().encode(value);
  let binary = "";
  for (const byte of bytes) {
    binary += String.fromCharCode(byte);
  }
  return btoa(binary);
}
