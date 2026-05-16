import { sitePath } from "../site-path";

export type ReviewDeck = {
  deck: string;
  revs: string[];
  source: "local" | "mock";
};

export type ContentFile = {
  text: string;
  sha: string | undefined;
  source: "github" | "local";
  path: string;
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

const mockDecks: ReviewDeck[] = [
  {
    deck: "260329-seminar-curriculum-proposal",
    revs: ["017"],
    source: "mock",
  },
];

export async function listReviewDecks(): Promise<ReviewDeck[]> {
  try {
    return await listReviewDecksFromLocalDev();
  } catch {
    return mockDecks;
  }
}

export async function fetchDecisionTsv(deck: string, rev: string): Promise<ContentFile> {
  const filePath = `doc/reviews/${deck}/rev-${rev}-decisions.tsv`;
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

export async function fetchJsonFile<T>(path: string): Promise<JsonContentFile<T> | undefined> {
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

function extractRevs(files: string[]): string[] {
  return files
    .map((file) => file.match(/^rev-(\d+)-decisions\.tsv$/)?.[1])
    .filter((rev): rev is string => Boolean(rev))
    .sort();
}
