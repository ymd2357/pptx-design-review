import {
  serializeFindingDispositions,
  type FindingDisposition,
} from "./decisions-tsv";
import { isDispositionStatus, type ReviewStatus } from "./enums";
import type { LintFinding } from "./lint-json";

export type FindingJudgement = {
  review_status: ReviewStatus;
  judgement_reason: string | null;
  rationale?: string;
  updated_at?: string;
  updated_by?: string;
};

export type FindingJudgementsFile = {
  rev: string;
  deck: string;
  judgements: Record<string, FindingJudgement>;
};

export function findingJudgementsPath(deck: string, rev: string): string {
  return `doc/reviews/${deck}/rev-${rev}-finding-judgements.json`;
}

export function parseFindingJudgementsJson(
  value: unknown,
  deck: string,
  rev: string,
): FindingJudgementsFile {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return emptyFindingJudgements(deck, rev);
  }
  const record = value as Record<string, unknown>;
  const judgements: Record<string, FindingJudgement> = {};
  const rawJudgements = record.judgements;
  if (rawJudgements && typeof rawJudgements === "object" && !Array.isArray(rawJudgements)) {
    for (const [key, rawJudgement] of Object.entries(rawJudgements)) {
      const judgement = parseJudgement(rawJudgement);
      if (judgement) judgements[key] = judgement;
    }
  }
  return {
    rev: typeof record.rev === "string" ? record.rev : rev,
    deck: typeof record.deck === "string" ? record.deck : deck,
    judgements,
  };
}

export function serializeFindingJudgementsJson(file: FindingJudgementsFile): string {
  const ordered = Object.fromEntries(
    Object.entries(file.judgements).sort(([left], [right]) => left.localeCompare(right)),
  );
  return `${JSON.stringify({ ...file, judgements: ordered }, null, 2)}\n`;
}

export function emptyFindingJudgements(deck: string, rev: string): FindingJudgementsFile {
  return { deck, rev, judgements: {} };
}

export function initializeFindingJudgements(
  deck: string,
  rev: string,
  findings: readonly LintFinding[],
  existing?: FindingJudgementsFile,
): FindingJudgementsFile {
  const file: FindingJudgementsFile = {
    deck: existing?.deck ?? deck,
    rev: existing?.rev ?? rev,
    judgements: { ...(existing?.judgements ?? {}) },
  };
  for (const finding of findings) {
    file.judgements[finding.key] ??= {
      review_status: "unreviewed",
      judgement_reason: null,
    };
  }
  return file;
}

export function mergeLocalJudgementDrafts(
  file: FindingJudgementsFile,
  deck: string,
  rev: string,
  findings: readonly LintFinding[],
): FindingJudgementsFile {
  const next = initializeFindingJudgements(deck, rev, findings, file);
  for (const finding of findings) {
    const stored = localStorage.getItem(findingJudgementStorageKey(deck, rev, finding.key));
    if (!stored) continue;
    const judgement = parseJudgementFromJson(stored);
    if (judgement) next.judgements[finding.key] = judgement;
  }
  return next;
}

export function storeFindingJudgementDraft(
  deck: string,
  rev: string,
  groupKey: string,
  judgement: FindingJudgement,
): void {
  localStorage.setItem(
    findingJudgementStorageKey(deck, rev, groupKey),
    JSON.stringify(judgement),
  );
}

export function clearFindingJudgementDrafts(deck: string, rev: string): void {
  const prefix = `pptx-review:finding-judgements:${deck}:${rev}:`;
  for (let index = localStorage.length - 1; index >= 0; index -= 1) {
    const key = localStorage.key(index);
    if (key?.startsWith(prefix)) localStorage.removeItem(key);
  }
}

export function aggregateFindingDispositions(
  findings: readonly LintFinding[],
  judgements: FindingJudgementsFile,
  checkId: string,
): string {
  const counts = new Map<string, FindingDisposition>();
  for (const finding of findings) {
    if (finding.check !== checkId) continue;
    const judgement = judgements.judgements[finding.key];
    if (!judgement?.judgement_reason || !isDispositionStatus(judgement.review_status)) continue;
    const key = `${judgement.review_status}:${judgement.judgement_reason}`;
    const current = counts.get(key);
    if (current) {
      current.count += 1;
    } else {
      counts.set(key, {
        review_status: judgement.review_status,
        judgement_reason: judgement.judgement_reason,
        count: 1,
      });
    }
  }
  return serializeFindingDispositions(Array.from(counts.values()));
}

export function judgedFindingCount(
  findings: readonly LintFinding[],
  judgements: FindingJudgementsFile,
): number {
  return findings.filter((finding) => {
    const judgement = judgements.judgements[finding.key];
    return Boolean(judgement?.judgement_reason && judgement.review_status !== "unreviewed");
  }).length;
}

function findingJudgementStorageKey(deck: string, rev: string, groupKey: string): string {
  return `pptx-review:finding-judgements:${deck}:${rev}:${groupKey}`;
}

function parseJudgementFromJson(value: string): FindingJudgement | null {
  try {
    return parseJudgement(JSON.parse(value));
  } catch {
    return null;
  }
}

function parseJudgement(value: unknown): FindingJudgement | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const reviewStatus = record.review_status;
  if (
    reviewStatus !== "unreviewed" &&
    reviewStatus !== "fix_required" &&
    reviewStatus !== "accepted" &&
    reviewStatus !== "false_positive" &&
    reviewStatus !== "fixed" &&
    reviewStatus !== "out_of_scope"
  ) {
    return null;
  }
  return {
    review_status: reviewStatus,
    judgement_reason:
      typeof record.judgement_reason === "string" ? record.judgement_reason : null,
    rationale: typeof record.rationale === "string" ? record.rationale : undefined,
    updated_at: typeof record.updated_at === "string" ? record.updated_at : undefined,
    updated_by: typeof record.updated_by === "string" ? record.updated_by : undefined,
  };
}
