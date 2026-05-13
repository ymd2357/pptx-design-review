import {
  isDispositionStatus,
  isObservationDecision,
  type DispositionStatus,
  type ObservationDecision,
} from "./enums";

export const DECISION_COLUMNS = [
  "review_no",
  "check_id",
  "priority",
  "latest_lint_count",
  "observation_decision",
  "finding_dispositions",
  "rationale",
  "related_artifacts",
] as const;

export type DecisionColumn = (typeof DECISION_COLUMNS)[number];

export type FindingDisposition = {
  review_status: DispositionStatus;
  judgement_reason: string;
  count: number;
};

export type DecisionRow = Record<DecisionColumn, string> & {
  latest_lint_count: string;
  observation_decision: ObservationDecision | "";
};

export type ParsedDecisions = {
  rows: DecisionRow[];
  errors: string[];
};

export function parseDecisionTsv(tsv: string): ParsedDecisions {
  const lines = tsv.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
  const nonEmptyLines = lines.filter((line) => line.length > 0);
  if (nonEmptyLines.length === 0) {
    return { rows: [], errors: ["TSV is empty."] };
  }

  const header = nonEmptyLines[0].split("\t");
  const errors = DECISION_COLUMNS.filter((column) => !header.includes(column)).map(
    (column) => `Missing column: ${column}`,
  );

  const rows = nonEmptyLines.slice(1).map((line) => {
    const cells = line.split("\t");
    const row = Object.fromEntries(
      DECISION_COLUMNS.map((column) => [column, cells[header.indexOf(column)] ?? ""]),
    ) as DecisionRow;
    if (!isObservationDecision(row.observation_decision)) {
      row.observation_decision = row.observation_decision === "" ? "" : "not_recorded";
    }
    return row;
  });

  return { rows, errors };
}

export function serializeDecisionTsv(rows: DecisionRow[]): string {
  const body = rows.map((row) =>
    DECISION_COLUMNS.map((column) => sanitizeCell(row[column] ?? "")).join("\t"),
  );
  return `${DECISION_COLUMNS.join("\t")}\n${body.join("\n")}\n`;
}

export function parseFindingDispositions(value: string): FindingDisposition[] {
  if (!value.trim()) return [];
  return value
    .split(";")
    .map((part) => part.trim())
    .filter(Boolean)
    .map((part) => {
      const match = part.match(/^([^:]+):([^\s]+)\s+x(\d+)$/);
      if (!match || !isDispositionStatus(match[1])) {
        return null;
      }
      return {
        review_status: match[1],
        judgement_reason: match[2],
        count: Number.parseInt(match[3], 10),
      };
    })
    .filter((item): item is FindingDisposition => item !== null);
}

export function serializeFindingDispositions(dispositions: FindingDisposition[]): string {
  return dispositions
    .filter((item) => item.review_status && item.judgement_reason && item.count > 0)
    .map((item) => `${item.review_status}:${item.judgement_reason} x${item.count}`)
    .join("; ");
}

export function dispositionCount(row: DecisionRow): number {
  return parseFindingDispositions(row.finding_dispositions).reduce(
    (sum, item) => sum + item.count,
    0,
  );
}

export function validateDecisionRow(row: DecisionRow): string[] {
  const errors: string[] = [];
  if (row.observation_decision === "remaining") {
    const expected = Number.parseInt(row.latest_lint_count, 10);
    const actual = dispositionCount(row);
    if (!Number.isFinite(expected)) {
      errors.push("latest_lint_count is not a number.");
    } else if (actual !== expected) {
      errors.push(`finding_dispositions total must be ${expected}; currently ${actual}.`);
    }
  }
  return errors;
}

export function countChangedDecisionRows(before: DecisionRow[], after: DecisionRow[]): number {
  const beforeByKey = new Map(before.map((row) => [decisionRowKey(row), rowSignature(row)]));
  return after.filter((row) => beforeByKey.get(decisionRowKey(row)) !== rowSignature(row)).length;
}

function sanitizeCell(value: string): string {
  return String(value).replace(/\t/g, " ").replace(/\r?\n/g, " ").trim();
}

function decisionRowKey(row: DecisionRow): string {
  return `${row.review_no}\t${row.check_id}`;
}

function rowSignature(row: DecisionRow): string {
  return DECISION_COLUMNS.map((column) => row[column] ?? "").join("\t");
}
