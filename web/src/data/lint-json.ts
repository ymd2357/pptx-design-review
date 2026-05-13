import { sitePath } from "../site-path";

export type BBoxPt = readonly [number, number, number, number];

export type LintFinding = {
  key: string;
  check: string;
  severity: string;
  slideIndex: number;
  slideNo: number;
  slideId: number | null;
  shapeId: number | string | null;
  shapeName: string;
  message: string;
  bboxPt: BBoxPt | null;
  actualBBoxPt: BBoxPt | null;
  measuredValue: unknown;
  detail: Record<string, unknown>;
  raw: Record<string, unknown>;
};

export type LintData = {
  findings: LintFinding[];
  slideNumbers: number[];
};

export type SlideSizePt = {
  w: number;
  h: number;
};

export async function loadSnapshotLint(deck: string, rev: string): Promise<LintData> {
  const lintPath = sitePath(
    `tmp/review-snapshot/${encodeURIComponent(deck)}/rev-${encodeURIComponent(rev)}/lint.json`,
  );
  const response = await fetch(lintPath);
  if (!response.ok) {
    throw new Error(`Failed to load lint.json: ${response.status}`);
  }
  return parseLintJson(await response.json());
}

export async function loadSlideSizePt(): Promise<SlideSizePt> {
  const response = await fetch(sitePath("doc/slide-guideline-v1.yml"));
  if (!response.ok) {
    throw new Error(`Failed to load slide guideline: ${response.status}`);
  }
  return parseSlideSizePt(await response.text());
}

export function parseLintJson(value: unknown): LintData {
  if (!Array.isArray(value)) {
    throw new Error("lint.json must be an array of findings.");
  }
  const findings = value.map(parseFinding).filter((finding): finding is LintFinding => Boolean(finding));
  const slideNumbers = Array.from(new Set(findings.map((finding) => finding.slideNo))).sort(
    (a, b) => a - b,
  );
  return { findings, slideNumbers };
}

export function findingsForObservation(
  findings: readonly LintFinding[],
  checkId: string,
): LintFinding[] {
  return findings.filter((finding) => finding.check === checkId);
}

export function parseSlideSizePt(yamlText: string): SlideSizePt {
  const lines = yamlText.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
  for (let index = 0; index < lines.length; index += 1) {
    if (!/^\s*size_pt:\s*$/.test(lines[index] ?? "")) continue;
    const indent = leadingSpaces(lines[index] ?? "");
    let width: number | undefined;
    let height: number | undefined;
    for (let child = index + 1; child < lines.length; child += 1) {
      const line = lines[child] ?? "";
      if (!line.trim() || line.trimStart().startsWith("#")) continue;
      if (leadingSpaces(line) <= indent) break;
      const wMatch = line.match(/^\s*w:\s*([0-9.]+)\s*$/);
      const hMatch = line.match(/^\s*h:\s*([0-9.]+)\s*$/);
      if (wMatch) width = Number.parseFloat(wMatch[1]);
      if (hMatch) height = Number.parseFloat(hMatch[1]);
      if (width !== undefined && height !== undefined) {
        return { w: width, h: height };
      }
    }
  }
  throw new Error("Could not read rules.slide.size_pt from guideline YAML.");
}

function parseFinding(value: unknown): LintFinding | null {
  if (!value || typeof value !== "object") return null;
  const raw = value as Record<string, unknown>;
  const detail = objectValue(raw.detail);
  const evidence = objectValue(detail.evidence);
  const target = objectValue(raw.target) ?? objectValue(detail.target) ?? objectValue(evidence.target);
  const check = stringValue(raw.check) ?? stringValue(detail.check_id);
  if (!check) return null;

  const slideIndex = numberValue(raw.slide_index) ?? numberValue(detail.example_slide_index) ?? 0;
  const slideNo =
    slideNumberFromPath(stringValue(detail.rendered_image_path) ?? stringValue(evidence.rendered_image_path)) ??
    slideIndex + 1;
  const shapeId =
    primitiveId(raw.shape_id) ??
    primitiveId(detail.example_shape_id) ??
    primitiveId(detail.shape_id) ??
    null;
  const groupKey =
    stringValue(raw.group_key) ??
    stringValue(detail.group_key) ??
    stringValue(evidence.group_key) ??
    fallbackKey(check, slideIndex, shapeId);

  return {
    key: groupKey,
    check,
    severity: stringValue(raw.severity) ?? "",
    slideIndex,
    slideNo,
    slideId: numberValue(raw.slide_id),
    shapeId,
    shapeName: stringValue(raw.shape_name) ?? "",
    message: stringValue(raw.message) ?? "",
    bboxPt:
      bboxValue(target?.bbox_pt) ??
      bboxValue(detail.bbox_pt) ??
      bboxValue(evidence.bbox_pt) ??
      bboxValue(detail.overlap_bbox_pt),
    actualBBoxPt:
      bboxValue(target?.actual_bbox_pt) ??
      bboxValue(detail.actual_bbox_pt) ??
      bboxValue(evidence.actual_bbox_pt),
    measuredValue: detail.measured_value ?? evidence.measured_value ?? detail.contrast_ratio ?? null,
    detail,
    raw,
  };
}

function fallbackKey(check: string, slideIndex: number, shapeId: number | string | null): string {
  return `${check}:${slideIndex}:${shapeId ?? "unknown"}`;
}

function bboxValue(value: unknown): BBoxPt | null {
  if (!Array.isArray(value) || value.length !== 4) return null;
  const numbers = value.map((item) => (typeof item === "number" ? item : Number.NaN));
  if (!numbers.every(Number.isFinite)) return null;
  return numbers as unknown as BBoxPt;
}

function objectValue(value: unknown): Record<string, unknown> {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : {};
}

function stringValue(value: unknown): string | undefined {
  return typeof value === "string" ? value : undefined;
}

function numberValue(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

function primitiveId(value: unknown): number | string | null {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "string" && value) return value;
  return null;
}

function leadingSpaces(value: string): number {
  return value.length - value.trimStart().length;
}

function slideNumberFromPath(value: string | undefined): number | null {
  const match = value?.match(/slide[-_ ]?0*(\d+)\.png/i);
  return match ? Number.parseInt(match[1], 10) : null;
}
