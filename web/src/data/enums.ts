export const REVIEW_STATUSES = [
  "unreviewed",
  "fix_required",
  "accepted",
  "false_positive",
  "fixed",
  "out_of_scope",
] as const;

export const OBSERVATION_DECISIONS = [
  "done",
  "inferred_done",
  "remaining",
  "not_recorded",
  "not_applicable",
] as const;

export const JUDGEMENT_REASONS = {
  accepted: [
    "intentional_template_design",
    "within_visual_tolerance",
    "decorative_only",
    "brand_approved_exception",
  ],
  fix_required: [
    "auto_fixable",
    "manual_layout_fix",
    "manual_content_fix",
    "master_template_fix",
    "requires_design_decision",
  ],
  fixed: ["fixed_in_later_artifact", "fixed_by_rule_update"],
  false_positive: ["lint_rule_too_strict", "measurement_error", "missing_context"],
  out_of_scope: [
    "master_owns",
    "different_distribution",
    "legacy_asset_frozen",
    "partner_owned",
  ],
} as const;

export type ReviewStatus = (typeof REVIEW_STATUSES)[number];
export type ObservationDecision = (typeof OBSERVATION_DECISIONS)[number];
export type JudgementReasonByStatus = typeof JUDGEMENT_REASONS;
export type DispositionStatus = keyof JudgementReasonByStatus;
export type JudgementReason = JudgementReasonByStatus[DispositionStatus][number];

export function isObservationDecision(value: string): value is ObservationDecision {
  return OBSERVATION_DECISIONS.includes(value as ObservationDecision);
}

export function isDispositionStatus(value: string): value is DispositionStatus {
  return Object.hasOwn(JUDGEMENT_REASONS, value);
}

export function reasonsForStatus(status: string): readonly string[] {
  return isDispositionStatus(status) ? JUDGEMENT_REASONS[status] : [];
}
