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

export const REVIEW_STATUS_LABELS: Record<string, string> = {
  unreviewed: "未判定",
  fix_required: "要修正",
  accepted: "許容",
  false_positive: "誤検知",
  fixed: "修正済",
  out_of_scope: "対象外",
};

export const OBSERVATION_DECISION_LABELS: Record<string, string> = {
  done: "完了",
  inferred_done: "推定完了 (lint 0 件)",
  remaining: "残あり",
  not_recorded: "未記入",
  not_applicable: "対象外",
};

export const JUDGEMENT_REASON_LABELS: Record<string, string> = {
  // accepted
  intentional_template_design: "テンプレ意図通り",
  within_visual_tolerance: "視覚的に許容範囲",
  decorative_only: "装飾のみ",
  brand_approved_exception: "ブランド承認済例外",
  // fix_required
  auto_fixable: "自動修正可",
  manual_layout_fix: "手動レイアウト修正",
  manual_content_fix: "手動本文修正",
  master_template_fix: "マスター/テンプレ修正",
  requires_design_decision: "デザイン判断要",
  // fixed
  fixed_in_later_artifact: "後続成果物で修正済",
  fixed_by_rule_update: "ルール更新で解消",
  // false_positive
  lint_rule_too_strict: "lint ルール過剰",
  measurement_error: "計測誤差",
  missing_context: "文脈不足の検出",
  // out_of_scope
  master_owns: "マスター起因",
  different_distribution: "別配布物の問題",
  legacy_asset_frozen: "凍結資産",
  partner_owned: "外部担当",
};

export function labelForReviewStatus(value: string): string {
  return REVIEW_STATUS_LABELS[value] ?? value;
}

export function labelForObservationDecision(value: string): string {
  return OBSERVATION_DECISION_LABELS[value] ?? value;
}

export function labelForJudgementReason(value: string): string {
  return JUDGEMENT_REASON_LABELS[value] ?? value;
}
