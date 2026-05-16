import { Encrypter, armor } from "age-encryption";

const FEEDBACK_ENDPOINT = "https://pptx-visual-review.pages.dev/api/feedback";
const APP_NAMESPACE = "pptx-design-review";
const AGE_PUBLIC_KEY =
  "age1l3vcgm69nwx5mgug58shf0qmhjtmg2frcawrz2jfagpq0vxawcvsxjc70v";

export type FeedbackPayload = {
  deck: string;
  rev: string;
  submittedBy?: string;
  decisions: unknown;
  findingJudgements: unknown;
};

export type SubmitResult = { key: string };

export async function submitFeedback(
  payload: FeedbackPayload,
): Promise<SubmitResult> {
  const plaintext = JSON.stringify({
    ...payload,
    submittedAt: new Date().toISOString(),
  });

  const enc = new Encrypter();
  enc.addRecipient(AGE_PUBLIC_KEY);
  const cipherBytes = await enc.encrypt(plaintext);
  const ciphertext = armor.encode(cipherBytes);

  const response = await fetch(FEEDBACK_ENDPOINT, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      app: APP_NAMESPACE,
      deck: payload.deck,
      rev: payload.rev,
      ciphertext,
    }),
  });

  if (!response.ok) {
    throw new Error(`Feedback submit failed: HTTP ${response.status}`);
  }
  const body = (await response.json()) as { ok?: boolean; key?: string; error?: string };
  if (!body.ok || !body.key) {
    throw new Error(`Feedback submit rejected: ${body.error ?? "unknown"}`);
  }
  return { key: body.key };
}
