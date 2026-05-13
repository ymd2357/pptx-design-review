const storagePrefix = "pptx-review:";
const tokenKey = `${storagePrefix}github-token`;

export type DeviceCodeResponse = {
  device_code: string;
  user_code: string;
  verification_uri: string;
  expires_in: number;
  interval?: number;
};

export type DeviceFlowState =
  | { status: "code"; code: DeviceCodeResponse }
  | { status: "pending"; message: string }
  | { status: "authenticated" }
  | { status: "error"; message: string };

type AccessTokenResponse =
  | { access_token: string; token_type: string; scope: string }
  | { error: string; error_description?: string };

export function getStoredToken(): string | null {
  return localStorage.getItem(tokenKey);
}

export function clearStoredToken(): void {
  localStorage.removeItem(tokenKey);
}

export async function startDeviceFlow(
  onState: (state: DeviceFlowState) => void,
): Promise<void> {
  const clientId = import.meta.env.VITE_GITHUB_CLIENT_ID;
  if (!clientId) {
    onState({
      status: "error",
      message: "VITE_GITHUB_CLIENT_ID is not configured.",
    });
    return;
  }

  const codeResponse = await postForm<DeviceCodeResponse>(
    "https://github.com/login/device/code",
    {
      client_id: clientId,
      scope: "repo",
    },
  );
  onState({ status: "code", code: codeResponse });
  await pollForAccessToken(clientId, codeResponse, onState);
}

async function pollForAccessToken(
  clientId: string,
  code: DeviceCodeResponse,
  onState: (state: DeviceFlowState) => void,
): Promise<void> {
  let interval = code.interval ?? 5;
  const expiresAt = Date.now() + code.expires_in * 1000;

  while (Date.now() < expiresAt) {
    await sleep(interval * 1000);
    onState({ status: "pending", message: "Waiting for GitHub authorization..." });
    const response = await postForm<AccessTokenResponse>(
      "https://github.com/login/oauth/access_token",
      {
        client_id: clientId,
        device_code: code.device_code,
        grant_type: "urn:ietf:params:oauth:grant-type:device_code",
      },
    );

    if ("access_token" in response) {
      localStorage.setItem(tokenKey, response.access_token);
      onState({ status: "authenticated" });
      return;
    }

    if (response.error === "authorization_pending") continue;
    if (response.error === "slow_down") {
      interval += 5;
      continue;
    }

    onState({
      status: "error",
      message: response.error_description ?? response.error,
    });
    return;
  }

  onState({ status: "error", message: "GitHub device code expired." });
}

async function postForm<T>(url: string, params: Record<string, string>): Promise<T> {
  const body = new URLSearchParams(params);
  const response = await fetch(url, {
    method: "POST",
    headers: {
      Accept: "application/json",
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body,
  });
  if (!response.ok) {
    throw new Error(`GitHub OAuth request failed: ${response.status}`);
  }
  return (await response.json()) as T;
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}
