const storagePrefix = "pptx-review:";
const tokenKey = `${storagePrefix}github-token`;

export function getStoredToken(): string | null {
  return localStorage.getItem(tokenKey);
}

export function setStoredToken(token: string): void {
  localStorage.setItem(tokenKey, token);
}

export function clearStoredToken(): void {
  localStorage.removeItem(tokenKey);
}

export async function verifyToken(token: string): Promise<string> {
  const response = await fetch("https://api.github.com/user", {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28",
    },
  });
  if (!response.ok) {
    throw new Error(`GitHub /user returned ${response.status}`);
  }
  const data = (await response.json()) as { login?: string };
  if (!data.login) throw new Error("GitHub /user response missing login");
  return data.login;
}
