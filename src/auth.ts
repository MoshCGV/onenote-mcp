import { spawn } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";
import {
  InteractionRequiredAuthError,
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
  type ICachePlugin,
  type TokenCacheContext,
} from "@azure/msal-node";
import { cliCommand, type AuthFlow, type Config } from "./config.js";

export type LoginPrompt =
  | { kind: "device"; userCode: string; verificationUri: string; expiresAt: Date }
  | { kind: "browser"; url: string; opened: boolean; expiresAt: Date };

/** Thrown when no usable token exists. The message tells the user how to finish signing in. */
export class AuthRequiredError extends Error {
  constructor(
    message: string,
    public readonly prompt?: LoginPrompt,
  ) {
    super(message);
    this.name = "AuthRequiredError";
  }
}

export class ConfigError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ConfigError";
  }
}

/** Persists the MSAL token cache to a user-only JSON file. */
class FileCachePlugin implements ICachePlugin {
  constructor(private readonly file: string) {}

  async beforeCacheAccess(ctx: TokenCacheContext): Promise<void> {
    try {
      const data = await fs.readFile(this.file, "utf8");
      ctx.tokenCache.deserialize(data);
    } catch (err) {
      if ((err as NodeJS.ErrnoException).code !== "ENOENT") throw err;
    }
  }

  async afterCacheAccess(ctx: TokenCacheContext): Promise<void> {
    if (!ctx.cacheHasChanged) return;
    await fs.mkdir(path.dirname(this.file), { recursive: true, mode: 0o700 });
    const tmp = `${this.file}.${process.pid}.tmp`;
    await fs.writeFile(tmp, ctx.tokenCache.serialize(), { mode: 0o600 });
    await fs.rename(tmp, this.file);
  }
}

export function describeLogin(prompt: LoginPrompt): string {
  if (prompt.kind === "device") {
    const minutes = Math.max(1, Math.round((prompt.expiresAt.getTime() - Date.now()) / 60_000));
    return (
      `Microsoft sign-in required for OneNote.\n` +
      `1. Open ${prompt.verificationUri}\n` +
      `2. Enter the code: ${prompt.userCode}\n` +
      `3. Sign in with the Microsoft account that owns the notebooks.\n` +
      `The code expires in about ${minutes} minute(s). After signing in, retry the request.`
    );
  }
  return (
    `Microsoft sign-in required for OneNote.\n` +
    (prompt.opened
      ? "A browser window has opened on this computer. Sign in there with your work or school account, then retry the request.\n"
      : "Open this link in a browser on this computer and sign in, then retry the request:\n") +
    `If no window appeared, use this link: ${prompt.url}`
  );
}

/** Hints for Microsoft identity platform errors that usually mean an app-registration or policy issue. */
const AAD_HINTS: [RegExp, string][] = [
  [/AADSTS700016|AADSTS50059/, "The app ID was not found for this tenant. Check the client ID, and set the tenant to your organisation's tenant ID or domain (configure --tenant ...)."],
  [/AADSTS7000218|client_assertion|client_secret/, "Enable \"Allow public client flows\" in the app registration (Authentication page), then try again."],
  [/AADSTS50011|AADSTS500113|redirect/i, "Add the redirect URI http://localhost under \"Mobile and desktop applications\" in the app registration (Authentication page)."],
  [/AADSTS53003|AADSTS530033|conditional access/i, "Blocked by your organisation's Conditional Access policy. If you used device-code sign-in, switch to browser sign-in (configure --auth browser). Otherwise ask IT to allow this app or device."],
  [/AADSTS50105/, "Your account is not assigned to this app. Ask the app owner or IT to assign you, or turn off \"Assignment required\" on the enterprise app."],
  [/AADSTS90094|AADSTS90008|admin consent|AADSTS65001|consent/i, "Admin consent is required. Ask IT to grant admin consent for the Notes permission on this app registration."],
  [/AADSTS50020|AADSTS50194|AADSTS9002331/, "This account type is not allowed by the app registration. Use the tenant that owns the app, or change the app's supported account types."],
  [/AADSTS70011|invalid_scope/i, "A requested scope is invalid. Check the scopes setting (default Notes.ReadWrite)."],
  [/expired_token|code_expired/i, "The sign-in code expired before it was used. Start sign-in again."],
  [/authorization_declined|access_denied/i, "Sign-in was declined."],
];

export function hintFor(text: string): string | undefined {
  return AAD_HINTS.find(([re]) => re.test(text))?.[1];
}

/**
 * MSAL reports device-code failures as a bare "post_request_failed". Ask the endpoint directly so
 * the user sees Microsoft's actual reason (e.g. AADSTS700016) and a fix.
 */
async function explainDeviceCodeFailure(cfg: Config, err: unknown): Promise<string> {
  const base = err instanceof Error ? err.message : String(err);
  let detail = "";
  try {
    const res = await fetch(`https://login.microsoftonline.com/${encodeURIComponent(cfg.tenantId)}/oauth2/v2.0/devicecode`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ client_id: cfg.clientId ?? "", scope: [...cfg.scopes, "offline_access"].join(" ") }),
      signal: AbortSignal.timeout(15_000),
    });
    const body = (await res.json()) as { error?: string; error_description?: string };
    if (body.error) detail = (body.error_description ?? body.error).split(/\s*Trace ID:/)[0].trim();
  } catch {
    // Diagnostics are best effort.
  }
  const hint = hintFor(`${detail} ${base}`);
  return [`Microsoft sign-in failed: ${detail || base}`, hint].filter(Boolean).join("\n");
}

/** Opens a URL in the default browser without going through a shell. */
function openInBrowser(url: string): Promise<boolean> {
  const [cmd, args] =
    process.platform === "darwin"
      ? ["open", [url]]
      : process.platform === "win32"
        ? ["rundll32", ["url.dll,FileProtocolHandler", url]]
        : ["xdg-open", [url]];
  return new Promise((resolve) => {
    try {
      const child = spawn(cmd, args, { stdio: "ignore", detached: true });
      child.once("error", () => resolve(false));
      child.once("spawn", () => {
        child.unref();
        resolve(true);
      });
    } catch {
      resolve(false);
    }
  });
}

const SUCCESS_PAGE =
  "<html><body style=\"font-family:system-ui;padding:3em\"><h2>Signed in to OneNote MCP</h2>" +
  "<p>You can close this tab and go back to Claude.</p></body></html>";
const ERROR_PAGE =
  "<html><body style=\"font-family:system-ui;padding:3em\"><h2>OneNote MCP sign-in failed</h2>" +
  "<p>Go back to Claude for details.</p></body></html>";

const BROWSER_LOGIN_WINDOW_MS = 5 * 60_000;

export class Auth {
  private pca?: PublicClientApplication;
  private pending?: { prompt: LoginPrompt; done: Promise<AuthenticationResult | null> };
  private lastLoginError?: string;

  constructor(private readonly cfg: Config) {}

  get usesStaticToken(): boolean {
    return Boolean(this.cfg.accessToken);
  }

  get flow(): AuthFlow {
    return this.cfg.authFlow;
  }

  private client(): PublicClientApplication {
    if (!this.cfg.clientId) {
      throw new ConfigError(
        "OneNote MCP is not configured yet: no app (client) ID. Register an app in Microsoft Entra " +
          `(README: "Register an app"), then run: ${cliCommand()} configure --client-id <id> --tenant <tenant>. ` +
          "Alternatively set ONENOTE_CLIENT_ID in the MCP server config.",
      );
    }
    if (!this.pca) {
      this.pca = new PublicClientApplication({
        auth: {
          clientId: this.cfg.clientId,
          authority: `https://login.microsoftonline.com/${this.cfg.tenantId}`,
        },
        cache: { cachePlugin: new FileCachePlugin(path.join(this.cfg.cacheDir, "msal-cache.json")) },
      });
    }
    return this.pca;
  }

  private async account(): Promise<AccountInfo | undefined> {
    const accounts = await this.client().getTokenCache().getAllAccounts();
    return accounts[0];
  }

  /** Returns a Graph access token, or throws AuthRequiredError after starting a sign-in. */
  async getToken(forceRefresh = false): Promise<string> {
    if (this.cfg.accessToken) return this.cfg.accessToken;

    const account = await this.account();
    if (account) {
      try {
        const result = await this.client().acquireTokenSilent({
          account,
          scopes: this.cfg.scopes,
          forceRefresh,
        });
        if (result?.accessToken) return result.accessToken;
      } catch (err) {
        if (!(err instanceof InteractionRequiredAuthError)) throw err;
      }
    }

    const prompt = await this.startLogin();
    throw new AuthRequiredError(describeLogin(prompt), prompt);
  }

  /** Starts (or reuses) a sign-in. It completes in the background once the user finishes it. */
  async startLogin(flow: AuthFlow = this.cfg.authFlow): Promise<LoginPrompt> {
    if (this.pending && this.pending.prompt.kind === flow && this.pending.prompt.expiresAt.getTime() - Date.now() > 30_000) {
      return this.pending.prompt;
    }
    return flow === "browser" ? this.startBrowserLogin() : this.startDeviceLogin();
  }

  private track(done: Promise<AuthenticationResult | null>, reportFailures: () => boolean) {
    const tracked = done
      .then((result) => {
        this.lastLoginError = undefined;
        if (result?.account) console.error(`[onenote-mcp] Signed in as ${result.account.username}`);
        return result;
      })
      .catch((err: unknown) => {
        const msg = err instanceof Error ? err.message : String(err);
        const hint = hintFor(msg);
        this.lastLoginError = hint ? `${msg} (${hint})` : msg;
        if (reportFailures()) console.error(`[onenote-mcp] Sign-in did not complete: ${this.lastLoginError}`);
        return null;
      })
      .finally(() => {
        if (this.pending?.done === tracked) this.pending = undefined;
      });
    return tracked;
  }

  private async startBrowserLogin(): Promise<LoginPrompt> {
    const pca = this.client();
    let urlIssued = false;
    let resolvePrompt!: (p: LoginPrompt) => void;
    const promptPromise = new Promise<LoginPrompt>((resolve) => (resolvePrompt = resolve));

    const done = pca.acquireTokenInteractive({
      scopes: this.cfg.scopes,
      prompt: "select_account",
      successTemplate: SUCCESS_PAGE,
      errorTemplate: ERROR_PAGE,
      openBrowser: async (url: string) => {
        urlIssued = true;
        const opened = this.cfg.openBrowser ? await openInBrowser(url) : false;
        resolvePrompt({ kind: "browser", url, opened, expiresAt: new Date(Date.now() + BROWSER_LOGIN_WINDOW_MS) });
      },
    });
    const tracked = this.track(done, () => urlIssued);

    const prompt = await Promise.race([
      promptPromise,
      done.then(
        () => {
          throw new Error("Sign-in finished before a browser was opened.");
        },
        (err: unknown) => {
          const msg = err instanceof Error ? err.message : String(err);
          const hint = hintFor(msg);
          throw new ConfigError(`Microsoft sign-in failed: ${msg}${hint ? `\n${hint}` : ""}`);
        },
      ),
    ]);
    this.pending = { prompt, done: tracked };
    return prompt;
  }

  private async startDeviceLogin(): Promise<LoginPrompt> {
    const pca = this.client();
    let codeIssued = false;
    let resolvePrompt!: (p: LoginPrompt) => void;
    const promptPromise = new Promise<LoginPrompt>((resolve) => (resolvePrompt = resolve));

    const done = pca.acquireTokenByDeviceCode({
      scopes: this.cfg.scopes,
      deviceCodeCallback: (response) => {
        // MSAL invokes the callback with an empty object when the endpoint returns an error.
        if (!response?.userCode || !response.verificationUri) return;
        codeIssued = true;
        resolvePrompt({
          kind: "device",
          userCode: response.userCode,
          verificationUri: response.verificationUri,
          expiresAt: new Date(Date.now() + response.expiresIn * 1000),
        });
      },
    });
    const tracked = this.track(done, () => codeIssued);

    // If MSAL fails before issuing a code (bad client ID, blocked flow), surface the reason.
    const prompt = await Promise.race([
      promptPromise,
      done.then(
        () => {
          throw new Error("Sign-in finished before a device code was issued.");
        },
        async (err: unknown) => {
          throw new ConfigError(await explainDeviceCodeFailure(this.cfg, err));
        },
      ),
    ]);
    this.pending = { prompt, done: tracked };
    return prompt;
  }

  /** Waits for an in-progress sign-in (CLI use). */
  async waitForLogin(): Promise<AccountInfo | null> {
    if (!this.pending) return (await this.account()) ?? null;
    const result = await this.pending.done;
    return result?.account ?? null;
  }

  async status(): Promise<{
    signedIn: boolean;
    username?: string;
    pending?: LoginPrompt;
    lastError?: string;
    mode: "static-token" | "msal";
  }> {
    if (this.cfg.accessToken) return { signedIn: true, username: "(static access token)", mode: "static-token" };
    const account = await this.account();
    return {
      signedIn: Boolean(account),
      username: account?.username,
      pending: this.pending?.prompt,
      lastError: this.lastLoginError,
      mode: "msal",
    };
  }

  async signOut(): Promise<number> {
    if (this.cfg.accessToken) return 0;
    const cache = this.client().getTokenCache();
    const accounts = await cache.getAllAccounts();
    for (const account of accounts) await cache.removeAccount(account);
    this.pending = undefined;
    return accounts.length;
  }
}
