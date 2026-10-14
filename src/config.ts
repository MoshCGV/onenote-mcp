import fs from "node:fs";
import os from "node:os";
import path from "node:path";

export type AuthFlow = "browser" | "device";

export interface Config {
  /** Entra (Azure AD) application (client) ID. Required unless accessToken is set. */
  clientId?: string;
  /** Tenant: "common", "consumers", "organizations", or a tenant ID/domain. */
  tenantId: string;
  /** How to sign in: system browser (auth code + PKCE) or device code. */
  authFlow: AuthFlow;
  /** Whether browser sign-in may launch the default browser (false: only print the link). */
  openBrowser: boolean;
  /** Delegated Graph scopes requested at sign-in. */
  scopes: string[];
  /** Directory holding the settings file and the MSAL token cache. */
  cacheDir: string;
  /** Graph API root, e.g. https://graph.microsoft.com/v1.0 */
  graphRoot: string;
  /** OneNote path under the Graph root, e.g. me/onenote or groups/{id}/onenote */
  basePath: string;
  /** Pre-obtained bearer token (skips MSAL). Mainly for testing. */
  accessToken?: string;
  /** Longest edge for a single, full-detail image returned to the model. */
  maxImagePx: number;
  /** Longest edge for images shown inline when reading a whole page. */
  overviewImagePx: number;
  /** Default number of images returned inline when reading a page. */
  maxImages: number;
  /** Byte budget for image data in a single tool response. */
  maxResponseImageBytes: number;
  /** Root folder that downloaded images/attachments are written under. */
  downloadDir: string;
  /** Hosts the bearer token may be sent to. */
  allowedHosts: Set<string>;
  /** Settings file that was read, if any. */
  settingsFile: string;
}

/** Settings that can be stored in ~/.onenote-mcp/config.json (environment variables override them). */
export interface FileSettings {
  clientId?: string;
  tenantId?: string;
  authFlow?: AuthFlow;
  scopes?: string;
  basePath?: string;
  downloadDir?: string;
}

const FILE_KEYS: (keyof FileSettings)[] = ["clientId", "tenantId", "authFlow", "scopes", "basePath", "downloadDir"];

function int(value: string | undefined, fallback: number, min: number, max: number): number {
  if (value === undefined || value.trim() === "") return fallback;
  const n = Number.parseInt(value, 10);
  if (!Number.isFinite(n)) return fallback;
  return Math.min(max, Math.max(min, n));
}

function expandHome(p: string): string {
  if (p === "~") return os.homedir();
  if (p.startsWith("~/")) return path.join(os.homedir(), p.slice(2));
  return p;
}

function pick(...values: (string | undefined)[]): string | undefined {
  for (const v of values) if (v !== undefined && v.trim() !== "") return v.trim();
  return undefined;
}

/** How to invoke this CLI, for messages: `onenote-mcp` when installed, else `node /path/dist/index.js`. */
export function cliCommand(): string {
  const entry = process.argv[1] ?? "";
  if (/[\\/]dist[\\/]index\.js$/.test(entry)) return `node "${entry}"`;
  return "onenote-mcp";
}

export function settingsPath(cacheDir: string): string {
  return path.join(cacheDir, "config.json");
}

export function resolveCacheDir(env: NodeJS.ProcessEnv = process.env): string {
  return path.resolve(expandHome(env.ONENOTE_CACHE_DIR ?? path.join(os.homedir(), ".onenote-mcp")));
}

export function readSettings(cacheDir: string): FileSettings {
  try {
    const raw = JSON.parse(fs.readFileSync(settingsPath(cacheDir), "utf8")) as Record<string, unknown>;
    const out: FileSettings = {};
    for (const key of FILE_KEYS) {
      const v = raw[key];
      if (typeof v === "string" && v.trim()) (out as Record<string, string>)[key] = v.trim();
    }
    return out;
  } catch (err) {
    if ((err as NodeJS.ErrnoException).code === "ENOENT") return {};
    throw new Error(`Could not read ${settingsPath(cacheDir)}: ${(err as Error).message}`);
  }
}

export function writeSettings(cacheDir: string, settings: FileSettings): string {
  const file = settingsPath(cacheDir);
  fs.mkdirSync(cacheDir, { recursive: true, mode: 0o700 });
  const clean: FileSettings = {};
  for (const key of FILE_KEYS) {
    const v = settings[key];
    if (v) (clean as Record<string, string>)[key] = v;
  }
  fs.writeFileSync(file, `${JSON.stringify(clean, null, 2)}\n`, { mode: 0o600 });
  return file;
}

function parseFlow(value: string | undefined): AuthFlow {
  if (!value) return "browser";
  const v = value.toLowerCase();
  if (v === "browser" || v === "device") return v;
  throw new Error(`Invalid sign-in method "${value}". Use "browser" or "device".`);
}

export function loadConfig(env: NodeJS.ProcessEnv = process.env): Config {
  const cacheDir = resolveCacheDir(env);
  const file = readSettings(cacheDir);

  const graphRoot = (env.ONENOTE_GRAPH_ROOT ?? "https://graph.microsoft.com/v1.0").replace(/\/+$/, "");
  const basePath = (pick(env.ONENOTE_BASE_PATH, file.basePath) ?? "me/onenote").replace(/^\/+|\/+$/g, "");
  const scopes = (pick(env.ONENOTE_SCOPES, file.scopes) ?? "Notes.ReadWrite")
    .split(/[\s,]+/)
    .map((s) => s.trim())
    .filter(Boolean);

  const allowedHosts = new Set<string>(["graph.microsoft.com"]);
  allowedHosts.add(new URL(graphRoot).host);

  const downloadDir = pick(env.ONENOTE_DOWNLOAD_DIR, file.downloadDir) ?? path.join(os.homedir(), "Downloads", "OneNote MCP");

  return {
    clientId: pick(env.ONENOTE_CLIENT_ID, file.clientId),
    tenantId: pick(env.ONENOTE_TENANT_ID, file.tenantId) ?? "common",
    authFlow: parseFlow(pick(env.ONENOTE_AUTH_FLOW, file.authFlow)),
    openBrowser: !/^(1|true|yes)$/i.test(env.ONENOTE_NO_BROWSER ?? ""),
    scopes,
    cacheDir,
    graphRoot,
    basePath,
    accessToken: env.ONENOTE_ACCESS_TOKEN?.trim() || undefined,
    maxImagePx: int(env.ONENOTE_MAX_IMAGE_PX, 1568, 256, 4096),
    overviewImagePx: int(env.ONENOTE_OVERVIEW_IMAGE_PX, 1024, 256, 4096),
    maxImages: int(env.ONENOTE_MAX_IMAGES, 6, 0, 50),
    maxResponseImageBytes: int(env.ONENOTE_MAX_RESPONSE_IMAGE_BYTES, 700_000, 100_000, 50_000_000),
    downloadDir: path.resolve(expandHome(downloadDir)),
    allowedHosts,
    settingsFile: settingsPath(cacheDir),
  };
}
