import type { Auth } from "./auth.js";
import type { Config } from "./config.js";

export class GraphError extends Error {
  constructor(
    message: string,
    public readonly status: number,
    public readonly code?: string,
    public readonly requestId?: string,
  ) {
    super(message);
    this.name = "GraphError";
  }
}

export type Query = Record<string, string | number | boolean | undefined>;

export interface RequestOptions {
  query?: Query;
  /** JSON-serialisable body, or a raw body with an explicit contentType. */
  json?: unknown;
  body?: Buffer | string;
  contentType?: string;
  accept?: string;
  timeoutMs?: number;
}

const RETRYABLE = new Set([429, 502, 503, 504]);
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

/** Hints for OneNote error codes that users commonly hit. */
const ONENOTE_HINTS: Record<string, string> = {
  "10008":
    "This account has too many OneNote items for a cross-notebook query. Narrow the request with a section_id or notebook_id.",
  "10013":
    "Too many sections for a cross-notebook query. Narrow the request with a section_id or notebook_id.",
  "20102": "The requested OneNote item does not exist or you do not have access to it.",
  "40004": "The notebook or section is read-only or you lack permission to write to it.",
};

/** Builds an OData query string, keeping `$` literal and encoding values with %20 for spaces. */
export function buildQuery(query?: Query): string {
  if (!query) return "";
  const parts: string[] = [];
  for (const [key, value] of Object.entries(query)) {
    if (value === undefined || value === "") continue;
    parts.push(`${key}=${encodeURIComponent(String(value))}`);
  }
  return parts.length ? `?${parts.join("&")}` : "";
}

/** Escapes a value for use inside an OData single-quoted string literal. */
export function odataString(value: string): string {
  return `'${value.replace(/'/g, "''")}'`;
}

export class GraphClient {
  constructor(
    private readonly cfg: Config,
    private readonly auth: Auth,
  ) {}

  /** Resolves a OneNote-relative path ("notebooks") or an absolute Graph URL. */
  resolve(pathOrUrl: string, query?: Query): string {
    if (/^https?:\/\//i.test(pathOrUrl)) {
      this.assertAllowed(pathOrUrl);
      return pathOrUrl + buildQuery(query);
    }
    const clean = pathOrUrl.replace(/^\/+/, "");
    return `${this.cfg.graphRoot}/${this.cfg.basePath}/${clean}${buildQuery(query)}`;
  }

  /** Refuse to send the bearer token anywhere other than Microsoft Graph (or the configured root). */
  assertAllowed(url: string): void {
    let host: string;
    try {
      host = new URL(url).host;
    } catch {
      throw new GraphError(`Invalid URL: ${url}`, 0);
    }
    if (!this.cfg.allowedHosts.has(host)) {
      throw new GraphError(
        `Refusing to send credentials to ${host}. Only Microsoft Graph URLs are fetched.`,
        0,
      );
    }
  }

  async request(method: string, pathOrUrl: string, opts: RequestOptions = {}): Promise<Response> {
    const url = this.resolve(pathOrUrl, opts.query);
    let refreshed = false;

    for (let attempt = 0; ; attempt++) {
      const token = await this.auth.getToken(refreshed);
      const headers: Record<string, string> = { Authorization: `Bearer ${token}` };
      if (opts.accept) headers.Accept = opts.accept;

      let body: Uint8Array<ArrayBuffer> | string | undefined;
      if (opts.json !== undefined) {
        headers["Content-Type"] = "application/json";
        body = JSON.stringify(opts.json);
      } else if (opts.body !== undefined) {
        headers["Content-Type"] = opts.contentType ?? "application/octet-stream";
        body = typeof opts.body === "string" ? opts.body : new Uint8Array(opts.body);
      }

      let res: Response;
      try {
        res = await fetch(url, {
          method,
          headers,
          body,
          redirect: "manual",
          signal: AbortSignal.timeout(opts.timeoutMs ?? 60_000),
        });
      } catch (err) {
        if (attempt < 3) {
          await sleep(500 * 2 ** attempt);
          continue;
        }
        const msg = err instanceof Error ? err.message : String(err);
        throw new GraphError(`Network error calling Microsoft Graph: ${msg}`, 0);
      }

      // Follow redirects ourselves so the Authorization header never leaks off-host.
      if (res.status >= 300 && res.status < 400 && res.headers.get("location")) {
        const location = new URL(res.headers.get("location")!, url).toString();
        const redirected = await fetch(location, {
          method: "GET",
          signal: AbortSignal.timeout(opts.timeoutMs ?? 60_000),
        });
        if (!redirected.ok) throw await toGraphError(redirected);
        return redirected;
      }

      if (res.status === 401 && !refreshed && !this.auth.usesStaticToken) {
        refreshed = true;
        continue;
      }

      if (RETRYABLE.has(res.status) && attempt < 4) {
        const retryAfter = Number(res.headers.get("retry-after"));
        const delay = Number.isFinite(retryAfter) && retryAfter > 0 ? retryAfter * 1000 : 1000 * 2 ** attempt;
        await sleep(Math.min(delay, 30_000));
        continue;
      }

      if (!res.ok) throw await toGraphError(res);
      return res;
    }
  }

  async json<T>(method: string, pathOrUrl: string, opts: RequestOptions = {}): Promise<T> {
    const res = await this.request(method, pathOrUrl, { accept: "application/json", ...opts });
    if (res.status === 204) return undefined as T;
    const text = await res.text();
    return (text ? JSON.parse(text) : undefined) as T;
  }

  async get<T>(pathOrUrl: string, query?: Query): Promise<T> {
    return this.json<T>("GET", pathOrUrl, { query });
  }

  async text(pathOrUrl: string, query?: Query, accept = "text/html"): Promise<string> {
    const res = await this.request("GET", pathOrUrl, { query, accept });
    return res.text();
  }

  async bytes(url: string): Promise<{ data: Buffer; contentType: string }> {
    const res = await this.request("GET", url, { timeoutMs: 120_000 });
    const data = Buffer.from(await res.arrayBuffer());
    return { data, contentType: res.headers.get("content-type") ?? "application/octet-stream" };
  }

  /** Follows @odata.nextLink until `limit` items are collected. */
  async collect<T>(pathOrUrl: string, query: Query | undefined, limit: number): Promise<T[]> {
    const items: T[] = [];
    let next: string | undefined = this.resolve(pathOrUrl, query);
    while (next && items.length < limit) {
      const page: { value?: T[]; "@odata.nextLink"?: string } = await this.get(next);
      items.push(...(page.value ?? []));
      next = page["@odata.nextLink"];
    }
    return items.slice(0, limit);
  }
}

async function toGraphError(res: Response): Promise<GraphError> {
  const requestId = res.headers.get("request-id") ?? res.headers.get("client-request-id") ?? undefined;
  let code: string | undefined;
  let message = `${res.status} ${res.statusText}`;
  try {
    const body = (await res.json()) as {
      error?: { code?: string; message?: string };
    };
    if (body?.error) {
      code = body.error.code;
      if (body.error.message) message = `${res.status}: ${body.error.message}`;
    }
  } catch {
    // Non-JSON error body; keep the status line.
  }

  const hint = code && ONENOTE_HINTS[code];
  if (hint) message += ` (${hint})`;
  if (res.status === 403) {
    message +=
      " Check that the app registration has the Notes.ReadWrite delegated permission and that you consented to it.";
  }
  return new GraphError(message, res.status, code, requestId);
}
