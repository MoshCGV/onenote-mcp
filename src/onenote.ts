import { GraphClient, GraphError, odataString, type Query } from "./graph.js";
import { parsePageHtml, type ParsedPage } from "./html.js";

export interface Link {
  href?: string;
}

export interface Notebook {
  id: string;
  displayName: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  isDefault?: boolean;
  isShared?: boolean;
  links?: { oneNoteClientUrl?: Link; oneNoteWebUrl?: Link };
  sections?: Section[];
  sectionGroups?: SectionGroup[];
}

export interface SectionGroup {
  id: string;
  displayName: string;
  sections?: Section[];
  sectionGroups?: SectionGroup[];
  parentNotebook?: { id: string; displayName?: string };
}

export interface Section {
  id: string;
  displayName: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  isDefault?: boolean;
  links?: { oneNoteClientUrl?: Link; oneNoteWebUrl?: Link };
  parentNotebook?: { id: string; displayName?: string };
  parentSectionGroup?: { id: string; displayName?: string } | null;
}

export interface Page {
  id: string;
  title?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  level?: number;
  order?: number;
  contentUrl?: string;
  links?: { oneNoteClientUrl?: Link; oneNoteWebUrl?: Link };
  parentSection?: { id: string; displayName?: string };
  parentNotebook?: { id: string; displayName?: string };
}

export interface PageFilter {
  sectionId?: string;
  notebookId?: string;
  titleContains?: string;
  modifiedSince?: string;
  modifiedBefore?: string;
  limit: number;
  order?: "modified_desc" | "modified_asc" | "created_desc" | "created_asc" | "title";
}

export interface BinaryPart {
  name: string;
  contentType: string;
  data: Buffer;
}

const PAGE_SELECT = "id,title,createdDateTime,lastModifiedDateTime,links,contentUrl";

const ORDER_BY: Record<NonNullable<PageFilter["order"]>, string> = {
  modified_desc: "lastModifiedDateTime desc",
  modified_asc: "lastModifiedDateTime asc",
  created_desc: "createdDateTime desc",
  created_asc: "createdDateTime asc",
  title: "title asc",
};

/** Small LRU cache with an approximate byte budget. */
class LruCache<V> {
  private map = new Map<string, { value: V; size: number; at: number }>();
  private total = 0;
  constructor(
    private readonly maxBytes: number,
    private readonly ttlMs: number,
  ) {}

  get(key: string): V | undefined {
    const hit = this.map.get(key);
    if (!hit) return undefined;
    if (Date.now() - hit.at > this.ttlMs) {
      this.delete(key);
      return undefined;
    }
    this.map.delete(key);
    this.map.set(key, hit);
    return hit.value;
  }

  set(key: string, value: V, size: number): void {
    this.delete(key);
    this.map.set(key, { value, size, at: Date.now() });
    this.total += size;
    while (this.total > this.maxBytes && this.map.size > 1) {
      const oldest = this.map.keys().next().value as string;
      this.delete(oldest);
    }
  }

  delete(key: string): void {
    const hit = this.map.get(key);
    if (!hit) return;
    this.total -= hit.size;
    this.map.delete(key);
  }

  clear(): void {
    this.map.clear();
    this.total = 0;
  }
}

/** Runs `fn` over items with bounded concurrency, preserving order. */
export async function mapLimit<T, R>(items: T[], limit: number, fn: (item: T, i: number) => Promise<R>): Promise<R[]> {
  const out = new Array<R>(items.length);
  let next = 0;
  const workers = Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (next < items.length) {
      const i = next++;
      out[i] = await fn(items[i], i);
    }
  });
  await Promise.all(workers);
  return out;
}

function toIsoDate(value: string, field: string): string {
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) throw new Error(`${field} must be an ISO date or date-time, got "${value}"`);
  return d.toISOString();
}

export class OneNote {
  private htmlCache = new LruCache<string>(40_000_000, 5 * 60_000);
  private parsedCache = new LruCache<ParsedPage>(40_000_000, 5 * 60_000);
  private resourceCache = new LruCache<{ data: Buffer; contentType: string }>(150_000_000, 15 * 60_000);

  constructor(readonly graph: GraphClient) {}

  // ---------- Structure ----------

  async listNotebooks(expandTree: boolean): Promise<Notebook[]> {
    if (!expandTree) return this.graph.collect<Notebook>("notebooks", { $top: 100 }, 1000);
    try {
      return await this.graph.collect<Notebook>(
        "notebooks",
        {
          $expand: "sections,sectionGroups($expand=sections,sectionGroups($levels=max;$expand=sections))",
        },
        1000,
      );
    } catch (err) {
      if (!(err instanceof GraphError) || err.status !== 400) throw err;
      // Some tenants reject $levels; fall back to two levels of nesting.
      return this.graph.collect<Notebook>(
        "notebooks",
        { $expand: "sections,sectionGroups($expand=sections)" },
        1000,
      );
    }
  }

  async listSections(opts: { notebookId?: string; sectionGroupId?: string; nameContains?: string }): Promise<Section[]> {
    const query: Query = { $top: 100 };
    if (opts.nameContains) {
      query.$filter = `contains(tolower(displayName),${odataString(opts.nameContains.toLowerCase())})`;
    }
    if (opts.sectionGroupId) {
      return this.graph.collect<Section>(`sectionGroups/${encodeURIComponent(opts.sectionGroupId)}/sections`, query, 2000);
    }
    if (opts.notebookId) {
      // Include sections nested inside section groups.
      const notebook = await this.getNotebookTree(opts.notebookId);
      const all = flattenSections(notebook);
      const needle = opts.nameContains?.toLowerCase();
      return needle ? all.filter((s) => s.displayName.toLowerCase().includes(needle)) : all;
    }
    return this.graph.collect<Section>("sections", { ...query, $expand: "parentNotebook,parentSectionGroup" }, 2000);
  }

  async getNotebookTree(notebookId: string): Promise<Notebook> {
    const id = encodeURIComponent(notebookId);
    try {
      return await this.graph.get<Notebook>(`notebooks/${id}`, {
        $expand: "sections,sectionGroups($expand=sections,sectionGroups($levels=max;$expand=sections))",
      });
    } catch (err) {
      if (!(err instanceof GraphError) || err.status !== 400) throw err;
      return this.graph.get<Notebook>(`notebooks/${id}`, { $expand: "sections,sectionGroups($expand=sections)" });
    }
  }

  async createNotebook(name: string): Promise<Notebook> {
    return this.graph.json<Notebook>("POST", "notebooks", { json: { displayName: name } });
  }

  async createSection(opts: { notebookId?: string; sectionGroupId?: string; name: string }): Promise<Section> {
    const parent = opts.sectionGroupId
      ? `sectionGroups/${encodeURIComponent(opts.sectionGroupId)}`
      : `notebooks/${encodeURIComponent(opts.notebookId!)}`;
    return this.graph.json<Section>("POST", `${parent}/sections`, { json: { displayName: opts.name } });
  }

  /** Finds a section by notebook/section display names, optionally creating the section. */
  async resolveSection(opts: {
    notebookName?: string;
    sectionName: string;
    create: boolean;
  }): Promise<{ section: Section; created: boolean }> {
    const wanted = opts.sectionName.trim().toLowerCase();

    if (!opts.notebookName) {
      const matches = (await this.listSections({})).filter((s) => s.displayName.trim().toLowerCase() === wanted);
      if (matches.length === 1) return { section: matches[0], created: false };
      if (matches.length > 1) {
        const where = matches.map((s) => `"${s.parentNotebook?.displayName ?? "?"}"`).join(", ");
        throw new Error(`Section "${opts.sectionName}" exists in several notebooks (${where}). Pass notebook_name or section_id.`);
      }
      throw new Error(`No section named "${opts.sectionName}". Pass notebook_name to create it, or list sections first.`);
    }

    const notebooks = await this.listNotebooks(false);
    const nbWanted = opts.notebookName.trim().toLowerCase();
    let notebook = notebooks.find((n) => n.displayName.trim().toLowerCase() === nbWanted);
    if (!notebook) {
      const partial = notebooks.filter((n) => n.displayName.toLowerCase().includes(nbWanted));
      if (partial.length === 1) notebook = partial[0];
    }
    if (!notebook) {
      const names = notebooks.map((n) => `"${n.displayName}"`).join(", ") || "(none)";
      throw new Error(`No notebook named "${opts.notebookName}". Available notebooks: ${names}`);
    }

    const tree = await this.getNotebookTree(notebook.id);
    const section = flattenSections(tree).find((s) => s.displayName.trim().toLowerCase() === wanted);
    if (section) return { section, created: false };
    if (!opts.create) {
      throw new Error(
        `Notebook "${notebook.displayName}" has no section "${opts.sectionName}". Set create_section_if_missing to true to create it.`,
      );
    }
    const created = await this.createSection({ notebookId: notebook.id, name: opts.sectionName.trim() });
    return { section: created, created: true };
  }

  // ---------- Pages ----------

  private pageQuery(f: PageFilter, top: number): Query {
    const filters: string[] = [];
    if (f.titleContains) filters.push(`contains(tolower(title),${odataString(f.titleContains.toLowerCase())})`);
    if (f.modifiedSince) filters.push(`lastModifiedDateTime ge ${toIsoDate(f.modifiedSince, "modified_since")}`);
    if (f.modifiedBefore) filters.push(`lastModifiedDateTime lt ${toIsoDate(f.modifiedBefore, "modified_before")}`);
    return {
      $select: PAGE_SELECT,
      $expand: "parentSection,parentNotebook",
      $orderby: ORDER_BY[f.order ?? "modified_desc"],
      $filter: filters.join(" and ") || undefined,
      $top: Math.min(100, Math.max(1, top)),
    };
  }

  async listPages(f: PageFilter): Promise<Page[]> {
    if (f.sectionId) {
      return this.graph.collect<Page>(
        `sections/${encodeURIComponent(f.sectionId)}/pages`,
        this.pageQuery(f, f.limit),
        f.limit,
      );
    }
    if (f.notebookId) {
      // Querying section-by-section avoids the cross-notebook limits on /pages.
      const tree = await this.getNotebookTree(f.notebookId);
      const sections = flattenSections(tree);
      const perSection = await mapLimit(sections, 4, (s) =>
        this.graph.collect<Page>(`sections/${encodeURIComponent(s.id)}/pages`, this.pageQuery(f, f.limit), f.limit),
      );
      return sortPages(perSection.flat(), f.order ?? "modified_desc").slice(0, f.limit);
    }
    return this.graph.collect<Page>("pages", this.pageQuery(f, f.limit), f.limit);
  }

  async getPageMeta(pageId: string): Promise<Page> {
    return this.graph.get<Page>(`pages/${encodeURIComponent(pageId)}`, {
      $select: PAGE_SELECT,
      $expand: "parentSection,parentNotebook",
    });
  }

  async getPageHtml(pageId: string, includeIds = false): Promise<string> {
    const key = `${pageId}|${includeIds}`;
    const cached = this.htmlCache.get(key);
    if (cached !== undefined) return cached;
    const html = await this.graph.text(
      `pages/${encodeURIComponent(pageId)}/content`,
      includeIds ? { includeIDs: "true" } : undefined,
    );
    this.htmlCache.set(key, html, html.length * 2);
    return html;
  }

  async getParsedPage(pageId: string): Promise<ParsedPage> {
    const cached = this.parsedCache.get(pageId);
    if (cached) return cached;
    const parsed = parsePageHtml(await this.getPageHtml(pageId));
    this.parsedCache.set(pageId, parsed, parsed.markdown.length * 2 + 4096);
    return parsed;
  }

  async getResource(url: string): Promise<{ data: Buffer; contentType: string }> {
    const cached = this.resourceCache.get(url);
    if (cached) return cached;
    const res = await this.graph.bytes(url);
    this.resourceCache.set(url, res, res.data.length);
    return res;
  }

  async createPage(sectionId: string, html: string, parts: BinaryPart[]): Promise<Page> {
    const path = `sections/${encodeURIComponent(sectionId)}/pages`;
    const page =
      parts.length === 0
        ? await this.graph.json<Page>("POST", path, { body: html, contentType: "text/html" })
        : await this.postMultipart<Page>("POST", path, { name: "Presentation", contentType: "text/html", data: Buffer.from(html, "utf8") }, parts);
    return page;
  }

  async updatePage(pageId: string, commands: unknown[], parts: BinaryPart[]): Promise<void> {
    const path = `pages/${encodeURIComponent(pageId)}/content`;
    if (parts.length === 0) {
      await this.graph.json<void>("PATCH", path, { json: commands });
    } else {
      await this.postMultipart<void>(
        "PATCH",
        path,
        { name: "Commands", contentType: "application/json", data: Buffer.from(JSON.stringify(commands), "utf8") },
        parts,
      );
    }
    this.invalidate(pageId);
  }

  clearCaches(): void {
    this.htmlCache.clear();
    this.parsedCache.clear();
    this.resourceCache.clear();
  }

  invalidate(pageId: string): void {
    this.htmlCache.delete(`${pageId}|false`);
    this.htmlCache.delete(`${pageId}|true`);
    this.parsedCache.delete(pageId);
  }

  private async postMultipart<T>(method: string, path: string, first: BinaryPart, parts: BinaryPart[]): Promise<T> {
    const boundary = `onenote-mcp-${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
    const chunks: Buffer[] = [];
    for (const part of [first, ...parts]) {
      chunks.push(
        Buffer.from(
          `--${boundary}\r\nContent-Disposition: form-data; name="${part.name}"\r\nContent-Type: ${part.contentType}\r\n\r\n`,
          "utf8",
        ),
        part.data,
        Buffer.from("\r\n", "utf8"),
      );
    }
    chunks.push(Buffer.from(`--${boundary}--\r\n`, "utf8"));
    return this.graph.json<T>(method, path, {
      body: Buffer.concat(chunks),
      contentType: `multipart/form-data; boundary=${boundary}`,
      timeoutMs: 180_000,
    });
  }
}

export function flattenSections(node: { sections?: Section[]; sectionGroups?: SectionGroup[] }, trail: string[] = []): (Section & { path: string })[] {
  const out: (Section & { path: string })[] = [];
  for (const s of node.sections ?? []) out.push({ ...s, path: [...trail, s.displayName].join(" / ") });
  for (const g of node.sectionGroups ?? []) out.push(...flattenSections(g, [...trail, g.displayName]));
  return out;
}

export function sortPages(pages: Page[], order: NonNullable<PageFilter["order"]>): Page[] {
  const by = (p: Page, k: "lastModifiedDateTime" | "createdDateTime") => (p[k] ? Date.parse(p[k]!) : 0);
  const sorted = [...pages];
  switch (order) {
    case "modified_asc":
      return sorted.sort((a, b) => by(a, "lastModifiedDateTime") - by(b, "lastModifiedDateTime"));
    case "created_desc":
      return sorted.sort((a, b) => by(b, "createdDateTime") - by(a, "createdDateTime"));
    case "created_asc":
      return sorted.sort((a, b) => by(a, "createdDateTime") - by(b, "createdDateTime"));
    case "title":
      return sorted.sort((a, b) => (a.title ?? "").localeCompare(b.title ?? ""));
    default:
      return sorted.sort((a, b) => by(b, "lastModifiedDateTime") - by(a, "lastModifiedDateTime"));
  }
}
