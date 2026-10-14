import * as cheerio from "cheerio";
import { marked } from "marked";
import TurndownService from "turndown";

export interface PageImage {
  /** 1-based position on the page, in reading order. */
  index: number;
  src?: string;
  srcType?: string;
  fullresSrc?: string;
  fullresType?: string;
  alt?: string;
  width?: number;
  height?: number;
  dataId?: string;
  /** Set for printouts / split images rendered from one source. */
  splitIndex?: number;
  options?: string;
}

export interface PageAttachment {
  index: number;
  name: string;
  url?: string;
  type?: string;
}

export interface ParsedPage {
  title: string;
  created?: string;
  markdown: string;
  images: PageImage[];
  attachments: PageAttachment[];
}

const NOTE_TAG_LABELS: Record<string, string> = {
  "to-do": "[ ]",
  "to-do:completed": "[x]",
  important: "[important]",
  question: "[question]",
  critical: "[critical]",
  idea: "[idea]",
  "remember-for-later": "[remember]",
  definition: "[definition]",
  highlight: "[highlight]",
  contact: "[contact]",
  address: "[address]",
  "phone-number": "[phone]",
  "web-site-to-visit": "[website]",
  password: "[password]",
  project: "[project]",
};

function numAttr(value: string | undefined): number | undefined {
  if (!value) return undefined;
  const n = Number.parseFloat(value);
  return Number.isFinite(n) ? n : undefined;
}

function stylePx(style: string | undefined, prop: string): number | undefined {
  if (!style) return undefined;
  const m = new RegExp(`(?:^|;)\\s*${prop}\\s*:\\s*(-?[\\d.]+)px`, "i").exec(style);
  return m ? Number.parseFloat(m[1]) : undefined;
}

function hasStyle(style: string | null | undefined, re: RegExp): boolean {
  return Boolean(style && re.test(style));
}

function makeTurndown(): TurndownService {
  const td = new TurndownService({
    headingStyle: "atx",
    bulletListMarker: "-",
    codeBlockStyle: "fenced",
    emDelimiter: "_",
    strongDelimiter: "**",
  });
  // Output goes to a model, not a markdown renderer: keep text verbatim instead of backslash-escaping.
  td.escape = (s: string) => s;

  td.remove(["script", "style", "head", "meta", "title"]);

  td.addRule("onenoteImage", {
    filter: "img",
    replacement: (_content, node) => {
      const el = node as HTMLElement;
      const index = el.getAttribute("data-mcp-index");
      const alt = el.getAttribute("alt")?.trim();
      const w = el.getAttribute("width");
      const h = el.getAttribute("height");
      const size = w && h ? ` (${w}x${h})` : "";
      return `\n\n[Image ${index ?? "?"}${alt ? `: ${alt}` : ""}${size}]\n\n`;
    },
  });

  td.addRule("onenoteVideo", {
    filter: "iframe",
    replacement: (_content, node) => {
      const el = node as HTMLElement;
      const src = el.getAttribute("data-original-src") ?? el.getAttribute("src") ?? "";
      return `\n\n[Embedded video: ${src}]\n\n`;
    },
  });

  td.addRule("styledSpan", {
    filter: (node) => {
      if (node.nodeName !== "SPAN") return false;
      const style = (node as HTMLElement).getAttribute("style");
      return hasStyle(style, /font-weight\s*:\s*bold|font-style\s*:\s*italic|line-through/i);
    },
    replacement: (content, node) => {
      const style = (node as HTMLElement).getAttribute("style") ?? "";
      const inner = content.trim();
      if (!inner) return content;
      let out = inner;
      if (/line-through/i.test(style)) out = `~~${out}~~`;
      if (/font-style\s*:\s*italic/i.test(style)) out = `_${out}_`;
      if (/font-weight\s*:\s*bold/i.test(style)) out = `**${out}**`;
      const lead = content.match(/^\s*/)?.[0] ?? "";
      const trail = content.match(/\s*$/)?.[0] ?? "";
      return lead + out + trail;
    },
  });

  td.addRule("strike", {
    filter: ["del", "s", "strike"] as (keyof HTMLElementTagNameMap)[],
    replacement: (content) => (content.trim() ? `~~${content.trim()}~~` : content),
  });

  td.addRule("tableCell", {
    filter: ["th", "td"],
    replacement: (content, node) => {
      let prev = node.previousSibling;
      while (prev && prev.nodeType !== 1) prev = prev.previousSibling;
      const prefix = prev ? " " : "| ";
      const text = content.trim().replace(/\n+/g, "<br>").replace(/\|/g, "\\|");
      return `${prefix}${text} |`;
    },
  });

  td.addRule("tableRow", {
    filter: "tr",
    replacement: (content, node) => {
      let out = `\n${content}`;
      if ((node as HTMLElement).getAttribute("data-mcp-first") === "1") {
        let cells = 0;
        for (const child of Array.from(node.childNodes)) {
          if (child.nodeName === "TD" || child.nodeName === "TH") cells++;
        }
        out += `\n|${" --- |".repeat(Math.max(1, cells))}`;
      }
      return out;
    },
  });

  td.addRule("tableSection", {
    filter: ["thead", "tbody", "tfoot"],
    replacement: (content) => content,
  });

  td.addRule("table", {
    filter: "table",
    replacement: (content) => `\n\n${content.trim()}\n\n`,
  });

  return td;
}

const turndown = makeTurndown();

/** Parses OneNote output HTML into markdown plus image/attachment inventories. */
export function parsePageHtml(html: string): ParsedPage {
  const $ = cheerio.load(html);
  const title = $("head > title").first().text().trim();
  const created = $('meta[name="created"]').attr("content");
  const body = $("body");

  // OneNote positions outlines absolutely; order them top-to-bottom, then left-to-right.
  const blocks = body.children().toArray();
  const keyed = blocks.map((el, i) => {
    const style = $(el).attr("style");
    return { el, i, top: stylePx(style, "top") ?? Number.MAX_SAFE_INTEGER, left: stylePx(style, "left") ?? 0 };
  });
  const positioned = keyed.some((k) => k.top !== Number.MAX_SAFE_INTEGER);
  if (positioned && blocks.length > 1) {
    keyed.sort((a, b) => a.top - b.top || a.left - b.left || a.i - b.i);
    for (const k of keyed) body.append(k.el);
  }

  // Note tags (to-do checkboxes, important, etc.).
  body.find("[data-tag]").each((_, el) => {
    const labels = ($(el).attr("data-tag") ?? "")
      .split(",")
      .map((t) => t.trim())
      .filter(Boolean)
      .map((t) => NOTE_TAG_LABELS[t] ?? `[${t}]`);
    if (labels.length) $(el).prepend(`${labels.join(" ")} `);
  });

  const images: PageImage[] = [];
  body.find("img").each((i, el) => {
    const $el = $(el);
    const index = i + 1;
    $el.attr("data-mcp-index", String(index));
    images.push({
      index,
      src: $el.attr("src"),
      srcType: $el.attr("data-src-type"),
      fullresSrc: $el.attr("data-fullres-src"),
      fullresType: $el.attr("data-fullres-src-type"),
      alt: $el.attr("alt")?.trim() || undefined,
      width: numAttr($el.attr("width")),
      height: numAttr($el.attr("height")),
      dataId: $el.attr("data-id"),
      splitIndex: numAttr($el.attr("data-index")),
      options: $el.attr("data-options"),
    });
  });

  const attachments: PageAttachment[] = [];
  body.find("object").each((i, el) => {
    const $el = $(el);
    const att: PageAttachment = {
      index: i + 1,
      name: $el.attr("data-attachment") ?? `attachment-${i + 1}`,
      url: $el.attr("data"),
      type: $el.attr("type"),
    };
    attachments.push(att);
    // Turndown drops empty elements, so swap in a text placeholder.
    const label = `[Attachment ${att.index}: ${att.name}${att.type ? ` (${att.type})` : ""}]`;
    $el.replaceWith($("<p></p>").text(label));
  });

  body.find("table").each((_, t) => {
    $(t).find("tr").first().attr("data-mcp-first", "1");
  });

  const bodyHtml = body.html() ?? "";
  const markdown = turndown
    .turndown(bodyHtml)
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  return { title, created, markdown, images, attachments };
}

/** Returns page HTML with generated element IDs but without noisy styling, for targeted edits. */
export function compactHtmlWithIds(html: string): string {
  const $ = cheerio.load(html);
  $("[style]").removeAttr("style");
  $("img").each((_, el) => {
    const $el = $(el);
    const keep = ["id", "data-id", "alt", "width", "height"];
    for (const name of Object.keys(el.attribs)) if (!keep.includes(name)) $el.removeAttr(name);
  });
  $("object").removeAttr("data");
  const title = $("head > title").text();
  return `<title>${escapeHtml(title)}</title>\n${($("body").html() ?? "").trim()}`;
}

export function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/**
 * Loads an HTML fragment for rewriting outgoing content. htmlparser2 mode avoids the implied
 * <tbody> that spec parsers insert, and serialises void tags as <img ... /> like OneNote's examples.
 */
export function loadFragment(html: string): cheerio.CheerioAPI {
  return cheerio.load(html, { xml: { xmlMode: false, decodeEntities: true, selfClosingTags: true } }, false);
}

/** Converts markdown to HTML using only elements and attributes OneNote accepts. */
export function markdownToOneNoteHtml(md: string): string {
  const raw = marked.parse(md, { gfm: true, breaks: false, async: false }) as string;
  return adaptHtmlForOneNote(raw);
}

/** Rewrites generic HTML into the subset OneNote renders well. */
export function adaptHtmlForOneNote(html: string): string {
  const $ = loadFragment(html);

  // Task lists -> OneNote to-do tags.
  $("li").each((_, li) => {
    const box = $(li).find('input[type="checkbox"]').first();
    if (!box.length) return;
    $(li).attr("data-tag", box.is("[checked]") ? "to-do:completed" : "to-do");
    const next = box[0].next;
    if (next && next.type === "text") next.data = next.data.replace(/^\s+/, "");
    box.remove();
  });
  $("input").remove();

  // Tables: OneNote supports table/tr/td only, with a border attribute.
  $("table").each((_, t) => {
    $(t).attr("border", "1");
  });
  $("th, td").each((_, cell) => {
    const $cell = $(cell);
    const align = $cell.attr("align");
    if (align) {
      $cell.removeAttr("align");
      $cell.attr("style", `text-align:${align}`);
    }
    if (cell.tagName === "th") {
      const inner = $cell.html() ?? "";
      const style = $cell.attr("style");
      $cell.replaceWith(`<td${style ? ` style="${style}"` : ""}><b>${inner}</b></td>`);
    }
  });
  $("thead, tbody, tfoot").each((_, el) => {
    $(el).replaceWith($(el).contents());
  });

  // Code blocks -> monospace paragraphs that keep line breaks and indentation.
  $("pre").each((_, pre) => {
    const text = $(pre).text().replace(/\n$/, "");
    const lines = text
      .split("\n")
      .map((line) => escapeHtml(line).replace(/^ +/, (m) => "&nbsp;".repeat(m.length)) || "&nbsp;");
    $(pre).replaceWith(
      `<p style="font-family:Consolas;font-size:10pt;background-color:#f2f2f2">${lines.join("<br />")}</p>`,
    );
  });
  $("code").each((_, code) => {
    $(code).replaceWith(`<span style="font-family:Consolas">${escapeHtml($(code).text())}</span>`);
  });

  // Block quotes -> indented-looking italic grey paragraphs.
  $("blockquote").each((_, bq) => {
    $(bq)
      .find("p")
      .each((__, p) => {
        $(p).attr("style", "font-style:italic;color:#595959");
      });
    $(bq).replaceWith($(bq).contents());
  });

  $("hr").replaceWith('<p style="color:#a6a6a6">────────────────────────</p>');

  // Markdown wraps standalone images in <p>; OneNote places images best as direct block children.
  $("p").each((_, p) => {
    const $p = $(p);
    const kids = $p.contents().toArray();
    const meaningful = kids.filter((k) => !(k.type === "text" && !k.data.trim()));
    if (meaningful.length === 1 && meaningful[0].type === "tag" && meaningful[0].tagName === "img") {
      $p.replaceWith(meaningful[0]);
    }
  });
  $("s").each((_, el) => {
    $(el).replaceWith(`<del>${$(el).html() ?? ""}</del>`);
  });

  return $.html();
}

/** Wraps body HTML in the document OneNote expects when creating a page. */
export function buildPageDocument(title: string, bodyHtml: string, created = new Date()): string {
  // Numeric entities keep the whole document ASCII, so no charset negotiation is needed.
  const safeTitle = escapeHtml(title).replace(/[^\x00-\x7f]/gu, (c) => `&#x${c.codePointAt(0)!.toString(16)};`);
  return (
    `<!DOCTYPE html>\n<html>\n<head>\n<title>${safeTitle}</title>\n` +
    `<meta name="created" content="${created.toISOString()}" />\n</head>\n<body>\n${bodyHtml}\n</body>\n</html>`
  );
}
