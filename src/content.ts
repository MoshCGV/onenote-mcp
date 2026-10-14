import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { normalizeForUpload } from "./images.js";
import { loadFragment, type PageImage } from "./html.js";
import type { BinaryPart, OneNote } from "./onenote.js";

/** Display width cap for large screenshots inserted into a page (the full-res file is still stored). */
const MAX_DISPLAY_WIDTH = 960;

export type ImageQuality = "fullres" | "web";

/** Downloads one image from a page, preferring full resolution and falling back to the web copy. */
export async function loadPageImage(
  onenote: OneNote,
  pageId: string,
  index: number,
  quality: ImageQuality = "fullres",
): Promise<{ data: Buffer; contentType: string; image: PageImage; total: number }> {
  const page = await onenote.getParsedPage(pageId);
  const image = page.images.find((i) => i.index === index);
  if (!image) {
    throw new Error(
      page.images.length
        ? `Page has ${page.images.length} image(s); image_index must be between 1 and ${page.images.length}.`
        : "This page has no images.",
    );
  }
  const candidates = quality === "fullres" ? [image.fullresSrc, image.src] : [image.src, image.fullresSrc];
  let lastErr: unknown;
  for (const url of candidates) {
    if (!url) continue;
    try {
      const res = await onenote.getResource(url);
      return { ...res, image, total: page.images.length };
    } catch (err) {
      lastErr = err;
    }
  }
  throw lastErr instanceof Error ? lastErr : new Error("Image has no downloadable source.");
}

/** Parses onenote-image://{pageId}/{n} (also onenote-image:{pageId}/{n}). */
function parseOneNoteImageRef(src: string): { pageId: string; index: number } | undefined {
  const m = /^onenote-image:(?:\/\/)?(.+)\/(\d+)$/i.exec(src.trim());
  if (!m) return undefined;
  return { pageId: decodeURIComponent(m[1]), index: Number.parseInt(m[2], 10) };
}

function localPath(src: string): string | undefined {
  if (/^file:\/\//i.test(src)) return fileURLToPath(src);
  const decoded = decodeURI(src);
  if (decoded.startsWith("~/")) return path.join(os.homedir(), decoded.slice(2));
  if (path.isAbsolute(decoded)) return decoded;
  return undefined;
}

export interface PreparedContent {
  html: string;
  parts: BinaryPart[];
  notes: string[];
}

/**
 * Rewrites <img> sources so OneNote can ingest them:
 * - local paths, file:// URLs and data: URLs become multipart parts ("name:imgN")
 * - onenote-image://{pageId}/{n} copies an image from another OneNote page
 * - https:// URLs are left for OneNote to download (page creation only)
 */
export async function prepareOutgoingHtml(
  html: string,
  opts: { onenote: OneNote; allowRemoteImages: boolean },
): Promise<PreparedContent> {
  const $ = loadFragment(html);
  const parts: BinaryPart[] = [];
  const notes: string[] = [];
  const imgs = $("img").toArray();

  for (const el of imgs) {
    const $el = $(el);
    const src = ($el.attr("src") ?? "").trim();
    if (!src) {
      $el.remove();
      continue;
    }
    if (/^name:/i.test(src)) continue;

    if (/^https?:\/\//i.test(src)) {
      if (!opts.allowRemoteImages) {
        throw new Error(
          `OneNote does not download web images when updating a page (${src}). ` +
            "Save the image locally and reference its absolute path, or create a new page instead.",
        );
      }
      continue;
    }

    let raw: Buffer;
    let label: string;
    const ref = parseOneNoteImageRef(src);
    if (ref) {
      raw = (await loadPageImage(opts.onenote, ref.pageId, ref.index, "fullres")).data;
      label = `image ${ref.index} from page ${ref.pageId}`;
    } else if (/^data:/i.test(src)) {
      const m = /^data:([^;,]+)?(;base64)?,(.*)$/is.exec(src);
      if (!m || !m[2]) throw new Error("Only base64 data: URLs are supported for images.");
      raw = Buffer.from(m[3], "base64");
      label = "inline data URL";
    } else {
      const file = localPath(src);
      if (!file) {
        throw new Error(
          `Image source "${src}" is a relative path. Use an absolute path (e.g. /Users/me/shot.png), ` +
            "a file:// URL, a data: URL, an https:// URL, or onenote-image://{pageId}/{n}.",
        );
      }
      const stat = await fs.stat(file).catch(() => undefined);
      if (!stat?.isFile()) throw new Error(`Image file not found: ${file}`);
      if (stat.size > 25 * 1024 * 1024) throw new Error(`Image file is larger than 25 MB: ${file}`);
      raw = await fs.readFile(file);
      label = file;
    }

    let normalized;
    try {
      normalized = await normalizeForUpload(raw);
    } catch (err) {
      throw new Error(`${label}: ${(err as Error).message}`);
    }

    const name = `img${parts.length + 1}`;
    parts.push({ name, contentType: normalized.contentType, data: normalized.data });
    $el.attr("src", `name:${name}`);

    if (!$el.attr("width") && !$el.attr("height") && normalized.width > MAX_DISPLAY_WIDTH) {
      $el.attr("width", String(MAX_DISPLAY_WIDTH));
      $el.attr("height", String(Math.round((normalized.height * MAX_DISPLAY_WIDTH) / normalized.width)));
    }
    notes.push(`Embedded ${label} (${normalized.width}x${normalized.height}).`);
  }

  return { html: $.html(), parts, notes };
}
