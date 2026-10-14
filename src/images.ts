import sharp from "sharp";

/** A crop region expressed as fractions (0-1) of the image width/height. */
export interface Region {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface ModelImage {
  /** Base64-encoded image data. */
  data: string;
  mimeType: "image/png" | "image/jpeg";
  width: number;
  height: number;
  sourceWidth: number;
  sourceHeight: number;
  bytes: number;
}

const LOSSLESS_FORMATS = new Set(["png", "gif", "tiff", "bmp", "svg", "webp"]);
const MAX_INPUT_PIXELS = 268_402_689;

/** Orientation-corrected dimensions of an image. */
async function orientedSize(buf: Buffer): Promise<{ width: number; height: number; format?: string }> {
  const meta = await sharp(buf, { limitInputPixels: MAX_INPUT_PIXELS }).metadata();
  if (!meta.width || !meta.height) throw new Error("Could not read image dimensions.");
  const swap = (meta.orientation ?? 1) >= 5;
  return {
    width: swap ? meta.height : meta.width,
    height: swap ? meta.width : meta.height,
    format: meta.format,
  };
}

function regionToPixels(region: Region, width: number, height: number) {
  const clamp = (v: number) => Math.min(1, Math.max(0, v));
  const x = clamp(region.x);
  const y = clamp(region.y);
  const w = clamp(Math.min(region.width, 1 - x));
  const h = clamp(Math.min(region.height, 1 - y));
  const left = Math.round(x * width);
  const top = Math.round(y * height);
  return {
    left,
    top,
    width: Math.max(1, Math.min(width - left, Math.round(w * width))),
    height: Math.max(1, Math.min(height - top, Math.round(h * height))),
  };
}

/**
 * Prepares an image for a vision model: honours EXIF orientation, optionally crops,
 * downsizes so the longest edge is at most maxPx, and keeps the encoded size under maxBytes.
 * Screenshots (lossless sources) stay PNG when they fit, so small text stays crisp.
 */
export async function toModelImage(
  buf: Buffer,
  opts: { maxPx: number; maxBytes: number; region?: Region },
): Promise<ModelImage> {
  const { width: srcW, height: srcH, format } = await orientedSize(buf);
  const lossless = format ? LOSSLESS_FORMATS.has(format) : true;

  const crop = opts.region ? regionToPixels(opts.region, srcW, srcH) : undefined;
  const baseW = crop?.width ?? srcW;
  const baseH = crop?.height ?? srcH;

  let maxPx = opts.maxPx;
  for (let attempt = 0; attempt < 6; attempt++) {
    const scale = Math.min(1, maxPx / Math.max(baseW, baseH));
    const targetW = Math.max(1, Math.round(baseW * scale));
    const targetH = Math.max(1, Math.round(baseH * scale));

    const pipeline = () => {
      let p = sharp(buf, { limitInputPixels: MAX_INPUT_PIXELS, animated: false }).rotate();
      if (crop) p = p.extract(crop);
      return p.resize({ width: targetW, height: targetH, fit: "fill" });
    };

    if (lossless && attempt === 0) {
      const png = await pipeline().png({ compressionLevel: 9, adaptiveFiltering: true }).toBuffer();
      if (png.length <= opts.maxBytes) {
        return finish(png, "image/png", targetW, targetH, srcW, srcH);
      }
    }

    const quality = attempt === 0 ? 85 : attempt === 1 ? 80 : 72;
    const jpg = await pipeline().flatten({ background: "#ffffff" }).jpeg({ quality, mozjpeg: true }).toBuffer();
    if (jpg.length <= opts.maxBytes || maxPx <= 320) {
      return finish(jpg, "image/jpeg", targetW, targetH, srcW, srcH);
    }
    maxPx = Math.round(Math.max(baseW, baseH) * scale * 0.8);
  }
  throw new Error("Could not compress the image under the size budget.");
}

function finish(
  out: Buffer,
  mimeType: ModelImage["mimeType"],
  width: number,
  height: number,
  sourceWidth: number,
  sourceHeight: number,
): ModelImage {
  return { data: out.toString("base64"), mimeType, width, height, sourceWidth, sourceHeight, bytes: out.length };
}

const UPLOAD_KEEP = new Map([
  ["png", "image/png"],
  ["jpeg", "image/jpeg"],
  ["gif", "image/gif"],
]);

/**
 * Validates that a buffer is a real image and returns it in a format OneNote accepts.
 * PNG, JPEG and GIF pass through; other decodable formats (WebP, HEIC, TIFF, SVG...) become PNG.
 */
export async function normalizeForUpload(
  buf: Buffer,
): Promise<{ data: Buffer; contentType: string; width: number; height: number }> {
  if (buf.length > 25 * 1024 * 1024) throw new Error("Image is larger than 25 MB.");
  let size: { width: number; height: number; format?: string };
  try {
    size = await orientedSize(buf);
  } catch {
    throw new Error("File is not a readable image (PNG, JPEG, GIF, WebP, TIFF, HEIC or SVG).");
  }
  const keep = size.format ? UPLOAD_KEEP.get(size.format) : undefined;
  if (keep) return { data: buf, contentType: keep, width: size.width, height: size.height };
  const png = await sharp(buf, { limitInputPixels: MAX_INPUT_PIXELS }).rotate().png().toBuffer();
  return { data: png, contentType: "image/png", width: size.width, height: size.height };
}

const EXT_BY_MIME: Record<string, string> = {
  "image/png": "png",
  "image/jpeg": "jpg",
  "image/jpg": "jpg",
  "image/gif": "gif",
  "image/bmp": "bmp",
  "image/tiff": "tiff",
  "image/webp": "webp",
  "image/svg+xml": "svg",
  "image/x-emf": "emf",
  "image/x-wmf": "wmf",
  "application/pdf": "pdf",
  "text/plain": "txt",
};

export function extensionFor(mime: string | undefined): string {
  if (!mime) return "bin";
  const clean = mime.split(";")[0].trim().toLowerCase();
  return EXT_BY_MIME[clean] ?? clean.split("/")[1]?.replace(/[^a-z0-9]/g, "") ?? "bin";
}
