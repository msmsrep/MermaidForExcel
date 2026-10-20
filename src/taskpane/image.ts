// Raster output is rendered at this multiple of the SVG size for sharper images.
const RASTER_SCALE = 2;
// Upper bound on canvas area; larger diagrams are rendered at a lower scale.
const MAX_CANVAS_PIXELS = 4096 * 4096;
const JPEG_QUALITY = 0.92;

export type RasterFormat = "png" | "jpeg";
export type ImageFormat = RasterFormat | "svg";
/** Base64 image (without data: prefix) and its display size in CSS pixels. */
export type RasterImage = { base64: string; width: number; height: number };

export function serializeSvg(svgEl: SVGSVGElement): string {
  return new XMLSerializer().serializeToString(svgEl);
}

export function svgToDataUrl(svgStr: string): string {
  return "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svgStr);
}

function loadImage(src: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error("Failed to load SVG image."));
    img.src = src;
  });
}

/**
 * Converts an SVG element to a PNG or JPEG via Canvas, rendered at
 * RASTER_SCALE (reduced for very large diagrams). JPEG gets a white
 * background since it does not support transparency.
 */
export async function svgToRaster(
  svgEl: SVGSVGElement,
  format: RasterFormat,
): Promise<RasterImage> {
  const img = await loadImage(svgToDataUrl(serializeSvg(svgEl)));

  const vb = svgEl.viewBox?.baseVal;
  const w = (vb && vb.width > 0 ? vb.width : img.naturalWidth) || 800;
  const h = (vb && vb.height > 0 ? vb.height : img.naturalHeight) || 600;
  const scale = Math.min(RASTER_SCALE, Math.sqrt(MAX_CANVAS_PIXELS / (w * h)));

  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(w * scale));
  canvas.height = Math.max(1, Math.round(h * scale));
  const ctx2d = canvas.getContext("2d")!;
  if (format === "jpeg") {
    ctx2d.fillStyle = "#ffffff";
    ctx2d.fillRect(0, 0, canvas.width, canvas.height);
  }
  ctx2d.drawImage(img, 0, 0, canvas.width, canvas.height);

  const dataUrl =
    format === "jpeg"
      ? canvas.toDataURL("image/jpeg", JPEG_QUALITY)
      : canvas.toDataURL("image/png");
  return { base64: dataUrl.split(",")[1], width: w, height: h };
}
