import mermaid from "mermaid";

const ALT_TEXT_TITLE = "Mermaid Diagram";
// Raster output is rendered at this multiple of the SVG size for sharper images.
const RASTER_SCALE = 2;
// Upper bound on canvas area; larger diagrams are rendered at a lower scale.
const MAX_CANVAS_PIXELS = 4096 * 4096;
const JPEG_QUALITY = 0.92;
// Excel shape sizes are in points (1px = 0.75pt at 96 DPI).
const PX_TO_PT = 0.75;
const MIN_INPUT_HEIGHT = 40;
const MIN_PREVIEW_HEIGHT = 120;

type ImageFormat = "png" | "jpeg" | "svg";
type DiagramItem = { id: string; name: string; altTextDescription: string };
type Placement = { left: number; top: number; name?: string };
/** Base64 image (without data: prefix) and its display size in CSS pixels. */
type RasterImage = { base64: string; width: number; height: number };

// https://github.com/mermaid-js/mermaid-cli/issues/112
Office.onReady(() => {
    mermaid.initialize({
    startOnLoad: false,
    theme: "default",
    htmlLabels: false,
    flowchart: {
      useMaxWidth: false,
      htmlLabels: false,
    },
  });

  const input = document.getElementById("mermaid-input") as HTMLTextAreaElement;
  const splitter = document.getElementById("splitter") as HTMLDivElement;
  const preview = document.getElementById("preview") as HTMLDivElement;
  const errDiv = document.getElementById("error") as HTMLDivElement;
  const insertBtn = document.getElementById("insert-btn") as HTMLButtonElement;
  const downloadBtn = document.getElementById(
    "download-btn",
  ) as HTMLButtonElement;
  const diagramSelect = document.getElementById(
    "sheet-diagram-select",
  ) as HTMLSelectElement;
  const formatSelect = document.getElementById(
    "format-select",
  ) as HTMLSelectElement;
  const previewHintHtml = '<span class="hint">Type Mermaid code</span>';
  let renderDebounceTimer: number | undefined;
  let renderRequestId = 0;

  let selectedFormat: ImageFormat = "png";
  let diagramItems: DiagramItem[] = [];

  formatSelect.addEventListener("change", () => {
    selectedFormat = formatSelect.value as ImageFormat;
  });

  function updateDiagramOptions(items: DiagramItem[], selectedId = "") {
    diagramItems = items;
    diagramSelect.innerHTML = "";

    const defaultOption = document.createElement("option");
    defaultOption.value = "";
    defaultOption.textContent = "(No diagram selected)";
    diagramSelect.appendChild(defaultOption);

    for (const item of items) {
      const option = document.createElement("option");
      option.value = item.id;
      option.textContent = item.name;
      diagramSelect.appendChild(option);
    }

    diagramSelect.value =
      selectedId && items.some((item) => item.id === selectedId) ? selectedId : "";
  }

  /** Lists images on the active sheet that were inserted by this add-in. */
  async function loadDiagramItems(selectedId = diagramSelect.value) {
    try {
      const items = await Excel.run(async (ctx) => {
        const shapes = ctx.workbook.worksheets.getActiveWorksheet().shapes;
        shapes.load(
          "items/id,items/name,items/type,items/altTextTitle,items/altTextDescription",
        );
        await ctx.sync();

        return shapes.items
          .filter(
            (shape) =>
              shape.type === "Image" && shape.altTextTitle === ALT_TEXT_TITLE,
          )
          .map((shape) => ({
            id: shape.id,
            name: shape.name,
            altTextDescription: shape.altTextDescription ?? "",
          }));
      });

      updateDiagramOptions(items, selectedId);
    } catch {
      // Ignore load failures so the main render/insert flow remains available.
      updateDiagramOptions([], "");
    }
  }

  diagramSelect.addEventListener("focus", () => {
    void loadDiagramItems();
  });

  diagramSelect.addEventListener("change", () => {
    const selected = diagramItems.find((item) => item.id === diagramSelect.value);
    if (!selected) {
      return;
    }
    input.value = selected.altTextDescription;
    void renderPreview();
  });

  async function renderPreview() {
    const requestId = ++renderRequestId;
    errDiv.textContent = "";
    insertBtn.disabled = true;
    downloadBtn.disabled = true;

    const code = input.value.trim();
    if (!code) {
      preview.innerHTML = previewHintHtml;
      return;
    }

    try {
      // Remove element with the same ID to avoid errors
      document.getElementById("mermaid-graph")?.remove();

      const { svg } = await mermaid.render("mermaid-graph", code);
      if (requestId !== renderRequestId) return;
      preview.innerHTML = svg;
      insertBtn.disabled = false;
      downloadBtn.disabled = false;
    } catch (e) {
      if (requestId !== renderRequestId) return;
      preview.innerHTML = "";
      insertBtn.disabled = true;
      downloadBtn.disabled = true;
      errDiv.textContent = `Render error: ${e instanceof Error ? e.message : String(e)}`;
    }
  }

  input.addEventListener("input", () => {
    if (renderDebounceTimer !== undefined) {
      window.clearTimeout(renderDebounceTimer);
    }
    renderDebounceTimer = window.setTimeout(() => {
      void renderPreview();
    }, 250);
  });

  // ── Splitter between the input and the preview ─────────────────────────────
  function setInputHeight(height: number) {
    const max = Math.max(
      MIN_INPUT_HEIGHT,
      input.offsetHeight + preview.offsetHeight - MIN_PREVIEW_HEIGHT,
    );
    input.style.height = `${Math.min(max, Math.max(MIN_INPUT_HEIGHT, height))}px`;
  }

  splitter.addEventListener("pointerdown", (e) => {
    e.preventDefault();
    const startY = e.clientY;
    const startHeight = input.offsetHeight;
    splitter.setPointerCapture(e.pointerId);
    splitter.classList.add("dragging");

    const onMove = (ev: PointerEvent) => {
      setInputHeight(startHeight + ev.clientY - startY);
    };
    const onEnd = () => {
      splitter.classList.remove("dragging");
      splitter.removeEventListener("pointermove", onMove);
      splitter.removeEventListener("pointerup", onEnd);
      splitter.removeEventListener("pointercancel", onEnd);
    };
    splitter.addEventListener("pointermove", onMove);
    splitter.addEventListener("pointerup", onEnd);
    splitter.addEventListener("pointercancel", onEnd);
  });

  splitter.addEventListener("keydown", (e) => {
    if (e.key !== "ArrowUp" && e.key !== "ArrowDown") return;
    e.preventDefault();
    setInputHeight(input.offsetHeight + (e.key === "ArrowUp" ? -10 : 10));
  });

  // Render initial Mermaid code when the task pane opens.
  void renderPreview();
  void loadDiagramItems();

  // Reset dropdown when the active sheet changes.
  Excel.run(async (ctx) => {
    ctx.workbook.worksheets.onActivated.add(() => {
      updateDiagramOptions([], "");
      return Promise.resolve();
    });
    await ctx.sync();
  }).catch(() => {/* ignore */});

  downloadBtn.addEventListener("click", async () => {
    errDiv.textContent = "";
    const svgEl = preview.querySelector<SVGSVGElement>("svg");
    if (!svgEl) {
      errDiv.textContent = "Please render first";
      return;
    }

    try {
      let dataUrl: string;
      let filename: string;
      if (selectedFormat === "svg") {
        const svgStr = new XMLSerializer().serializeToString(svgEl);
        dataUrl = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svgStr);
        filename = "diagram.svg";
      } else {
        const { base64 } = await svgToRaster(svgEl, selectedFormat);
        dataUrl = `data:image/${selectedFormat};base64,` + base64;
        filename = selectedFormat === "jpeg" ? "diagram.jpg" : "diagram.png";
      }
      const a = document.createElement("a");
      a.href = dataUrl;
      a.download = filename;
      a.click();
    } catch (e) {
      errDiv.textContent = `Download error: ${e instanceof Error ? e.message : String(e)}`;
    }
  });

  insertBtn.addEventListener("click", async () => {
    errDiv.textContent = "";
    const svgEl = preview.querySelector<SVGSVGElement>("svg");
    if (!svgEl) {
      errDiv.textContent = "Please render first";
      return;
    }

    try {
      const mermaidCode = input.value.trim();
      const replaceId = diagramSelect.value || undefined;
      let insertedShapeId: string;

      if (selectedFormat === "svg") {
        if (
          !Office.context.requirements?.isSetSupported("ImageCoercion", "1.2")
        ) {
          throw new Error(
            "SVG insertion is not supported in this version of Excel. Use PNG or JPEG.",
          );
        }
        const svgStr = new XMLSerializer().serializeToString(svgEl);
        insertedShapeId = await insertSvgShape(svgStr, mermaidCode, replaceId);
      } else {
        const image = await svgToRaster(svgEl, selectedFormat);
        insertedShapeId = await insertImageShape(image, mermaidCode, replaceId);
      }

      await loadDiagramItems(insertedShapeId);
    } catch (e) {
      errDiv.textContent = `Insert error: ${e instanceof Error ? e.message : String(e)}`;
    }
  });
});

/**
 * Loads where a new shape should go: the replacement target's position and
 * name, or the active cell's top-left corner when there is no target.
 */
async function loadPlacement(
  ctx: Excel.RequestContext,
  replaceId?: string,
): Promise<Placement> {
  const cell = ctx.workbook.getActiveCell();
  cell.load(["left", "top"]);
  const target = replaceId
    ? ctx.workbook.worksheets.getActiveWorksheet().shapes.getItemOrNullObject(replaceId)
    : undefined;
  target?.load(["isNullObject", "left", "top", "name"]);
  await ctx.sync();

  if (!target) {
    return { left: cell.left, top: cell.top };
  }
  if (target.isNullObject) {
    throw new Error("Selected diagram no longer exists.");
  }
  return { left: target.left, top: target.top, name: target.name };
}

function applyShapeProps(
  shape: Excel.Shape,
  placement: Placement,
  mermaidCode: string,
) {
  shape.left = placement.left;
  shape.top = placement.top;
  if (placement.name) {
    shape.name = placement.name;
  }
  shape.altTextTitle = ALT_TEXT_TITLE;
  shape.altTextDescription = mermaidCode;
}

/** Adds a PNG/JPEG image and, when replacing, deletes the old shape in the same sync. */
function insertImageShape(
  image: RasterImage,
  mermaidCode: string,
  replaceId?: string,
): Promise<string> {
  return Excel.run(async (ctx) => {
    const placement = await loadPlacement(ctx, replaceId);
    const shapes = ctx.workbook.worksheets.getActiveWorksheet().shapes;

    const shape = shapes.addImage(image.base64);
    shape.width = image.width * PX_TO_PT;
    shape.height = image.height * PX_TO_PT;
    applyShapeProps(shape, placement, mermaidCode);
    shape.load("id");
    if (replaceId) {
      shapes.getItem(replaceId).delete();
    }
    await ctx.sync();
    return shape.id;
  });
}

/**
 * Inserts an SVG through setSelectedDataAsync. The new shape is found by
 * diffing shape IDs before and after insertion, and the old shape is deleted
 * only after the new one is in place.
 */
async function insertSvgShape(
  svgStr: string,
  mermaidCode: string,
  replaceId?: string,
): Promise<string> {
  const before = await Excel.run(async (ctx) => {
    const shapes = ctx.workbook.worksheets.getActiveWorksheet().shapes;
    shapes.load("items/id");
    const placement = await loadPlacement(ctx, replaceId);
    return { ids: new Set(shapes.items.map((s) => s.id)), placement };
  });

  await insertSvgToSelection(svgStr);

  return Excel.run(async (ctx) => {
    const shapes = ctx.workbook.worksheets.getActiveWorksheet().shapes;
    shapes.load("items/id");
    await ctx.sync();

    const added = shapes.items.filter((s) => !before.ids.has(s.id));
    if (added.length !== 1) {
      throw new Error("Could not identify the inserted shape.");
    }
    const shape = added[0];
    applyShapeProps(shape, before.placement, mermaidCode);
    await ctx.sync();

    if (replaceId) {
      const target = shapes.getItemOrNullObject(replaceId);
      target.load("isNullObject");
      await ctx.sync();
      if (!target.isNullObject) {
        target.delete();
        await ctx.sync();
      }
    }
    return shape.id;
  });
}

function insertSvgToSelection(svgStr: string): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(
      svgStr,
      { coercionType: Office.CoercionType.XmlSvg },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(asyncResult.error.message));
        } else {
          resolve();
        }
      },
    );
  });
}

/**
 * Converts an SVG element to a PNG or JPEG via Canvas, rendered at
 * RASTER_SCALE (reduced for very large diagrams). JPEG gets a white
 * background since it does not support transparency.
 */
async function svgToRaster(
  svgEl: SVGSVGElement,
  format: "png" | "jpeg",
): Promise<RasterImage> {
  const svgStr = new XMLSerializer().serializeToString(svgEl);
  const url = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svgStr);

  const img = await new Promise<HTMLImageElement>((resolve, reject) => {
    const i = new Image();
    i.onload = () => resolve(i);
    i.onerror = () => reject(new Error("Failed to load SVG image."));
    i.src = url;
  });

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
