import mermaid from "mermaid";

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
  const preview = document.getElementById("preview") as HTMLDivElement;
  const errDiv = document.getElementById("error") as HTMLDivElement;
  const insertBtn = document.getElementById("insert-btn") as HTMLButtonElement;
  const downloadBtn = document.getElementById(
    "download-btn",
  ) as HTMLButtonElement;
  const sheetSvgSelect = document.getElementById(
    "sheet-svg-select",
  ) as HTMLSelectElement;
  const formatSelect = document.getElementById(
    "format-select",
  ) as HTMLSelectElement;
  const previewHintHtml = '<span class="hint">Type Mermaid code</span>';
  let renderDebounceTimer: number | undefined;
  let renderRequestId = 0;

  let selectedFormat: "png" | "jpeg" | "svg" = "png";
  let sheetSvgItems: Array<{
    id: string;
    name: string;
    altTextDescription: string;
  }> = [];

  formatSelect.addEventListener("change", () => {
    selectedFormat = formatSelect.value as "png" | "jpeg" | "svg";
  });

  function updateSheetSvgOptions(
    items: Array<{ id: string; name: string; altTextDescription: string }>,
    selectedId = "",
  ) {
    sheetSvgItems = items;
    sheetSvgSelect.innerHTML = "";

    const defaultOption = document.createElement("option");
    defaultOption.value = "";
    defaultOption.textContent = "(No SVG selected)";
    sheetSvgSelect.appendChild(defaultOption);

    for (const item of items) {
      const option = document.createElement("option");
      option.value = item.id;
      option.textContent = item.name;
      sheetSvgSelect.appendChild(option);
    }

    sheetSvgSelect.value =
      selectedId && items.some((item) => item.id === selectedId) ? selectedId : "";
  }

  async function loadSheetSvgItems() {
    const selectedId = sheetSvgSelect.value;
    try {
      const items = await Excel.run(async (ctx) => {
        const shapes = ctx.workbook.worksheets.getActiveWorksheet().shapes;
        shapes.load("items/id,items/name,items/type,items/altTextDescription");
        await ctx.sync();

        const imageShapes = shapes.items.filter((shape) => shape.type === "Image");
        for (const shape of imageShapes) {
          shape.image.load("format");
        }
        await ctx.sync();

        return imageShapes
          .filter((shape) => shape.image.format === "SVG")
          .map((shape) => ({
            id: shape.id,
            name: shape.name,
            altTextDescription: shape.altTextDescription ?? "",
          }));
      });

      updateSheetSvgOptions(items, selectedId);
    } catch {
      // Ignore load failures so the main render/insert flow remains available.
      updateSheetSvgOptions([], "");
    }
  }

  sheetSvgSelect.addEventListener("focus", () => {
    void loadSheetSvgItems();
  });

  sheetSvgSelect.addEventListener("change", () => {
    const selected = sheetSvgItems.find((item) => item.id === sheetSvgSelect.value);
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

  // Render initial Mermaid code when the task pane opens.
  void renderPreview();
  void loadSheetSvgItems();

  async function applyInsertedShapeMetadata(
    mermaidCode: string,
    position?: { left: number; top: number },
    shapeName?: string,
  ): Promise<string> {
    return Excel.run(async (ctx) => {
      const shapes = ctx.workbook.worksheets.getActiveWorksheet().shapes;
      const countResult = shapes.getCount();
      await ctx.sync();
      if (countResult.value <= 0) {
        throw new Error("No inserted shape found.");
      }

      const lastShape = shapes.getItemAt(countResult.value - 1);
      if (position) {
        lastShape.left = position.left;
        lastShape.top = position.top;
      }
      if (shapeName) {
        lastShape.name = shapeName;
      }
      lastShape.altTextTitle = "Mermaid Diagram";
      lastShape.altTextDescription = mermaidCode;
      lastShape.load("id");
      await ctx.sync();
      return lastShape.id;
    });
  }

  async function replaceSvgShape(
    targetShapeId: string,
    svgStr: string,
    mermaidCode: string,
  ): Promise<string> {
    const { position, shapeName } = await Excel.run(async (ctx) => {
      const target = ctx.workbook.worksheets
        .getActiveWorksheet()
        .shapes.getItemOrNullObject(targetShapeId);
      target.load(["isNullObject", "left", "top", "name"]);
      await ctx.sync();
      if (target.isNullObject) {
        throw new Error("Selected SVG no longer exists.");
      }

      const result = {
        position: { left: target.left, top: target.top },
        shapeName: target.name,
      };
      target.delete();
      await ctx.sync();
      return result;
    });

    await insertSvgToSelection(svgStr);
    return applyInsertedShapeMetadata(mermaidCode, position, shapeName);
  }

  async function replaceImageShape(
    targetShapeId: string,
    base64: string,
    mermaidCode: string,
  ): Promise<string> {
    return Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const target = sheet.shapes.getItemOrNullObject(targetShapeId);
      target.load(["isNullObject", "left", "top", "name"]);
      await ctx.sync();
      if (target.isNullObject) {
        throw new Error("Selected SVG no longer exists.");
      }

      const shape = sheet.shapes.addImage(base64);
      shape.name = target.name;
      shape.left = target.left;
      shape.top = target.top;
      shape.altTextTitle = "Mermaid Diagram";
      shape.altTextDescription = mermaidCode;
      shape.load("id");

      target.delete();
      await ctx.sync();
      return shape.id;
    });
  }

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
      if (selectedFormat === "jpeg") {
        const base64 = await svgToBase64Jpeg(svgEl);
        dataUrl = "data:image/jpeg;base64," + base64;
        filename = "diagram.jpg";
      } else if (selectedFormat === "svg") {
        const svgStr = new XMLSerializer().serializeToString(svgEl);
        dataUrl = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svgStr);
        filename = "diagram.svg";
      } else {
        const base64 = await svgToBase64Png(svgEl);
        dataUrl = "data:image/png;base64," + base64;
        filename = "diagram.png";
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
      let insertedShapeId = "";
      const replacementTargetId = sheetSvgSelect.value;

      if (selectedFormat === "svg") {
        const svgStr = new XMLSerializer().serializeToString(svgEl);
        if (replacementTargetId) {
          insertedShapeId = await replaceSvgShape(
            replacementTargetId,
            svgStr,
            mermaidCode,
          );
        } else {
          await insertSvgToSelection(svgStr);
          insertedShapeId = await applyInsertedShapeMetadata(mermaidCode);
        }
      } else {
        const base64 =
          selectedFormat === "jpeg"
            ? await svgToBase64Jpeg(svgEl)
            : await svgToBase64Png(svgEl);
        if (replacementTargetId) {
          insertedShapeId = await replaceImageShape(
            replacementTargetId,
            base64,
            mermaidCode,
          );
        } else {
          insertedShapeId = await Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();
            const activeCell = ctx.workbook.getActiveCell();
            activeCell.load(["left", "top"]);
            await ctx.sync();

            const shape = sheet.shapes.addImage(base64);
            shape.left = activeCell.left;
            shape.top = activeCell.top;
            shape.altTextTitle = "Mermaid Diagram";
            shape.altTextDescription = mermaidCode;
            shape.load("id");
            await ctx.sync();
            return shape.id;
          });
        }
      }

      await loadSheetSvgItems();
      if (insertedShapeId) {
        sheetSvgSelect.value = insertedShapeId;
      }
    } catch (e) {
      errDiv.textContent = `Insert error: ${e instanceof Error ? e.message : String(e)}`;
    }
  });
});

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
 * Converts an SVG element to a base64 PNG string (without data: prefix).
 * Uses Canvas to produce the PNG format required by Excel's addImage API.
 */
async function svgToBase64Png(svgEl: SVGSVGElement): Promise<string> {
  const svgStr = new XMLSerializer().serializeToString(svgEl);
  const url = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svgStr);

  const img = await new Promise<HTMLImageElement>((resolve, reject) => {
    const i = new Image();
    i.onload = () => resolve(i);
    i.onerror = reject;
    i.src = url;
  });

  const vb = svgEl.viewBox?.baseVal;
  const w = (vb && vb.width > 0 ? vb.width : img.naturalWidth) || 800;
  const h = (vb && vb.height > 0 ? vb.height : img.naturalHeight) || 600;

  const canvas = document.createElement("canvas");
  canvas.width = w;
  canvas.height = h;
  canvas.getContext("2d")!.drawImage(img, 0, 0, w, h);

  return canvas.toDataURL("image/png").replace("data:image/png;base64,", "");
}

/**
 * Converts an SVG element to a base64 JPEG string (without data: prefix).
 */
async function svgToBase64Jpeg(
  svgEl: SVGSVGElement,
  quality = 0.92,
): Promise<string> {
  const svgStr = new XMLSerializer().serializeToString(svgEl);
  const url = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svgStr);

  const img = await new Promise<HTMLImageElement>((resolve, reject) => {
    const i = new Image();
    i.onload = () => resolve(i);
    i.onerror = reject;
    i.src = url;
  });

  const vb = svgEl.viewBox?.baseVal;
  const w = (vb && vb.width > 0 ? vb.width : img.naturalWidth) || 800;
  const h = (vb && vb.height > 0 ? vb.height : img.naturalHeight) || 600;

  const canvas = document.createElement("canvas");
  canvas.width = w;
  canvas.height = h;
  const ctx2d = canvas.getContext("2d")!;
  // Fill with white background since JPEG does not support transparency
  ctx2d.fillStyle = "#ffffff";
  ctx2d.fillRect(0, 0, w, h);
  ctx2d.drawImage(img, 0, 0, w, h);

  return canvas
    .toDataURL("image/jpeg", quality)
    .replace("data:image/jpeg;base64,", "");
}
