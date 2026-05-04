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
  const formatSelect = document.getElementById(
    "format-select",
  ) as HTMLSelectElement;
  const previewHintHtml = '<span class="hint">Type Mermaid code</span>';
  let renderDebounceTimer: number | undefined;
  let renderRequestId = 0;

  let selectedFormat: "png" | "jpeg" | "svg" = "png";

  formatSelect.addEventListener("change", () => {
    selectedFormat = formatSelect.value as "png" | "jpeg" | "svg";
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
      if (selectedFormat === "svg") {
        const svgStr = new XMLSerializer().serializeToString(svgEl);
        await insertSvgToSelection(svgStr);
        // Set alt text on the newly inserted shape (last item in collection)
        await Excel.run(async (ctx) => {
          const shapes = ctx.workbook.worksheets.getActiveWorksheet().shapes;
          const countResult = shapes.getCount();
          await ctx.sync();
          if (countResult.value > 0) {
            const lastShape = shapes.getItemAt(countResult.value - 1);
            lastShape.altTextDescription = mermaidCode;
            await ctx.sync();
          }
        });
      } else {
        const base64 =
          selectedFormat === "jpeg"
            ? await svgToBase64Jpeg(svgEl)
            : await svgToBase64Png(svgEl);
        await Excel.run(async (ctx) => {
          const sheet = ctx.workbook.worksheets.getActiveWorksheet();
          const activeCell = ctx.workbook.getActiveCell();
          activeCell.load(["left", "top"]);
          await ctx.sync();

          const shape = sheet.shapes.addImage(base64);
          shape.left = activeCell.left;
          shape.top = activeCell.top;
          shape.altTextTitle = "Mermaid Diagram";
          shape.altTextDescription = mermaidCode;
          await ctx.sync();
        });
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
