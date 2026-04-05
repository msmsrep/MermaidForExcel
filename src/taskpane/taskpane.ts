import mermaid from "mermaid";

Office.onReady(() => {
  mermaid.initialize({ startOnLoad: false, theme: "default" });

  const input = document.getElementById("mermaid-input") as HTMLTextAreaElement;
  const preview = document.getElementById("preview") as HTMLDivElement;
  const errDiv = document.getElementById("error") as HTMLDivElement;
  const renderBtn = document.getElementById("render-btn") as HTMLButtonElement;
  const insertBtn = document.getElementById("insert-btn") as HTMLButtonElement;
  const downloadBtn = document.getElementById(
    "download-btn",
  ) as HTMLButtonElement;
  const formatSelect = document.getElementById(
    "format-select",
  ) as HTMLSelectElement;

  let selectedFormat: "png" | "jpeg" = "png";

  formatSelect.addEventListener("change", () => {
    selectedFormat = formatSelect.value as "png" | "jpeg";
  });

  renderBtn.addEventListener("click", async () => {
    errDiv.textContent = "";
    insertBtn.disabled = true;

    const code = input.value.trim();
    if (!code) return;

    try {
      // 同一 ID が残っているとエラーになるため削除
      document.getElementById("mermaid-graph")?.remove();

      const { svg } = await mermaid.render("mermaid-graph", code);
      preview.innerHTML = svg;
      insertBtn.disabled = false;
      downloadBtn.disabled = false;
    } catch (e) {
      preview.innerHTML = "";
      insertBtn.disabled = true;
      downloadBtn.disabled = true;
      errDiv.textContent = `レンダリングエラー: ${e instanceof Error ? e.message : String(e)}`;
    }
  });

  downloadBtn.addEventListener("click", async () => {
    errDiv.textContent = "";
    const svgEl = preview.querySelector<SVGSVGElement>("svg");
    if (!svgEl) {
      errDiv.textContent = "先にレンダリングしてください";
      return;
    }

    try {
      let dataUrl: string;
      let filename: string;
      if (selectedFormat === "jpeg") {
        const base64 = await svgToBase64Jpeg(svgEl);
        dataUrl = "data:image/jpeg;base64," + base64;
        filename = "diagram.jpg";
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
      errDiv.textContent = `ダウンロードエラー: ${e instanceof Error ? e.message : String(e)}`;
    }
  });

  insertBtn.addEventListener("click", async () => {
    errDiv.textContent = "";
    const svgEl = preview.querySelector<SVGSVGElement>("svg");
    if (!svgEl) {
      errDiv.textContent = "先にレンダリングしてください";
      return;
    }

    try {
      const base64 =
        selectedFormat === "jpeg"
          ? await svgToBase64Jpeg(svgEl)
          : await svgToBase64Png(svgEl);
      await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        sheet.shapes.addImage(base64);
        await ctx.sync();
      });
    } catch (e) {
      errDiv.textContent = `挿入エラー: ${e instanceof Error ? e.message : String(e)}`;
    }
  });
});

/**
 * SVG 要素を PNG の base64 文字列（data: プレフィックスなし）に変換する。
 * Canvas を経由することで Excel の addImage API が要求する PNG フォーマットに対応する。
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
 * SVG 要素を JPEG の base64 文字列（data: プレフィックスなし）に変換する。
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
  // JPEG は透過をサポートしないため白背景で塗りつぶす
  ctx2d.fillStyle = "#ffffff";
  ctx2d.fillRect(0, 0, w, h);
  ctx2d.drawImage(img, 0, 0, w, h);

  return canvas
    .toDataURL("image/jpeg", quality)
    .replace("data:image/jpeg;base64,", "");
}
