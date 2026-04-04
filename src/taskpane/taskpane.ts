import mermaid from "mermaid";

Office.onReady(() => {
  mermaid.initialize({ startOnLoad: false, theme: "default" });

  const input = document.getElementById("mermaid-input") as HTMLTextAreaElement;
  const preview = document.getElementById("preview") as HTMLDivElement;
  const errDiv = document.getElementById("error") as HTMLDivElement;
  const renderBtn = document.getElementById("render-btn") as HTMLButtonElement;
  const insertBtn = document.getElementById("insert-btn") as HTMLButtonElement;

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
    } catch (e) {
      preview.innerHTML = "";
      errDiv.textContent = `レンダリングエラー: ${e instanceof Error ? e.message : String(e)}`;
    }
  });

  insertBtn.addEventListener("click", async () => {
    errDiv.textContent = "";
    const svgEl = preview.querySelector("svg");
    if (!svgEl) {
      errDiv.textContent = "先にレンダリングしてください";
      return;
    }

    try {
      const base64 = await svgToBase64Png(svgEl);
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
async function svgToBase64Png(svgEl: SVGElement): Promise<string> {
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
