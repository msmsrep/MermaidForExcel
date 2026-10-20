import {
  type DiagramItem,
  insertImageShape,
  insertSvgShape,
  isSvgInsertSupported,
  listDiagrams,
  onSheetActivated,
} from "./excelShapes";
import { type ImageFormat, serializeSvg, svgToDataUrl, svgToRaster } from "./image";
import { initMermaid, renderMermaid } from "./render";
import { initSplitter } from "./splitter";
import { byId, errorMessage } from "./util";

const RENDER_DEBOUNCE_MS = 250;
const PREVIEW_HINT_HTML = '<span class="hint">Type Mermaid code</span>';
const FILE_EXTENSIONS: Record<ImageFormat, string> = {
  png: "png",
  jpeg: "jpg",
  svg: "svg",
};

Office.onReady(() => {
  initMermaid();

  const input = byId<HTMLTextAreaElement>("mermaid-input");
  const splitter = byId<HTMLDivElement>("splitter");
  const preview = byId<HTMLDivElement>("preview");
  const errDiv = byId<HTMLDivElement>("error");
  const insertBtn = byId<HTMLButtonElement>("insert-btn");
  const downloadBtn = byId<HTMLButtonElement>("download-btn");
  const diagramSelect = byId<HTMLSelectElement>("sheet-diagram-select");
  const formatSelect = byId<HTMLSelectElement>("format-select");

  let renderDebounceTimer: number | undefined;
  let renderRequestId = 0;
  let diagramItems: DiagramItem[] = [];

  const selectedFormat = () => formatSelect.value as ImageFormat;

  function setActionsEnabled(enabled: boolean) {
    insertBtn.disabled = !enabled;
    downloadBtn.disabled = !enabled;
  }

  // ── Sheet diagram dropdown ─────────────────────────────────────────────────
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

  async function refreshDiagramOptions(selectedId = diagramSelect.value) {
    try {
      updateDiagramOptions(await listDiagrams(), selectedId);
    } catch {
      // Ignore load failures so the main render/insert flow remains available.
      updateDiagramOptions([], "");
    }
  }

  diagramSelect.addEventListener("focus", () => {
    void refreshDiagramOptions();
  });

  diagramSelect.addEventListener("change", () => {
    const selected = diagramItems.find((item) => item.id === diagramSelect.value);
    if (!selected) {
      return;
    }
    input.value = selected.altTextDescription;
    void renderPreview();
  });

  // Reset dropdown when the active sheet changes.
  onSheetActivated(() => updateDiagramOptions([], "")).catch(() => {/* ignore */});

  // ── Preview ────────────────────────────────────────────────────────────────
  async function renderPreview() {
    const requestId = ++renderRequestId;
    errDiv.textContent = "";
    setActionsEnabled(false);

    const code = input.value.trim();
    if (!code) {
      preview.innerHTML = PREVIEW_HINT_HTML;
      return;
    }

    try {
      const svg = await renderMermaid(code);
      if (requestId !== renderRequestId) return;
      preview.innerHTML = svg;
      setActionsEnabled(true);
    } catch (e) {
      if (requestId !== renderRequestId) return;
      preview.innerHTML = "";
      errDiv.textContent = `Render error: ${errorMessage(e)}`;
    }
  }

  input.addEventListener("input", () => {
    window.clearTimeout(renderDebounceTimer);
    renderDebounceTimer = window.setTimeout(() => {
      void renderPreview();
    }, RENDER_DEBOUNCE_MS);
  });

  initSplitter(splitter, input, preview);

  // ── Insert / Download ──────────────────────────────────────────────────────
  /** Runs a button's action on the rendered SVG and reports any failure. */
  function bindSvgAction(
    button: HTMLButtonElement,
    errorPrefix: string,
    action: (svgEl: SVGSVGElement) => Promise<void>,
  ) {
    button.addEventListener("click", async () => {
      errDiv.textContent = "";
      const svgEl = preview.querySelector<SVGSVGElement>("svg");
      if (!svgEl) {
        errDiv.textContent = "Please render first";
        return;
      }
      try {
        await action(svgEl);
      } catch (e) {
        errDiv.textContent = `${errorPrefix}: ${errorMessage(e)}`;
      }
    });
  }

  bindSvgAction(downloadBtn, "Download error", async (svgEl) => {
    const format = selectedFormat();
    const dataUrl =
      format === "svg"
        ? svgToDataUrl(serializeSvg(svgEl))
        : `data:image/${format};base64,` + (await svgToRaster(svgEl, format)).base64;

    const a = document.createElement("a");
    a.href = dataUrl;
    a.download = `diagram.${FILE_EXTENSIONS[format]}`;
    a.click();
  });

  bindSvgAction(insertBtn, "Insert error", async (svgEl) => {
    const format = selectedFormat();
    const mermaidCode = input.value.trim();
    const replaceId = diagramSelect.value || undefined;
    let insertedShapeId: string;

    if (format === "svg") {
      if (!isSvgInsertSupported()) {
        throw new Error(
          "SVG insertion is not supported in this version of Excel. Use PNG or JPEG.",
        );
      }
      insertedShapeId = await insertSvgShape(serializeSvg(svgEl), mermaidCode, replaceId);
    } else {
      const image = await svgToRaster(svgEl, format);
      insertedShapeId = await insertImageShape(image, mermaidCode, replaceId);
    }

    await refreshDiagramOptions(insertedShapeId);
  });

  // Render initial Mermaid code when the task pane opens.
  void renderPreview();
  void refreshDiagramOptions();
});
