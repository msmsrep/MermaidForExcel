import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { DiagramItem } from "../src/taskpane/excelShapes";

const m = vi.hoisted(() => ({
  listDiagrams: vi.fn(),
  onSheetActivated: vi.fn(),
  isSvgInsertSupported: vi.fn(),
  insertImageShape: vi.fn(),
  insertSvgShape: vi.fn(),
  initMermaid: vi.fn(),
  renderMermaid: vi.fn(),
  svgToRaster: vi.fn(),
}));

vi.mock("../src/taskpane/excelShapes", () => ({
  listDiagrams: m.listDiagrams,
  onSheetActivated: m.onSheetActivated,
  isSvgInsertSupported: m.isSvgInsertSupported,
  insertImageShape: m.insertImageShape,
  insertSvgShape: m.insertSvgShape,
}));
vi.mock("../src/taskpane/render", () => ({
  initMermaid: m.initMermaid,
  renderMermaid: m.renderMermaid,
}));
vi.mock("../src/taskpane/image", async (importOriginal) => ({
  ...(await importOriginal<typeof import("../src/taskpane/image")>()),
  svgToRaster: m.svgToRaster,
}));

const HTML = readFileSync(join(__dirname, "../src/taskpane/taskpane.html"), "utf8");
const INITIAL_CODE = `graph LR
    A[Start] --> B{Condition}
    B -->|Yes| C[Process A]
    B -->|No| D[Process B]
    C --> E[End]
    D --> E`;
const SVG = '<svg id="mermaid-graph" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"><text>A</text></svg>';
const RASTER = { base64: "B64", width: 10, height: 10 };
const DIAGRAMS: DiagramItem[] = [
  { id: "id-1", name: "Diagram 1", altTextDescription: "graph TD\nX-->Y" },
  { id: "id-2", name: "Diagram 2", altTextDescription: "pie\n\"a\": 1" },
];

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const input = () => $<HTMLTextAreaElement>("mermaid-input");
const preview = () => $<HTMLDivElement>("preview");
const errorText = () => $<HTMLDivElement>("error").textContent;
const insertBtn = () => $<HTMLButtonElement>("insert-btn");
const downloadBtn = () => $<HTMLButtonElement>("download-btn");
const diagramSelect = () => $<HTMLSelectElement>("sheet-diagram-select");
const formatSelect = () => $<HTMLSelectElement>("format-select");
const optionLabels = () => Array.from(diagramSelect().options, (o) => [o.value, o.textContent]);

/** Lets pending promise chains (mocked async calls) settle. */
async function flush() {
  for (let i = 0; i < 20; i++) await Promise.resolve();
}

function deferred<T>() {
  let resolve!: (v: T) => void;
  let reject!: (e: unknown) => void;
  const promise = new Promise<T>((res, rej) => ((resolve = res), (reject = rej)));
  return { promise, resolve, reject };
}

function type(code: string) {
  input().value = code;
  input().dispatchEvent(new Event("input"));
}

function selectDiagram(id: string) {
  diagramSelect().value = id;
  diagramSelect().dispatchEvent(new Event("change"));
}

function setFormat(format: "png" | "jpeg" | "svg") {
  formatSelect().value = format;
}

async function click(button: HTMLButtonElement) {
  button.click();
  await flush();
}

/** Loads taskpane.html and runs taskpane.ts as if Office had just become ready. */
async function start() {
  document.documentElement.innerHTML = HTML.replace(/^[\s\S]*?<html[^>]*>|<\/html>[\s\S]*$/g, "");
  vi.stubGlobal("Office", { onReady: (cb: () => void) => cb() });
  vi.resetModules();
  await import("../src/taskpane/taskpane");
  await flush();
}

let downloads: Array<{ href: string; download: string }>;

beforeEach(() => {
  vi.useFakeTimers();
  for (const fn of Object.values(m)) fn.mockReset();
  m.listDiagrams.mockResolvedValue([]);
  m.onSheetActivated.mockResolvedValue(undefined);
  m.isSvgInsertSupported.mockReturnValue(true);
  m.renderMermaid.mockResolvedValue(SVG);
  m.svgToRaster.mockResolvedValue(RASTER);
  m.insertImageShape.mockResolvedValue("new-id");
  m.insertSvgShape.mockResolvedValue("new-svg-id");

  downloads = [];
  vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(function (this: HTMLAnchorElement) {
    downloads.push({ href: this.href, download: this.download });
  });
});

afterEach(() => {
  vi.useRealTimers();
});

describe("startup", () => {
  it("initializes mermaid and renders the sample code", async () => {
    await start();

    expect(m.initMermaid).toHaveBeenCalledOnce();
    expect(m.initMermaid.mock.invocationCallOrder[0]).toBeLessThan(
      m.renderMermaid.mock.invocationCallOrder[0],
    );
    expect(m.renderMermaid).toHaveBeenCalledWith(INITIAL_CODE);
    expect(preview().querySelector("svg")).not.toBeNull();
    expect(insertBtn().disabled).toBe(false);
    expect(downloadBtn().disabled).toBe(false);
    expect(errorText()).toBe("");
  });

  it("loads the sheet diagram list and listens for sheet activation", async () => {
    m.listDiagrams.mockResolvedValue(DIAGRAMS);
    await start();

    expect(optionLabels()).toEqual([
      ["", "(No diagram selected)"],
      ["id-1", "Diagram 1"],
      ["id-2", "Diagram 2"],
    ]);
    expect(diagramSelect().value).toBe("");
    expect(m.onSheetActivated).toHaveBeenCalledOnce();
  });

  it("keeps working when registering the sheet activation handler fails", async () => {
    m.onSheetActivated.mockRejectedValue(new Error("no events"));
    await start();
    expect(errorText()).toBe("");
    expect(insertBtn().disabled).toBe(false);
  });
});

describe("preview", () => {
  beforeEach(start);

  it("renders 250ms after the last input", async () => {
    m.renderMermaid.mockClear();
    type("graph TD\nA");
    await vi.advanceTimersByTimeAsync(200);
    type("graph TD\nA-->B");
    await vi.advanceTimersByTimeAsync(249);
    expect(m.renderMermaid).not.toHaveBeenCalled();

    await vi.advanceTimersByTimeAsync(1);
    expect(m.renderMermaid).toHaveBeenCalledOnce();
    expect(m.renderMermaid).toHaveBeenCalledWith("graph TD\nA-->B");
  });

  it("trims the code before rendering", async () => {
    type("\n  pie  \n");
    await vi.advanceTimersByTimeAsync(250);
    expect(m.renderMermaid).toHaveBeenLastCalledWith("pie");
  });

  it("shows the hint and disables the actions when the input is empty", async () => {
    m.renderMermaid.mockClear();
    type("   \n ");
    await vi.advanceTimersByTimeAsync(250);

    expect(m.renderMermaid).not.toHaveBeenCalled();
    expect(preview().textContent).toBe("Type Mermaid code");
    expect(insertBtn().disabled).toBe(true);
    expect(downloadBtn().disabled).toBe(true);
  });

  it("shows render errors, clears the preview and keeps the actions disabled", async () => {
    m.renderMermaid.mockRejectedValue(new Error("Parse error"));
    type("bad");
    await vi.advanceTimersByTimeAsync(250);

    expect(errorText()).toBe("Render error: Parse error");
    expect(preview().innerHTML).toBe("");
    expect(insertBtn().disabled).toBe(true);
    expect(downloadBtn().disabled).toBe(true);
  });

  it("clears the previous error on the next render", async () => {
    m.renderMermaid.mockRejectedValueOnce("oops");
    type("bad");
    await vi.advanceTimersByTimeAsync(250);
    expect(errorText()).toBe("Render error: oops");

    type("good");
    await vi.advanceTimersByTimeAsync(250);
    expect(errorText()).toBe("");
    expect(insertBtn().disabled).toBe(false);
  });

  it("disables the actions while a render is in progress", async () => {
    const pending = deferred<string>();
    m.renderMermaid.mockReturnValue(pending.promise);
    type("graph TD\nA-->B");
    await vi.advanceTimersByTimeAsync(250);
    expect(insertBtn().disabled).toBe(true);

    pending.resolve(SVG);
    await flush();
    expect(insertBtn().disabled).toBe(false);
  });

  it("discards a stale render result that finishes after a newer one", async () => {
    const first = deferred<string>();
    m.renderMermaid.mockReturnValueOnce(first.promise).mockResolvedValueOnce("<svg><text>new</text></svg>");

    type("old");
    await vi.advanceTimersByTimeAsync(250);
    type("new");
    await vi.advanceTimersByTimeAsync(250);
    first.resolve("<svg><text>old</text></svg>");
    await flush();

    expect(preview().textContent).toBe("new");
  });

  it("discards a stale render error", async () => {
    const first = deferred<string>();
    m.renderMermaid.mockReturnValueOnce(first.promise).mockResolvedValueOnce(SVG);

    type("old");
    await vi.advanceTimersByTimeAsync(250);
    type("new");
    await vi.advanceTimersByTimeAsync(250);
    first.reject(new Error("late failure"));
    await flush();

    expect(errorText()).toBe("");
    expect(preview().querySelector("svg")).not.toBeNull();
  });

  it("discards a pending render when the input is cleared", async () => {
    const first = deferred<string>();
    m.renderMermaid.mockReturnValueOnce(first.promise);

    type("old");
    await vi.advanceTimersByTimeAsync(250);
    type("");
    await vi.advanceTimersByTimeAsync(250);
    first.resolve(SVG);
    await flush();

    expect(preview().textContent).toBe("Type Mermaid code");
    expect(insertBtn().disabled).toBe(true);
  });
});

describe("sheet diagram dropdown", () => {
  beforeEach(async () => {
    m.listDiagrams.mockResolvedValue(DIAGRAMS);
    await start();
  });

  it("loads the selected diagram's code into the input and renders it immediately", async () => {
    m.renderMermaid.mockClear();
    selectDiagram("id-2");
    await flush();

    expect(input().value).toBe("pie\n\"a\": 1");
    expect(m.renderMermaid).toHaveBeenCalledWith("pie\n\"a\": 1");
  });

  it("leaves the input alone when '(No diagram selected)' is chosen", async () => {
    selectDiagram("id-1");
    await flush();
    m.renderMermaid.mockClear();
    input().value = "edited";

    selectDiagram("");
    await flush();
    expect(input().value).toBe("edited");
    expect(m.renderMermaid).not.toHaveBeenCalled();
  });

  it("reloads the list on focus, keeping the selection when it still exists", async () => {
    selectDiagram("id-2");
    await flush();
    m.listDiagrams.mockResolvedValue([
      DIAGRAMS[1],
      { id: "id-3", name: "Diagram 3", altTextDescription: "" },
    ]);

    diagramSelect().dispatchEvent(new Event("focus"));
    await flush();

    expect(optionLabels()).toEqual([
      ["", "(No diagram selected)"],
      ["id-2", "Diagram 2"],
      ["id-3", "Diagram 3"],
    ]);
    expect(diagramSelect().value).toBe("id-2");
  });

  it("falls back to '(No diagram selected)' when the selected diagram is gone", async () => {
    selectDiagram("id-1");
    await flush();
    m.listDiagrams.mockResolvedValue([DIAGRAMS[1]]);

    diagramSelect().dispatchEvent(new Event("focus"));
    await flush();
    expect(diagramSelect().value).toBe("");
  });

  it("empties the list without showing an error when loading fails", async () => {
    m.listDiagrams.mockRejectedValue(new Error("Excel busy"));

    diagramSelect().dispatchEvent(new Event("focus"));
    await flush();
    expect(optionLabels()).toEqual([["", "(No diagram selected)"]]);
    expect(errorText()).toBe("");
  });

  it("resets the list when another sheet is activated", async () => {
    selectDiagram("id-1");
    await flush();

    const handler = m.onSheetActivated.mock.calls[0][0] as () => void;
    handler();
    expect(optionLabels()).toEqual([["", "(No diagram selected)"]]);
    expect(diagramSelect().value).toBe("");
  });
});

describe("insert", () => {
  beforeEach(async () => {
    m.listDiagrams.mockResolvedValue(DIAGRAMS);
    await start();
  });

  it.each(["png", "jpeg"] as const)("inserts a %s image of the preview with the trimmed code", async (format) => {
    setFormat(format);
    input().value = "  graph TD\nA-->B \n";
    await click(insertBtn());

    expect(m.svgToRaster).toHaveBeenCalledWith(preview().querySelector("svg"), format);
    expect(m.insertImageShape).toHaveBeenCalledWith(RASTER, "graph TD\nA-->B", undefined);
    expect(m.insertSvgShape).not.toHaveBeenCalled();
    expect(errorText()).toBe("");
  });

  it("replaces the selected sheet diagram", async () => {
    selectDiagram("id-1");
    await flush();
    await click(insertBtn());

    expect(m.insertImageShape).toHaveBeenCalledWith(RASTER, "graph TD\nX-->Y", "id-1");
  });

  it("reloads the list and selects the inserted shape", async () => {
    m.listDiagrams.mockResolvedValue([...DIAGRAMS, { id: "new-id", name: "New", altTextDescription: "" }]);
    await click(insertBtn());

    expect(diagramSelect().value).toBe("new-id");
  });

  it("inserts the serialized SVG when SVG is selected and supported", async () => {
    setFormat("svg");
    selectDiagram("id-2");
    await flush();
    await click(insertBtn());

    const svg = new XMLSerializer().serializeToString(preview().querySelector("svg")!);
    expect(m.insertSvgShape).toHaveBeenCalledWith(svg, "pie\n\"a\": 1", "id-2");
    expect(m.svgToRaster).not.toHaveBeenCalled();
  });

  it("refuses SVG insertion when ImageCoercion 1.2 is unsupported", async () => {
    setFormat("svg");
    m.isSvgInsertSupported.mockReturnValue(false);
    await click(insertBtn());

    expect(m.insertSvgShape).not.toHaveBeenCalled();
    expect(errorText()).toBe(
      "Insert error: SVG insertion is not supported in this version of Excel. Use PNG or JPEG.",
    );
  });

  it("reports insertion failures", async () => {
    m.insertImageShape.mockRejectedValue(new Error("Selected diagram no longer exists."));
    await click(insertBtn());

    expect(errorText()).toBe("Insert error: Selected diagram no longer exists.");
    expect(m.listDiagrams).toHaveBeenCalledTimes(1);
  });

  it("clears the previous error when clicked again", async () => {
    m.insertImageShape.mockRejectedValueOnce(new Error("fail"));
    await click(insertBtn());
    expect(errorText()).toBe("Insert error: fail");

    await click(insertBtn());
    expect(errorText()).toBe("");
  });

  it("asks to render first when the preview has no SVG", async () => {
    m.renderMermaid.mockResolvedValue("<p>no svg</p>");
    type("graph TD");
    await vi.advanceTimersByTimeAsync(250);
    await click(insertBtn());

    expect(errorText()).toBe("Please render first");
    expect(m.insertImageShape).not.toHaveBeenCalled();
  });
});

describe("download", () => {
  beforeEach(start);

  it("downloads PNG as diagram.png", async () => {
    await click(downloadBtn());

    expect(m.svgToRaster).toHaveBeenCalledWith(preview().querySelector("svg"), "png");
    expect(downloads).toEqual([{ href: "data:image/png;base64,B64", download: "diagram.png" }]);
  });

  it("downloads JPEG as diagram.jpg", async () => {
    setFormat("jpeg");
    await click(downloadBtn());

    expect(m.svgToRaster).toHaveBeenCalledWith(preview().querySelector("svg"), "jpeg");
    expect(downloads).toEqual([{ href: "data:image/jpeg;base64,B64", download: "diagram.jpg" }]);
  });

  it("downloads SVG as diagram.svg without rasterizing", async () => {
    setFormat("svg");
    await click(downloadBtn());

    const svg = new XMLSerializer().serializeToString(preview().querySelector("svg")!);
    expect(downloads).toEqual([
      { href: "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg), download: "diagram.svg" },
    ]);
    expect(m.svgToRaster).not.toHaveBeenCalled();
  });

  it("does not touch Excel", async () => {
    await click(downloadBtn());
    expect(m.insertImageShape).not.toHaveBeenCalled();
    expect(m.insertSvgShape).not.toHaveBeenCalled();
  });

  it("reports conversion failures", async () => {
    m.svgToRaster.mockRejectedValue(new Error("Failed to load SVG image."));
    await click(downloadBtn());

    expect(errorText()).toBe("Download error: Failed to load SVG image.");
    expect(downloads).toEqual([]);
  });

  it("asks to render first when the preview has no SVG", async () => {
    m.renderMermaid.mockResolvedValue("<p>no svg</p>");
    type("graph TD");
    await vi.advanceTimersByTimeAsync(250);
    await click(downloadBtn());

    expect(errorText()).toBe("Please render first");
    expect(downloads).toEqual([]);
  });
});
