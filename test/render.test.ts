import { beforeEach, describe, expect, it, vi } from "vitest";

const mermaid = vi.hoisted(() => ({ initialize: vi.fn(), render: vi.fn() }));
vi.mock("mermaid", () => ({ default: mermaid }));

import { initMermaid, renderMermaid } from "../src/taskpane/render";

beforeEach(() => {
  mermaid.initialize.mockReset();
  mermaid.render.mockReset();
  document.body.innerHTML = "";
});

describe("initMermaid", () => {
  it("disables auto rendering and HTML labels, and renders at natural size", () => {
    initMermaid();
    expect(mermaid.initialize).toHaveBeenCalledWith({
      startOnLoad: false,
      theme: "default",
      htmlLabels: false,
      flowchart: { useMaxWidth: false, htmlLabels: false },
    });
  });
});

describe("renderMermaid", () => {
  it("renders the code under the fixed graph ID and returns the SVG", async () => {
    mermaid.render.mockResolvedValue({ svg: "<svg>ok</svg>" });

    expect(await renderMermaid("graph TD\nA-->B")).toBe("<svg>ok</svg>");
    expect(mermaid.render).toHaveBeenCalledWith("mermaid-graph", "graph TD\nA-->B");
  });

  it("removes a leftover element with the graph ID before rendering", async () => {
    document.body.innerHTML = '<div id="mermaid-graph"></div>';
    mermaid.render.mockImplementation(async () => {
      expect(document.getElementById("mermaid-graph")).toBeNull();
      return { svg: "<svg/>" };
    });

    await renderMermaid("pie");
    expect(mermaid.render).toHaveBeenCalledOnce();
  });

  it("propagates render errors", async () => {
    mermaid.render.mockRejectedValue(new Error("Parse error on line 1"));
    await expect(renderMermaid("bad")).rejects.toThrow("Parse error on line 1");
  });
});
