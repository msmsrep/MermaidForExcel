import mermaid from "mermaid";

const GRAPH_ID = "mermaid-graph";

export function initMermaid() {
  // htmlLabels is disabled so labels are plain SVG text, which survives
  // rasterization and SVG insertion into Excel.
  // https://github.com/mermaid-js/mermaid-cli/issues/112
  mermaid.initialize({
    startOnLoad: false,
    theme: "default",
    htmlLabels: false,
    flowchart: {
      useMaxWidth: false,
      htmlLabels: false,
    },
  });
}

/** Renders Mermaid code to an SVG string. */
export async function renderMermaid(code: string): Promise<string> {
  // Remove element with the same ID to avoid errors
  document.getElementById(GRAPH_ID)?.remove();
  const { svg } = await mermaid.render(GRAPH_ID, code);
  return svg;
}
