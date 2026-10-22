// Fraction of the input's height left above a line scrolled into view.
const SCROLL_MARGIN_RATIO = 1 / 3;

export type LineHighlight = {
  /** Highlights the 0-based line, keeping it highlighted while the text is edited. */
  show(line: number): void;
  clear(): void;
  /** Selects the line in the input and scrolls it into view. */
  goTo(line: number): void;
};

/**
 * Highlights one line of a textarea, which cannot style its own text. The
 * backdrop behind the (transparent) textarea holds a copy of the text with
 * one block per line, so it wraps the same way, and follows its scrolling.
 */
export function initLineHighlight(
  input: HTMLTextAreaElement,
  backdrop: HTMLElement,
): LineHighlight {
  let highlighted: number | undefined;

  function sync() {
    // clientWidth excludes the scrollbar, so the copy wraps at the same width.
    backdrop.style.width = `${input.clientWidth}px`;
    backdrop.style.height = `${input.clientHeight}px`;
    backdrop.scrollTop = input.scrollTop;
  }

  function draw() {
    backdrop.replaceChildren();
    if (highlighted === undefined) {
      return;
    }
    input.value.split("\n").forEach((text, i) => {
      const line = document.createElement("div");
      // A space keeps empty lines one line tall.
      line.textContent = text || " ";
      if (i === highlighted) {
        line.className = "error-line";
      }
      backdrop.appendChild(line);
    });
    sync();
  }

  input.addEventListener("input", draw);
  input.addEventListener("scroll", sync);
  window.addEventListener("resize", sync);

  return {
    show(line) {
      highlighted = line;
      draw();
    },
    clear() {
      highlighted = undefined;
      draw();
    },
    goTo(line) {
      const lines = input.value.split("\n");
      if (line >= lines.length) {
        return;
      }
      const start = lines.slice(0, line).reduce((sum, text) => sum + text.length + 1, 0);
      input.focus();
      input.setSelectionRange(start, start + lines[line].length);
      const block = backdrop.children[line] as HTMLElement | undefined;
      if (block) {
        input.scrollTop = Math.max(0, block.offsetTop - input.clientHeight * SCROLL_MARGIN_RATIO);
      }
    },
  };
}
