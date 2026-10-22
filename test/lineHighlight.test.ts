import { beforeEach, describe, expect, it } from "vitest";
import { initLineHighlight, type LineHighlight } from "../src/taskpane/lineHighlight";

let input: HTMLTextAreaElement;
let backdrop: HTMLDivElement;
let highlight: LineHighlight;

const lines = () => Array.from(backdrop.children, (el) => [el.textContent, el.className]);

beforeEach(() => {
  document.body.innerHTML = '<div id="backdrop"></div><textarea id="input"></textarea>';
  input = document.querySelector("#input")!;
  backdrop = document.querySelector("#backdrop")!;
  input.value = "graph TD\n\nA-->B";
  highlight = initLineHighlight(input, backdrop);
});

function type(value: string) {
  input.value = value;
  input.dispatchEvent(new Event("input"));
}

describe("show / clear", () => {
  it("copies the text into the backdrop, one block per line, marking the line", () => {
    highlight.show(2);
    expect(lines()).toEqual([
      ["graph TD", ""],
      [" ", ""],
      ["A-->B", "error-line"],
    ]);
  });

  it("empties the backdrop when cleared", () => {
    highlight.show(0);
    highlight.clear();
    expect(backdrop.children).toHaveLength(0);
  });

  it("keeps the same line highlighted while the text is edited", () => {
    highlight.show(1);
    type("graph LR\nX\nY");
    expect(lines()).toEqual([
      ["graph LR", ""],
      ["X", "error-line"],
      ["Y", ""],
    ]);
  });

  it("draws nothing on edits when no line is highlighted", () => {
    type("graph LR");
    expect(backdrop.children).toHaveLength(0);
  });

  it("matches the backdrop to the textarea's size and scrolling", () => {
    Object.defineProperty(input, "clientWidth", { value: 180 });
    Object.defineProperty(input, "clientHeight", { value: 90 });
    highlight.show(0);
    expect(backdrop.style.width).toBe("180px");
    expect(backdrop.style.height).toBe("90px");

    input.scrollTop = 40;
    input.dispatchEvent(new Event("scroll"));
    expect(backdrop.scrollTop).toBe(40);
  });
});

describe("goTo", () => {
  it("focuses the input and selects the line", () => {
    highlight.goTo(2);
    expect(document.activeElement).toBe(input);
    expect([input.selectionStart, input.selectionEnd]).toEqual([10, 15]);
  });

  it("scrolls the highlighted line into view", () => {
    Object.defineProperty(input, "clientHeight", { value: 90 });
    highlight.show(2);
    Object.defineProperty(backdrop.children[2], "offsetTop", { value: 100 });
    highlight.goTo(2);
    expect(input.scrollTop).toBe(70);
  });

  it("ignores lines past the end", () => {
    highlight.goTo(5);
    expect(document.activeElement).not.toBe(input);
  });
});
