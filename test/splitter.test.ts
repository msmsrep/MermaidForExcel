import { beforeEach, describe, expect, it, vi } from "vitest";
import { initSplitter } from "../src/taskpane/splitter";

// Input and preview share a fixed amount of vertical space, as in the task pane.
const TOTAL_HEIGHT = 400;
const INITIAL_INPUT_HEIGHT = 100;

let splitter: HTMLDivElement;
let input: HTMLTextAreaElement;
let preview: HTMLDivElement;

const inputHeight = () => parseFloat(input.style.height) || INITIAL_INPUT_HEIGHT;

beforeEach(() => {
  document.body.innerHTML =
    '<textarea id="input"></textarea><div id="splitter"></div><div id="preview"></div>';
  input = document.querySelector("#input")!;
  splitter = document.querySelector("#splitter")!;
  preview = document.querySelector("#preview")!;

  // jsdom has no layout, so derive offsetHeight from the inline style.
  Object.defineProperty(input, "offsetHeight", { get: inputHeight });
  Object.defineProperty(preview, "offsetHeight", { get: () => TOTAL_HEIGHT - inputHeight() });
  splitter.setPointerCapture = vi.fn();

  initSplitter(splitter, input, preview);
});

function key(k: string) {
  const e = new KeyboardEvent("keydown", { key: k, cancelable: true });
  splitter.dispatchEvent(e);
  return e;
}

function pointer(type: string, clientY: number) {
  const e = new PointerEvent(type, { clientY, pointerId: 7, cancelable: true });
  splitter.dispatchEvent(e);
  return e;
}

describe("keyboard", () => {
  it("shrinks the input by 10px with ArrowUp and grows it with ArrowDown", () => {
    key("ArrowUp");
    expect(input.style.height).toBe("90px");
    key("ArrowDown");
    key("ArrowDown");
    expect(input.style.height).toBe("110px");
  });

  it("prevents the default action only for arrow keys", () => {
    expect(key("ArrowUp").defaultPrevented).toBe(true);
    expect(key("Enter").defaultPrevented).toBe(false);
    expect(input.style.height).toBe("90px");
  });

  it("does not shrink the input below 40px", () => {
    for (let i = 0; i < 10; i++) key("ArrowUp");
    expect(input.style.height).toBe("40px");
  });

  it("does not grow the input past the point where the preview drops below 120px", () => {
    for (let i = 0; i < 30; i++) key("ArrowDown");
    expect(input.style.height).toBe(`${TOTAL_HEIGHT - 120}px`);
  });
});

describe("pointer drag", () => {
  it("resizes the input by the distance dragged", () => {
    const down = pointer("pointerdown", 200);
    expect(down.defaultPrevented).toBe(true);
    expect(splitter.setPointerCapture).toHaveBeenCalledWith(7);

    pointer("pointermove", 250);
    expect(input.style.height).toBe("150px");
    // Measured from the drag start, not the previous move.
    pointer("pointermove", 180);
    expect(input.style.height).toBe("80px");
  });

  it("clamps the height while dragging", () => {
    pointer("pointerdown", 200);
    pointer("pointermove", 0);
    expect(input.style.height).toBe("40px");
    pointer("pointermove", 1000);
    expect(input.style.height).toBe(`${TOTAL_HEIGHT - 120}px`);
  });

  it.each(["pointerup", "pointercancel"])("marks the splitter as dragging until %s", (endType) => {
    pointer("pointerdown", 200);
    expect(splitter.classList.contains("dragging")).toBe(true);

    pointer(endType, 220);
    expect(splitter.classList.contains("dragging")).toBe(false);

    pointer("pointermove", 300);
    expect(input.style.height).toBe("");
  });

  it("ignores pointer moves without a preceding pointerdown", () => {
    pointer("pointermove", 300);
    expect(input.style.height).toBe("");
  });
});
