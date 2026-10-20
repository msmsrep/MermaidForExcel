const MIN_INPUT_HEIGHT = 40;
const MIN_PREVIEW_HEIGHT = 120;
// Pixels moved per arrow key press.
const KEYBOARD_STEP = 10;

/**
 * Makes the splitter resize the input by pointer drag or arrow keys, keeping
 * the preview at least MIN_PREVIEW_HEIGHT tall.
 */
export function initSplitter(
  splitter: HTMLElement,
  input: HTMLElement,
  preview: HTMLElement,
) {
  function setInputHeight(height: number) {
    const max = Math.max(
      MIN_INPUT_HEIGHT,
      input.offsetHeight + preview.offsetHeight - MIN_PREVIEW_HEIGHT,
    );
    input.style.height = `${Math.min(max, Math.max(MIN_INPUT_HEIGHT, height))}px`;
  }

  splitter.addEventListener("pointerdown", (e) => {
    e.preventDefault();
    const startY = e.clientY;
    const startHeight = input.offsetHeight;
    splitter.setPointerCapture(e.pointerId);
    splitter.classList.add("dragging");

    const onMove = (ev: PointerEvent) => {
      setInputHeight(startHeight + ev.clientY - startY);
    };
    const onEnd = () => {
      splitter.classList.remove("dragging");
      splitter.removeEventListener("pointermove", onMove);
      splitter.removeEventListener("pointerup", onEnd);
      splitter.removeEventListener("pointercancel", onEnd);
    };
    splitter.addEventListener("pointermove", onMove);
    splitter.addEventListener("pointerup", onEnd);
    splitter.addEventListener("pointercancel", onEnd);
  });

  splitter.addEventListener("keydown", (e) => {
    if (e.key !== "ArrowUp" && e.key !== "ArrowDown") return;
    e.preventDefault();
    const delta = e.key === "ArrowUp" ? -KEYBOARD_STEP : KEYBOARD_STEP;
    setInputHeight(input.offsetHeight + delta);
  });
}
