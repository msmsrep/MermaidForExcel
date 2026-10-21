type Box = { x: number; y: number; width: number; height: number };

/** Sum of `translate(x, y)` in the element's own transform attribute. */
function ownTranslate(el: Element): { x: number; y: number } {
  let x = 0;
  let y = 0;
  for (const m of (el.getAttribute("transform") ?? "").matchAll(/translate\(\s*([-\d.]+)[ ,]+([-\d.]+)\s*\)/g)) {
    x += Number(m[1]);
    y += Number(m[2]);
  }
  return { x, y };
}

function polyline(el: Element): Array<{ x: number; y: number }> {
  if (el.tagName.toLowerCase() === "line") {
    return [{ x: attr(el, "x1"), y: attr(el, "y1") }, { x: attr(el, "x2"), y: attr(el, "y2") }];
  }
  const nums = (el.getAttribute("d") ?? "").match(/-?[\d.]+/g)?.map(Number) ?? [];
  const points = [];
  for (let i = 0; i + 1 < nums.length; i += 2) points.push({ x: nums[i], y: nums[i + 1] });
  return points;
}

const attr = (el: Element, name: string) => Number(el.getAttribute(name) ?? 0);

/**
 * getBBox in the element's own user space: rect/circle/polygon from their
 * attributes, `data-bbox="x y w h"` for anything else, and the union of the
 * children (with their transforms) for groups.
 */
function bbox(el: Element): Box | null {
  const fixed = el.getAttribute("data-bbox");
  if (fixed) {
    const [x, y, width, height] = fixed.split(/\s+/).map(Number);
    return { x, y, width, height };
  }
  switch (el.tagName.toLowerCase()) {
    case "rect":
      return { x: attr(el, "x"), y: attr(el, "y"), width: attr(el, "width"), height: attr(el, "height") };
    case "circle": {
      const r = attr(el, "r");
      return { x: attr(el, "cx") - r, y: attr(el, "cy") - r, width: 2 * r, height: 2 * r };
    }
    case "polygon": {
      const nums = (el.getAttribute("points") ?? "").split(/[\s,]+/).filter(Boolean).map(Number);
      const xs = nums.filter((_, i) => i % 2 === 0);
      const ys = nums.filter((_, i) => i % 2 === 1);
      const x = Math.min(...xs);
      const y = Math.min(...ys);
      return { x, y, width: Math.max(...xs) - x, height: Math.max(...ys) - y };
    }
  }
  const boxes = Array.from(el.children).flatMap((child) => {
    const b = bbox(child);
    if (!b) return [];
    const t = ownTranslate(child);
    return [{ ...b, x: b.x + t.x, y: b.y + t.y }];
  });
  if (boxes.length === 0) return null;
  const x = Math.min(...boxes.map((b) => b.x));
  const y = Math.min(...boxes.map((b) => b.y));
  return {
    x,
    y,
    width: Math.max(...boxes.map((b) => b.x + b.width)) - x,
    height: Math.max(...boxes.map((b) => b.y + b.height)) - y,
  };
}

/**
 * Installs getBBox, getScreenCTM (translations only) and the text/path
 * measurement methods, which jsdom lacks.
 * The root <svg> sits at `screenOffset` so code must convert to root
 * coordinates.
 */
export function stubSvgGeometry(screenOffset = { x: 5, y: 7 }) {
  Object.assign(SVGElement.prototype, {
    getBBox(this: Element): Box {
      return bbox(this) ?? { x: 0, y: 0, width: 0, height: 0 };
    },
    /** From `data-start="x y"`; jsdom does no text layout. */
    getStartPositionOfChar(this: Element) {
      const [x, y] = (this.getAttribute("data-start") ?? "0 0").split(/\s+/).map(Number);
      return { x, y };
    },
    /** From `data-extent="x y w h"`, else a 16px cell on an alphabetic baseline at 0 (ascent 12). */
    getExtentOfChar(this: Element) {
      const [x, y, width, height] = (this.getAttribute("data-extent") ?? "0 -12 8 16").split(/\s+/).map(Number);
      return { x, y, width, height };
    },
    getNumberOfChars(this: Element) {
      return (this.textContent ?? "").length;
    },
    /** Paths are treated as polylines through the numbers in `d`. */
    getTotalLength(this: Element) {
      const p = polyline(this);
      return p.slice(1).reduce((sum, q, i) => sum + Math.hypot(q.x - p[i].x, q.y - p[i].y), 0);
    },
    getPointAtLength(this: Element, length: number) {
      const p = polyline(this);
      let rest = length;
      for (let i = 1; i < p.length; i++) {
        const seg = Math.hypot(p[i].x - p[i - 1].x, p[i].y - p[i - 1].y);
        if (rest <= seg) {
          const t = seg ? rest / seg : 0;
          return { x: p[i - 1].x + (p[i].x - p[i - 1].x) * t, y: p[i - 1].y + (p[i].y - p[i - 1].y) * t };
        }
        rest -= seg;
      }
      return p[p.length - 1];
    },
    getScreenCTM(this: Element) {
      let e = screenOffset.x;
      let f = screenOffset.y;
      for (let el: Element | null = this; el && el.tagName.toLowerCase() !== "svg"; el = el.parentElement) {
        const t = ownTranslate(el);
        e += t.x;
        f += t.y;
      }
      return { a: 1, b: 0, c: 0, d: 1, e, f };
    },
  });
}
