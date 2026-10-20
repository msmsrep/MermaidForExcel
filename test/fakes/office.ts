import { vi } from "vitest";

/** In-memory stand-in for a shape on the active worksheet. */
export type ShapeState = {
  id: string;
  name: string;
  type: string;
  altTextTitle: string;
  altTextDescription: string;
  left: number;
  top: number;
  width: number;
  height: number;
  base64?: string;
};

/** Operation log entry; `batch` is the number of syncs completed before it. */
export type Op = { op: "addImage" | "delete" | "setSelectedData"; id?: string; batch: number };

/**
 * Minimal fake of the Excel / Office APIs used by excelShapes.ts. Property
 * writes take effect immediately; `ops` records which sync batch each
 * mutating call belongs to so tests can check batching.
 */
export function createFakeOffice(init: Partial<ShapeState>[] = []) {
  let nextId = 1;
  const newId = () => `shape-${nextId++}`;

  const shapes: ShapeState[] = init.map((s) => ({
    id: newId(),
    name: "Picture",
    type: "Image",
    altTextTitle: "",
    altTextDescription: "",
    left: 0,
    top: 0,
    width: 100,
    height: 100,
    ...s,
  }));
  const ops: Op[] = [];
  const activeCell = { left: 10, top: 20 };
  const activatedHandlers: Array<() => unknown> = [];
  let syncCount = 0;

  /** How setSelectedDataAsync behaves: fail, or add this many shapes. */
  const svgInsert: { result: "ok" | "fail"; errorMessage: string; addShapes: number } = {
    result: "ok",
    errorMessage: "Insert failed",
    addShapes: 1,
  };
  let svgSupported = true;

  function shapeProxy(state: ShapeState) {
    return {
      get id() { return state.id; },
      get type() { return state.type; },
      get name() { return state.name; },
      set name(v: string) { state.name = v; },
      get altTextTitle() { return state.altTextTitle; },
      set altTextTitle(v: string) { state.altTextTitle = v; },
      get altTextDescription() { return state.altTextDescription; },
      set altTextDescription(v: string) { state.altTextDescription = v; },
      get left() { return state.left; },
      set left(v: number) { state.left = v; },
      get top() { return state.top; },
      set top(v: number) { state.top = v; },
      get width() { return state.width; },
      set width(v: number) { state.width = v; },
      get height() { return state.height; },
      set height(v: number) { state.height = v; },
      isNullObject: false,
      load: vi.fn(),
      delete() {
        ops.push({ op: "delete", id: state.id, batch: syncCount });
        const i = shapes.indexOf(state);
        if (i >= 0) shapes.splice(i, 1);
      },
    };
  }

  function find(id: string) {
    return shapes.find((s) => s.id === id);
  }

  const shapeCollection = {
    get items() { return shapes.map(shapeProxy); },
    load: vi.fn(),
    getItem(id: string) {
      const s = find(id);
      if (!s) throw new Error(`ItemNotFound: ${id}`);
      return shapeProxy(s);
    },
    getItemOrNullObject(id: string) {
      const s = find(id);
      return s ? shapeProxy(s) : { isNullObject: true, load: vi.fn(), delete: vi.fn() };
    },
    addImage(base64: string) {
      const id = newId();
      const state: ShapeState = {
        id, name: `Default ${id}`, type: "Image",
        altTextTitle: "", altTextDescription: "",
        left: 0, top: 0, width: 50, height: 50, base64,
      };
      shapes.push(state);
      ops.push({ op: "addImage", id: state.id, batch: syncCount });
      return shapeProxy(state);
    },
  };

  const ctx = {
    workbook: {
      worksheets: {
        getActiveWorksheet: () => ({ shapes: shapeCollection }),
        onActivated: { add: vi.fn((h: () => unknown) => { activatedHandlers.push(h); }) },
      },
      getActiveCell: () => ({ ...activeCell, load: vi.fn() }),
    },
    sync: vi.fn(async () => { syncCount++; }),
  };

  const Excel = {
    run: vi.fn(async (fn: (c: typeof ctx) => unknown) => fn(ctx)),
  };

  const Office = {
    CoercionType: { XmlSvg: "xmlSvg" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    context: {
      requirements: {
        isSetSupported: vi.fn((name: string, version: string) =>
          svgSupported && name === "ImageCoercion" && version === "1.2"),
      },
      document: {
        setSelectedDataAsync: vi.fn(
          (data: string, _opts: unknown, cb: (r: unknown) => void) => {
            ops.push({ op: "setSelectedData", batch: syncCount });
            if (svgInsert.result === "fail") {
              cb({ status: "failed", error: { message: svgInsert.errorMessage } });
              return;
            }
            for (let i = 0; i < svgInsert.addShapes; i++) {
              const id = newId();
              shapes.push({
                id, name: `Default ${id}`, type: "Image",
                altTextTitle: "", altTextDescription: "",
                left: 300, top: 300, width: 80, height: 80, base64: data,
              });
            }
            cb({ status: "succeeded" });
          },
        ),
      },
    },
  };

  return {
    Excel,
    Office,
    ctx,
    shapes,
    ops,
    activeCell,
    svgInsert,
    activatedHandlers,
    get syncCount() { return syncCount; },
    setSvgSupported(v: boolean) { svgSupported = v; },
    /** Installs the fakes as the `Excel` and `Office` globals. */
    install() {
      vi.stubGlobal("Excel", Excel);
      vi.stubGlobal("Office", Office);
    },
  };
}

export type FakeOffice = ReturnType<typeof createFakeOffice>;
