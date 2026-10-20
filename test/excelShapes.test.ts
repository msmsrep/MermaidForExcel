import { beforeEach, describe, expect, it } from "vitest";
import {
  insertImageShape,
  insertSvgShape,
  isSvgInsertSupported,
  listDiagrams,
  onSheetActivated,
} from "../src/taskpane/excelShapes";
import { createFakeOffice, type FakeOffice, type ShapeState } from "./fakes/office";

const TITLE = "Mermaid Diagram";
const image = { base64: "AAAA", width: 200, height: 100 };

let fake: FakeOffice;

function setup(...shapes: Partial<ShapeState>[]) {
  fake = createFakeOffice(shapes);
  fake.install();
}

describe("listDiagrams", () => {
  it("returns only images whose alt text title marks them as Mermaid diagrams", async () => {
    setup(
      { name: "Diagram 1", altTextTitle: TITLE, altTextDescription: "graph TD\nA-->B" },
      { name: "Photo", altTextTitle: "Other" },
      { name: "Rect", type: "GeometricShape", altTextTitle: TITLE },
      { name: "Diagram 2", altTextTitle: TITLE, altTextDescription: "pie" },
    );

    expect(await listDiagrams()).toEqual([
      { id: "shape-1", name: "Diagram 1", altTextDescription: "graph TD\nA-->B" },
      { id: "shape-4", name: "Diagram 2", altTextDescription: "pie" },
    ]);
  });

  it("returns an empty description when the shape has none", async () => {
    setup({ name: "D", altTextTitle: TITLE, altTextDescription: undefined as unknown as string });
    expect(await listDiagrams()).toEqual([{ id: "shape-1", name: "D", altTextDescription: "" }]);
  });

  it("returns an empty list when the sheet has no diagrams", async () => {
    setup();
    expect(await listDiagrams()).toEqual([]);
  });
});

describe("onSheetActivated", () => {
  it("registers a handler that is called when a worksheet is activated", async () => {
    setup();
    let calls = 0;
    await onSheetActivated(() => calls++);

    expect(fake.activatedHandlers).toHaveLength(1);
    await fake.activatedHandlers[0]();
    expect(calls).toBe(1);
  });
});

describe("isSvgInsertSupported", () => {
  beforeEach(() => setup());

  it("is true when ImageCoercion 1.2 is supported", () => {
    expect(isSvgInsertSupported()).toBe(true);
    expect(fake.Office.context.requirements.isSetSupported).toHaveBeenCalledWith("ImageCoercion", "1.2");
  });

  it("is false when ImageCoercion 1.2 is not supported", () => {
    fake.setSvgSupported(false);
    expect(isSvgInsertSupported()).toBe(false);
  });
});

describe("insertImageShape", () => {
  it("adds the image at the active cell with the SVG size converted to points", async () => {
    setup();
    const id = await insertImageShape(image, "graph TD\nA-->B");

    const shape = fake.shapes.find((s) => s.id === id)!;
    expect(shape).toMatchObject({
      base64: "AAAA",
      left: 10,
      top: 20,
      width: 150,
      height: 75,
      altTextTitle: TITLE,
      altTextDescription: "graph TD\nA-->B",
      name: `Default ${id}`,
    });
  });

  it("replaces the target shape, taking over its position and name", async () => {
    setup({ name: "My diagram", altTextTitle: TITLE, left: 111, top: 222 });
    const id = await insertImageShape(image, "pie", "shape-1");

    expect(fake.shapes.map((s) => s.id)).toEqual([id]);
    expect(fake.shapes[0]).toMatchObject({ left: 111, top: 222, name: "My diagram", altTextDescription: "pie" });
  });

  it("adds the new shape and deletes the old one in a single sync", async () => {
    setup({ altTextTitle: TITLE });
    await insertImageShape(image, "pie", "shape-1");

    const add = fake.ops.find((o) => o.op === "addImage")!;
    const del = fake.ops.find((o) => o.op === "delete")!;
    expect(add.batch).toBe(del.batch);
    expect(fake.syncCount).toBe(add.batch + 1);
  });

  it("fails without adding anything when the target no longer exists", async () => {
    setup({ altTextTitle: TITLE });
    await expect(insertImageShape(image, "pie", "missing")).rejects.toThrow(
      "Selected diagram no longer exists.",
    );
    expect(fake.shapes.map((s) => s.id)).toEqual(["shape-1"]);
    expect(fake.ops).toEqual([]);
  });
});

describe("insertSvgShape", () => {
  it("inserts through setSelectedDataAsync as XmlSvg and moves the new shape to the active cell", async () => {
    setup({ name: "Unrelated" });
    const id = await insertSvgShape("<svg/>", "graph TD\nA-->B");

    expect(fake.Office.context.document.setSelectedDataAsync).toHaveBeenCalledWith(
      "<svg/>",
      { coercionType: "xmlSvg" },
      expect.any(Function),
    );
    const shape = fake.shapes.find((s) => s.id === id)!;
    expect(shape).toMatchObject({
      left: 10,
      top: 20,
      altTextTitle: TITLE,
      altTextDescription: "graph TD\nA-->B",
      name: `Default ${id}`,
    });
    expect(fake.shapes.find((s) => s.name === "Unrelated")).toMatchObject({ altTextTitle: "" });
  });

  it("replaces the target shape after the new one is in place", async () => {
    setup({ name: "My diagram", altTextTitle: TITLE, left: 111, top: 222 });
    const id = await insertSvgShape("<svg/>", "pie", "shape-1");

    expect(fake.shapes.map((s) => s.id)).toEqual([id]);
    expect(fake.shapes[0]).toMatchObject({ left: 111, top: 222, name: "My diagram", altTextDescription: "pie" });
    const kinds = fake.ops.map((o) => o.op);
    expect(kinds).toEqual(["setSelectedData", "delete"]);
  });

  it("does not delete anything when the target disappeared during insertion", async () => {
    setup({ altTextTitle: TITLE });
    fake.Office.context.document.setSelectedDataAsync.mockImplementationOnce(
      (_data: string, _opts: unknown, cb: (r: unknown) => void) => {
        fake.shapes.splice(0, 1);
        fake.shapes.push({
          id: "new", name: "Graphic", type: "Image", altTextTitle: "", altTextDescription: "",
          left: 0, top: 0, width: 1, height: 1,
        });
        cb({ status: "succeeded" });
      },
    );

    expect(await insertSvgShape("<svg/>", "pie", "shape-1")).toBe("new");
    expect(fake.ops.filter((o) => o.op === "delete")).toEqual([]);
  });

  it("fails before inserting when the target no longer exists", async () => {
    setup();
    await expect(insertSvgShape("<svg/>", "pie", "missing")).rejects.toThrow(
      "Selected diagram no longer exists.",
    );
    expect(fake.Office.context.document.setSelectedDataAsync).not.toHaveBeenCalled();
  });

  it("reports the setSelectedDataAsync error and keeps the original shape", async () => {
    setup({ altTextTitle: TITLE });
    fake.svgInsert.result = "fail";
    fake.svgInsert.errorMessage = "Not allowed";

    await expect(insertSvgShape("<svg/>", "pie", "shape-1")).rejects.toThrow("Not allowed");
    expect(fake.shapes.map((s) => s.id)).toEqual(["shape-1"]);
  });

  it.each([0, 2])("fails when %i shapes appear instead of one, keeping the original", async (n) => {
    setup({ altTextTitle: TITLE });
    fake.svgInsert.addShapes = n;

    await expect(insertSvgShape("<svg/>", "pie", "shape-1")).rejects.toThrow(
      "Could not identify the inserted shape.",
    );
    expect(fake.shapes.some((s) => s.id === "shape-1")).toBe(true);
  });
});
