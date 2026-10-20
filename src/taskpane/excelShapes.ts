import type { RasterImage } from "./image";

// Marks images inserted by this add-in; the Mermaid code goes in the description.
const ALT_TEXT_TITLE = "Mermaid Diagram";
// Excel shape sizes are in points (1px = 0.75pt at 96 DPI).
const PX_TO_PT = 0.75;

export type DiagramItem = { id: string; name: string; altTextDescription: string };
type Placement = { left: number; top: number; name?: string };

function activeShapes(ctx: Excel.RequestContext): Excel.ShapeCollection {
  return ctx.workbook.worksheets.getActiveWorksheet().shapes;
}

/** Lists images on the active sheet that were inserted by this add-in. */
export async function listDiagrams(): Promise<DiagramItem[]> {
  return Excel.run(async (ctx) => {
    const shapes = activeShapes(ctx);
    shapes.load(
      "items/id,items/name,items/type,items/altTextTitle,items/altTextDescription",
    );
    await ctx.sync();

    return shapes.items
      .filter(
        (shape) =>
          shape.type === "Image" && shape.altTextTitle === ALT_TEXT_TITLE,
      )
      .map((shape) => ({
        id: shape.id,
        name: shape.name,
        altTextDescription: shape.altTextDescription ?? "",
      }));
  });
}

export async function onSheetActivated(handler: () => void): Promise<void> {
  return Excel.run(async (ctx) => {
    ctx.workbook.worksheets.onActivated.add(async () => handler());
    await ctx.sync();
  });
}

export function isSvgInsertSupported(): boolean {
  return !!Office.context.requirements?.isSetSupported("ImageCoercion", "1.2");
}

/**
 * Loads where a new shape should go: the replacement target's position and
 * name, or the active cell's top-left corner when there is no target.
 */
async function loadPlacement(
  ctx: Excel.RequestContext,
  replaceId?: string,
): Promise<Placement> {
  const cell = ctx.workbook.getActiveCell();
  cell.load(["left", "top"]);
  const target = replaceId
    ? activeShapes(ctx).getItemOrNullObject(replaceId)
    : undefined;
  target?.load(["isNullObject", "left", "top", "name"]);
  await ctx.sync();

  if (!target) {
    return { left: cell.left, top: cell.top };
  }
  if (target.isNullObject) {
    throw new Error("Selected diagram no longer exists.");
  }
  return { left: target.left, top: target.top, name: target.name };
}

function applyShapeProps(
  shape: Excel.Shape,
  placement: Placement,
  mermaidCode: string,
) {
  shape.left = placement.left;
  shape.top = placement.top;
  if (placement.name) {
    shape.name = placement.name;
  }
  shape.altTextTitle = ALT_TEXT_TITLE;
  shape.altTextDescription = mermaidCode;
}

/** Adds a PNG/JPEG image and, when replacing, deletes the old shape in the same sync. */
export function insertImageShape(
  image: RasterImage,
  mermaidCode: string,
  replaceId?: string,
): Promise<string> {
  return Excel.run(async (ctx) => {
    const placement = await loadPlacement(ctx, replaceId);
    const shapes = activeShapes(ctx);

    const shape = shapes.addImage(image.base64);
    shape.width = image.width * PX_TO_PT;
    shape.height = image.height * PX_TO_PT;
    applyShapeProps(shape, placement, mermaidCode);
    shape.load("id");
    if (replaceId) {
      shapes.getItem(replaceId).delete();
    }
    await ctx.sync();
    return shape.id;
  });
}

/**
 * Inserts an SVG through setSelectedDataAsync. The new shape is found by
 * diffing shape IDs before and after insertion, and the old shape is deleted
 * only after the new one is in place.
 */
export async function insertSvgShape(
  svgStr: string,
  mermaidCode: string,
  replaceId?: string,
): Promise<string> {
  const before = await Excel.run(async (ctx) => {
    const shapes = activeShapes(ctx);
    shapes.load("items/id");
    const placement = await loadPlacement(ctx, replaceId);
    return { ids: new Set(shapes.items.map((s) => s.id)), placement };
  });

  await insertSvgToSelection(svgStr);

  return Excel.run(async (ctx) => {
    const shapes = activeShapes(ctx);
    shapes.load("items/id");
    await ctx.sync();

    const added = shapes.items.filter((s) => !before.ids.has(s.id));
    if (added.length !== 1) {
      throw new Error("Could not identify the inserted shape.");
    }
    const shape = added[0];
    applyShapeProps(shape, before.placement, mermaidCode);
    await ctx.sync();

    if (replaceId) {
      const target = shapes.getItemOrNullObject(replaceId);
      target.load("isNullObject");
      await ctx.sync();
      if (!target.isNullObject) {
        target.delete();
        await ctx.sync();
      }
    }
    return shape.id;
  });
}

function insertSvgToSelection(svgStr: string): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(
      svgStr,
      { coercionType: Office.CoercionType.XmlSvg },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(asyncResult.error.message));
        } else {
          resolve();
        }
      },
    );
  });
}
