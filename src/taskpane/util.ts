// Excel shape sizes are in points (1px = 0.75pt at 96 DPI).
export const PX_TO_PT = 0.75;

/** Returns the element with the given ID, throwing if it does not exist. */
export function byId<T extends HTMLElement>(id: string): T {
  const el = document.getElementById(id);
  if (!el) {
    throw new Error(`Element #${id} not found.`);
  }
  return el as T;
}

export function errorMessage(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
}
