/**
 * Z-index clearance for FinSheets' body-portaled floating UI.
 *
 * Menus, dropdowns, toolbar tooltips and modal dialogs in FinSheets render through
 * React portals to `document.body`. That takes them OUT of FinSheets' own DOM
 * subtree, so their z-index competes with the HOST application's stacking context —
 * not just FinSheets' internal layers.
 *
 * When FinSheets is embedded inside a host overlay (e.g. an in-place editor mounted
 * full-screen at a high z-index over the rest of the app), a body-portaled popup at
 * a low z-index (menus were `1000`) renders BEHIND that opaque overlay and is
 * invisible — the "click File, nothing opens" bug. Elements that are NOT portaled to
 * `document.body` (rendered inline in the editor, even when `position: fixed`) ride
 * the editor's own stacking context and do not need this clearance.
 *
 * `Z_BASE` lifts every body-portaled layer into a high band, comfortably above any
 * reasonable host overlay, while each call site keeps adding its original in-band
 * offset so FinSheets' floating layers preserve the exact relative order they always
 * had among themselves. (The transient "Saved ✓" toast is the one exception: it
 * intentionally sits at the max 32-bit int so it clears even this band.)
 */
export const Z_BASE = 2_000_000;
