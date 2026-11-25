# Architecture

This file describes how the code is organised for version 6, and the steps that get it there from the structure of
version 5. AGENTS.md describes the code as it is today; this file says where it is going. Update both as the steps land.

## Why

Version 5 builds a worksheet as a plain object that about twenty modules add properties to. About fifty state
properties have no owner: any module reads and writes any of them. About 140 public methods are bound to each worksheet
instance. There were 17 import cycles, which only work because no module calls an imported function while it loads. The
keyboard and clipboard events find their worksheet through one global, `jspreadsheet.current`, which the code reads in
more than 300 places.

Version 6 may break the public API (see [UPGRADING.md](UPGRADING.md)). The goal is not classes for their own sake. The
goal is that each piece of state has one owner, that the modules form a tree instead of a web, and that a change to one
feature touches one place.

## Target

```
src/
  index.ts            jspreadsheet(element, options) and the static helpers
  Spreadsheet.ts      class Spreadsheet: the worksheets, the tabs, the toolbar, the context menu, the plugins
  Worksheet.ts        class Worksheet: the DOM of one worksheet, its options, and its components
  components/         one class per feature, each owning its state
    History.ts        the undo history
    Clipboard.ts      copy, cut and paste
    Editor.ts         opening and closing cell editors, the editor types
    Selection.ts      the selected range, the highlight, aria-selected and the roving tabindex
    Formulas.ts       the formula chain and the formula results
    ...
  events/             the document event handlers, split by input (keyboard, mouse, touch, clipboard, context menu)
  utils/              functions without state (cell names, CSV, HTML helpers)
```

-   **A worksheet owns its components.** `worksheet.history` is a `History`, `worksheet.selection` a `Selection`, and so on.
    A component gets the worksheet in its constructor and keeps its own state in its own fields. Another component asks it
    through its methods, not through its fields.
-   **The public methods stay on the worksheet.** `worksheet.undo()` calls `this.history.undo()`. Applications keep one
    object to talk to.
-   **Methods are on the prototype.** They are not bound to each instance any more.
-   **Imports form a tree.** A component imports the types of the others with `import type` and calls them through the
    worksheet at run time. CI fails on an import cycle.
-   **The document events route through the spreadsheet.** The handlers find the worksheet of the event target, instead of
    reading one global.

## Steps

Each step is one pull request. Each keeps `npm test`, the type check and the accessibility tests passing, adds tests for
what it changes, and lists its breaking changes in UPGRADING.md.

1. **Worksheet and Spreadsheet classes.** One constructor sets up all the state of a worksheet; the public methods move
   from each instance to the prototype. `instanceof` replaces the check of `parent` that tells a worksheet from a
   spreadsheet.
2. **History.** The undo history and the `ignoreHistory` flag move into `components/History.ts`. Applying a record
   (`undo` and `redo` in `utils/history.ts`) stays a worksheet method until the functions it calls are components.
3. **Clipboard.** Copy, cut, paste and the copied data move into `components/Clipboard.ts`.
4. **Editor.** `edition` and the editor types move into `components/Editor.ts`.
5. **Selection.** `selectedCell`, `highlighted`, the borders, the active cell and the focus state move into
   `components/Selection.ts`, together with the accessibility code that follows the selection.
6. **Formulas.** The formula chain, the results and the footer formulas move into `components/Formulas.ts`.
7. **The rest of the worksheet state**, one component at a time: merges, filters and search, pagination and lazy
   loading, frozen columns, columns and rows.
8. **Events.** `events.ts` splits by input under `events/`, and the handlers route through the spreadsheet.
9. **No import cycles.** The last cycles go, and `npm run lint` checks for new ones.
10. **The public typings.** The fields that only the library uses leave `WorksheetInstance`; the typings describe the
    classes.

## Rules for each step

-   A step moves code; it does not rewrite it. Fix a bug that the move finds in its own commit, with a test.
-   Keep the accessibility behaviour (`test/accessibility.js`) and the safe cell rendering described in AGENTS.md.
-   A breaking change is fine in version 6, but it must be on purpose: list it in UPGRADING.md with what an application
    does instead.
