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
  events/             the document event handlers, split by input (keyboard, mouse, touch, clipboard, context menu, focus)
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
-   **The document events go to the active worksheet.** A press or the focus makes a worksheet the active one, and the
    handlers act on it. A key on an element outside the spreadsheet does not reach the grid.

## Steps

Each step is one pull request. Each keeps `npm test`, the type check and the accessibility tests passing, adds tests for
what it changes, and lists its breaking changes in UPGRADING.md.

1. **Worksheet and Spreadsheet classes.** One constructor sets up all the state of a worksheet; the public methods move
   from each instance to the prototype. `instanceof` replaces the check of `parent` that tells a worksheet from a
   spreadsheet.
2. **History.** The undo history and the `ignoreHistory` flag move into `components/History.ts`. Applying a record
   (`undo` and `redo` in `utils/history.ts`) stays a worksheet method until the functions it calls are components.
3. **Clipboard.** Copy, cut, paste and the copied data move into `components/Clipboard.ts`. The copied text, its styles,
   the textarea and the copy border have moved; `copy` and `paste` stay worksheet methods in `utils/copyPaste.ts` for
   the same reason as `undo` and `redo`.
4. **Editor.** `edition` and the editor types move into `components/Editor.ts`. The cell, its position and its HTML
   have moved; `openEditor`, `closeEditor` and the editor types stay in `utils/editor.ts` for the same reason as `undo`
   and `redo`.
5. **Selection.** `selectedCell`, `highlighted`, the borders, the active cell and the focus state move into
   `components/Selection.ts`, together with the accessibility code that follows the selection. The range, its
   bounds, the highlighted cells, the cells of the fill area, the active cell and the focus state have moved;
   `updateSelection`, `resetSelection` and the accessibility functions stay in `utils/` for the same reason as `undo`
   and `redo`.
6. **Formulas.** The formula chain, the results and the footer formulas move into `components/Formulas.ts`. The chain
   was the only formula state: `executeFormula` keeps the results of one calculation in local variables, and the footer
   formulas are calculated from `options.footers` each time. The chain has moved, with adding to it and renaming its
   cells; `executeFormula`, `updateFormulaChain` and the footers stay in `utils/` for the same reason as `undo` and
   `redo`.
7. **The rest of the worksheet state**, one component at a time: merges, filters and search, pagination and lazy
   loading, frozen columns, columns and rows.
   Search, filters and pagination had state of their own: the rows found, the filter values and the page. These moved
   into `components/View.ts`, which the lazy loading also reads. The other features have none to move: the merges and
   the frozen columns are in the options, and the columns and rows are the DOM of the worksheet (`headers`, `cols`,
   `rows`, `records`), which `Worksheet` keeps. The state of a mouse drag (`resizing`, `dragging`, `selectedHeader`,
   `selectedRow`, `selectedCorner`, `cursor`) goes with the events in step 8.
8. **Events.** `events.ts` splits by input under `events/`, and the active worksheet follows the focus.
   The split is done: `src/events/` has one module per input, and `updateFreezePosition`, which places the frozen
   columns, moved to `utils/freezePosition.ts`, so the modules that change columns no longer import the event handlers.
   That removed two import cycles.
   The second part keeps one active worksheet, `jspreadsheet.current`, and makes it follow the user. A drag goes on
   while the mouse leaves the worksheet, and a key reaches the document from wherever the focus is, so a handler cannot
   always find the worksheet from the event target; the active worksheet is what they share. Before, only a press set
   it. Now the focus moving into a worksheet sets it too (`events/focus.ts`), and a key on an element outside the
   spreadsheet stays out of the grid. The state of a mouse gesture moved into `components/Pointer.ts` (`resizing`,
   `dragging`, `cursor`, `corner`), and the header or row number that a press selected into `Selection` (`column`,
   `row`).
9. **No import cycles.** The last cycles go, and `npm run lint` checks for new ones.
10. **The public typings.** The fields that only the library uses leave `WorksheetInstance`; the typings describe the
    classes.

## Rules for each step

-   A step moves code; it does not rewrite it. Fix a bug that the move finds in its own commit, with a test.
-   Keep the accessibility behaviour (`test/accessibility.js`) and the safe cell rendering described in AGENTS.md.
-   A breaking change is fine in version 6, but it must be on purpose: list it in UPGRADING.md with what an application
    does instead.
