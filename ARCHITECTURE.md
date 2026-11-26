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
   Done. Two groups of modules were left: `helpers ↔ internalHelpers`, and `internal`, `selection`, `footer`,
   `pagination`, `freezePosition` and `toolbar`. Three moves broke them: the value that a cell shows (`parseValue`,
   `executeFormula` and the functions they need) went to `utils/values.ts`, the fill with the selection corner
   (`copyData`) to `utils/fill.ts`, and the worksheet that a spreadsheet shows (`getWorksheetActive`,
   `getWorksheetInstance`) to `utils/activeWorksheet.ts`. None of these imports the module it left. `check-cycles.cjs`
   finds the cycles with the TypeScript parser, so the check needs no new package.
10. **The public typings.** The fields that only the library uses leave `WorksheetInstance`; the typings describe the
    classes.
    Done. `ads`, `cols`, `colgroupContainer`, `corner` and `headerContainer` left `WorksheetInstance`, and
    `timeControl`, `timeControlLoading` and `isMouseAction` left `JSpreadsheet`: they are in `WorksheetState` and
    `Library` (`src/types/internal.d.ts`). `WorksheetInstance` and `SpreadsheetInstance`
    say that they are the instances of the classes. Only the typings changed: the fields are still on the worksheet.

The steps above give every piece of state an owner and remove the cycles. The next steps clean up what they left: code
that is hard to read but does not need a new owner. They change no behaviour.

11. **Split `internal.ts`.** It held the functions that were left when the features moved out, and many modules needed it.
    Done. Each function went to the module of its feature: the creation and the rendering of the cells to
    `utils/render.ts`, the formula chain and the references to `utils/references.ts`, `updateTable` to `utils/table.ts`,
    `fullscreen` to `utils/fullscreen.ts`, `getCell`, `getCellFromCoords` and `getLabel` to `cells.ts`, `updateScroll`
    to `keys.ts`, `updateResult` to `pagination.ts`, `showIndex` and `hideIndex` to `rows.ts`, and
    `createNestedHeader` to `headers.ts`.
12. **The event handlers.** Each handler reads the active worksheet once, and the long handlers (`mouseDownControls`,
    `keyDownControls`) are split into a function for each target or key.
    Done. Each handler reads `jspreadsheet.current` once, after a press has made its worksheet the active one, and calls
    its functions with that worksheet as `this`. `mouseDownControls` calls a function for a press on a header, a row
    number, a cell and the pagination; `mouseUpControls` and `mouseMoveControls` one for the resize, the move and the
    fill; `keyDownControls` one for the keys of an open editor and one for the keys on the grid, which calls one for
    Delete, Enter, Tab, the shortcuts and the keys that start an edition. A timer still reads the active worksheet when
    it runs. The handlers read the global in 26 places instead of 361.
13. **The tabs.** `createTabs` builds the tabs in one long closure. The tabs become a class, with the methods on the
    prototype.
    Done. `Tabs` in `utils/tabs.ts` is a class: `new Tabs(element, options)` builds the tabs and keeps them on
    `element.tabs`. The variables of the closure are fields, and its inner functions are private methods.
14. **The anys in the public typings.** The `any` types left in `src/jspreadsheet.d.ts` get real types.
    Done. The values that the application or a formula gives (the meta information, the properties of a history record,
    the data of `onsave` and the result of `executeFormula`) are `unknown`. `onevent` gets `EventArguments`, the name of
    an event and the arguments of its handler, and `dispatch` takes only those. The toolbar and picker instances have
    the interfaces `ToolbarInstance` and `PickerInstance`, written from what jsuites puts on them; the typings of
    jsuites are not used, because they fail a strict check. A context menu item gets the element of the item. The
    library object has no index signature. The bundle is the same as before.

15. **The editor types.** `openEditor` opens each column type in its own branch of one long function. Each type gets its
    own function.
    Done. `openEditor` finds the column and starts the edition, then calls `openCustomEditor`, `toggleCheckbox`,
    `openDropdownEditor`, `openPickerEditor` (calendar and color), `openHtmlEditor`, `openImageEditor` or
    `openTextEditor`. `createEditor`, which puts the editor element in the cell, is a function of the module.
16. **The default toolbar.** `getDefault` in `utils/toolbar.ts` builds all the items of the default toolbar in one
    function of more than 400 lines. Each item, or group of items, gets its own function.
    Done. `getDefault` lists the items: `getHistoryItems` (undo, redo, save), `getFontItems` (the font, the size, the
    alignment, bold and the colors), `getVerticalAlignItem`, `getMergeItem`, `getBordersItem` and `getFullscreenItem`.
    The borders set the style in `setBorders` and add their controls in `addBorderControls`. Five items styled the
    selected cells with the same code, which is now `setSelectedStyle`, and the merge and the borders share
    `getSelectedRange`. The items and their order are the same.
17. **The long builders.** `createTable`, `insertColumn`, `insertRow`, `copy` and `paste` are each more than 200 lines.
    They are split into the steps they do.
    Done. `createTable` calls `createContent`, `createSearch`, `createHead` (the nested headers, the headers and the
    filters), `createGrid`, `createCorner`, `createAbout` and `createPagination`, puts them in order, and then calls
    `setLayout` and `loadContent` (the data, the styles, the comments and the classes). `insertColumn` gets the columns for
    the events from `getInsertedColumns` and adds the headers, the footers, the cells and the nested headers in a function
    each; `insertRow` uses `getInsertedRows` and `addRows`. `copy` reads the cells in `getCopiedRows` and the nested
    headers in `getNestedHeadersText`, and quotes a value in `quote`; `paste` calls `repeatToSelection`,
    `insertSpaceForPaste` and `pasteRows`. Code that had no effect went: a second table and a container that
    `createTable` created and did not use, and a loop in `copy` that worked out a value that nothing read. The functions
    are 41 to 121 lines long instead of 150 to 258.

The steps above leave the state in components, but the behaviour is still in `src/utils/`: about 8000 lines of
functions that get the worksheet as `this` through `fn.call(worksheet, ...)`, and that reach each other the same way. The
next steps move the behaviour of each feature into a component, so that a feature is one class: its state and the code
that changes it.

-   **A component gets its worksheet.** The worksheet constructor creates each component with `new Merges(this)`, and the
    component keeps it as `this.worksheet`. A moved function becomes a method that starts with
    `const obj = this.worksheet;`, so its code stays the same.
-   **Components call each other through the worksheet**, as `obj.merges.destroy()`, and import only the modules that
    stay functions. Only the builders (`utils/worksheets.ts` and `utils/factory.ts`) import `Worksheet.ts` at run
    time; the other modules tell a spreadsheet from a worksheet with `instanceof Spreadsheet`. So `Worksheet.ts` can
    import the components, and the components the modules they need, without a cycle.
-   **The public methods stay on the worksheet.** Each entry of `worksheetPublicMethods` calls the method of its
    component, for example `getMerge` calls `this.merges.get(cellName)`. The public API does not change.
-   **The components are internal.** The new components are not in `WorksheetInstance`, and the methods that move into
    the existing components are not in their public interfaces. Applications keep calling the worksheet methods.
-   **What stays functions.** Code without state of its own stays in `src/utils/`: the helpers for cell names and CSV
    (`helpers.ts`, `internalHelpers.ts`), `dispatch`, the rendering of a cell (`render.ts`), `updateTable`, the value
    that a cell shows (`values.ts`, without the formulas), the builders, `config`, `download`, `fullscreen`, the roles
    and labels of the grid (`accessibility.ts`, without the selection) and the worksheet that a spreadsheet shows.

18. **The plan and the groundwork.** This section, and the modules that told a worksheet from a spreadsheet with
    `instanceof Worksheet` (`dispatch`, `config`, `activeWorksheet`) use `instanceof Spreadsheet`.
    Done.
19. **The data of the cells.** `Styles` (`style.ts`), `Comments` (`comments.ts`), `Meta` (`meta.ts`) and `Cells`
    (`cells.ts`: finding a cell, its label, read-only).
    Done. The four modules are the components `worksheet.styles`, `worksheet.comments`, `worksheet.meta` and
    `worksheet.cells`. The methods have shorter names in their class (`setStyle` is `styles.set`, `getCellFromCoords` is
    `cells.getFromCoords`), and the public methods call them. The references rename the cells of the meta information
    with `obj.meta.update`.
20. **Merges.** `Merges` (`merges.ts`).
    Done. `worksheet.merges` gets, sets and removes the merged ranges (`get`, `set`, `remove`, `destroy`), tells which
    ranges a row or a column is in (`isRowMerged`, `isColMerged`), and updates the first cell of a range when its
    columns are hidden or shown (`updateColumns`).
21. **Formulas.** The references (`references.ts`), the footers (`footer.ts`) and the calculation of a formula
    (`executeFormula` and what only it uses in `values.ts`) move into `Formulas`.
    Done. `worksheet.formulas` calculates a formula (`execute`, with the private `parseNumber`), calculates the chain of
    a cell again (`updateChain`), renames the references (`updateFormula`, `updateTableReferences`) and shows the
    footers (`setFooter`). `parseValue` stays in `values.ts` and calls `obj.formulas.execute`.
22. **The view.** Search, the filters, pagination, the order and the lazy loading (`search.ts`, `filter.ts`,
    `pagination.ts`, `orderBy.ts`, `lazyLoading.ts`) move into `View`.
    Done. The functions of the five modules are methods of `worksheet.view`, with the same names, except
    `quantiyOfPages`, which is `countPages` in the class. The public methods keep their names.
23. **The editor.** `editor.ts` moves into `Editor`.
    Done. `worksheet.editor` opens the editor of the column type (`open`, with the private `openCustom`,
    `toggleCheckbox`, `openDropdown`, `openPicker`, `openHtml`, `openImage` and `openText`), closes it and saves the
    value (`close`), and toggles the highlighted checkboxes (`setCheckRadioValue`). `openEditor` and `closeEditor` call it.
24. **The clipboard.** `copyPaste.ts` moves into `Clipboard`.
    Done. `worksheet.clipboard` copies the cells (`copy`, with the private `getNestedHeadersText` and `getCopiedRows`)
    and pastes text (`paste`, with the private `insertSpaceForPaste` and `pasteRows`). `quote` and `repeatToSelection`
    use no worksheet and stay functions in `Clipboard.ts`. The public `copy`, `paste` and `download` call it.
25. **The selection.** `selection.ts`, the moves of `keys.ts` and the accessibility code that follows the selection move
    into `Selection`; the fill with the selection corner (`fill.ts`) into `Pointer`.
    Done. `worksheet.selection` selects cells (`updateFromCoords`, `update`, `reset`, `refresh`, `conditionalUpdate`,
    `selectAll`), places the corner (`updateCorner`), tells which cells are selected (`get`, `getSelected`,
    `getSelectedColumns`, `getSelectedRows`, `getRange`, `isSelected`, `getHighlighted`), moves the selection with the
    keys (`up`, `down`, `left`, `right`, `first`, `last`, and `rightGet` and `downGet`, which the paste uses), and keeps
    the selected cell the tab stop of the grid (`updateActiveCell`, `updateTabStop`, `restoreActiveCell`,
    `setFocusEvents`). The roles and the labels of the grid stay in `accessibility.ts`. `worksheet.pointer.fill` fills
    the cells after a drag or a double click of the corner (`copyData` before).
26. **Columns.** `Columns` (`columns.ts`, the headers in `headers.ts`, and the frozen columns in `freeze.ts` and
    `freezePosition.ts`).
    Done. `worksheet.columns` counts the columns (`count`), creates their headers (`createHeader`,
    `createNestedHeader`), inserts, moves, deletes, shows and hides them (`insert`, `move`, `delete`, `show`, `hide`),
    gets and sets their widths, titles and data (`getWidth`, `setWidth`, `getHeader`, `getHeaders`, `setHeader`,
    `getData`, `setData`), and places the frozen columns (`getIndexWidth`, `getFreezeWidth`, `updateFreezePosition`).
27. **Rows.** `Rows` (`rows.ts`).
    Done. `worksheet.rowManager` creates a row (`create`), inserts, moves, deletes, shows and hides the rows (`insert`,
    with the private `getInserted` and `add`, `move`, `delete`, `show`, `hide`), gets and sets their heights and data
    (`getHeight`, `setHeight`, `getData`, `setData`) and shows or hides the index column (`showIndex`, `hideIndex`). It
    is not `worksheet.rows`, because that is the list of row records, which applications use.
28. **The data.** `Data` (`data.ts`: getting and setting the data and the values).
    Done. `worksheet.data` gets and sets all the data (`get`, `set`), the data of a range (`getFromRange`) and the value
    of a cell (`getValue`, `getValueFromCoords`, `setValue`, `setValueFromCoords`). The data stays in `options.data`.
29. **History.** Applying a record to undo or redo it (`history.ts`) moves into `History`. It comes last because it calls
    every other feature.
    Done. `worksheet.history.undo` and `worksheet.history.redo` apply a record, with the private `processRow` and
    `processColumn` for the records that insert or delete rows and columns. `History` gets its worksheet in its
    constructor, as the other components do.
30. **The toolbar.** `toolbar.ts` becomes the class `Toolbar` of the spreadsheet, as `Tabs` did.
    Done. `spreadsheet.toolbarManager` builds the default items (`getDefault`), creates the jSuites toolbar (`create`,
    with the private `adjustForJSuites`), shows and hides it (`show`, `hide`) and updates its items for a worksheet
    (`update`). `spreadsheet.toolbar` stays the element, which applications use. `Spreadsheet.ts` imports no module at
    run time, so the factory creates the `Toolbar` right after the spreadsheet. The functions that build the items stay
    module functions in `toolbar.ts`.

After step 30 about 150 functions outside a class still get a worksheet or a spreadsheet as `this`: the rendering of a
cell, the roles and labels of the grid, the event handlers for each worksheet, `dispatch`, the build of a worksheet and
of a spreadsheet, and the public methods themselves, which `worksheetPublicMethods` and `spreadsheetPublicMethods` put
on the prototypes. The next steps move each of them into a class, so that no function outside a class reads `this`.

-   **A function that reads `this` belongs to a class.** It becomes a method of the component of its feature, or of
    `Worksheet` or `Spreadsheet` when it is about the whole worksheet or spreadsheet. A function that gets what it needs
    as arguments, such as the helpers for cell names, stays a function.
-   **The public methods are methods of the classes.** `Worksheet` and `Spreadsheet` declare their public methods in the
    class body, and each one calls its component or the code of the class. The two lists go. The public typings stay the
    source of truth: a method whose type does not match `jspreadsheet.d.ts` fails the type check, because the class
    merges with the public interface.
-   **The classes may import modules.** `Spreadsheet.ts` imports nothing at run time so that `dispatch`, `config` and
    `activeWorksheet` can tell a spreadsheet from a worksheet with `instanceof`. Once those are methods of the classes,
    they need no `instanceof`, and the restriction goes.
-   **What stays.** The document event handlers (`mousedown`, `keydown`, …) get the event, not a worksheet, and stay
    functions. `jspreadsheet.helpers.getCaretIndex` is a public helper that reads `this` only when an application calls it
    with a spreadsheet; it stays as it is.

31. **The cells.** Creating, rendering and updating a cell (`render.ts`) and the value that a cell shows (`values.ts`)
    move into `Cells`; `updateTable` (`table.ts`), which runs after the data changes, into `Data`.
    Done. `worksheet.cells` creates the td of a cell (`create`), renders it again when a merge changes the column whose
    options it uses (`render`), updates its content (`update`, with the private `getDropDownValue`), and gives the value
    that a cell shows (`parseValue`) and the column whose options it uses (`getColumnIndex`). `worksheet.data.updateTable`
    adds the spare rows and columns. `values.ts` keeps `isFormula` and `getMask`, which get what they need as arguments.
32. **Accessibility.** The roles, the labels and the size of the grid (`accessibility.ts`) become the component
    `Accessibility`.
    Done. `worksheet.accessibility` sets the roles of the grid and its rows (`setGridRoles`, `setRowRole`), the accessible
    names (`setGridLabels`, `setCellLabel`, `getLabelText`), the position of a row (`updateRowIndex`) and the size of the
    grid (`updateGridSize`). `getLabelText` reads the labels of its spreadsheet, which the code passed as `this` before.
    `setHeaderRole` gets what it needs as arguments and stays a function, exported by `components/Accessibility.ts`.
33. **The event handlers of a worksheet.** The mouse gestures (`events/mouse.ts`) move into `Pointer`, the keys of the grid
    and of the editor (`events/keyboard.ts`) into a new component `Keyboard`, and the scrolling (`events/scroll.ts`) into
    `View`.
34. **The spreadsheet.** `dispatch`, the active worksheet (`activeWorksheet.ts`), `fullscreen`, the configuration of a
    spreadsheet (`config.ts`), `setPlugins` and the build of a spreadsheet (`factory.ts`) become methods of `Spreadsheet`.
35. **The worksheet.** The build of a worksheet (`worksheets.ts`), its configuration, and creating, opening and deleting
    worksheets become methods of `Worksheet`; `download` moves into `Clipboard`.
36. **The public methods.** `Worksheet` and `Spreadsheet` declare their public methods, and `worksheetPublicMethods` and
    `spreadsheetPublicMethods` go.

## Rules for each step

-   A step moves code; it does not rewrite it. Fix a bug that the move finds in its own commit, with a test.
-   Keep the accessibility behaviour (`test/accessibility.js`) and the safe cell rendering described in AGENTS.md.
-   A breaking change is fine in version 6, but it must be on purpose: list it in UPGRADING.md with what an application
    does instead.
