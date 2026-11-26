# Upgrading to version 6

Version 6 reorganises the code into classes that own their state (see [ARCHITECTURE.md](ARCHITECTURE.md)). This file
lists the changes that can break an application, and what to do instead. Each step of the reorganisation adds to it.

## Worksheets and spreadsheets are class instances

`jspreadsheet()` still returns the array of worksheets, and each worksheet still has `parent`, its spreadsheet. Both
are now instances of the `Worksheet` and `Spreadsheet` classes, and their methods are on the prototype instead of on
each instance.

-   **A method needs its worksheet as `this`.** The methods used to be bound to their worksheet, so a method taken off
    the worksheet still worked. Now it does not:

    ```js
    // Version 5
    const getData = worksheet.getData;
    getData();

    button.onclick = worksheet.undo;

    // Version 6
    const getData = () => worksheet.getData();

    button.onclick = () => worksheet.undo();
    ```

    The same holds for the methods of a spreadsheet, such as `getConfig` and `setConfig`.

-   **The methods are not own properties.** `Object.keys(worksheet)`, `{ ...worksheet }` and `Object.assign({}, worksheet)`
    no longer include the methods. `worksheet.getData` still finds them, and assigning a method to one worksheet still
    replaces it for that worksheet only.

-   **A worksheet has its methods from the start.** Version 5 added them while it built the worksheet. Calling a method
    before the `onload` event is still not supported: the worksheet has no table yet.

## The undo history is an object

`worksheet.history` used to be the array of records, with the position in `worksheet.historyIndex` and the flag that
stops new records in `worksheet.ignoreHistory`. It is now an object that holds all three:

| Version 5                                               | Version 6                     |
| ------------------------------------------------------- | ----------------------------- |
| `worksheet.history`                                     | `worksheet.history.records`   |
| `worksheet.historyIndex`                                | `worksheet.history.index`     |
| `worksheet.ignoreHistory`                               | `worksheet.history.ignore`    |
| `worksheet.historyIndex >= 0`                           | `worksheet.history.canUndo()` |
| `worksheet.historyIndex < worksheet.history.length - 1` | `worksheet.history.canRedo()` |
| `worksheet.history = []; worksheet.historyIndex = -1;`  | `worksheet.history.clear()`   |

`worksheet.history.add(record)` keeps a record, as the library does for each change. `worksheet.undo()` and
`worksheet.redo()` work as before, and the `onundo` and `onredo` events still get the record.

## What a worksheet copied is in its clipboard

The text that a worksheet copied last, the styles of the copied cells and the hidden textarea were fields of the
worksheet. They are now in `worksheet.clipboard`:

| Version 5              | Version 6                      |
| ---------------------- | ------------------------------ |
| `worksheet.data`       | `worksheet.clipboard.data`     |
| `worksheet.style`      | `worksheet.clipboard.style`    |
| `worksheet.hashString` | `worksheet.clipboard.hash`     |
| `worksheet.textarea`   | `worksheet.clipboard.textarea` |

`worksheet.clipboard.data` is an empty string before the first copy, where `worksheet.data` was undefined. A cancelled
copy and `download()` no longer change it. `worksheet.copy()` and `worksheet.paste()` work as before.

## The cell editor is an object

`worksheet.edition` was `null`, or an array with the cell that is being edited, its HTML before the editor opened, and
its column and row as text. It is now `worksheet.editor`:

| Version 5                        | Version 6                   |
| -------------------------------- | --------------------------- |
| `worksheet.edition` (to test it) | `worksheet.editor.isOpen()` |
| `worksheet.edition[0]`           | `worksheet.editor.cell`     |
| `worksheet.edition[1]`           | `worksheet.editor.html`     |
| `parseInt(worksheet.edition[2])` | `worksheet.editor.x`        |
| `parseInt(worksheet.edition[3])` | `worksheet.editor.y`        |

`x` and `y` are numbers. `worksheet.editor.cell` is `null` when no editor is open, where `worksheet.edition` was `null`
or, before the first editor, undefined. `openEditor()`, `closeEditor()`, the editor events and the custom editors work as
before.

## The selection is an object

The fields of the selection are now in `worksheet.selection`, a `Selection` object. The array that `worksheet.selection`
held, the cells of the area that the corner fills, is `worksheet.selection.fill`.

| Version 5                     | Version 6                         |
| ----------------------------- | --------------------------------- |
| `worksheet.selectedCell`      | `worksheet.selection.range`       |
| `worksheet.selectedContainer` | `worksheet.selection.bounds`      |
| `worksheet.highlighted`       | `worksheet.selection.highlighted` |
| `worksheet.selection`         | `worksheet.selection.fill`        |
| `worksheet.activeCell`        | `worksheet.selection.activeCell`  |
| `worksheet.focusInGrid`       | `worksheet.selection.focusInGrid` |

`worksheet.selection.range` is `null` before a cell is selected, where `worksheet.selectedCell` was undefined, and
`worksheet.selection.highlighted` is an empty array before the first selection. `getSelection()`, `getSelected()`,
`updateSelectionFromCoords()`, `resetSelection()` and the selection events work as before.

## The formula chain is an object

`worksheet.formula`, the list of the cells whose formulas refer to each cell, is now `worksheet.formulas.chain`. It is
an object, where `worksheet.formula` was an array used as an object, so it has no `length` and no array methods. An
entry that a reference loop ended is an empty array, where it was an empty text. `worksheet.formulas.add(cell,
dependent)` adds to the chain, and `worksheet.formulas.rename(references)` renames its cells. `executeFormula()` and the
calculation of formulas work as before.

## The rows that a worksheet shows are in its view

The state of the search, the filters and the pagination is now in `worksheet.view`, a `View` object:

| Version 5              | Version 6                   |
| ---------------------- | --------------------------- |
| `worksheet.results`    | `worksheet.view.results`    |
| `worksheet.filters`    | `worksheet.view.filters`    |
| `worksheet.pageNumber` | `worksheet.view.pageNumber` |

`worksheet.view.pageNumber` is `0` from the start, also without pagination, where `worksheet.pageNumber` was undefined
until the worksheet had pagination. `worksheet.view.results` is `null` from the start. `search()`, `resetSearch()`,
`openFilter()`, `resetFilters()`, `page()` and the `onupdateresult` and `onchangepage` events work as before.
The search input, the filter row and the pagination controls stay on the worksheet (`searchInput`, `filter`,
`pagination`).

## The state of a mouse gesture is in its pointer

The state of a press on a header, a row number or the selection corner is now in `worksheet.pointer`, a `Pointer`
object, and in `worksheet.selection`:

| Version 5                  | Version 6                    |
| -------------------------- | ---------------------------- |
| `worksheet.resizing`       | `worksheet.pointer.resizing` |
| `worksheet.dragging`       | `worksheet.pointer.dragging` |
| `worksheet.cursor`         | `worksheet.pointer.cursor`   |
| `worksheet.selectedCorner` | `worksheet.pointer.corner`   |
| `worksheet.selectedHeader` | `worksheet.selection.column` |
| `worksheet.selectedRow`    | `worksheet.selection.row`    |

Each is `null`, or `false` for `corner`, before the first press, where the fields of version 5 were undefined. The
resize, the move of columns and rows, the fill with the selection corner and their events work as before.

## The keys go to the grid that has the focus

The keys go to the active worksheet, `jspreadsheet.current`. A press on a worksheet made it the active worksheet, and
now the focus does too: when the focus moves into another worksheet, for example with Tab, that worksheet becomes the
active worksheet. The worksheet that was active closes its editor and loses its selection, as after a press on the
other worksheet, and a focused cell becomes the selection of a worksheet without one. The focus on the tabs or the
toolbar of a spreadsheet changes nothing.

The keys on an element outside the spreadsheet, such as an input or a button next to the grid, no longer go to the
grid when no editor is open. Before, only the keys on a `jss_object` control stayed out of the grid, and typing in an
input after Tab out of the grid opened the editor of the selected cell. The keys with the focus on the body still go
to the grid.

## The typings declare only the public fields

The TypeScript typings no longer declare the fields that only the library uses. The fields are still on the worksheet,
so JavaScript code works as before, but TypeScript code that reads them no longer compiles:

| Field                             | Instead                                   |
| --------------------------------- | ----------------------------------------- |
| `worksheet.cols`                  | `getWidth()` and `setWidth()`             |
| `worksheet.colgroupContainer`     | `getWidth()` and `setWidth()`             |
| `worksheet.corner`                | none: the corner is part of the selection |
| `worksheet.headerContainer`       | `worksheet.headers`, the header cells     |
| `worksheet.ads`                   | none                                      |
| `jspreadsheet.timeControl`        | none                                      |
| `jspreadsheet.timeControlLoading` | none                                      |
| `jspreadsheet.isMouseAction`      | none                                      |

Code that still needs one of these fields can cast the worksheet, for example `(worksheet as any).corner`, but the
fields can change in any version. `WorksheetInstance` and `SpreadsheetInstance` are the types of the `Worksheet` and
`Spreadsheet` instances.

## The methods of the tabs are on the prototype

`spreadsheet.element.tabs`, the tabs of the worksheets, is now an instance of a class, and its methods (`open`,
`create`, `remove`, `getActive` and the others) are on the prototype, as for the worksheets. A method taken off the
object needs the object as `this`: use `() => tabs.open(1)` instead of `const open = tabs.open`. `Object.keys(tabs)` no
longer includes the methods. Assigning a method to the object still replaces it, also for the calls that the tabs make.

## The typings have no any

The TypeScript typings no longer use `any`. JavaScript code works as before, but TypeScript code can need a check or a
cast where it relied on `any`:

| Typing                                          | Version 5                         | Version 6                                               |
| ----------------------------------------------- | --------------------------------- | ------------------------------------------------------- |
| `MetaInformation`                               | `Record<string, any>`             | `Record<string, unknown>`                               |
| The other properties of `HistoryRecord`         | `any`                             | `unknown`                                               |
| The data of `onsave`                            | `any`                             | `unknown`: the rows, or what `onbeforesave` returned    |
| `executeFormula()`                              | `any`                             | `unknown`                                               |
| `dispatch()`                                    | any event name, returns any       | the events of `SpreadsheetOptions` with their arguments |
| `onevent` of the options and of plugins         | `(event: string, ...rest: any[])` | `(...args: EventArguments)`                             |
| The first argument of a context menu `onclick`  | `any`                             | `HTMLDivElement`, the element of the item               |
| The toolbar and picker instances of the toolbar | `Record<string, any>`             | `ToolbarInstance` and `PickerInstance`                  |
| The other properties of a toolbar item          | `any`                             | `unknown`                                               |
| Other properties of `jspreadsheet`              | `any`                             | not declared                                            |

-   **onevent.** `EventArguments` is the name of an event, then the arguments of its handler. A test of `args[0]`
    narrows the other arguments:

    ```ts
    onevent: function (...args) {
        if (args[0] === 'onchange') {
            const worksheet = args[1]; // WorksheetInstance
        }
    },
    ```

    A handler declared as `(event: string, ...rest: unknown[])` still fits. A handler that declares a worksheet as its
    second parameter no longer fits, because some events, such as `onbeforesave`, give the spreadsheet there.

-   **dispatch.** It takes the name of an event of `SpreadsheetOptions` and the arguments of its handler. To dispatch an
    event of your own, cast the name, for example `worksheet.dispatch('oncustom' as jspreadsheet.EventName, worksheet)`.

## The boolean options must be booleans

The library compares the boolean options with `true` and `false` strictly (`===`). Version 5 compared them loosely, so `1`
and `'1'` counted as `true`, and `0`, `''` and `'0'` as `false`. Give these options `true` or `false`, as the typings
declare. This applies to the options of a spreadsheet, a worksheet and a column, such as `editable`, `search`,
`filters`, `lazyLoading`, `tableOverflow`, `wordWrap`, `readOnly`, `fullscreen`, `allowInsertRow`, `allowInsertColumn`,
`allowDeleteRow`, `allowDeleteColumn`, `selectionCopy`, `columnSorting`, `columnDrag`, `columnResize`, `rowDrag`,
`rowResize`, `secureFormulas`, `parseFormulas`, `autoCasting`, `autoIncrement`, `allowExport`, `allowComments`,
`csvHeaders`, `persistence` and `about`.

```js
// Version 5: 0 disabled editing
jspreadsheet(element, { worksheets: [{ editable: 0 }] });
// Version 6
jspreadsheet(element, { worksheets: [{ editable: false }] });
```

The values of the cells are still compared loosely where the library reads them: `1`, `'1'`, `true` and `'true'` check a
checkbox, and a text that is a number equals the number.

## getValue returns null for a cell that is missing

`getValue` and `getValueFromCoords` return `null` for a cell outside the data, also when the row exists and the column
does not. Version 5 returned `undefined` for a column that the row did not have, and `null` for a cell that held the
text `undefined` (#36, #49). An application that checked the result with `=== undefined` checks it with `== null`.

## Any handler can cancel an action

An event calls the `onevent` option, then the handler of the event in the options, then the `onevent` of each plugin.
Version 5 returned the result of the last one, so a plugin whose `onevent` returned nothing made an `onbefore…` handler
that returned `false` cancel nothing (#37). The result of an event is now `false` when any of them returns `false`, and
otherwise the last result that is not `undefined`. A plugin's `onevent` that returns `false` cancels the action too.

## A read-only cell gives no change

`setValue`, `setValueFromCoords`, a paste and a fill leave a read-only cell as it is. Version 5 still put a record for
the cell in the history and in `onafterchanges`, with the coordinates and without a value (#58). Now a cell that did
not change has no record: `onafterchanges` has only the cells that changed, and when no cell changed it does not occur
and the history gets no entry.

## onbeforechange can cancel a change

Returning `false` from `onbeforechange` cancels the change of the cell, as `false` cancels the other `onbefore…` events.
Version 5 wrote `false` into the cell instead (#38). Another value that the handler returns is still written instead of
the new value. A handler that wants the cell to hold `false` returns nothing when the new value is `false`; it cannot
replace another value with `false`.

## The change events get the column and row as numbers

`onbeforechange`, `onchange`, the records of `onafterchanges` and the records of the history get the column and row of
the cell as numbers. Version 5 gave them as text after a typed edit, a click on a checkbox, the Delete key, Ctrl+X and
`setValue` with a cell element, and as numbers after the other changes (#52). An application that compared them with
text, such as `x === '2'`, compares them with numbers. In the typings, `CellChange.x` and `CellChange.y` are numbers,
and `CellChange.value` can be `undefined`, as the new value of `onchange` can.

## The persistence events occur for an undo and a redo

With `persistence`, an undo or a redo posts the cells it changed. Version 5 posted them without `onbeforesave` and
`onsave`, because undo and redo set `ignoreEvents` (#57). `onbeforesave` and `onsave` now occur while `ignoreEvents` is
set, so `onbeforesave` can change or cancel every post.

## Loading the comments option is not a change

The `comments` option of a worksheet is loaded without an entry in the undo history and without `oncomments`, as the
`style` option is. Version 5 added a `setComments` entry while it built the worksheet, so an undo right after loading
removed the comments, and it called `oncomments` for them (#87).

## A footer formula runs without coordinates

A formula in a footer runs as `executeFormula` without coordinates: `onbeforeformula` and `formulaEngine` get `undefined`
as `x` and `y`. Version 5 gave them the position of a data cell, the number of rows plus the column, so a footer formula
that referred to that cell showed `#ERROR`, and the footer was in the formula chain under the name of that cell (#18).
A footer shows the result of its formula, or its text, without the mask of a column.

## setStyle trims the values

`setStyle` with an object and the `style` option trim the value of each declaration, as they trim the property, so
`color: red` sets `red`. Version 5 set `' red'`: setting the same style again did not toggle it off (#67). The history
and `onchangestyle` get the trimmed value, for example `color:red` for `color: red`.

## Tab moves the focus out of the grid at the end of a row

Tab on the last column and Shift+Tab on the first column keep their default action, so the focus moves to the control
after or before the grid, as the WAI-ARIA grid pattern asks (#28). Version 5 kept the focus in the grid, and Tab on the
last column inserted a column. The option `allowManualInsertColumn` has no effect any more. To add a column with the
keyboard, an application handles a key itself and calls `insertColumn`.
