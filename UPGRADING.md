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
