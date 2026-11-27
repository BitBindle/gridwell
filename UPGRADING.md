# Upgrading to version 6

Version 6 reorganises the code into classes that own their state (see [ARCHITECTURE.md](ARCHITECTURE.md)). This file
lists the changes that can break an application, and what to do instead. Each step of the reorganisation adds to it.

## The package is @bitbindle/gridwell

The fork is called Gridwell, and its package is `@bitbindle/gridwell`, not `jspreadsheet-ce`. Install it and change the
imports; the API keeps its names:

```js
// Version 5
import jspreadsheet from 'jspreadsheet-ce';
import 'jspreadsheet-ce/dist/jspreadsheet.css';

// Version 6
import jspreadsheet from '@bitbindle/gridwell';
import '@bitbindle/gridwell/dist/jspreadsheet.css';
```

- The script of the bundle sets the global `gridwell` as well as `jspreadsheet`.
- There is no Vue wrapper. `@jspreadsheet-ce/vue` was a component that called `jspreadsheet()` in `mounted`: call
  `gridwell()` on an element of your component instead. The repository no longer has the `<jspreadsheet>` web
  component example (`src/webcomponent.js`) or the documentation of version 4 (`docs/`).
- "About" in the context menu shows `Gridwell`, and the link below the table (`worksheet.ads`) goes to the repository of
  the fork.
- The package has only `dist/` and `resources/lang/`, with the README and the license.

## Worksheets and spreadsheets are class instances

`jspreadsheet()` still returns the array of worksheets, and each worksheet still has `parent`, its spreadsheet. Both
are now instances of the `Worksheet` and `Spreadsheet` classes, and their methods are on the prototype instead of on
each instance.

- **A method needs its worksheet as `this`.** The methods used to be bound to their worksheet, so a method taken off
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

- **The methods are not own properties.** `Object.keys(worksheet)`, `{ ...worksheet }` and `Object.assign({}, worksheet)`
  no longer include the methods. `worksheet.getData` still finds them, and assigning a method to one worksheet still
  replaces it for that worksheet only.

- **A worksheet has its methods from the start.** Version 5 added them while it built the worksheet. Calling a method
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

## Insert, delete and move keep the selection and the editor on their cells

After `insertRow`, `insertColumn`, `deleteRow`, `deleteColumn`, `moveRow`, `moveColumn`, and their undo and redo, the
selection and an open editor stay on their cells. After a move, each end of a larger selection stays on its cell, and a
move inside the selection keeps its rows and columns, so a selection of whole rows or columns stays whole. The active
cell can then be another cell. An insert after the last column or row adds the new columns or rows to a selection of
whole rows or columns. Version 5 kept the row and
column numbers, so the selection and the editor moved to other cells, and a value typed into a deleted row went into the
row that moved up.

- `getSelection()` gives the new position of the selected cells.
- After an insert at the end, a selection of whole rows or columns includes the new columns or rows.
- A delete removes the selection if a selected cell is deleted, also when the selection is inside the deleted rows or
  columns. Version 5 kept such a selection.
- An editor in a deleted row or column closes without a save: `oneditionend` gets `false` for the save, and no
  `onchange` fires. Version 5 left it open.

No option keeps the old behavior, because it put values into the wrong cells.

## Shift, Ctrl and Delete use only a current selection of whole rows or columns

This changes the result for users who press Shift or Ctrl on row numbers or headers, or press Delete, and for
applications that read `worksheet.selection.row` or `worksheet.selection.column`.

A press on a row number or a header starts a selection of whole rows or columns. Version 5 remembered that row or
column after a selection of other cells with the keyboard, `updateSelectionFromCoords` or `selectAll`.

- Delete deletes rows or columns only while the selection is of whole rows or columns that a press started. Version 5
  deleted the row of a single cell after a press on its row number and an arrow key.
- A Shift or Ctrl press on a row number or a header, after a selection that does not start at a pressed row number or
  header, extends that selection from its first row or column. Version 5 extended from the row or column of an earlier
  press, or selected only the pressed row or column. On macOS, a Ctrl press is the usual secondary press.
- A drag over the row numbers or headers extends whole rows or columns only from a current selection of whole rows or
  columns.
- `worksheet.selection.row` and `worksheet.selection.column` can still keep an earlier press. The new methods
  `rowAnchor()` and `columnAnchor()` give them only while they count.

No option keeps the old behavior, because it deleted rows that the user did not select.

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
| `onevent` of the options and of plugins         | `(event: string, ...rest: any[])` | `EventListener`: one signature for each event           |
| The first argument of a context menu `onclick`  | `any`                             | `HTMLDivElement`, the element of the item               |
| The toolbar and picker instances of the toolbar | `Record<string, any>`             | `ToolbarInstance` and `PickerInstance`                  |
| The other properties of a toolbar item          | `any`                             | `unknown`                                               |
| Other properties of `jspreadsheet`              | `any`                             | not declared                                            |

- **onevent.** `EventListener` has one signature for each event: the name of the event, then the arguments of its
  handler. A handler can name fewer parameters than an event gives (#133). The parameters have the types of all the
  events together: `event` is an `EventName`, and the second parameter is a `WorksheetInstance` or, for events such as
  `onbeforesave`, the `SpreadsheetInstance`:

    ```ts
    onevent: function (event, instance) {
        console.log(event, instance);
    },
    ```

    A parameter that the shortest events do not give, such as the third, must be optional:
    `(event, instance, cell?) => …`. To narrow the arguments with a test of the name, take them as `EventArguments`:

    ```ts
    onevent: function (...args: jspreadsheet.EventArguments) {
        if (args[0] === 'onchange') {
            const [, worksheet, cell, x, y, value] = args; // WorksheetInstance, HTMLTableCellElement, number, …
        }
    },
    ```

    Without the type, `(...args) => …` gets no narrowing. A handler declared as `(event: string, ...rest: unknown[])`
    still fits. A handler that declares a worksheet as its second parameter does not fit, because some events give the
    spreadsheet there.

- **dispatch.** It takes the name of an event of `SpreadsheetOptions` and the arguments of its handler. To dispatch an
  event of your own, cast the name, for example `worksheet.dispatch('oncustom' as jspreadsheet.EventName, worksheet)`.

## setConfig takes the options to change

The typings of `setConfig` take only the options to change. On a worksheet they are the options of a worksheet, or with
`spreadsheetLevel` set to `true` the options of the spreadsheet. A TypeScript call that gives a worksheet an option of
the spreadsheet without `spreadsheetLevel` no longer compiles, because the worksheet would keep the option without using
it.

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

## The clipboard textarea has no id

The hidden textarea that a worksheet uses for the clipboard has no `id` any more. Version 5 gave the textarea of every
worksheet the id `jss_textarea`, so a page with two worksheets had duplicate ids (#68). Select it by its class,
`.jss_textarea`, or through `worksheet.clipboard.textarea`.

## jspreadsheet throws for a table without a parent

`jspreadsheet(table, options)` replaces the table with a div in its parent. When the table has no parent, for example
before it is added to the document, it now throws an `Error` (#55). Version 5 failed in a promise: it returned an empty
array, created no spreadsheet and gave an unhandled rejection. Add the table to an element before you call
`jspreadsheet`.

## A formula reads only formatted numbers as numbers

A formula reads the text of a referenced cell as a number when the text is a number with symbols around it, such as
`$ 1,234.50`, `1 234,5 €` (with `decimal: ','` on the column) or `50%` (#129). Version 5 kept every digit of any text,
so `Room 12b` was `12` and the date `01/02/2024` was `1022024`. Such texts are now texts, as in Excel. A formatted zero,
such as `$ 0`, is now the number `0` instead of its text. Set `autoCasting: false` to read every text as text.

## The application registers the formula engine

`@jspreadsheet/formula` is no longer a dependency of `jspreadsheet-ce`, and the bundle no longer looks for it in the
global `formula` or with `require`. It is an optional peer dependency. An application that uses formulas installs it and
registers it once, before it creates the spreadsheets:

```javascript
import formula from '@jspreadsheet/formula';

jspreadsheet.setFormulaEngine(formula);
```

With the scripts of a CDN, call `jspreadsheet.setFormulaEngine(formula)` after the script of the engine. Without an
engine, a formula shows `#ENGINE`, which Excel never shows, and the console says once why. `executeFormula` also returns
`'#ENGINE'`. The `formulaEngine` option of a spreadsheet still replaces the engine, and needs none to be registered. A
spreadsheet without formulas no longer needs the 140 KB of the engine.

## Translate with jspreadsheet.setDictionary

Jspreadsheet keeps its own dictionary and no longer reads the dictionary of jSuites. An application that called
`jSuites.setDictionary`, or set `document.dictionary`, to translate the texts of the spreadsheet calls
`jspreadsheet.setDictionary` instead. It takes the same object. An empty translation shows the English text again.

## The requests are sent with fetch

The worksheet loads the `csv` and `url` options and the `url` of a dropdown column, and posts the changes of the
`persistence` option, with `fetch` in place of `jSuites.ajax`. The requests have the same method, body and
`X-Requested-With` header, but a GET request no longer has the header `Content-Type: text/json`. The `beforeSend`
handlers of `jSuites.ajax` and `document.ajax` no longer see these requests: to add a header, such as a token, use the
new `onbeforerequest` event, which gets the worksheet, the url and the options of `fetch`. It can change the options,
return others, or return `false` to cancel the request.

```javascript
jspreadsheet(element, {
    worksheets: [{ url: '/data' }],
    onbeforerequest(worksheet, url, options) {
        options.headers.Authorization = `Bearer ${token}`;
    },
});
```

A request that fails, or that the server answers with a status that is not 2xx, is logged with `console.error`. A
worksheet whose `csv` or `url` request fails is now built with its other options, such as `data` and `minDimensions`;
before, it was never built. A dropdown column whose `url` fails is now created without its source, where a network error
left the worksheet without its table before.

## The masks show numbers and dates without jSuites

The masks of the columns (`mask`, `format`, and `locale` with `options`) and the formats of the calendar columns are read
by Jspreadsheet, not by `jSuites.mask` and `jSuites.calendar`. The number formats follow Excel, and some cells show
other texts than before:

- A column with a `locale` shows what `Intl.NumberFormat` gives with its `options`: `$3.00` for 3 in US dollars, where
  jSuites showed `$3` (issue #136). Its `decimal` option no longer replaces the `options` of the locale.
- The second section of a mask, for the negative numbers, shows the number without its minus, as in Excel: `0;[Red]0`
  shows 5 in red for -5, and `0;[Red](0)` shows `(5)`. The parentheses are those of the section; Jspreadsheet no longer
  adds them when the section has one. A color, such as `[Red]`, now applies to the section it is in, also the first.
- Without a section for the negative numbers, the minus goes before the symbols: `-$ 5.00` for -5 with `$ #,##0.00`.
- A text that is not a number shows as it is, where jSuites showed an empty text or only its digits (`12` for `a1b2`).
- A number in scientific notation shows as in Excel, such as `1.23E+03` for `0.00E+00`, where jSuites showed `1.23e+3`.

The editor of a cell with a mask shows the number without the symbols of the mask, such as `1234.5` for `$ 1,234.50`,
with the decimal separator of the mask, and a percentage with its percent sign, such as `12.5%`. The editor no longer
masks the text as it is typed, and the column option `disabledMaskOnEdition` has no effect. A number that is typed with
or without the symbols of the mask is saved as a number, and a text that is not a number is saved as it is typed, where
jSuites saved an empty value or the digits of the text.

The dates of the calendar columns, and of the date masks, are written and read with the same tokens. These change:

- `HH12` shows midnight as `12:00 AM`, where jSuites showed `00:00 AM`.
- `YY` reads 00 to 29 as 2000 to 2029 and 30 to 99 as 1930 to 1999, as Excel does, where jSuites read all of them in
  this century.
- A text that is not a date in the format of the column shows as it is, and a pasted text that is not a date in the
  format empties the cell, as before. A pasted date in the ISO format, such as `2024-03-05`, is now kept, where jSuites
  read a wrong date from it in a column with another format.
- `jSuites.calendar.extractDateFromString` and `jSuites.calendar.getDateString` are no longer called, so a change to them
  does not change the cells.

## The context menu works without jSuites

The context menu is shown by Jspreadsheet, not by `jSuites.contextmenu`. It keeps the class names of jSuites
(`jcontextmenu`, `jcontextmenu-focus`, `jcontextmenu-disabled`, `jcontexthassubmenu`), its place and its look, and the
element still keeps the menu as `contextmenu`, with `open(e, items)`, `create(items)`, `close()` and `isOpened()`. These change:

- The title of an item is shown as text. jSuites wrote it with `innerHTML`, so a title with HTML now shows the tags.
- The menu has the roles of an ARIA menu (`menu`, `menuitem`, `separator`), and a disabled item has `aria-disabled`. When
  the menu opens, the focus moves to its first item; when it closes after an item is chosen, or with Escape, the focus
  goes back to the element that had it.
- The keys move in the menu: the arrows, Home and End move between the items, Enter and Space choose an item, the right
  arrow opens a submenu and the left arrow closes it, Escape closes the menu and Tab closes it and moves the focus on.
  These keys no longer go to the grid while the menu has the focus.
- A menu opened with the keyboard (the menu key or Shift+F10) opens under the selected cell, not in the corner of the
  window.
- The release of the right button no longer chooses the item under the pointer.
- The `onclick` of an item gets the release of the mouse, or the `KeyboardEvent` of the key that chose it. Its type is
  `(element: HTMLDivElement, e: MouseEvent | KeyboardEvent) => void`. The menu closes before `onclick` runs.
- The menu has no `options.onclick`, which jSuites called after each chosen item.

## The toolbar and the color picker work without jSuites

The toolbar, the pickers of its select items and the color picker are shown by Jspreadsheet, not by `jSuites.toolbar`,
`jSuites.picker` and `jSuites.color`. They keep the class names of jSuites (`jtoolbar`, `jtoolbar-item`, `jpicker`,
`jpicker-focus`, `jcolor`, `jcolor-focus`, …) and their look, and the elements still keep them as `toolbar`, `picker`
and `color`. These change:

- The toolbar has the methods `create`, `update`, `selectItem`, `setReadonly`, `get`, `show`, `hide` and `destroy`. It
  no longer moves the items that do not fit into a menu, so `refresh`, `open` and `close` are gone, and so are
  `setBadge` and the options `responsive`, `badge`, `title`, `app` and `container`. `hide` and `show` no longer slide
  the toolbar.
- The content of a `label` item, and the icon name of an item, are shown as text. jSuites wrote them with `innerHTML`.
- An option of a select item without `render` is shown as text. `render` still gives HTML, or an element.
- The typings of `render` of a select item are corrected: its second argument is the element of the option in the list, or `null` for the option that the picker shows, not the picker, and it can give an element. A `render` that declares the picker as its second argument fails the type check.
- The options of a picker have `aria-selected`, and the picker has `isOpened()`. The pickers no longer have the
  options `header`, `grid`, `height`, `onmouseover` and `onselect`, nor `setOptions`, `getLabel` and `setLabel`, nor
  the `val` and `change` properties on their element.
- The color picker has a grid of colors with the ARIA grid pattern (the arrows, Home and End move between the colors,
  and Enter and Space choose one), an input for any other color, and Reset and Done buttons. The Spectrum and Sliders
  tabs are replaced by that input, which opens the color dialog of the browser. The picker no longer fills the bottom
  of the screen on a small screen, and has no `fullscreen` option, no `setOptions` and `toHex` methods, and no `val`
  and `change` properties on its element.
- When the color picker opens, the focus moves to the chosen color, or to the first; when it closes, the focus goes
  back to its element. The focus leaving the picker closes it.
- The color items of the toolbar open their picker with Enter and Space each time; jSuites opened it from the keyboard
  only the first time.
- In the editor of a color column, Escape on the picker closes the editor without saving, as on the input. The options
  of the column (`palette`, `placeholder`, `closeOnChange`, `doneLabel`, `resetLabel`, `onopen`, `onchange`) are still
  given to the picker. The texts of the buttons are translated with `jspreadsheet.setDictionary` (`Reset`, `Done`).

## The dropdown works without jSuites

The editor of a dropdown column and the filters of the columns are shown by Jspreadsheet, not by `jSuites.dropdown`.
They keep the class names of jSuites (`jdropdown`, `jdropdown-header`, `jdropdown-item`, `jdropdown-selected`,
`jdropdown-cursor`, …) and the look of its default mode, and the element still keeps the dropdown as `dropdown`, with
`getValue(asArray)`, `getText()`, `setValue(value)`, `open()`, `close(ignoreEvents)`, `isOpened()`, `reset()` and
`find(text)`. These change:

- The dropdown has the roles of an ARIA combobox: the header is a `combobox` with `aria-expanded` and
  `aria-activedescendant`, the list is a `listbox` (with `aria-multiselectable` for a column with `multiple`), the
  items are `option`s with `aria-selected`, and a group is a `group` named by its title. The header of a filter is named
  "Filter column" and the title of the column, with the new label `labels.filterColumn`.
- The keys: the arrows move between the items, Page Up and Page Down by seven, Home and End to the first and the last
  (not with `autocomplete`, where they move in the text), and a letter to the next item that starts with it (not with
  `autocomplete`, where it types). Enter chooses the item, or with `multiple` adds it or removes it; Space does the same
  with `multiple`. In the editor, Enter on a single dropdown saves the value and keeps the cell selected, and Tab saves it
  and selects the next cell. Escape closes the editor without saving, and closes a filter without changing it; jSuites
  applied the filter.
- Enter on the chosen item keeps it chosen; jSuites cleared it. A press on the chosen item still clears it.
- A value that is not the id of an item is shown as itself in the header, and kept.
- The items have `id`, `name`, `title`, `image` and `group`. `icon`, `color`, `tooltip` and `disabled` are ignored, a
  press on the title of a group no longer chooses its items, and the groups can no longer be collapsed.
- With `autocomplete`, the items are filtered as you type, without a delay, by the text that their name, title or group
  contains; jSuites read the text as a regular expression.
- The `options.type` of a dropdown column (`picker`, `searchbar`, `list`) is ignored: the list opens under the cell, or
  over it when there is no room under it in the window or in the element that scrolls (a list that fits on neither side
  opens on the side with more room, and scrolls), also in a window narrower than 800 pixels, where jSuites showed a picker at the bottom of the screen or a search
  bar over the whole screen.
- The dropdown has none of the other methods of jSuites (`setData`, `setUrl`, `add`, `getPosition`, `setId`, …), no
  sorting, no lazy loading and no `newOptions`.
- A key whose element it removed, such as Tab on the editor of a dropdown, goes to the grid.

## The calendar works without jSuites

The editor of a calendar column is shown by Jspreadsheet, not by `jSuites.calendar`. It keeps the class names of jSuites
(`jcalendar`, `jcalendar-container`, `jcalendar-input`, `jcalendar-selected`, `jcalendar-disabled`, …) and its look,
and the input still keeps the calendar as `calendar`, with `open()`, `close(ignoreEvents)`, `isOpened()`, `getValue()`,
`setValue(value)` and `reset()`. These change:

- The input is an ARIA combobox and the days are an ARIA grid: the input has `aria-expanded`, `aria-controls` and
  `aria-activedescendant`, the days have an accessible name with the weekday and the month, the chosen day has
  `aria-selected`, and a day out of the `validRange` has `aria-disabled`. The focus stays in the input.
- The keys: the arrows move the active day by a day or a week, Page Up and Page Down by a month (with Shift by a year),
  and Home and End to the start and the end of the week. With `readonly: false`, the left and the right arrows, Home and
  End move in the text. The arrows open the calendar again after it closes.
- Enter and Tab choose the active day and close the calendar; Enter keeps the cell selected and moves down, and Tab
  selects the next cell. With `time`, Tab goes through the hour and the minute first. Escape closes the editor without
  saving. A press outside saves the date without the active day, and no longer changes a date that was not changed.
- With `readonly: false`, an empty input saves an empty cell.
- The `fullscreen` option is gone, and so is the full-screen calendar in a window narrower than 800 pixels: the calendar
  always opens under the cell, or over it when there is no room under it.
- The calendar no longer has the type `year-month-picker`, the options `months`, `monthsFull`, `weekdays`,
  `textDone`, `textReset`, `textUpdate`, `controls`, `dataType`, `autoSelect` and `onupdate`, nor the warning class on
  an input with a date out of the `validRange`. The years view shows 25 years. The swipe between the months is gone.
- The texts are translated with `jspreadsheet.setDictionary` (`Reset`, `Done`, `Update`, and the new `Previous`, `Next`,
  `Hour` and `Minute`). The names of the months and the weekdays are translated with it too, as in the cells.

## The rich text and image editors work without jSuites

The editors of the `html` and `image` columns are shown by Jspreadsheet, not by `jSuites.editor` and `jSuites.image`.
The columns keep their types and their values, and the cells show them as before. These change:

- The rich text editor is an element with `contenteditable`, the class `jss_richtext`, and the role `textbox`. It has no
  `editor` property and no `getData()`. The formatting keys of the browser work in it, such as Ctrl+B. Enter adds a
  line, Tab saves the value and selects the next cell, and Escape closes it without saving.
- The rich text editor saves its HTML without scripts, event handlers, forms, images and frames. It keeps the tags of
  formatting (`b`, `i`, `u`, `p`, `div`, `span`, the headings, the lists, `a`, …), the `href` of a link to a web
  address, an email address or a phone number, and these properties of the style: `color`, `background-color`,
  `font-weight`, `font-style`, `font-size`, `text-decoration` and `text-align`. It pastes HTML the same way, and a text
  as lines. jSuites kept more tags and attributes, such as images, tables and `width`.
- The image editor is a dialog with the image and the buttons "Choose an image", "Remove" and "Done". Tab moves between
  the buttons, Enter and Space press them, Done saves the image and keeps the cell selected, and Escape closes the
  dialog without saving. An image file can be dropped on it. It no longer opens the file dialog on a press on the image,
  has no close icon and no `upload` options: the options of the column, such as `maxWidth`, `maxHeight` and
  `maxJpegSizeBytes`, are not read. The element of the dialog keeps it as `upload`, with `getValue()` and `setValue(value)`.
- The new texts are translated with `jspreadsheet.setDictionary`: `Choose an image`, `Remove` and `Done`.

## Jspreadsheet does not need jSuites

Jspreadsheet no longer depends on jSuites. The package is not installed with it, and the bundle no longer reads the
global `jSuites` or calls `require('jsuites')`. A page needs only the script and the stylesheet of Jspreadsheet:

- Remove the script and the stylesheet of jSuites, unless the application uses jSuites itself. The stylesheet of
  Jspreadsheet has the styles of the tabs, the context menu, the toolbar, the pickers, the color picker, the dropdown, the
  calendar and the image editor. The stylesheet of jSuites still works next to it, but it is not needed.
- `jspreadsheet.setDictionary` no longer passes the translations to `jSuites.setDictionary`. An application that shows
  jSuites controls of its own translates them with `jSuites.setDictionary`.
- A press near the top right corner of an element with the class `jclose`, the close icon of the jSuites image editor, no
  longer closes the editor of the cell.

## The pagination puts a space before "entries"

The text after the number of rows on a page starts with a space: "Show 10 entries", where it showed "Show 10entries".
A translation of `entries` with `jspreadsheet.setDictionary` needs no space at its start any more.

## The language files translate by English text

The files in `resources/lang` have the English texts as their keys, as `jspreadsheet.setDictionary` reads them, such as
`"Insert a new column before": "Eine neue Spalte davor einfügen"`. Their keys were the names of version 4, such as
`insertANewColumnBefore`, so they translated nothing (issue #153). They also translate the texts that were added since,
such as those of the toolbar, the calendar, the filters and the image editor. An application that read the old keys
reads the English texts instead. The `Copy`, `Paste` and `Save as` items add the dots after the translation, so their
translations have none.

## The borders of the toolbar open a dialog

The Borders item of the toolbar is a button that opens a dialog, where it was a select item with a picker (issue #146).
The dialog has the types of borders, then the border color, the border thickness and the border style, and the keyboard
reaches each of them: the arrows move between the types, Tab moves between the types, the color, the thickness and the
style, and Escape closes the dialog. The types are buttons with names, such as "Outer borders", which are new texts for
`jspreadsheet.setDictionary`.

- In the items that the `toolbar` option gets, the borders are an icon item (`content: 'border_all'`), not an item of
  type `select` with `options`, `columns` and `onchange`.
- The element of the item has no `picker` property. The dialog is the element after it, with the class `jss_borders`,
  and it is added the first time the dialog opens.
- Choosing a type of borders still closes the dialog, and the button shows the icon of the type chosen last.

## The short names of the months and the days have their own translations

The short names of the months and the days, which the tokens `MMM`, `MON`, `DDD` and `DY` show and the calendar shows,
are texts of their own: `Jan` to `Dec` and `Sun` to `Sat`. Jspreadsheet cut the first three letters of the full names,
so a language whose short names are not those letters showed wrong names, such as `Thá` for `Thg 6` in Vietnamese,
and could not read its own short names in a pasted date (issue #155). A dictionary that translated only the full names
shows the English short names now: add the translations of the short names. The short name of May is the same text as
its full name in English, so the dictionary has it as `May (short)`. A date is read with the full and the short names,
in English and as the dictionary translates them.

The header of the days in the calendar shows their short names, such as `Mon`, where it showed their first letters. The
full names are still their accessible names. The language files in `resources/lang` translate the names of the months
and the days, full and short.

## The merged ranges of the options select no cell

The build of a worksheet with `mergeCells`, and `setData` on such a worksheet, select no cell and dispatch no
`onbeforeselection`, `onselection` or `onfocus`. Version 5 selected the first cell of each range in turn, so the last
range stayed selected and was the tab stop. The worksheet now acts as a worksheet without merged ranges: the selection
does not change, and the first visible cell is the tab stop. To select a cell after the build, call
`updateSelectionFromCoords`. `setMerge` still selects the first cell of its range.

## Lazy loading loads rows in the direction of the scroll

With `lazyLoading`, a load keeps the rows in the view in place, and rows load only in the direction of the scroll. The
direction comes from `deltaY` of the last wheel event, else from the change of `scrollTop`. A wheel event without
`deltaY`, for example a horizontal scroll, and a scroll event that does not change `scrollTop`, load no rows. Version 5
loaded rows for such an event when the view was near the bottom. To load rows from code, use the keyboard methods, such
as `down`.

## The createCell of a custom type can set the alignment

The column `align` applies before the content of a cell, so that the `formatValue` option can change it. A custom
column type that sets `textAlign` in `createCell` now keeps it. Version 5 set the column alignment after `createCell`,
so the alignment of the column replaced it. To keep the column alignment, do not set `textAlign` in `createCell`, or set
it from `column.align`.

## deleteWorksheet removes the options of the worksheet

`deleteWorksheet` removes the options of the worksheet from `getConfig().worksheets`. Version 5 left them there, so the
list named a worksheet that did not exist. `getConfig()` gives the options object that the application passed, so the
worksheet also leaves the `worksheets` array of the application. An `ondeleteworksheet` handler sees the list without the
deleted worksheet, and gets its options as `worksheet.options`. An application that removed the options itself must stop,
or it removes the options of another worksheet.

## oncreateworksheet gives the index of the new worksheet

`oncreateworksheet` occurs when the new worksheet is built, and gives its index in the worksheets at that time. Version 5
gave the number of worksheets minus one. When the application called `createWorksheet` again before the first worksheet
was built, both events got the index of the last worksheet.

## The text editor of a cell is a textarea

The editor of a text or numeric cell is a `textarea` in each cell, so it keeps the line breaks of the value. Version 5
used an `input`, except with `wordWrap`, and the input removed the line breaks when it saved the value. No option keeps
the input, because an input cannot keep the value as it is. Without word wrap, the textarea has `rows="1"` and
`wrap="off"`, and the font of the cell, and it shows the line of the caret, as the input showed its line.

- Code, CSS or tests that find the open editor with `input`, for example `td > input`, must look for `td > textarea`.
  The calendar and color editors are still inputs.
- Alt+Enter adds a line break at the caret in each text cell, and in a custom editor whose first element is a
  textarea. In version 5, it did this only with `wordWrap` or a value of more than 200 characters, and in other cells
  it saved the value and moved the selection down. Alt+Enter is the key for a line break, so no option keeps the old
  behavior. An application can stop the key with a `keydown` listener in the capture phase.
- Text with line breaks that the user pastes into an open editor keeps them. The input changed them to spaces.

## The table is in a wrapper in the content

The content of a worksheet (`worksheet.content`, `div.jss_content`) holds one element, the new wrapper
`worksheet.zoomed` (`div.jss_zoom`), which holds the table, the selection corner and the clipboard helper. Version 5 put
these three elements directly in the content. The wrapper has the zoom of the spreadsheet option `zoom`. CSS or code
that selects `.jss_content > table` must select `.jss_content > .jss_zoom > table`, or `.jss_content table`.

## A column width of 0 stays 0

A column with `width: 0`, and each column without a width when `defaultColWidth` is 0, has no width. `getWidth` gives
0, and `setWidth` sets a width of 0, also for a 0 in an array of widths. Version 5 gave such a column `defaultColWidth`, or 100 pixels, and
`setWidth` with 0 did nothing. No option keeps the old behavior, because `defaultColWidth` is the width of the columns
with no width, and the frozen columns already counted such a column as 0. To keep a visible column, remove `width: 0`, or
set the width that you want.

## Insert and delete keep the merged ranges that they do not cross

`insertRow`, `insertColumn`, `deleteRow` and `deleteColumn` remove only the merged ranges that they cross, after the
dialog. An insert crosses a range when the new rows or columns split it. A delete crosses a range when the range has a
cell in the deleted rows or columns. The other ranges stay and move with their cells. Version 5 removed all merged ranges
of the worksheet.

`deleteRow` also asks before it deletes a row of a range that version 5 did not count, and removes that range: a range
of one row, or the last row of a taller range. Version 5 did not ask. It kept the range of one row on the next row in
`getMerge`, and kept the rowspan of the taller range, which then covered the next row.

No option keeps the old behavior, because the removal of ranges that the action does not touch lost data. The dialog
text stays the same, so that the translations of applications still apply.

## The build leaves out the merged ranges of the options that do not fit

The build of a worksheet, `setData` and `loadWorksheet` leave out a range of `mergeCells` that goes past the last row or
column, or that crosses an earlier range of the option. The console shows the first cells of these ranges, and
`getMerge` does not give them. Version 5 merged them, and the worksheet then threw on a later delete, or kept merged
cells after `destroyMerge`. A range whose first cell was in an earlier range showed the alert "Cell already merged". It
now goes to the console.

The build also leaves out, with the same message, a range that is not valid: a range of one cell, a span that is not a
whole number, or a cell name that is not in capitals. Version 5 kept a range of one cell in `getMerge`. A missing span,
a span of `null`, or a span below 1, is 1, as in `setMerge` and in version 5. The row source mode left out all these
ranges without a message, and it now reads such a span as 1 too.

No option keeps the old behavior, because the ranges that go past the sheet or cross another range broke the worksheet,
and the other ranges merged no cells. To keep a range, give it spans that fit in the sheet and do not cross another
range.

## The editor closes when its cell leaves the table

The open editor saves and closes when its cell leaves the table or no longer shows: a page change, a search, a filter,
an insert, a delete, a move or an undo that takes the row off the page, `hideRow`, `hideColumn` and `setWidth` with 0.
`oneditordetach` occurs first, as for lazy loading, and its handler can take the edit. `setData` closes the editor
without a save. Version 5 kept the editor open on a cell that was not in the table, and a later close saved the text into
the cell that then had the same position. A load of the lazy loading keeps the editor of a row that stays in the table.

The tab stop and the active cell do not go to a cell that does not show: a cell of a hidden row, or of a column with
the width 0. `setWidth` with 0 moves the tab stop, as `hideColumn` does. A selection that starts in such a cell, for
example `selectAll` with a first column of width 0, makes the cell that `down` moves from the active cell: the next
column with a width, in the same row or in the next row that shows. `up` and `down` move in that column. No option
keeps the old behavior, because the user could not see or reach the cell.

## Moves keep the merged ranges that they do not cross, and undo merges the removed ranges again

`moveRow` and `moveColumn` remove only the merged ranges that they cross: a range that has the moved row or column, or a
range that the row or column goes into. The other ranges stay and move with their cells, and `onunmerge` fires only for
the removed ranges. Version 5 removed all merged ranges of the worksheet, and sent `onunmerge` for each. A `moveColumn`
next to a range that it does not cross no longer throws.

An undo of `insertRow`, `deleteRow`, `insertColumn`, `deleteColumn`, `moveRow` or `moveColumn` merges again the ranges
that the action removed, and a redo removes them again, with no question. An editor in a cell that a range hides again
closes without a save, and a selection that starts in such a cell goes to the first cell of the range. The history
records of these actions have a `merges` property with the removed ranges. Version 5 did not give the ranges back.

The actions remove the ranges only after the question about the search results. A cancel of that question keeps the
ranges, also for an undo or a redo of a move. A cancelled undo or redo keeps its place in the history, so the next undo
or redo applies it again. Version 5 removed the ranges before the question, and went past the record of a cancelled
move.

No option keeps the old behavior, because the old behavior lost merged ranges that the user did not change.

## A right press on a row number or a header selects as a left press

A right press on a row number selects the whole row, and a right press on a header selects the whole column, as a left
press does. A selection of whole rows that has the pressed row stays, also one of `selectAll` or of
`updateSelectionFromCoords`. The same applies to whole columns and a header. After the press, `selection.row` (or
`selection.column`) is the first row (or column) of the selection, and the anchor of the other axis is null.

So Delete then deletes the selected rows or columns, as the menu item "Delete selected rows" does: `onbeforedeleterow`
and `ondeleterow` (or the column events) fire instead of `onchange`, and with the `confirm` option false Delete asks no
question. When the right press changes the selection, a press with Shift on a row number extends the selection from the
pressed row. Version 5 kept each selection that had the pressed row or column, also a selection of cells, and Delete
cleared the cells.

No option keeps the old behavior, because Delete did not act on the rows or columns that the menu names.
