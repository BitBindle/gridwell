# AGENTS.md

Guidance for coding agents working in this repository.

## Project

This is Gridwell (`@bitbindle/gridwell` on npm), BitBindle's fork of [Jspreadsheet CE](https://github.com/jspreadsheet/ce) v5: a vanilla JavaScript data grid with spreadsheet controls. It runs in the browser and has no dependencies. The code still calls the library `jspreadsheet` (the function, the `jss_` classes, the files), and the bundle sets both the globals `gridwell` and `jspreadsheet`.

- **No jSuites.** The library used to show its controls with [jSuites](https://jsuites.net). They were replaced one part at a time, and the package is no longer a dependency: the texts are translated with `src/utils/dictionary.ts`, which `jspreadsheet.setDictionary` fills, the requests are sent with `fetch` in `src/utils/request.ts`, `src/utils/ids.ts` gives the ids of the elements, the tabs (`src/utils/tabs.ts`) have their styles in `src/jspreadsheet.css`, and `src/utils/mask.ts` and `src/utils/dates.ts` show the numbers and the dates with the masks of the columns. The context menu is the class `ContextMenu` in `src/utils/contextMenu.ts`, with the roles and the keys of an ARIA menu, and the toolbar is `ToolbarControl` (`src/utils/toolbarControl.ts`), with the pickers of `src/utils/picker.ts`, the color picker of `src/utils/color.ts`, which the editor of a color column uses too, and the dialog of the borders, `BordersDialog` (`src/utils/borders.ts`). The editor of a dropdown column and the column filters use the `Dropdown` of `src/utils/dropdown.ts`, an ARIA combobox, and the editor of a calendar column uses the `Calendar` of `src/utils/calendar.ts`, a combobox with a grid of the days. The editor of an html column is a contenteditable (`src/utils/richText.ts`, which also cleans the HTML that it saves and pastes), and the editor of an image column is the dialog `ImageUpload` of `src/utils/imageUpload.ts`. These controls keep the class names of jSuites (`jcontextmenu`, `jdropdown`, `jcalendar`, …), and `src/jspreadsheet.css` has their styles.
- **`@jspreadsheet/formula`**: the formula engine, an optional peer dependency. The library does not import it: the application registers it with `jspreadsheet.setFormulaEngine(formula)`, and without an engine a formula shows `#ENGINE`. `src/components/Formulas.ts` calls it, and `test/support/jspreadsheet.ts` registers it for the tests.

The fork has its own fixes (pull requests #5 and later), many of them for accessibility. Changes from upstream are not merged any more.

The code is being reorganised into classes for version 6, one step per pull request. [ARCHITECTURE.md](ARCHITECTURE.md) has the target and the steps, and [UPGRADING.md](UPGRADING.md) lists the breaking changes. Add each breaking change you make to UPGRADING.md.

The source is TypeScript (see [Type checking](#type-checking)). The modules were moved from JavaScript without changing the code, so they still use `any` in many places. The public typings in `src/jspreadsheet.d.ts` are written by hand.

## Commands

The tools need Node 22.22 or later (the library itself runs in the browser). No `package-lock.json` is committed, so use `npm install`.

| Task                    | Command                                                                |
| ----------------------- | ---------------------------------------------------------------------- |
| Dev server (port 8000)  | `npm start` (entry `src/test.js`, page `public/index.html`)            |
| Build `dist/`           | `npm run build`                                                        |
| All tests               | `npm test` (builds first, then runs mocha)                             |
| One test file           | `npx mocha test/data.ts` (build first)                                 |
| Lint / format check     | `npm run lint` (also checks for import cycles), `npm run prettier`     |
| Type check              | `npm run typecheck`                                                    |
| Test coverage           | `npm run coverage` (thresholds in `.c8rc.json`, report in `coverage/`) |
| Fix lint and formatting | `npm run format`                                                       |

Building and testing do not change the source. CI checks the formatting and lint separately, so run `npm run format` before you commit.

`npm run coverage` builds `dist/index.js` without minifying it and with a source map, so c8 reports the coverage of the files in `src/`. It fails when the coverage of a file is under the thresholds in `.c8rc.json`: 85% of the lines, the statements and the functions, and 75% of the branches. Add tests with your change so that it stays above them. Run `npm run build` again afterwards for a minified bundle.

The tests import the built bundle (through `test/support/jspreadsheet.ts`), not `src/`. After you change the source, build again before you run a single test file, or the tests run against old code.

## Layout

- `src/index.ts`: the entry point. It defines `jspreadsheet(element, options)` and its static helpers.
- `src/Spreadsheet.ts`, `src/Worksheet.ts`: the `Spreadsheet` and `Worksheet` classes. Their constructors set up the state an instance has before it is built. `Worksheet.build` loads the data and creates the DOM of a worksheet, and `Worksheet` creates, opens and deletes worksheets (`createWorksheet`, `openWorksheet`, `deleteWorksheet`). `Spreadsheet.build` builds the worksheets in their tabs, the context menu and the toolbar, and the class has the methods of a spreadsheet (`dispatch`, `getConfig`, `setConfig`, `setPlugins`, `fullscreen`, …).
- `src/components/*.ts`: the components of a worksheet, one class per feature, each owning its state and, for the features that have moved, the code that changes it (`History`, `Clipboard`, `Editor`, `Selection`, `Formulas`, `View`, `Pointer`, `Cells`, `Styles`, `Comments`, `Meta`, `Merges`, `Columns`, `Rows`, `Data`, `Accessibility`, `Keyboard`). A component gets its worksheet in its constructor and keeps it as `this.worksheet`. See [ARCHITECTURE.md](ARCHITECTURE.md).
- `src/utils/factory.ts`: replaces a table with a div, then creates the spreadsheet and builds it.
- `src/utils/*.ts`: one module per feature (`helpers`, `worksheets`, `toolbar`, …). `toolbar.ts` and `tabs.ts` have the classes `Toolbar` and `Tabs` of a spreadsheet: `Toolbar` builds the items and creates a `ToolbarControl`. `values.ts` tells whether a value is a formula (`isFormula`) and gives the mask of a column (`getMask`).
- `src/events/*.ts`: the document event handlers, one module per input (`mouse`, `keyboard`, `touch`, `clipboard`, `contextMenu`, `focus`). `index.ts` adds and removes them (`setEvents`, `destroyEvents`), and `helpers.ts` finds the worksheet and the cell of an event target. A handler finds the worksheet and calls its component: the mouse gestures are methods of `worksheet.pointer`, the keys of `worksheet.keyboard`, and the scrolling of the content (`scroll`, `wheel`) of `worksheet.view`.
- `src/jspreadsheet.d.ts`: the public typings, copied unchanged to `dist/index.d.ts`.
- `src/types/`: types for the type check only, not published (see [Type checking](#type-checking)).
- `src/jspreadsheet.css`, `src/jspreadsheet.themes.css`: the styles, also copied to `dist/`.
- `src/test.js`: the entry of the dev server, in JavaScript and not part of the bundle.
- `resources/lang/*.json`: translations for `jspreadsheet.setDictionary`, by English text. `en_GB.json` lists every text that the library translates, and each file has the same texts (`test/languages.ts` checks both). A new text given to `translate` goes in each file. A text that is the same in English as another, with another meaning, gets a context: `translate('May', 'short')` reads the key `May (short)`.
- `test/*.ts`: mocha + chai tests in jsdom, one file per feature. `.mocharc.json` runs them with tsx.
- `test/support/`: the typed library for the tests (`jspreadsheet.ts`) and the globals of the tests (`globals.d.ts`).
- `test/typings.ts`: checks the public typings against what the library returns. `npm run typecheck` checks it; mocha does not run it.

## How the code is organised

- **Two classes.** A _spreadsheet_ (`Spreadsheet`) holds `config`, `element`, `worksheets` and the toolbar: `spreadsheet.toolbar` is its element, and `spreadsheet.toolbarManager`, a `Toolbar`, shows, hides and updates it. A _worksheet_ (`Worksheet`) holds `options`, the DOM (`table`, `thead`, `tbody`, `headers`, `rows`, `records`) and the state (`selection`, …). `worksheet.parent` is its spreadsheet. `worksheet.build` adds the DOM and most of the state. The state that has moved into a component is reached through it: `worksheet.history` is a `History`, a change is recorded with `worksheet.history.add(record)`, and it undoes and redoes the records (`undo`, `redo`); `worksheet.clipboard` is a `Clipboard`, which copies the cells (`copy`), pastes text (`paste`) and keeps what the worksheet copied; `worksheet.editor` is an `Editor`, which opens the editor of a cell (`open`), closes it and saves the value (`close`), and `worksheet.editor.isOpen()` tells whether a cell is being edited; `worksheet.selection` is a `Selection`, which selects cells (`updateFromCoords`, `reset`), moves the selection with the keys (`up`, `down`, …), tells which cells are selected, keeps the selected cell the tab stop of the grid (`updateTabStop`), and holds the selected range, the highlighted cells and the header or row number that a press selected; `worksheet.formulas` is a `Formulas`, which calculates the formulas (`execute`), holds the formula chain and calculates its cells again (`updateChain`), renames the references after rows or columns move (`updateTableReferences`) and shows the footers (`setFooter`); `worksheet.view` is a `View`, which searches, filters, sorts and pages the rows, loads the rows of the lazy loading, and holds the rows that a search or the filters found and the page; `worksheet.pointer` is a `Pointer`, which handles the mouse gestures (`pressCell`, `moveResize`, `endDrag`, …), holds the column or row that the mouse resizes or moves and the drag of the selection corner, and fills the cells when the drag ends (`fill`); `worksheet.keyboard` is a `Keyboard`, which handles the keys of an open editor (`editorKeys`) and of the grid (`gridKeys`). The features whose code has moved are components too: `worksheet.accessibility` sets the roles, the accessible names and the size of the grid (`setRowRole`, `setCellLabel`, `updateGridSize`), `worksheet.cells` creates and updates the td of a cell (`create`, `update`), gives the value that a cell shows (`parseValue`), finds the td of a cell and its label and makes cells read-only, `worksheet.styles` gets and sets the styles of the cells, `worksheet.comments` their comments, `worksheet.meta` their meta information, `worksheet.merges` merges cells and tells which merged ranges a row or a column is in, `worksheet.columns` inserts, moves, deletes, hides and shows the columns, gets and sets their widths, titles and data, and places the frozen columns, and `worksheet.rowManager` (a `Rows`; `worksheet.rows` is the list of row records) creates, inserts, moves, deletes, hides and shows the rows and gets and sets their heights and data, and `worksheet.data` gets and sets the data and the values of the cells (the data itself stays in `options.data`) and updates the table after the cells change (`updateTable`).
- **Classes, not `this` functions.** The code that works on a worksheet or a spreadsheet is a method of a class: of `Worksheet`, `Spreadsheet`, or a component. The methods of a component start with `const worksheet = this.worksheet`, and are called as `worksheet.styles.set(...)`; the methods of `Worksheet` and `Spreadsheet` use `this`, and their callbacks are arrow functions, so no method keeps `this` in a variable (`@typescript-eslint/no-this-alias`). A handler that needs its element reads it from a constant, and a callback in an options object that needs the worksheet is an arrow property (`onclick: () => { ... }`), not a method. A function outside a class does not read `this`: it gets what it needs as arguments. (`jspreadsheet.helpers.getCaretIndex` is the one exception, a public helper that applications call with a spreadsheet as `this`.) The public methods of a worksheet are methods of the `Worksheet` class, most of them one line that calls the method of a component, and the public methods of a spreadsheet are methods of `Spreadsheet`. They are not bound: they need the worksheet or the spreadsheet as `this`.
- **One global active worksheet.** `libraryBase.jspreadsheet.current` is the worksheet that keyboard and clipboard events go to. A press (`mouse.ts`, `touch.ts`), the focus moving into a worksheet (`focus.ts`) and `worksheet.build` set it. A key on an element outside the spreadsheet does not go to it while no editor is open. Each event handler reads it once, into `const current`, and calls the functions for each target or key with it as `this`; a timer reads it again when it runs.
- **Events** go through `dispatch('oneventname', ...args)`, on a worksheet or a spreadsheet. `dispatch` calls the spreadsheet's `config.onevent`, then the specific handler in `config`, then each plugin's `onevent` (`Spreadsheet.callHandlers`), with the worksheet or the spreadsheet as `this`, and returns the last result. On a worksheet, `onafterchanges` also saves the changes when the worksheet has the `persistence` option. Returning `false` from an `onbefore…` handler cancels the action.
- **No import cycles.** The modules under `src/` do not import each other in a cycle at run time. `npm run lint` runs `check-cycles.cjs`, which fails on a cycle; type-only imports do not count. When a new import would close a cycle, move the function that both modules need into a module that imports neither, or make it a method of a class that both modules already have, as `getWorksheetActive` of `Spreadsheet` did.

## Adding or changing a public method

1. Declare it on `WorksheetInstance` in `src/jspreadsheet.d.ts`, with a JSDoc comment.
2. Write the code as a method of the component of its feature (or of a new component).
3. Add a method with the same name and parameters to the `Worksheet` class in `src/Worksheet.ts` that calls it, as `getStyle` calls `this.styles.get(cell, key)`. The class merges with `WorksheetInstance`, so a method whose type does not match the declaration fails the type check. When the declaration has overloads, copy them above the method.
4. Add tests that go through the public API. When the return type depends on the arguments, also add the cases to `test/typings.ts`.

New options also go in `src/jspreadsheet.d.ts` (`SpreadsheetOptions` or `WorksheetOptions`).

## Type checking

`npm run typecheck` runs TypeScript 6.0 over `src/` with `tsconfig.json` and over the tests with `test/tsconfig.json`, and CI runs it too. (TypeScript 7 has no JavaScript API yet, and typescript-eslint needs it. Keep the code valid for TypeScript 7 as well.)

Modules import each other with the `.js` extension, although the files are `.ts` files: TypeScript and webpack find the `.ts` file. esbuild-loader removes the types in the build without checking them, so a type-only import must use `import type`.

- **Settings.** `strict` is on, `strictNullChecks` included.
- **Null and undefined.** Give a value that can be missing its real type (`HTMLElement | null`, `string | undefined`). Where the code reads a value without a check because it is always there at that point, mark it with `!` (`worksheet.filter!.children`, `parentNode!.removeChild`). A comparison with an optional number, such as `worksheet.options.pagination! > 0`, keeps its meaning: `undefined > 0` is false. When a `!` hides a real bug, open an issue and name it in a comment next to the `!`. Do not use `!` to avoid a check that the code needs.
- **The selection.** `selection.range` is `null` until a cell is selected. The private methods of `Selection` that move the selection and the keys on the grid (`src/components/Keyboard.ts`) use `SelectedWorksheet`, a worksheet that has a selection. The keys of an open editor use `EditingWorksheet`, a worksheet whose editor has a cell.
- **Internal types.** `src/types/internal.d.ts` has `WorksheetState` and `SpreadsheetState`: the fields of the classes, which are the public instance types and the state only the library uses. Add new internal fields there, with a comment, and give them their first value in the constructor when they have one before the build. A field goes in `WorksheetInstance` only when applications use it; the DOM and the state that only the library uses (such as `cols`, `corner` and `headerContainer`) stay internal. The file also exports the classes as types: import them with `import type { Worksheet } from '../types/internal';`, and call their methods through the instance. `src/Spreadsheet.ts` and `src/Worksheet.ts` import the modules that build them, so only the builders (`utils/factory.ts` imports `Spreadsheet`, and `Spreadsheet.ts` imports `Worksheet`) import them at run time; other modules use `import type`. The history records have the union `WorksheetHistoryRecord`, by action: a new action needs a member there. DOM event handlers take `TargetEvent<MouseEvent>` and similar, an event whose target the code reads as an element.
- **Other declarations.** `src/types/modules.d.ts` declares `@jspreadsheet/formula`, which has no typings, for the tests and the dev page. `src/types/globals.d.ts` lets `parseInt` take numbers, null and undefined, and `setAttribute` take numbers, as the code does in many places.
- **The public methods.** A method of `Worksheet` or `Spreadsheet` must have the type of its declaration in `src/jspreadsheet.d.ts`: the class merges with the public interface, so the type check compares them. Use rest parameters, not `arguments`.
- **Casts.** Where the type check cannot follow the code, cast with `as X` (or `as any`) and say why in a comment when it is not obvious. Wrap the expression in parentheses rather than add a local variable.
- **No behaviour changes.** A commit that adds or changes types changes no behaviour. Fix a bug that the type check finds in a separate commit, with a test. To show that a typing commit changes no code, build without minifying before and after, format both bundles the same way (for example with esbuild and terser `--format beautify=true,comments=false`) and compare them.

## Build quirks

- A webpack plugin wraps the bundle in a UMD header. The bundle imports no package: do not import `@jspreadsheet/formula` in `src/`, except in the dev page `src/test.js`: the application registers it.
- The build produces one minified UMD file, `dist/index.js`. There is no ES module build.
- `jspreadsheet.version()` gives the version in `package.json`: webpack replaces `__VERSION__` in `src/utils/version.ts` with it. The major and minor version are set by hand in `package.json`, and the publish workflow gives the patch (see [Releases](#releases)).
- The npm package has only `dist/` and `resources/lang/` (`files` in `package.json`), with the README and the license.

## Things to keep

- **Safe cell rendering.** Cell values are written with `textContent`. `innerHTML` is used only when the application sets `parseHTML: true`, and then after `stripScript`. Do not add new `innerHTML` writes of user data.
- **Accessibility.** The grid uses the ARIA grid pattern: roles, accessible names, `aria-selected`, and a roving `tabindex` that moves the keyboard focus to the selected cell (`src/components/Accessibility.ts`, `src/components/Selection.ts`, `test/accessibility.ts`). Changes to selection, editing, hiding or merging must keep these working.
- **Public API compatibility.** Applications use the worksheet methods, options and event signatures. Call out any change to them as a breaking change.

## Tests

- `mocha.config.js` sets up jsdom, makes its window the global object, and adds a global `root` element. It empties `root` after each test.
- The tests are TypeScript, checked as strictly as `src/`. Import the library with `import jspreadsheet, { type Jspreadsheet, type Worksheet } from './support/jspreadsheet';`: it is the built bundle with the public typings, and it returns worksheets of the `Worksheet` class, so a test can read the state that only the library uses. `Jspreadsheet` is the namespace of the public types (`Jspreadsheet.SpreadsheetOptions`, …).
- The tests do not load the Node types: `test/support/globals.d.ts` declares `root`, `global` and `require`.
- When a test passes an argument that the types forbid on purpose, mark the line with `// @ts-expect-error -- reason`. When a test needs a cast because a public typing is wrong, fix the typing instead, in its own commit with a case in `test/typings.ts`.
- Create a spreadsheet with `jspreadsheet(root, { worksheets: [...] })`. It returns the array of worksheet instances.
- The spreadsheet adds its document listeners after it builds its worksheets, a microtask after `jspreadsheet()` returns. Before a test presses or types, wait for them (`await new Promise((resolve) => setTimeout(resolve))`, or resolve in `onload`). Otherwise the test passes only when an earlier test left a listener on the document. `root` is the element of the spreadsheet, so an element that a test puts outside the spreadsheet goes in `document.body`.
- Test through the public API and the DOM, as the existing tests do. For a bug fix, add a test that fails before the fix.
- jsdom has no layout, so sizes and scroll positions are zero. Test the attributes and the state instead.
- For a fix to the public typings, add a case to `test/typings.ts` that fails before the fix: `assertType<Equals<typeof result, Expected>>()`, or `// @ts-expect-error` for a property that must not exist.

## Style

- Prettier (pinned to one version, because its formatting changes between versions): 4 spaces, single quotes, semicolons, a print width of 160 (`.prettierrc`). ESLint uses the recommended rules (`eslint.config.mjs`).
- Write a function outside a class as a function declaration (`export function getCellName(...) { ... }`), a callback as an arrow function, and a function in an object literal as a method (`onclick() { ... }`). A handler that reads its element as `this` stays a function expression. ESLint checks these (`func-style`, `prefer-arrow-callback`, `object-shorthand`).
- Build strings with template literals (`` `${width}px` ``), not `+`. ESLint checks this (`prefer-template`).
- Compare with `===` and `!==`. `x == null` tests for null and undefined. Where a loose comparison is wanted, such as a cell value that can be a number or its text, keep `==` with an `eslint-disable-next-line eqeqeq -- reason` comment. ESLint checks this (`eqeqeq`).
- Read a property of a value that can be missing with optional chaining (`column?.type`), not `column && column.type`. ESLint checks this (`@typescript-eslint/prefer-optional-chain`).
- Write JSDoc comments as a sentence, then `@param name - text` for each parameter, without types: the types are in the code.
- Loop over the items with `for (const item of items)`, and over an object with `for (const [key, value] of Object.entries(o))`; keep an index loop when the index is needed. Search a list with `includes`. ESLint checks these (`@typescript-eslint/prefer-for-of`, `@typescript-eslint/prefer-includes`).
- Follow the style of the code around your change. Keep changes small and focused.

## Commits and pull requests

- Write commit subjects in the imperative, describing the behaviour for users, for example "Keep the selection when a tap is on a jss_object control". Use plain, short sentences.
- When you find a bug that you do not fix in your change, open an issue for it. Follow the existing issues: what goes wrong, a reproduction with the expected and actual result, the branch and commit where you checked it, and the cause in the code. Say why it was not fixed.
- Pull request descriptions follow this order: the problem, a reproduction (code, expected and actual result), the fix, breaking changes, and the tests that fail before and pass after.
- CI (`.github/workflows/ci.yml`) runs `npm test` on Node 22 and 24.
