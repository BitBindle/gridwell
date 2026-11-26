# AGENTS.md

Guidance for coding agents working in this repository.

## Project

This is BitBindle's fork of [Jspreadsheet CE](https://github.com/jspreadsheet/ce) v5, a vanilla JavaScript data grid with spreadsheet controls. It runs in the browser and depends on two packages at runtime:

-   `jsuites`: UI components (calendar, dropdown, context menu, tabs, `ajax`, `translate`)
-   `@jspreadsheet/formula`: the formula engine

The fork has its own fixes (pull requests #5 and later), many of them for accessibility. Changes from upstream are not merged any more.

The code is being reorganised into classes for version 6, one step per pull request. [ARCHITECTURE.md](ARCHITECTURE.md) has the target and the steps, and [UPGRADING.md](UPGRADING.md) lists the breaking changes. Add each breaking change you make to UPGRADING.md.

The source is TypeScript (see [Type checking](#type-checking)). The modules were moved from JavaScript without changing the code, so they still use `any` in many places. The public typings in `src/jspreadsheet.d.ts` are written by hand.

## Commands

No `package-lock.json` is committed, so use `npm install`.

| Task                    | Command                                                            |
| ----------------------- | ------------------------------------------------------------------ |
| Dev server (port 8000)  | `npm start` (entry `src/test.js`, page `public/index.html`)        |
| Build `dist/`           | `npm run build`                                                    |
| All tests               | `npm test` (builds first, then runs mocha)                         |
| One test file           | `npx mocha --require=mocha.config.js test/data.js` (build first)   |
| Lint / format check     | `npm run lint` (also checks for import cycles), `npm run prettier` |
| Type check              | `npm run typecheck`                                                |
| Fix lint and formatting | `npm run format`                                                   |

Building and testing do not change the source. CI checks the formatting and lint separately, so run `npm run format` before you commit.

The tests import the built bundle (`require('../dist/index.js')`), not `src/`. After you change the source, build again before you run a single test file, or the tests run against old code.

## Layout

-   `src/index.ts`: the entry point. It defines `jspreadsheet(element, options)` and its static helpers.
-   `src/Spreadsheet.ts`, `src/Worksheet.ts`: the `Spreadsheet` and `Worksheet` classes. Their constructors set up the state an instance has before it is built.
-   `src/components/*.ts`: the components of a worksheet, one class per feature, each owning its state and, for the features that have moved, the code that changes it (`History`, `Clipboard`, `Editor`, `Selection`, `Formulas`, `View`, `Pointer`, `Cells`, `Styles`, `Comments`, `Meta`, `Merges`). A component gets its worksheet in its constructor and keeps it as `this.worksheet`. See [ARCHITECTURE.md](ARCHITECTURE.md).
-   `src/utils/factory.ts`: builds the spreadsheet and its worksheets. `spreadsheetPublicMethods` lists the methods of a spreadsheet.
-   `src/utils/worksheets.ts`: builds a worksheet's DOM. `worksheetPublicMethods` lists the methods of a worksheet.
-   `src/utils/*.ts`: one module per feature (`columns`, `rows`, `history`, `accessibility`, …). `render.ts` creates and renders the cells (`createCell`, `updateCell`), and `table.ts` has `updateTable`, which runs after the cells change. `values.ts` has the value that a cell shows (`parseValue`, `getCellColumnIndex`), which many modules need.
-   `src/events/*.ts`: the document event handlers, one module per input (`mouse`, `keyboard`, `touch`, `clipboard`, `contextMenu`, `scroll`, `focus`). `index.ts` adds and removes them (`setEvents`, `destroyEvents`), and `helpers.ts` finds the worksheet and the cell of an event target.
-   `src/jspreadsheet.d.ts`: the public typings, copied unchanged to `dist/index.d.ts`.
-   `src/types/`: types for the type check only, not published (see [Type checking](#type-checking)).
-   `src/jspreadsheet.css`, `src/jspreadsheet.themes.css`: the styles, also copied to `dist/`.
-   `src/test.js`: the entry of the dev server. `src/webcomponent.js`: a `<jspreadsheet>` custom element. Both are JavaScript and not part of the bundle.
-   `resources/lang/*.json`: translations for `jspreadsheet.setDictionary`.
-   `test/*.js`: mocha + chai tests in jsdom, one file per feature.
-   `test/typings.ts`: checks the public typings against what the library returns. `npm run typecheck` checks it; mocha does not run it.
-   `packages/vue`: a small Vue wrapper, published separately.
-   `docs/`: old v4 documentation. Prettier skips it.

## How the code is organised

-   **Two classes.** A _spreadsheet_ (`Spreadsheet`) holds `config`, `element`, `worksheets` and the toolbar. A _worksheet_ (`Worksheet`) holds `options`, the DOM (`table`, `thead`, `tbody`, `headers`, `rows`, `records`) and the state (`selection`, …). `worksheet.parent` is its spreadsheet. Modules still add most of the state while they build the worksheet. The state that has moved into a component is reached through it: `worksheet.history` is a `History`, and a change is recorded with `obj.history.add(record)`; `worksheet.clipboard` is a `Clipboard`, which copies the cells (`copy`), pastes text (`paste`) and keeps what the worksheet copied; `worksheet.editor` is an `Editor`, which opens the editor of a cell (`open`), closes it and saves the value (`close`), and `obj.editor.isOpen()` tells whether a cell is being edited; `worksheet.selection` is a `Selection`, which selects cells (`updateFromCoords`, `reset`), moves the selection with the keys (`up`, `down`, …), tells which cells are selected, keeps the selected cell the tab stop of the grid (`updateTabStop`), and holds the selected range, the highlighted cells and the header or row number that a press selected; `worksheet.formulas` is a `Formulas`, which calculates the formulas (`execute`), holds the formula chain and calculates its cells again (`updateChain`), renames the references after rows or columns move (`updateTableReferences`) and shows the footers (`setFooter`); `worksheet.view` is a `View`, which searches, filters, sorts and pages the rows, loads the rows of the lazy loading, and holds the rows that a search or the filters found and the page; `worksheet.pointer` is a `Pointer`, which holds the column or row that the mouse resizes or moves and the drag of the selection corner, and fills the cells when the drag ends (`fill`). The features whose code has moved are components too: `worksheet.cells` finds the td of a cell and its label and makes cells read-only, `worksheet.styles` gets and sets the styles of the cells, `worksheet.comments` their comments, `worksheet.meta` their meta information and `worksheet.merges` merges cells and tells which merged ranges a row or a column is in. To tell a worksheet from a spreadsheet, use `instanceof Spreadsheet` (see [Type checking](#type-checking)).
-   **The worksheet is `this`.** Most functions start with `const obj = this` and are called with `fn.call(worksheet, ...)`; the methods of a component start with `const obj = this.worksheet` instead, and are called as `obj.styles.set(...)`. The public methods are these functions, or a function that calls the method of a component, put on `Worksheet.prototype` from `worksheetPublicMethods` (and on `Spreadsheet.prototype` from `spreadsheetPublicMethods`). They are not bound: they need the worksheet as `this`.
-   **One global active worksheet.** `libraryBase.jspreadsheet.current` is the worksheet that keyboard and clipboard events go to. A press (`mouse.ts`, `touch.ts`), the focus moving into a worksheet (`focus.ts`) and `buildWorksheet` set it. A key on an element outside the spreadsheet does not go to it while no editor is open. Each event handler reads it once, into `const current`, and calls the functions for each target or key with it as `this`; a timer reads it again when it runs.
-   **Events** go through `dispatch.call(obj, 'oneventname', ...args)`. `dispatch` calls the spreadsheet's `config.onevent`, then the specific handler in `config`, then each plugin's `onevent`, and returns the last result. Returning `false` from an `onbefore…` handler cancels the action.
-   **No import cycles.** The modules under `src/` do not import each other in a cycle at run time. `npm run lint` runs `check-cycles.cjs`, which fails on a cycle; type-only imports do not count. When a new import would close a cycle, move the function that both modules need into a module that imports neither, as `values.ts` (the value that a cell shows) and `activeWorksheet.ts` (the worksheet that a spreadsheet shows) did.

## Adding or changing a public method

1. Declare it on `WorksheetInstance` in `src/jspreadsheet.d.ts`, with a JSDoc comment.
2. Write the function in the right `src/utils/*.ts` module, typed from the declaration: `export const name: WorksheetMethod<'name'> = function (a, b) { ... };`. The worksheet is `this`. When the feature is a component, write a method of the component instead.
3. Add it to `worksheetPublicMethods` in `src/utils/worksheets.ts`. For a method of a component, add a function that calls it, as `getStyle` calls `this.styles.get(cell, key)`.
4. Add tests that go through the public API. When the return type depends on the arguments, also add the cases to `test/typings.ts`.

New options also go in `src/jspreadsheet.d.ts` (`SpreadsheetOptions` or `WorksheetOptions`).

## Type checking

`npm run typecheck` runs TypeScript 6.0 over `src/` with `tsconfig.json`, and CI runs it too. (TypeScript 7 has no JavaScript API yet, and typescript-eslint needs it. Keep the code valid for TypeScript 7 as well.)

Modules import each other with the `.js` extension, although the files are `.ts` files: TypeScript and webpack find the `.ts` file. esbuild-loader removes the types in the build without checking them, so a type-only import must use `import type`.

-   **Settings.** `strict` is on, `strictNullChecks` included.
-   **Null and undefined.** Give a value that can be missing its real type (`HTMLElement | null`, `string | undefined`). Where the code reads a value without a check because it is always there at that point, mark it with `!` (`obj.filter!.children`, `parentNode!.removeChild`). A comparison with an optional number, such as `obj.options.pagination! > 0`, keeps its meaning: `undefined > 0` is false. When a `!` hides a real bug, open an issue and name it in a comment next to the `!`. Do not use `!` to avoid a check that the code needs.
-   **The selection.** `selection.range` is `null` until a cell is selected. The private methods of `Selection` that move the selection and the keys on the grid (`src/events/keyboard.ts`) use `SelectedWorksheet`, a worksheet that has a selection. The keys of an open editor use `EditingWorksheet`, a worksheet whose editor has a cell.
-   **Internal types.** `src/types/internal.d.ts` has `WorksheetState` and `SpreadsheetState`: the fields of the classes, which are the public instance types and the state only the library uses. Add new internal fields there, with a comment, and give them their first value in the constructor when they have one before the build. A field goes in `WorksheetInstance` only when applications use it; the DOM and the state that only the library uses (such as `cols`, `corner` and `headerContainer`) stay internal. The file also exports the classes as types: import them with `import type { Worksheet } from '../types/internal';`, or from `../Spreadsheet.js` where the code needs the class itself (for `instanceof`). `src/Spreadsheet.ts` imports nothing at run time, so any module can import it. `src/Worksheet.ts` imports the components, which import the modules they need, so only the builders (`utils/worksheets.ts` and `utils/factory.ts`) import it at run time: a module that tells a worksheet from a spreadsheet uses `instanceof Spreadsheet`. The history records have the union `WorksheetHistoryRecord`, by action: a new action needs a member there. DOM event handlers take `TargetEvent<MouseEvent>` and similar, an event whose target the code reads as an element.
-   **Other declarations.** `src/types/modules.d.ts` declares `@jspreadsheet/formula`, which has no typings. `tsconfig.json` maps `jsuites` to `src/types/jsuites-module.d.ts`, which corrects the jsuites typings where they do not match jsuites and adds the methods they miss (some declared in `src/types/jsuites.d.ts`). Correct a jsuites typing there instead of casting at the call. `src/types/globals.d.ts` lets `parseInt` take numbers, null and undefined, and `setAttribute` take numbers, as the code does in many places.
-   **The worksheet as `this`.** Type a public method from its declaration with `WorksheetMethod<'name'>`, so the code and the public typings agree. Give other functions that use the worksheet as `this` a `this: Worksheet` parameter. Use rest parameters, not `arguments`.
-   **Casts.** Where the type check cannot follow the code, cast with `as X` (or `as any`) and say why in a comment when it is not obvious. Wrap the expression in parentheses rather than add a local variable.
-   **No behaviour changes.** A commit that adds or changes types changes no behaviour. Fix a bug that the type check finds in a separate commit, with a test. To show that a typing commit changes no code, build without minifying before and after, format both bundles the same way (for example with esbuild and terser `--format beautify=true,comments=false`) and compare them.

## Build quirks

-   `build.cjs` uses regular expressions to remove `import jSuites from 'jsuites'` and `import formula from '@jspreadsheet/formula'` from every module. A webpack plugin then wraps the bundle in a UMD header that gets both libraries from globals or `require`. Keep these two import lines exactly as they are, or the bundle includes a second copy of each library.
-   The build produces one minified UMD file, `dist/index.js`. There is no ES module build.

## Things to keep

-   **Safe cell rendering.** Cell values are written with `textContent`. `innerHTML` is used only when the application sets `parseHTML: true`, and then after `stripScript`. Do not add new `innerHTML` writes of user data.
-   **Accessibility.** The grid uses the ARIA grid pattern: roles, accessible names, `aria-selected`, and a roving `tabindex` that moves the keyboard focus to the selected cell (`src/utils/accessibility.ts`, `src/components/Selection.ts`, `test/accessibility.js`). Changes to selection, editing, hiding or merging must keep these working.
-   **Public API compatibility.** Applications use the worksheet methods, options and event signatures. Call out any change to them as a breaking change.

## Tests

-   `mocha.config.js` sets up jsdom and a global `root` element. It empties `root` after each test.
-   Create a spreadsheet with `jspreadsheet(root, { worksheets: [...] })`. It returns the array of worksheet instances.
-   The spreadsheet adds its document listeners after it builds its worksheets, a microtask after `jspreadsheet()` returns. Before a test presses or types, wait for them (`await new Promise((resolve) => setTimeout(resolve))`, or resolve in `onload`). Otherwise the test passes only when an earlier test left a listener on the document. `root` is the element of the spreadsheet, so an element that a test puts outside the spreadsheet goes in `document.body`.
-   Test through the public API and the DOM, as the existing tests do. For a bug fix, add a test that fails before the fix.
-   jsdom has no layout, so sizes and scroll positions are zero. Test the attributes and the state instead.
-   For a fix to the public typings, add a case to `test/typings.ts` that fails before the fix: `assertType<Equals<typeof result, Expected>>()`, or `// @ts-expect-error` for a property that must not exist.

## Style

-   Prettier: 4 spaces, single quotes, semicolons, a print width of 160 (`.prettierrc`). ESLint uses the recommended rules (`eslint.config.mjs`).
-   Follow the style of the code around your change. Keep changes small and focused.

## Commits and pull requests

-   Write commit subjects in the imperative, describing the behaviour for users, for example "Keep the selection when a tap is on a jss_object control". Use plain, short sentences.
-   When you find a bug that you do not fix in your change, open an issue for it. Follow the existing issues: what goes wrong, a reproduction with the expected and actual result, the branch and commit where you checked it, and the cause in the code. Say why it was not fixed.
-   Pull request descriptions follow this order: the problem, a reproduction (code, expected and actual result), the fix, breaking changes, and the tests that fail before and pass after.
-   CI (`.github/workflows/ci.yml`) runs `npm test` on Node 20 and 22.
