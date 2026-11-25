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

| Task                    | Command                                                          |
| ----------------------- | ---------------------------------------------------------------- |
| Dev server (port 8000)  | `npm start` (entry `src/test.js`, page `public/index.html`)      |
| Build `dist/`           | `npm run build`                                                  |
| All tests               | `npm test` (builds first, then runs mocha)                       |
| One test file           | `npx mocha --require=mocha.config.js test/data.js` (build first) |
| Lint / format check     | `npm run lint`, `npm run prettier`                               |
| Type check              | `npm run typecheck`                                              |
| Fix lint and formatting | `npm run format`                                                 |

Building and testing do not change the source. CI checks the formatting and lint separately, so run `npm run format` before you commit.

The tests import the built bundle (`require('../dist/index.js')`), not `src/`. After you change the source, build again before you run a single test file, or the tests run against old code.

## Layout

-   `src/index.ts`: the entry point. It defines `jspreadsheet(element, options)` and its static helpers.
-   `src/Spreadsheet.ts`, `src/Worksheet.ts`: the `Spreadsheet` and `Worksheet` classes. Their constructors set up the state an instance has before it is built.
-   `src/components/*.ts`: the components of a worksheet, one class per feature, each owning its state (`History`, `Clipboard`, `Editor`). See [ARCHITECTURE.md](ARCHITECTURE.md).
-   `src/utils/factory.ts`: builds the spreadsheet and its worksheets. `spreadsheetPublicMethods` lists the methods of a spreadsheet.
-   `src/utils/worksheets.ts`: builds a worksheet's DOM. `worksheetPublicMethods` lists the methods of a worksheet.
-   `src/utils/*.ts`: one module per feature (`columns`, `rows`, `selection`, `editor`, `events`, `history`, `merges`, `accessibility`, …). `internal.ts` and `events.ts` are the largest and most connected.
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

-   **Two classes.** A _spreadsheet_ (`Spreadsheet`) holds `config`, `element`, `worksheets` and the toolbar. A _worksheet_ (`Worksheet`) holds `options`, the DOM (`table`, `thead`, `tbody`, `headers`, `rows`, `records`) and the state (`selectedCell`, …). `worksheet.parent` is its spreadsheet. Modules still add most of the state while they build the worksheet. The state that has moved into a component is reached through it: `worksheet.history` is a `History`, and a change is recorded with `obj.history.add(record)`; `worksheet.clipboard` is a `Clipboard`, which keeps what the worksheet copied; `worksheet.editor` is an `Editor`, and `obj.editor.isOpen()` tells whether a cell is being edited. To tell a worksheet from a spreadsheet, use `instanceof Worksheet`.
-   **The worksheet is `this`.** Most functions start with `const obj = this` and are called with `fn.call(worksheet, ...)`. The public methods are these functions, put on `Worksheet.prototype` from `worksheetPublicMethods` (and on `Spreadsheet.prototype` from `spreadsheetPublicMethods`). They are not bound: they need the worksheet as `this`.
-   **One global active worksheet.** `libraryBase.jspreadsheet.current` is the worksheet that keyboard and clipboard events go to. `events.ts` and `worksheets.ts` set it.
-   **Events** go through `dispatch.call(obj, 'oneventname', ...args)`. `dispatch` calls the spreadsheet's `config.onevent`, then the specific handler in `config`, then each plugin's `onevent`, and returns the last result. Returning `false` from an `onbefore…` handler cancels the action.
-   **Circular imports** exist (6 run-time cycles, for example `internal → pagination → selection`). They work because imported functions are only called at runtime. Do not call an imported function at the top level of a module.

## Adding or changing a public method

1. Declare it on `WorksheetInstance` in `src/jspreadsheet.d.ts`, with a JSDoc comment.
2. Write the function in the right `src/utils/*.ts` module, typed from the declaration: `export const name: WorksheetMethod<'name'> = function (a, b) { ... };`. The worksheet is `this`.
3. Add it to `worksheetPublicMethods` in `src/utils/worksheets.ts`.
4. Add tests that go through the public API. When the return type depends on the arguments, also add the cases to `test/typings.ts`.

New options also go in `src/jspreadsheet.d.ts` (`SpreadsheetOptions` or `WorksheetOptions`).

## Type checking

`npm run typecheck` runs TypeScript 6.0 over `src/` with `tsconfig.json`, and CI runs it too. (TypeScript 7 has no JavaScript API yet, and typescript-eslint needs it. Keep the code valid for TypeScript 7 as well.)

Modules import each other with the `.js` extension, although the files are `.ts` files: TypeScript and webpack find the `.ts` file. esbuild-loader removes the types in the build without checking them, so a type-only import must use `import type`.

-   **Settings.** `strict` is on, `strictNullChecks` included.
-   **Null and undefined.** Give a value that can be missing its real type (`HTMLElement | null`, `string | undefined`). Where the code reads a value without a check because it is always there at that point, mark it with `!` (`obj.filter!.children`, `parentNode!.removeChild`). A comparison with an optional number, such as `obj.options.pagination! > 0`, keeps its meaning: `undefined > 0` is false. When a `!` hides a real bug, open an issue and name it in a comment next to the `!`. Do not use `!` to avoid a check that the code needs.
-   **The selection.** `selectedCell` is `null` or `undefined` until a cell is selected. The functions that move the selection (`src/utils/keys.ts`) use `SelectedWorksheet`, a worksheet that has a selection.
-   **Internal types.** `src/types/internal.d.ts` has `WorksheetState` and `SpreadsheetState`: the fields of the classes, which are the public instance types and the state only the library uses. Add new internal fields there, with a comment, and give them their first value in the constructor when they have one before the build. The file also exports the classes as types: import them with `import type { Worksheet } from '../types/internal';`, or from `../Worksheet.js` where the code needs the class itself (for `new` or `instanceof`). `src/Worksheet.ts` and `src/Spreadsheet.ts` import only the components at run time, and the components import nothing at run time, so any module can import them without a cycle. The history records have the union `WorksheetHistoryRecord`, by action: a new action needs a member there. DOM event handlers take `TargetEvent<MouseEvent>` and similar, an event whose target the code reads as an element.
-   **Other declarations.** `src/types/modules.d.ts` declares `@jspreadsheet/formula`, which has no typings. `tsconfig.json` maps `jsuites` to `src/types/jsuites-module.d.ts`, which corrects the jsuites typings where they do not match jsuites and adds the methods they miss (some declared in `src/types/jsuites.d.ts`). Correct a jsuites typing there instead of casting at the call. `src/types/globals.d.ts` lets `parseInt` take numbers, null and undefined, and `setAttribute` take numbers, as the code does in many places.
-   **The worksheet as `this`.** Type a public method from its declaration with `WorksheetMethod<'name'>`, so the code and the public typings agree. Give other functions that use the worksheet as `this` a `this: Worksheet` parameter. Use rest parameters, not `arguments`.
-   **Casts.** Where the type check cannot follow the code, cast with `as X` (or `as any`) and say why in a comment when it is not obvious. Wrap the expression in parentheses rather than add a local variable.
-   **No behaviour changes.** A commit that adds or changes types changes no behaviour. Fix a bug that the type check finds in a separate commit, with a test. To show that a typing commit changes no code, build without minifying before and after, format both bundles the same way (for example with esbuild and terser `--format beautify=true,comments=false`) and compare them.

## Build quirks

-   `build.cjs` uses regular expressions to remove `import jSuites from 'jsuites'` and `import formula from '@jspreadsheet/formula'` from every module. A webpack plugin then wraps the bundle in a UMD header that gets both libraries from globals or `require`. Keep these two import lines exactly as they are, or the bundle includes a second copy of each library.
-   The build produces one minified UMD file, `dist/index.js`. There is no ES module build.

## Things to keep

-   **Safe cell rendering.** Cell values are written with `textContent`. `innerHTML` is used only when the application sets `parseHTML: true`, and then after `stripScript`. Do not add new `innerHTML` writes of user data.
-   **Accessibility.** The grid uses the ARIA grid pattern: roles, accessible names, `aria-selected`, and a roving `tabindex` that moves the keyboard focus to the selected cell (`src/utils/accessibility.ts`, `test/accessibility.js`). Changes to selection, editing, hiding or merging must keep these working.
-   **Public API compatibility.** Applications use the worksheet methods, options and event signatures. Call out any change to them as a breaking change.

## Tests

-   `mocha.config.js` sets up jsdom and a global `root` element. It empties `root` after each test.
-   Create a spreadsheet with `jspreadsheet(root, { worksheets: [...] })`. It returns the array of worksheet instances.
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
