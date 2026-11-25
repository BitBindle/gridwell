# AGENTS.md

Guidance for coding agents working in this repository.

## Project

This is BitBindle's fork of [Jspreadsheet CE](https://github.com/jspreadsheet/ce) v5, a vanilla JavaScript data grid with spreadsheet controls. It runs in the browser and depends on two packages at runtime:

-   `jsuites`: UI components (calendar, dropdown, context menu, tabs, `ajax`, `translate`)
-   `@jspreadsheet/formula`: the formula engine

The fork has its own fixes (pull requests #5 and later), many of them for accessibility. Changes from upstream are not merged automatically.

The code is moving to TypeScript step by step. For now the source is JavaScript, some modules are type-checked through JSDoc (see [Type checking](#type-checking)), and the public typings in `src/index.d.ts` are written by hand.

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

-   `src/index.js`: the entry point. It defines `jspreadsheet(element, options)` and its static helpers.
-   `src/utils/factory.js`: creates the spreadsheet object and its worksheets.
-   `src/utils/worksheets.js`: builds a worksheet's DOM. `worksheetPublicMethods` lists every method a worksheet instance gets.
-   `src/utils/*.js`: one module per feature (`columns`, `rows`, `selection`, `editor`, `events`, `history`, `merges`, `accessibility`, …). `internal.js` and `events.js` are the largest and most connected.
-   `src/index.d.ts`: the public typings, copied to `dist/` unchanged.
-   `src/types/`: types for the type check only, not published (see [Type checking](#type-checking)).
-   `src/jspreadsheet.css`, `src/jspreadsheet.themes.css`: the styles, also copied to `dist/`.
-   `src/webcomponent.js`: a `<jspreadsheet>` custom element. It is not part of the bundle.
-   `resources/lang/*.json`: translations for `jspreadsheet.setDictionary`.
-   `test/*.js`: mocha + chai tests in jsdom, one file per feature.
-   `packages/vue`: a small Vue wrapper, published separately.
-   `docs/`: old v4 documentation. Prettier skips it.

## How the code is organised

-   **Two objects.** A _spreadsheet_ holds `config`, `element`, `worksheets` and the toolbar. A _worksheet_ holds `options`, the DOM (`table`, `thead`, `tbody`, `headers`, `rows`, `records`) and the state (`selectedCell`, `history`, …). `worksheet.parent` is its spreadsheet. Both are plain objects that modules add properties to; there are no classes.
-   **The worksheet is `this`.** Most functions start with `const obj = this` and are called with `fn.call(worksheet, ...)`. Public methods are bound to each worksheet in `setWorksheetFunctions`.
-   **One global active worksheet.** `libraryBase.jspreadsheet.current` is the worksheet that keyboard and clipboard events go to. `events.js` and `worksheets.js` set it.
-   **Events** go through `dispatch.call(obj, 'oneventname', ...args)`. `dispatch` calls the spreadsheet's `config.onevent`, then the specific handler in `config`, then each plugin's `onevent`, and returns the last result. Returning `false` from an `onbefore…` handler cancels the action.
-   **Circular imports** exist (17 cycles, for example `internal → pagination → selection → history → orderBy`). They work because imported functions are only called at runtime. Do not call an imported function at the top level of a module.

## Adding or changing a public method

1. Write the function in the right `src/utils/*.js` module, with the worksheet as `this`.
2. Add it to `worksheetPublicMethods` in `src/utils/worksheets.js`.
3. Declare it on `WorksheetInstance` in `src/index.d.ts`, with a JSDoc comment.
4. Add tests that go through the public API.

New options also go in `src/index.d.ts` (`SpreadsheetOptions` or `WorksheetOptions`).

## Type checking

`npm run typecheck` runs TypeScript 6.0 over `src/` with `tsconfig.json`, and CI runs it too. (TypeScript 7 has no JavaScript API yet, and typescript-eslint needs it. Keep the code valid for TypeScript 7 as well.)

The modules are moving from `.js` to `.ts` files. A `.ts` module is always checked. A `.js` module is checked only when its first line is `// @ts-check`. Modules import each other with the `.js` extension, also when the file is a `.ts` file: TypeScript and webpack find the `.ts` file. esbuild-loader removes the types in the build, so a type-only import must use `import type`.

-   **Settings.** `strict` is on except `strictNullChecks`. Turn that on once the modules are TypeScript files.
-   **Internal types.** `src/types/internal.d.ts` has `Worksheet` and `Spreadsheet`: the public instance types and the state only the library uses. Add new internal fields there, with a comment.
-   **Other declarations.** `src/types/modules.d.ts` declares `@jspreadsheet/formula`, which has no typings. `src/types/jsuites.d.ts` adds the jsuites methods that its typings miss. `src/types/globals.d.ts` lets `parseInt` take numbers, as the code does in many places.
-   **In a checked module:**
    -   Import the internal types with `/** @import { Worksheet } from '../types/internal' */`.
    -   Give each function that uses the worksheet as `this` the tag `@this {Worksheet}`, and type every parameter with `@param`.
    -   Where the type check cannot follow the code, cast with `/** @type {X} */ (expression)` and say why in a comment.
-   **JSDoc and TypeScript 7.** TypeScript 7 ignores JSDoc `@overload` tags: declare overloads as `@type {{ (a: string): X; (a: number): Y }}` on the constant, as `getIdFromColumnName` does. A function that reads `arguments` is not variadic: give it a function type, as `dispatch` does.
-   **No behaviour changes.** A commit that adds types changes no behaviour. Fix a bug that the type check finds in a separate commit, with a test.
-   When you change a module that is not checked yet, consider checking it in the same pull request.

## Build quirks

-   `build.cjs` uses regular expressions to remove `import jSuites from 'jsuites'` and `import formula from '@jspreadsheet/formula'` from every module. A webpack plugin then wraps the bundle in a UMD header that gets both libraries from globals or `require`. Keep these two import lines exactly as they are, or the bundle includes a second copy of each library.
-   The build produces one minified UMD file, `dist/index.js`. There is no ES module build.

## Things to keep

-   **Safe cell rendering.** Cell values are written with `textContent`. `innerHTML` is used only when the application sets `parseHTML: true`, and then after `stripScript`. Do not add new `innerHTML` writes of user data.
-   **Accessibility.** The grid uses the ARIA grid pattern: roles, accessible names, `aria-selected`, and a roving `tabindex` that moves the keyboard focus to the selected cell (`src/utils/accessibility.js`, `test/accessibility.js`). Changes to selection, editing, hiding or merging must keep these working.
-   **Public API compatibility.** Applications use the worksheet methods, options and event signatures. Call out any change to them as a breaking change.

## Tests

-   `mocha.config.js` sets up jsdom and a global `root` element. It empties `root` after each test.
-   Create a spreadsheet with `jspreadsheet(root, { worksheets: [...] })`. It returns the array of worksheet instances.
-   Test through the public API and the DOM, as the existing tests do. For a bug fix, add a test that fails before the fix.
-   jsdom has no layout, so sizes and scroll positions are zero. Test the attributes and the state instead.

## Style

-   Prettier: 4 spaces, single quotes, semicolons, a print width of 160 (`.prettierrc`). ESLint uses the recommended rules (`eslint.config.mjs`).
-   Follow the style of the code around your change. Keep changes small and focused.

## Commits and pull requests

-   Write commit subjects in the imperative, describing the behaviour for users, for example "Keep the selection when a tap is on a jss_object control". Use plain, short sentences.
-   Pull request descriptions follow this order: the problem, a reproduction (code, expected and actual result), the fix, breaking changes, and the tests that fail before and pass after.
-   CI (`.github/workflows/ci.yml`) runs `npm test` on Node 20 and 22.
