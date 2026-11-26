# AGENTS.md

Guidance for coding agents working in this repository.

## Project

Gridwell (`@bitbindle/gridwell` on npm) is a vanilla JavaScript data grid with spreadsheet controls, forked from Jspreadsheet CE v5. It runs in the browser and has no runtime dependencies. Internally the library is still called `jspreadsheet` (the function, the `jss_` classes, the file names); the bundle sets the globals `gridwell` and `jspreadsheet`.

- The UI controls (context menu, toolbar, pickers, dropdown, calendar, rich-text and image editors) are our own, in `src/utils/`, with the ARIA roles and keys of their widget patterns. They keep the jSuites class names (`jcontextmenu`, `jdropdown`, …) for styling.
- The formula engine `@jspreadsheet/formula` is an optional peer dependency. The library never imports it: the application registers it with `jspreadsheet.setFormulaEngine(formula)`. Without one, a formula shows `#ENGINE`.

The code is being reorganised into classes for version 6, one step per pull request. [ARCHITECTURE.md](ARCHITECTURE.md) has the target. Record every breaking change in [UPGRADING.md](UPGRADING.md).

## Commands

Node 22.22 or later. Use `npm install` (no lockfile is committed). In the dev container use `pnpm` instead of `npm`, and never mix the two in one `node_modules`.

| Task                    | Command                                          |
| ----------------------- | ------------------------------------------------ |
| Dev server (port 8000)  | `npm start` (`src/test.js`, `public/index.html`) |
| Build `dist/`           | `npm run build`                                  |
| All tests               | `npm test` (builds first)                        |
| One test file           | `npx mocha test/data.ts` (build first)           |
| Lint (+ import cycles)  | `npm run lint`                                   |
| Format check            | `npm run prettier`                               |
| Type check              | `npm run typecheck`                              |
| Coverage                | `npm run coverage`                               |
| Fix lint and formatting | `npm run format`                                 |

- The tests run against the built bundle, not `src/`. Rebuild before running a single test file.
- `npm run coverage` fails below the thresholds in `.c8rc.json`. Add tests with your change to stay above them.
- Run `npm run format` before committing; CI checks formatting and lint.

## Layout

- `src/index.ts`: entry point, `jspreadsheet(element, options)` and its static helpers.
- `src/Spreadsheet.ts`, `src/Worksheet.ts`: the two main classes.
- `src/components/`: one class per worksheet feature (`Selection`, `Editor`, `History`, `Formulas`, …).
- `src/events/`: document event handlers, one module per input. They find the worksheet and call its components.
- `src/utils/`: stateless helpers and the UI controls.
- `src/jspreadsheet.d.ts`: hand-written public typings, shipped as `dist/index.d.ts`.
- `src/types/`: internal types for the type check only (`internal.d.ts` has the state of the classes).
- `resources/lang/*.json`: translations. Every file has the same keys as `en_GB.json`.
- `test/`: mocha + chai tests in jsdom, one file per feature. `test/typings.ts` checks the public typings.

## Architecture principles

- **A spreadsheet owns worksheets; a worksheet owns components.** Each component gets its worksheet in the constructor (`this.worksheet`) and owns its state. Other code goes through its methods, not its fields.
- **Public methods live on the classes.** A public worksheet method is a thin method on `Worksheet` that delegates to a component (`getStyle` → `this.styles.get(...)`). Methods are not bound.
- **No `this` outside classes.** Functions outside a class take what they need as arguments. Callbacks are arrow functions; no `this` aliases.
- **No import cycles.** `npm run lint` fails on one. Break it by moving the shared function to a module neither imports, or onto a class both already use. Type-only imports don't count.
- **Events go through `dispatch`.** `dispatch('oneventname', ...args)` calls `config.onevent`, then the specific handler, then the plugins. Returning `false` from an `onbefore…` handler cancels the action.
- **One active worksheet.** Keyboard and clipboard events go to `jspreadsheet.current`. Handlers read it once into a local constant.

## Adding a public method or option

1. Declare it in `src/jspreadsheet.d.ts` with a JSDoc comment (`WorksheetInstance`, `SpreadsheetOptions` or `WorksheetOptions`).
2. Implement it in the component of its feature.
3. Add a delegating method with the same signature (and overloads) to `Worksheet`. The class merges with the interface, so a mismatch fails the type check.
4. Test it through the public API; add `test/typings.ts` cases when the return type depends on the arguments.

## TypeScript

- `strict` is on. Code must stay valid for TypeScript 7.
- Import modules with the `.js` extension. Use `import type` for type-only imports (esbuild strips types without checking).
- Give possibly missing values their real type (`HTMLElement | null`). Use `!` only where the value is guaranteed; never to skip a needed check.
- Cast with `as X` where the checker can't follow the code, with a comment when the reason isn't obvious.
- Internal-only state goes in `src/types/internal.d.ts`, initialised in the constructor. Only state applications use goes in the public typings.
- Modules other than the builders import `Worksheet` and `Spreadsheet` with `import type`.
- Typing-only commits change no behaviour. Fix bugs found by the type checker in a separate commit, with a test.

## Things to keep

- **Safe rendering.** Write cell values with `textContent`. `innerHTML` only when `parseHTML: true`, after `stripScript`. No new `innerHTML` writes of user data.
- **Accessibility.** The grid follows the ARIA grid pattern: roles, accessible names, `aria-selected`, and a roving `tabindex` on the selected cell. Changes to selection, editing, hiding or merging must keep it working (`test/accessibility.ts`).
- **API compatibility.** Call out any change to public methods, options or event signatures as breaking.
- **No bundled packages.** The bundle imports nothing; `@jspreadsheet/formula` is imported only by the dev page.
- **Translations.** A new text passed to `translate` goes in every `resources/lang` file. Disambiguate same-spelled texts with a context: `translate('May', 'short')` → key `May (short)`.

## Tests

- Import with `import jspreadsheet, { type Jspreadsheet, type Worksheet } from './support/jspreadsheet';` and create sheets with `jspreadsheet(root, { worksheets: [...] })`.
- Document listeners attach a microtask after creation. Await `new Promise((resolve) => setTimeout(resolve))` before simulating input.
- Test through the public API and the DOM. jsdom has no layout, so assert attributes and state, not sizes.
- A bug fix comes with a test that fails before the fix.
- Mark deliberately invalid arguments with `// @ts-expect-error -- reason`. If a test needs a cast because a public typing is wrong, fix the typing instead.

## Style

Prettier and ESLint enforce most of this; run `npm run format`.

- Function declarations for top-level functions, arrow functions for callbacks, method shorthand in object literals.
- Template literals, `===` (`x == null` is fine), optional chaining, `for…of`, `includes`.
- JSDoc: a sentence, then `@param name - text` without types.
- Follow the surrounding code. Keep changes small and focused.

## Commits, pull requests, releases

- Commit subjects are imperative and describe user-visible behaviour, e.g. "Keep the selection when a tap is on a jss_object control".
- PR descriptions: problem, reproduction (code, expected, actual), fix, breaking changes, tests that fail before and pass after.
- Open an issue for any bug you find but don't fix: reproduction, branch and commit, cause, and why it wasn't fixed.
- Releases are published by hand from the Actions tab (`publish.yml`); the patch version comes from npm. Change `package.json`'s version only for a new major or minor. See [CONTRIBUTING.md](CONTRIBUTING.md#releases).
