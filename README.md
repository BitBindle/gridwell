# Gridwell

Gridwell is a JavaScript data grid with spreadsheet controls. It runs in the browser, has no dependencies, and works with
the keyboard and screen readers.

Gridwell started as a fork of [Jspreadsheet CE](https://github.com/jspreadsheet/ce) v5. It has its own code, releases
and issues now, and does not merge changes from upstream.

## Features

- Column types: text, numbers with masks, dropdowns, calendars, checkboxes, colors, images and rich text.
- Copy and paste with Excel and Google Sheets, undo and redo, and filling cells by dragging the selection corner.
- Sorting, search, column filters, pagination and lazy loading.
- Merged cells, comments, frozen columns, nested headers, a toolbar and a context menu.
- Several worksheets in tabs.
- An ARIA grid. The keyboard reaches every control, and the selected cell is the focus.
- Formulas through the optional `@jspreadsheet/formula` engine.
- Translations for nine languages.
- TypeScript typings.

## Compared with Jspreadsheet CE 5.0.4

|                                  | Jspreadsheet CE 5.0.4                  | Gridwell                           |
| -------------------------------- | -------------------------------------- | ---------------------------------- |
| Runtime dependencies             | 2 (`jsuites`, `@jspreadsheet/formula`) | 0 (the formula engine is optional) |
| Packages installed               | 3                                      | 1                                  |
| Installed size                   | 1.15 MB                                | 0.44 MB                            |
| JavaScript a page loads, gzipped | 86.3 kB                                | 59.5 kB (-31%)                     |
| CSS a page loads, gzipped        | 15.2 kB                                | 8.1 kB (-47%)                      |

The page sizes are for the library without the formula engine. Both versions use the same engine, which adds 44.7 kB
gzipped. All the files were minified the same way with esbuild before gzip.

The speed is about the same. These are the median times in milliseconds in Chromium, over 9 runs, with 10 columns and
10,000 rows:

| Task                                    | Jspreadsheet CE 5.0.4 | Gridwell |
| --------------------------------------- | --------------------: | -------: |
| Create the grid                         |                  1568 |     1580 |
| Create a grid of 2,000 rows of formulas |                   350 |      309 |
| Sort a column                           |                  1846 |     1776 |
| Search                                  |                   227 |      245 |
| Set 500 values                          |                    89 |       93 |
| Insert 100 rows                         |                   530 |      576 |

The numbers were measured in November 2026, against `jspreadsheet-ce` 5.0.4 with `jsuites` 5.13.5.

## Installation

```bash
npm install @bitbindle/gridwell
```

```javascript
import gridwell from '@bitbindle/gridwell';
import '@bitbindle/gridwell/dist/gridwell.css';
```

Or load the files from a CDN. The script sets the global `gridwell`:

```html
<script src="https://cdn.jsdelivr.net/npm/@bitbindle/gridwell/dist/index.js"></script>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@bitbindle/gridwell/dist/gridwell.css" />
```

The toolbar and some controls show [Material Icons](https://fonts.google.com/icons):

```html
<link rel="stylesheet" href="https://fonts.googleapis.com/css?family=Material+Icons" />
```

## Usage

```html
<div id="spreadsheet"></div>
```

```javascript
gridwell(document.getElementById('spreadsheet'), {
    worksheets: [
        {
            data: [
                ['Jazz', 'Honda', '2019-02-12', true, 2000],
                ['Civic', 'Honda', '2018-07-11', true, 4000.01],
            ],
            columns: [
                { type: 'text', title: 'Car', width: 120 },
                { type: 'dropdown', title: 'Make', width: 160, source: ['Alfa Romeo', 'Audi', 'BMW', 'Honda'] },
                { type: 'calendar', title: 'Available', width: 140 },
                { type: 'checkbox', title: 'Stock', width: 80 },
                { type: 'numeric', title: 'Price', width: 120, mask: '$ #,##0.00' },
            ],
        },
    ],
});
```

`gridwell()` returns the worksheets. Each worksheet has the methods that read and change it, such as `getData()`,
`setValue('A1', 'Text')` and `insertRow()`.

## Formulas

The formula engine is a separate package. Install it and register it once, before you create a grid. Without an engine,
a formula shows `#ENGINE`.

```javascript
import gridwell from '@bitbindle/gridwell';
import formula from '@jspreadsheet/formula';

gridwell.setFormulaEngine(formula);
```

## Translations

`gridwell.setDictionary` translates the texts of the grid: the context menu, the toolbar, the editors, the search, the
pagination, and the names of the months and the days. It takes the translations by English text. The package has
translations in `resources/lang`: German, Danish, French, Italian, Polish, Portuguese, Vietnamese and Chinese.
`en_GB.json` lists all the texts.

```javascript
import de from '@bitbindle/gridwell/resources/lang/de_DE.json';

gridwell.setDictionary(de);
```

The short name of May is `May (short)`, because it is the same as the full name in English.

## Upgrading from Jspreadsheet CE

The library is called `gridwell` instead of `jspreadsheet`, and its stylesheet is `gridwell.css`. The methods, options and
events keep the names of Jspreadsheet CE v5, but some of them have changed. [UPGRADING.md](UPGRADING.md) lists each
breaking change and what to do instead.

## Development

You need Node 22.22 or later.

```bash
npm install
npm start          # dev server on port 8000
npm test           # build, then run the tests
npm run lint       # ESLint and the import cycle check
npm run typecheck
```

The tests run in jsdom and cover 96% of the lines. CI fails when a file has less than 85% of its lines or 75% of its
branches covered.

[AGENTS.md](AGENTS.md) describes the code and the conventions, and [ARCHITECTURE.md](ARCHITECTURE.md) describes the
classes.

## License

MIT. See [LICENSE](LICENSE). Gridwell includes code from Jspreadsheet CE, © Jspreadsheet Ltd.
