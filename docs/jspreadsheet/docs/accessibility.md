title: Spreadsheet Accessibility
keywords: Jspreadsheet, data grid, JavaScript, accessibility, screen reader, ARIA grid, aria-label, assistive technology, translations
description: How Jspreadsheet exposes the worksheet to assistive technology, and how to change the texts of the accessible names.

# Spreadsheet Accessibility

Each worksheet table follows the [ARIA grid pattern](https://www.w3.org/WAI/ARIA/apg/patterns/grid/). The table has the role `grid`, and its rows, column headers, row numbers and cells have the roles `row`, `columnheader`, `rowheader` and `gridcell`. The attributes `aria-rowcount` and `aria-colcount` give the full size of the worksheet, also when lazy loading or pagination shows only some of the rows. Each row and cell gives its position in `aria-rowindex` and `aria-colindex`.

## Documentation

### Accessible names

The table has the name of the worksheet. The corner cell that selects all cells, the hidden text area of the copy action and the text, numeric, calendar, color, dropdown and rich text (`html`) editors have the English names below. A custom editor sets the accessible name of its own control.

### Settings

| Property                                                                     | Description                                                                                                                                         |
|------------------------------------------------------------------------------|-----------------------------------------------------------------------------------------------------------------------------------------------------|
| `labels: { selectAll?: string, editCell?: string, clipboardHelper?: string }` | Spreadsheet-level texts of the accessible names. `selectAll` default: `'Select all'`. `editCell` default: `'Edit cell {0}'`, where `{0}` is the address of the cell, for example `B4`. `clipboardHelper` default: `'Clipboard helper'`. |

A text that is not in `labels` is the English default. `jspreadsheet.setDictionary` can translate the English defaults for all spreadsheets on the page. Use `labels` when each spreadsheet needs its own texts, for example two languages on one page. Set `labels` when you create the spreadsheet.

## Examples

### Texts in the language of the page

{.ignore-execution}
```html
<div id="spreadsheet"></div>

<script>
jspreadsheet(document.getElementById('spreadsheet'), {
    labels: {
        selectAll: 'Alles auswählen',
        editCell: 'Zelle {0} bearbeiten',
        clipboardHelper: 'Zwischenablage',
    },
    worksheets: [{
        worksheetName: 'Tabelle1',
        minDimensions: [6, 6],
    }],
});
</script>
```
