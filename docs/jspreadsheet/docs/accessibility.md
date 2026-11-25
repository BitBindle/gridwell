title: Spreadsheet Accessibility
keywords: Jspreadsheet, data grid, JavaScript, accessibility, screen reader, ARIA grid, aria-label, assistive technology, translations
description: How Jspreadsheet exposes the worksheet to assistive technology, and how to change the texts of the accessible names.

# Spreadsheet Accessibility

Each worksheet table follows the [ARIA grid pattern](https://www.w3.org/WAI/ARIA/apg/patterns/grid/). The table has the role `grid`, and its rows, column headers, row numbers and cells have the roles `row`, `columnheader`, `rowheader` and `gridcell`. The attributes `aria-rowcount` and `aria-colcount` give the full size of the worksheet, also when lazy loading or pagination shows only some of the rows. Each row and cell gives its position in `aria-rowindex` and `aria-colindex`.

## Documentation

### Accessible names

The table has the name of the worksheet. The corner cell that selects all cells, the hidden text area of the copy action and the text, numeric, calendar, color, dropdown and rich text (`html`) editors have the English names below. A custom editor sets the accessible name of its own control.

### Keyboard focus

The cells use a roving tabindex. The spreadsheet container stays a tab stop, and the grid adds one tab stop after it. Before a selection, this is the first visible cell. After a selection, it is the selected cell, and the focus follows the selection when the focus is in the grid. When the focused cell leaves the table or moves, for example after a search or a sort, the new tab stop gets the focus. A selection or a change of the grid does not take the focus from an element outside the spreadsheet.

The focused cell has a name with its address and its text, for example `B4 Total`. A cell without text, for example a checkbox or a color square, gives its value, for example `B4 true`. The name of an image cell has only its address. The grid has `aria-multiselectable`, because a selection can have more than one cell. The selected cells have `aria-selected="true"`. The tab stop has `aria-selected="false"` when it is not selected.

Tab moves the focus to the tab stop of the grid. The keys act on the grid after a cell is selected, for example with a press on the cell.

When the editor of a cell is open, the editor has the focus, also for the dropdown editor. The cell is then not a tab stop and has no name. When the editor closes, for example with Escape, the cell gets the focus back.

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
