title: Jspreadsheet Freeze Columns
keywords: Jspreadsheet, freeze columns, data grid customization, JavaScript spreadsheets, Excel-like freeze functionality, floating columns, spreadsheet column freezing
description: This section provides comprehensive information on freezing columns in Jspreadsheet, including programmatic settings, methods, and related events.

# Freeze Columns

Freezing columns in a spreadsheet keeps specific columns visible as you scroll. This section covers the settings, methods, and events for managing frozen columns. 

{.pro}
> #### What you can find on the Pro Version
> 
> The **Pro version** of Jspreadsheet provides programmatic methods, events, and manual controls for `freezing columns`, enabling users to adjust this setting directly through the interface.    
> \
> [Learn more](https://jspreadsheet.com/docs/freeze-columns){.button}

## Documentation

### Initial Settings

Below are the initial settings related to freezing columns.

| Property                   | Description                                       |
|----------------------------|---------------------------------------------------|
| **freezeColumns**: number  | Defines the columns to freeze on initialization.  |

### Line of the frozen columns

When the worksheet is scrolled, the last frozen column has a line of 2 pixels on its right side. By default, the line is transparent, so only the shadow of the frozen columns shows. In a forced-colors mode, the browser removes the shadow and shows the line in the text color of the system. Set the CSS custom property `--jss-frozen-line-color` to show the line in all modes.

```css
.jss_worksheet {
    --jss-frozen-line-color: #595959;
}
```


> #### Limitations
> Freeze columns currently have limited functionality when used with filters and footers. These improvements are planned for the next release.

## Examples

### Basic Freeze Columns Example

A straightforward example of how to use frozen columns in Jspreadsheet.

```html
<html>
<script src="https://bossanova.uk/jspreadsheet/v5/jspreadsheet.js"></script>
<script src="https://jsuites.net/v5/jsuites.js"></script>
<link rel="stylesheet" href="https://bossanova.uk/jspreadsheet/v5/jspreadsheet.css" type="text/css" />
<link rel="stylesheet" href="https://jsuites.net/v5/jsuites.css" type="text/css" />

<link rel="stylesheet" href="https://fonts.googleapis.com/css?family=Material+Icons" />

<div id="spreadsheet"></div>

<script>
// Create a new spreadsheet
let spreadsheet = jspreadsheet(document.getElementById('spreadsheet'), {
    worksheets: [{
        minDimensions: [50,10],
        tableOverflow: true,
        tableWidth: '800px',
        freezeColumns: 1
    }]
});

</script>
</html>
```
```jsx
import React, { useRef } from "react";
import { Spreadsheet, Worksheet } from "@jspreadsheet-ce/react";
import "jsuites/dist/jsuites.css";
import "jspreadsheet-ce/style.css";

export default function App() {
    // Spreadsheet array of worksheets
    const spreadsheet = useRef();

    // Render component
    return (
        <>
            <Spreadsheet ref={spreadsheet}>
                <Worksheet minDimensions={[50,10]} tableOverflow={true} tableWidth={"800px"} freezeColumns={1} />
            </Spreadsheet>
        </>
    );
}
```
```vue
<template>
    <Spreadsheet ref="spreadsheet">
        <Worksheet :minDimensions="[50,10]" :tableOverflow="true" tableWidth="800px" :freezeColumns="1" />
    </Spreadsheet>
</template>

<script>
import { Spreadsheet, Worksheet } from "@jspreadsheet-ce/vue";
import "jsuites/dist/jsuites.css";
import "jspreadsheet-ce/dist/jspreadsheet.css";

export default {
    components: {
        Spreadsheet,
        Worksheet,
    }
};
</script>

```
```angularjs
import { Component, ViewChild, ElementRef } from "@angular/core";
import jspreadsheet from "jspreadsheet-ce";

import "jspreadsheet-ce/dist/jspreadsheet.css";
import "jsuites/dist/jsuites.css";

@Component({
    standalone: true,
    selector: "app-root",
    template: `<div #spreadsheet></div>`
})
export class AppComponent {
    @ViewChild("spreadsheet") spreadsheet: ElementRef;
    // Worksheets
    worksheets: jspreadsheet.worksheetInstance[];
    // Create a new data grid
    ngAfterViewInit() {
        // Create spreadsheet
        this.worksheets = jspreadsheet(this.spreadsheet.nativeElement, {
            worksheets: [
                minDimensions: [50,10],
                tableOverflow: true,
                tableWidth: '800px',
                freezeColumns: 1
            }]
        });
    }
}
```

## Related Tools

- [Jspreadsheet Pro Freeze Columns](https://jspreadsheet.com/docs/freeze-columns) - Advanced freeze with user controls (Pro)
- [Jspreadsheet Pro Freeze Rows](https://jspreadsheet.com/docs/freeze-rows) - Freeze header rows with interactive controls (Pro)
