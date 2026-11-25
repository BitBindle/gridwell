# Upgrading to version 6

Version 6 reorganises the code into classes that own their state (see [ARCHITECTURE.md](ARCHITECTURE.md)). This file
lists the changes that can break an application, and what to do instead. Each step of the reorganisation adds to it.

## Worksheets and spreadsheets are class instances

`jspreadsheet()` still returns the array of worksheets, and each worksheet still has `parent`, its spreadsheet. Both
are now instances of the `Worksheet` and `Spreadsheet` classes, and their methods are on the prototype instead of on
each instance.

-   **A method needs its worksheet as `this`.** The methods used to be bound to their worksheet, so a method taken off
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

-   **The methods are not own properties.** `Object.keys(worksheet)`, `{ ...worksheet }` and `Object.assign({}, worksheet)`
    no longer include the methods. `worksheet.getData` still finds them, and assigning a method to one worksheet still
    replaces it for that worksheet only.

-   **A worksheet has its methods from the start.** Version 5 added them while it built the worksheet. Calling a method
    before the `onload` event is still not supported: the worksheet has no table yet.
