/**
 * Declarations for runtime libraries that have no typings.
 */

declare module '@jspreadsheet/formula' {
    // The engine that an application registers with jspreadsheet.setFormulaEngine. The library does not import it.
    const formula: jspreadsheet.FormulaEngine;

    export default formula;
}
