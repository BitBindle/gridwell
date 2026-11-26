/**
 * Declarations for runtime libraries that have no typings.
 */

declare module '@jspreadsheet/formula' {
    // The engine that an application registers with gridwell.setFormulaEngine. The library does not import it.
    const formula: gridwell.FormulaEngine;

    export default formula;
}
