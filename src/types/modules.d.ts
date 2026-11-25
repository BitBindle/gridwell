/**
 * Declarations for runtime libraries that have no typings.
 */

declare module '@jspreadsheet/formula' {
    /**
     * Calculate a formula.
     * @param expression - The formula without the leading "=", with the references already replaced.
     * @param variables - The values of the references, by name.
     * @param x - The column of the cell.
     * @param y - The row of the cell.
     * @param instance - The worksheet.
     */
    function formula(expression: string, variables?: Record<string, unknown>, x?: number, y?: number, instance?: unknown): unknown;

    export default formula;
}
