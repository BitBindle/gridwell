/**
 * The formula chain of a worksheet: for each cell that a formula refers to, the cells whose formulas refer to it. When
 * the value of a cell changes, the worksheet calculates these cells again (updateFormulaChain in src/utils/internal.ts).
 *
 * Each worksheet has one, as worksheet.formulas. Calculating a formula reads and writes the cells of the worksheet, so
 * executeFormula and updateFormulaChain stay worksheet functions, and keep the chain here.
 */
export class Formulas {
    /**
     * The names of the cells whose formulas refer to a cell, by the name of that cell. An entry stays when the formula
     * changes; updateFormulaChain skips a cell that no longer has a formula (issue #43).
     */
    chain: Record<string, string[]> = {};

    /** Keep that the formula of the cell `dependent` refers to the cell `cell`. Both are cell names, such as A1. */
    add(cell: string, dependent: string) {
        if (!this.chain[cell]) {
            this.chain[cell] = [];
        }

        if (this.chain[cell].indexOf(dependent) < 0) {
            this.chain[cell].push(dependent);
        }
    }

    /**
     * Rename the cells in the chain after rows or columns moved. `references` gives the new name by the old name; the
     * cells that are not in it keep their names.
     */
    rename(references: Record<string, string>) {
        const chain: Record<string, string[]> = {};

        for (const name of Object.keys(this.chain)) {
            chain[references[name] || name] = this.chain[name].map((dependent) => references[dependent] || dependent);
        }

        this.chain = chain;
    }
}
