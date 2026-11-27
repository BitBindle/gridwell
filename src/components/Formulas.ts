import libraryBase from '../utils/libraryBase.js';
import { getColumnName } from '../utils/helpers.js';
import { getColumnNameFromId, getIdFromColumnName } from '../utils/internalHelpers.js';
import { isFormula } from '../utils/values.js';
import type { FormulaText, HistoryCellChange, MergeInfo, Worksheet } from '../types/internal';

/** The rows (axis 1) or the columns (axis 0) that a delete removed, at their positions before the delete. */
interface Removed {
    axis: 0 | 1;
    index: number;
    count: number;
}

/**
 * An insert (count above 0) or a delete (count below 0) of rows (axis 1) or columns (axis 0) at index, with the number of
 * rows and columns of the table before it.
 */
interface Shift {
    axis: 0 | 1;
    index: number;
    count: number;
    rows: number;
    columns: number;
}

/**
 * The new name of a cell outside the table after an insert or a delete. Undefined when the cell stays: a cell before the
 * change, a cell of the table, or any cell when the change moves no row or column of the table, for example when it adds
 * rows after the last row. Null when the delete removed the row or the column of the cell.
 */
function shiftName(name: string, shift: Shift): string | null | undefined {
    const first = shift.count > 0 ? shift.index : shift.index - shift.count;
    if (first >= (shift.axis === 1 ? shift.rows : shift.columns)) {
        return undefined;
    }
    const cell = getIdFromColumnName(name);
    if ((cell[0] < shift.columns && cell[1] < shift.rows) || cell[shift.axis] < shift.index) {
        return undefined;
    }
    if (cell[shift.axis] < first) {
        return null;
    }
    cell[shift.axis] += shift.count;
    return getColumnNameFromId(cell);
}

// The cells in progress of updateChain, by cell name
let chainLoopProtection: Record<string, boolean> = {};

// Whether the console already said that no formula engine is registered
let noEngineWarned = false;

/**
 * Escapes the characters of a text that have a meaning in a regular expression.
 * @param text - The text.
 */
function escapeRegExp(text: string) {
    return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Writes a text as a JavaScript string literal for the formula plugin. The plugin searches the value of each variable for
 * ranges, such as A:B in "Ratio A:B", so a colon is written as an escape, which gives the same text.
 * @param text - The text.
 */
function toStringLiteral(text: string) {
    return JSON.stringify(text).replace(/:/g, '\\u003a');
}

/**
 * The formulas of a worksheet: their calculation, the formula chain, the references to rename after columns or rows
 * move, are inserted or are deleted, and the footers.
 *
 * The formula chain keeps, for each cell that a formula refers to, the cells whose formulas refer to it. When the value
 * of a cell changes, the worksheet calculates these cells again (updateChain).
 *
 * Each worksheet has one, as worksheet.formulas. executeFormula calls execute, and the cells show the result of their
 * formula through parseValue (src/utils/values.ts), which calls it too.
 */
export class Formulas {
    readonly worksheet: Worksheet;

    /**
     * The names of the cells whose formulas refer to a cell, by the name of that cell. An entry stays when the formula
     * changes; updateChain skips a cell that no longer has a formula (issue #43).
     */
    chain: Record<string, string[]> = {};

    constructor(worksheet: Worksheet) {
        this.worksheet = worksheet;
    }

    /** True while keepChain runs a change: add does not change the chain then. */
    private paused = false;

    /** Keep that the formula of the cell `dependent` refers to the cell `cell`. Both are cell names, such as A1. */
    add(cell: string, dependent: string) {
        if (this.paused) {
            return;
        }
        if (!this.chain[cell]) {
            this.chain[cell] = [];
        }

        if (!this.chain[cell].includes(dependent)) {
            this.chain[cell].push(dependent);
        }
    }

    /**
     * Rename the cells in the chain after rows or columns moved. `references` gives the new name by the old name; the
     * cells that are not in it keep their names. Without keys, only the formula cells get new names, because the formula
     * texts keep the names of the cells that they refer to.
     */
    rename(references: Record<string, string>, keys = true) {
        const chain: Record<string, string[]> = {};

        for (const name of Object.keys(this.chain)) {
            // A name that a delete removed can get the same new name as a cell that moved, as the formula texts do. The
            // dependents of both stay, and a name without dependents leaves the chain.
            const key = (keys && references[name]) || name;
            const dependents = this.chain[name].map((dependent) => references[dependent] || dependent);
            if (dependents.length === 0) {
                continue;
            }
            if (!chain[key]) {
                chain[key] = dependents;
            } else {
                for (const dependent of dependents) {
                    if (!chain[key].includes(dependent)) {
                        chain[key].push(dependent);
                    }
                }
            }
        }

        this.chain = chain;
    }

    /**
     * Run a change and keep the chain as it was before it. Insert creates the cells of the new rows or columns before
     * the other cells move, so the formulas of the new cells would add their names before the names of the chain change.
     * updateTableReferences adds the new cells to the chain after the move.
     */
    keepChain<T>(change: () => T): T {
        const paused = this.paused;
        this.paused = true;
        try {
            return change();
        } finally {
            this.paused = paused;
        }
    }

    /**
     * Reads a formatted number, such as "$ 1,234.50", "1 234,5 €" or "-5%", with the decimal point of its column. Symbols
     * may come before and after the number, and its integer part may have thousands separators. Any other text, such as
     * "Room 12b" or a date, is not a number.
     * @param value - The value of a cell.
     * @param columnNumber - The column whose options the cell uses.
     * @returns The number, or null when the text is not a formatted number.
     */
    private parseNumber(value: unknown, columnNumber: number) {
        const worksheet = this.worksheet;

        // Decimal point. Column 0 has one too.
        const column = worksheet.options.columns[columnNumber];
        const decimal = column?.decimal ? column.decimal : '.';
        const thousands = decimal === ',' ? '.' : ',';

        // Symbols, such as a sign, a currency or a percent sign, then the integer part, then the decimal part, then symbols
        const symbols = '[^\\p{L}\\p{N}\\s]*';
        const pattern = new RegExp(`^(${symbols})\\s*(\\d[\\d\\s'${escapeRegExp(thousands)}]*?)(?:${escapeRegExp(decimal)}(\\d*))?\\s*${symbols}$`, 'u');
        const parts = pattern.exec(`${value}`.trim());

        if (!parts) {
            return null;
        }

        const sign = parts[1].includes('-') ? '-' : '';

        return Number(`${sign}${parts[2].replace(/\D/g, '')}.${parts[3] || '0'}`);
    }

    /**
     * Parse formulas
     */
    execute(expression: string, x?: number, y?: number): unknown {
        const worksheet = this.worksheet;

        // The results and the formulas in progress, by cell name
        const formulaResults: Record<string, unknown> = {};
        const formulaLoopProtection: Record<string, boolean> = {};

        // A formula engine of the application calculates the formula in place of the formula plugin
        const formulaEngine = typeof worksheet.parent.config.formulaEngine === 'function' ? worksheet.parent.config.formulaEngine : null;

        // Execute formula with loop protection
        function execute(expression: string, x?: number, y?: number) {
            // The formula text as typed, for the formula engine
            const text = expression;

            // A formula without x and y, of a footer or of executeFormula, is in no cell: it has no name, no loop and no
            // place in the chain
            const parentId = x != null && y != null ? getColumnNameFromId([x, y]) : null;

            // Code protection
            if (parentId) {
                if (formulaLoopProtection[parentId]) {
                    console.error('Reference loop detected');
                    return '#ERROR';
                }

                formulaLoopProtection[parentId] = true;
            }

            // The names of the cells of a range
            function getRangeNames(from: string, to: string) {
                const names = [];
                const e1 = getIdFromColumnName(from);
                const e2 = getIdFromColumnName(to);
                const x1 = Math.min(e1[0], e2[0]);
                const x2 = Math.max(e1[0], e2[0]);
                const y1 = Math.min(e1[1], e2[1]);
                const y2 = Math.max(e1[1], e2[1]);

                for (let j = y1; j <= y2; j++) {
                    for (let i = x1; i <= x2; i++) {
                        names.push(getColumnNameFromId([i, j]));
                    }
                }

                return names;
            }

            // Range with $ remove $ (string literals are kept untouched)
            expression = expression.replace(/("(?:[^"]|"")*")|\$?([A-Z]+)\$?([0-9]+)/g, (match, literal, column, row) => {
                return literal ? literal : column + row;
            });

            // A range becomes the list of its cells. A string literal has no references, so it is kept as it is.
            expression = expression.replace(/("(?:[^"]|"")*")|([A-Z]+[0-9]+):([A-Z]+[0-9]+)/g, (match, literal, from, to) => {
                return literal ? literal : getRangeNames(from, to).join(',');
            });

            // Get the references, outside the string literals
            const stripped = expression.replace(/"(?:[^"]|"")*"/g, '""');
            const tokens = stripped.match(/([A-Z]+[0-9]+)/g);

            // Direct self-reference protection
            if (parentId && tokens?.includes(parentId)) {
                console.error('Self Reference detected');
                return '#ERROR';
            } else {
                // Expressions to be used in the parsing
                const formulaExpressions: Record<string, unknown> = {};
                // A reference to a formula cell whose result is #REF! gives the formula the result #REF!
                let refError = false;

                if (tokens) {
                    for (const reference of tokens) {
                        // Keep chain
                        if (parentId) {
                            worksheet.formulas.add(reference, parentId);
                        }

                        // Do not calculate again. The formula engine reads the values of the references itself.
                        if (!formulaEngine && typeof (globalThis as unknown as Record<string, unknown>)[reference] === 'undefined') {
                            // Coords
                            const position = getIdFromColumnName(reference);
                            // Get value
                            let value: unknown;

                            if (typeof worksheet.options.data[position[1]]?.[position[0]] !== 'undefined') {
                                value = worksheet.options.data[position[1]][position[0]];
                            } else {
                                value = '';
                            }
                            // Get column data
                            if (`${value}`.substr(0, 1) === '=') {
                                if (typeof formulaResults[reference] !== 'undefined') {
                                    value = formulaResults[reference];
                                } else {
                                    value = execute(value as string, position[0], position[1]);
                                    formulaResults[reference] = value;
                                }
                                if (value === '#REF!') {
                                    refError = true;
                                }
                            }
                            // Type!
                            if (`${value}`.trim() === '') {
                                // Null
                                formulaExpressions[reference] = null;
                            } else {
                                // eslint-disable-next-line eqeqeq -- a text that is a number equals the number
                                if (value == Number(value) && worksheet.parent.config.autoCasting !== false) {
                                    // Number
                                    formulaExpressions[reference] = Number(value);
                                } else {
                                    // Trying any formatted number
                                    // The decimal point of the column whose options the cell uses
                                    const number = worksheet.formulas.parseNumber(value, worksheet.cells.getColumnIndex(position[0], position[1]));
                                    if (worksheet.parent.config.autoCasting !== false && number !== null) {
                                        formulaExpressions[reference] = number;
                                    } else {
                                        // Text as a string literal, with its quotes and backslashes escaped
                                        formulaExpressions[reference] = toStringLiteral(`${value}`);
                                    }
                                }
                            }
                        }
                    }
                }

                let source = text;
                let given = false;

                const ret = worksheet.dispatch('onbeforeformula', worksheet, expression, x, y);
                if (ret === false) {
                    return expression;
                } else if (ret) {
                    expression = ret;
                    source = ret;
                    given = true;
                }

                // A delete gave a reference the text #REF!, and the formula shows #REF!. So does a formula that names a
                // formula cell with the result #REF!. The error functions, such as IFERROR, do not catch it. A formula
                // engine of the application gets the text.
                if (!formulaEngine && (refError || (given ? expression.replace(/"(?:[^"]|"")*"/g, '""') : stripped).includes('#REF!'))) {
                    return '#REF!';
                }

                // Hand string literals to the formula parser as variables, so characters such as $ or & inside them are not
                // interpreted as part of the formula syntax
                let literalIndex = 0;
                const parsedExpression = expression.substr(1).replace(/"(?:[^"]|"")*"/g, (literal) => {
                    const name = `__JSSSTR${literalIndex++}__`;
                    // The text of the literal, with "" as a quote, as JavaScript code: a backslash is a character, as in Excel
                    formulaExpressions[name] = toStringLiteral(literal.slice(1, -1).replace(/""/g, '"'));
                    return name;
                });

                // Convert formula to javascript
                let res;

                try {
                    if (formulaEngine) {
                        res = formulaEngine(worksheet, source, x === undefined ? x : Number(x), y === undefined ? y : Number(y));
                    } else if (libraryBase.formulaEngine) {
                        res = libraryBase.formulaEngine(parsedExpression, formulaExpressions, x, y, worksheet);
                    } else {
                        // Excel has no such error, so the cell tells that the application did not register an engine
                        if (!noEngineWarned) {
                            noEngineWarned = true;
                            console.warn(
                                'Jspreadsheet: no formula engine is registered, so the formulas show #ENGINE. Register one with jspreadsheet.setFormulaEngine(formula), where formula is the default export of @jspreadsheet/formula.'
                            );
                        }

                        res = '#ENGINE';
                    }

                    if (typeof res === 'function') {
                        res = '#ERROR';
                    }
                } catch (e) {
                    res = '#ERROR';

                    if (worksheet.parent.config.debugFormulas === true) {
                        console.log(expression.substr(1), formulaExpressions, e);
                    }
                }

                return res;
            }
        }

        return execute(expression, x, y);
    }

    /**
     * Show again the result of each formula that names a cell at or after a start, and of the formulas that name these
     * formulas. Each formula cell shows its result once, with no data write and no event. A formula cell that a delete
     * removed leaves the chain.
     */
    private showResults(firstColumn: number, firstRow: number) {
        const worksheet = this.worksheet;
        function exists(name: string) {
            const [x, y] = getIdFromColumnName(name);
            return !!worksheet.records[y]?.[x];
        }

        const next: string[] = [];
        for (const name of Object.keys(this.chain)) {
            this.chain[name] = this.chain[name].filter(exists);
            const [x, y] = getIdFromColumnName(name);
            if (x >= firstColumn && y >= firstRow) {
                next.push(name);
            }
        }
        this.showDependents(next);
    }

    /** Show again the result of each formula that names one of the cells, directly or through other formulas. */
    private showDependents(next: string[]) {
        const worksheet = this.worksheet;
        const cells = new Set<string>();
        while (next.length) {
            for (const dependent of this.chain[next.pop()!] || []) {
                if (!cells.has(dependent)) {
                    cells.add(dependent);
                    next.push(dependent);
                }
            }
        }

        for (const name of cells) {
            const [x, y] = getIdFromColumnName(name);
            if (isFormula(worksheet.options.data[y][x])) {
                worksheet.cells.showValue(x, y);
            }
        }
    }

    /**
     * Update all related cells in the chain
     */
    updateChain(x: number, y: number, records: HistoryCellChange[]) {
        const worksheet = this.worksheet;

        const cellId = getColumnNameFromId([x, y]);
        if (worksheet.formulas.chain[cellId] && worksheet.formulas.chain[cellId].length > 0) {
            if (chainLoopProtection[cellId]) {
                worksheet.records[y][x].element.innerHTML = '#ERROR';
                // The chain ends here
                worksheet.formulas.chain[cellId] = [];
            } else {
                // Protection
                chainLoopProtection[cellId] = true;

                for (let i = 0; i < worksheet.formulas.chain[cellId].length; i++) {
                    const cell = getIdFromColumnName(worksheet.formulas.chain[cellId][i]);
                    // A cell outside the window of the row source mode has no record, and calculates when it shows
                    if (worksheet.rowSource && !worksheet.records[cell[1]]?.[cell[0]]) {
                        continue;
                    }
                    // Update cell
                    const value = `${worksheet.options.data[cell[1]][cell[0]]}`;
                    if (value.substr(0, 1) === '=') {
                        const record = worksheet.cells.update(cell[0], cell[1], value, true);
                        if (record) {
                            records.push(record);
                        }
                    } else {
                        // No longer a formula: remove it from the chain, and read the next cell at the same index
                        worksheet.formulas.chain[cellId].splice(i, 1);
                        i--;
                    }
                    this.updateChain(cell[0], cell[1], records);
                }
            }
        }

        chainLoopProtection = {};
    }

    /**
     * Update formula
     */
    updateFormula(formula: string, referencesToUpdate: Record<string, string>) {
        const testLetter = /[A-Z]/;
        const testNumber = /[0-9]/;

        let newFormula = '';
        let letter = null;
        let number = null;
        let token = '';

        // for...of would read the string by code points, and this loop reads its UTF-16 code units
        // eslint-disable-next-line @typescript-eslint/prefer-for-of
        for (let index = 0; index < formula.length; index++) {
            if (testLetter.exec(formula[index])) {
                letter = 1;
                number = 0;
                token += formula[index];
            } else if (testNumber.exec(formula[index])) {
                number = letter ? 1 : 0;
                token += formula[index];
            } else {
                if (letter && number) {
                    token = referencesToUpdate[token] ? referencesToUpdate[token] : token;
                }
                newFormula += token;
                newFormula += formula[index];
                letter = 0;
                number = 0;
                token = '';
            }
        }

        if (token) {
            if (letter && number) {
                token = referencesToUpdate[token] ? referencesToUpdate[token] : token;
            }
            newFormula += token;
        }

        return newFormula;
    }

    /**
     * Whether a cell or a footer of the worksheet holds a formula.
     */
    private hasFormulas() {
        for (const footer of this.worksheet.options.footers || []) {
            if (footer.some((value) => typeof value === 'string' && value[0] === '=')) {
                return true;
            }
        }
        for (const row of this.worksheet.options.data) {
            // The row source mode has rows without values
            if (!row) {
                continue;
            }
            for (const value of row) {
                if (typeof value === 'string' && value[0] === '=') {
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * Give a formula the new names of the cells that moved. A cell name keeps its $ signs: they fix a reference for a
     * copy, not for an insert or a delete. The string literals, the function names, such as LOG10, and the cells and
     * ranges after a worksheet name and a !, also a quoted name, do not change.
     *
     * With removed, a delete removed rows or columns. A reference to a removed cell becomes #REF!, with or without $
     * signs. A range that keeps some of its cells ends at the first and the last cell that stay, with the $ signs of its
     * ends. A range that loses all its cells becomes #REF!.
     *
     * referencesToUpdate gives the new names of the cells of the table. With shift, a reference to a cell outside the
     * table moves as the cells of its rows or columns would move. When the insert or the delete moves no row or column
     * of the table, for example when it adds rows after the last row, these references stay.
     * @param removed - The removed rows (axis 1) or columns (axis 0), at their positions before the delete.
     * @param shift - The insert or the delete.
     * @returns The new text, and if the formula named a removed cell, alone or in a range, so that its result changes.
     */
    private renameReferences(formula: string, referencesToUpdate: Record<string, string>, removed?: Removed, shift?: Shift) {
        // Give a cell name the column and the row of another cell, and keep its $ signs
        function withName(token: string, name: string) {
            const [, column, row] = /^([A-Z]+)([0-9]+)$/.exec(name)!;
            return token.replace(/^(\$?)[A-Z]+(\$?)[0-9]+$/, `$1${column}$2${row}`);
        }
        function rename(token: string) {
            const plain = token.replace(/\$/g, '');
            const name = referencesToUpdate[plain] || (shift && shiftName(plain, shift));
            return name ? withName(token, name) : token;
        }
        const end = removed ? removed.index + removed.count : 0;
        const axis = removed?.axis === 1 ? 1 : 0;
        function isRemoved(cell: number[]) {
            return !!removed && cell[axis] >= removed.index && cell[axis] < end;
        }
        let named = false;
        const text = formula.replace(
            // A string literal, or a worksheet name with the cell or the range after its !, stays as it is
            /("(?:[^"]|"")*"|(?:'(?:[^']|'')*'|[A-Za-z0-9_.]+)!(?:\$?[A-Z]+\$?[0-9]+(?::\$?[A-Z]+\$?[0-9]+)?)?)|(?<![A-Za-z0-9_.!$])(\$?[A-Z]+\$?[0-9]+)(?::(\$?[A-Z]+\$?[0-9]+))?(?![A-Za-z0-9_.!(])/g,
            (match, kept: string | undefined, from: string, to: string | undefined) => {
                if (kept) {
                    return kept;
                }
                if (!removed) {
                    return to ? `${rename(from)}:${rename(to)}` : rename(from);
                }
                const fromCell = getIdFromColumnName(from.replace(/\$/g, ''));
                const fromGone = isRemoved(fromCell);
                if (!to) {
                    named ||= fromGone;
                    return fromGone ? '#REF!' : rename(from);
                }
                const toCell = getIdFromColumnName(to.replace(/\$/g, ''));
                const toGone = isRemoved(toCell);
                named ||= Math.min(fromCell[axis], toCell[axis]) < end && Math.max(fromCell[axis], toCell[axis]) >= removed.index;
                if (fromGone && toGone) {
                    return '#REF!';
                }
                // A removed end of a range goes to the first cell after the removed cells, or to the last cell before
                // them, which is in the range
                const first = fromCell[axis] <= toCell[axis];
                function renameEnd(token: string, cell: number[], gone: boolean, low: boolean) {
                    if (!gone) {
                        return rename(token);
                    }
                    cell[axis] = low ? removed!.index : removed!.index - 1;
                    return withName(token, getColumnNameFromId(cell));
                }
                return `${renameEnd(from, fromCell, fromGone, first)}:${renameEnd(to, toCell, toGone, !first)}`;
            }
        );
        return { text, named };
    }

    /**
     * Update the formulas of the cells, and for an insert or a delete the formulas of the footers. With removed, a delete
     * removed rows or columns, and the references to its cells become #REF!.
     * @param keepText - Whether a cell keeps its text. A footer has no row, so it gets no y.
     * @param shift - The insert, the delete, or their undo or redo. Null for a move or a sort.
     * @returns The formulas whose texts the delete changed, at their positions before the delete, with their texts before
     * it.
     */
    private updateFormulas(referencesToUpdate: Record<string, string>, keepText: (x: number, y?: number) => boolean, shift: Shift | null, removed?: Removed) {
        const worksheet = this.worksheet;
        const changed: FormulaText[] = [];
        // The new names of the cells that the undo of a delete moves back, as the insert gives them
        const back = removed ? Object.fromEntries(Object.entries(referencesToUpdate).map(([from, to]) => [to, from])) : undefined;
        const undo = shift && {
            ...shift,
            count: -shift.count,
            rows: shift.rows + (shift.axis === 1 ? shift.count : 0),
            columns: shift.columns + (shift.axis === 0 ? shift.count : 0),
        };

        // Give a formula its new text. A formula that named a removed cell calculates again. The undo of the delete moves
        // the names back, and the history keeps the texts that it does not give back, also of a formula that the delete
        // did not change.
        const update = (value: string, i: number, j: number, footer: boolean) => {
            // A move or a sort renames the cells one by one, so it would cut a range. It keeps the rule of version 5.
            if (!shift) {
                return this.updateFormula(value, referencesToUpdate);
            }
            const result = this.renameReferences(value, referencesToUpdate, removed, shift);
            if (removed && (result.named || this.renameReferences(result.text, back!, undefined, undo!).text !== value)) {
                const moved = (removed.axis === 1 ? j : i) >= removed.index && !(footer && removed.axis === 1) ? removed.count : 0;
                const x = removed.axis === 0 ? i + moved : i;
                const y = removed.axis === 1 ? j + moved : j;
                changed.push(footer ? { x, y, text: value, footer } : { x, y, text: value });
            }
            return result.text;
        };

        // Update formulas. The formula of a cell that an undo or a redo adds already names the cells at their new positions.
        for (let j = 0; j < worksheet.options.data.length; j++) {
            const row = worksheet.options.data[j];
            for (let i = 0; i < worksheet.options.data[0].length; i++) {
                if (keepText(i, j)) {
                    continue;
                }
                const value = `${row[i]}`;
                if (value.substr(0, 1) === '=') {
                    const newFormula = update(value, i, j, false);
                    if (newFormula !== value) {
                        row[i] = newFormula;
                    }
                }
            }
        }

        // The footers name the cells too. setFooter shows their new results. A move or a sort keeps their texts.
        for (const [j, footer] of ((shift && worksheet.options.footers) || []).entries()) {
            for (let i = 0; i < footer.length; i++) {
                const value = `${footer[i]}`;
                if (value.substr(0, 1) === '=' && !keepText(i)) {
                    const newFormula = update(value, i, j, true);
                    if (newFormula !== value) {
                        footer[i] = newFormula;
                    }
                }
            }
        }

        // Update formula chain. The names of the cells outside the table move as the formula texts do, and the names of the
        // cells that the delete removed leave the chain.
        const outside: Record<string, string> = {};
        if (shift) {
            for (const name of Object.keys(this.chain)) {
                const shifted = referencesToUpdate[name] ? undefined : shiftName(name, shift);
                if (shifted) {
                    outside[name] = shifted;
                } else if (shifted === null && removed) {
                    delete this.chain[name];
                }
            }
        }
        worksheet.formulas.rename(Object.keys(outside).length ? { ...referencesToUpdate, ...outside } : referencesToUpdate);

        // A formula that names #REF! in place of a removed cell has another result now, and so have the formulas that
        // name it
        if (changed.length) {
            const names: string[] = [];
            for (const cell of changed) {
                if (cell.footer) {
                    continue;
                }
                const moved = removed!.axis === 1 ? cell.y >= removed!.index : cell.x >= removed!.index;
                const x = removed!.axis === 0 && moved ? cell.x - removed!.count : cell.x;
                const y = removed!.axis === 1 && moved ? cell.y - removed!.count : cell.y;
                if (worksheet.records[y]?.[x]) {
                    worksheet.cells.showValue(x, y);
                    names.push(getColumnNameFromId([x, y]));
                }
            }
            this.showDependents(names);
        }

        return changed;
    }

    /**
     * The insert or the delete of updateTableReferences, with the size of the table before it. Null for a move or a sort.
     */
    private shiftOf(inserted?: { axis: 0 | 1; index: number; count: number }, removed?: Removed): Shift | null {
        const change = inserted || removed;
        if (!change) {
            return null;
        }
        const count = inserted ? change.count : -change.count;
        return {
            axis: change.axis,
            index: change.index,
            count,
            rows: this.worksheet.options.data.length - (change.axis === 1 ? count : 0),
            columns: this.worksheet.headers.length - (change.axis === 0 ? count : 0),
        };
    }

    /**
     * Give formula cells and footers their texts again, for the undo of a delete, and show their results and the results
     * of the formulas that name them.
     */
    restoreTexts(texts: FormulaText[]) {
        const worksheet = this.worksheet;
        const names: string[] = [];
        let footers = false;
        for (const { x, y, text, footer } of texts) {
            if (footer) {
                if (worksheet.options.footers?.[y] && x < worksheet.options.footers[y].length) {
                    worksheet.options.footers[y][x] = text;
                    footers = true;
                }
            } else if (worksheet.options.data[y] && worksheet.records[y]?.[x]) {
                worksheet.options.data[y][x] = text;
                worksheet.cells.showValue(x, y);
                names.push(getColumnNameFromId([x, y]));
            }
        }
        this.showDependents(names);
        if (footers) {
            this.setFooter();
        }
    }

    /**
     * Update the positions of the rows, the columns and the cells, and the references of the formulas, after rows or
     * columns move.
     * @param from - The first row or column that moved. The rows or columns before it keep their positions, so the update
     * starts there. Without it, the update reads each row and each column. `inserted` gives the rows (axis 1) or the
     * columns (axis 0) that an insert, an undo or a redo added: they get their place in the chain after the other cells
     * moved. With `keepTexts`, their formulas keep their texts, else their references move as the other references do.
     * `removed` gives the rows (axis 1) or the columns (axis 0) that a delete removed, at their old positions: the formula
     * cells that it removed leave the chain. With `refErrors`, the references to the removed cells become #REF!.
     * @returns The formula cells whose texts the delete changed, at their positions before the delete, with their texts
     * before it. The undo of the delete gives them back.
     */
    updateTableReferences(
        from: {
            row?: number;
            column?: number;
            inserted?: { axis: 0 | 1; index: number; count: number; keepTexts?: boolean };
            removed?: Removed & { refErrors?: boolean };
        } = {}
    ): FormulaText[] {
        const worksheet = this.worksheet;
        if (worksheet.skipUpdateTableReferences) {
            return [];
        }
        let changed: FormulaText[] = [];

        let firstRow = Number.isInteger(from.row) ? Math.max(0, from.row!) : 0;
        let firstColumn = Number.isInteger(from.column) ? Math.max(0, from.column!) : 0;

        // A merged range that starts before the start and ends on or after it is read from its first cell, because the
        // record of a covered cell is the element of the first cell. A lower start can fall into another range that starts
        // before it, so the ranges are read from the last first row (or column) to the first one, in one pass.
        const ranges = Object.entries(worksheet.options.mergeCells || {}).map(([cellName, [colspan, rowspan]]) => {
            const [x, y] = getIdFromColumnName(cellName);
            return { x, y, colspan, rowspan };
        });
        for (const range of [...ranges].sort((a, b) => b.y - a.y)) {
            if (range.y < firstRow && range.y + range.rowspan > firstRow) {
                firstRow = range.y;
            }
        }
        for (const range of [...ranges].sort((a, b) => b.x - a.x)) {
            if (range.x < firstColumn && range.x + range.colspan > firstColumn) {
                firstColumn = range.x;
            }
        }

        // The new names of the moved cells are only for the formulas, the formula chain and the meta information
        const references = Object.keys(this.chain).length > 0 || Object.keys(worksheet.options.meta || {}).length > 0 || this.hasFormulas();

        // The cells that an insert or an undo added. Their position attributes can be of another position, for example
        // of the cell before a delete, so they give no old name.
        const inserted = from.inserted;
        function isInserted(x: number, y: number) {
            if (!inserted) {
                return false;
            }
            const position = inserted.axis === 1 ? y : x;
            return position >= inserted.index && position < inserted.index + inserted.count;
        }

        // A footer has no row: it keeps its text when the undo of a delete gives back its column
        function keepsText(x: number, y?: number) {
            if (y === undefined) {
                return !!inserted?.keepTexts && inserted.axis === 0 && isInserted(x, 0);
            }
            return !!inserted?.keepTexts && isInserted(x, y);
        }

        // Update headers
        for (let i = firstColumn; i < worksheet.headers.length; i++) {
            const x = parseInt(worksheet.headers[i].getAttribute('data-x'));

            if (x !== i) {
                // Update coords
                worksheet.headers[i].setAttribute('data-x', i);
                worksheet.headers[i].setAttribute('aria-colindex', i + 2);
                // Title
                if (!worksheet.headers[i].getAttribute('title')) {
                    worksheet.headers[i].innerHTML = getColumnName(i);
                }
            }
        }

        // Update all rows
        for (let j = firstRow; j < worksheet.rows.length; j++) {
            if (worksheet.rows[j]) {
                const y = parseInt(worksheet.rows[j].element.getAttribute('data-y'));

                if (y !== j) {
                    // Update coords
                    worksheet.rows[j].element.setAttribute('data-y', j);
                    worksheet.rows[j].element.children[0].setAttribute('data-y', j);
                    worksheet.accessibility.updateRowIndex(worksheet.rows[j].element, j);
                    // Row number
                    worksheet.rows[j].element.children[0].innerHTML = String(j + 1);
                }
            }
        }

        // Regular cells affected by this change: the new cell name by the old one
        const affectedTokens: Record<string, string> = {};
        // The merged cells affected by this change: false, or the new cell name and the column and row offsets
        const mergeCellUpdates: Record<string, false | [string, number, number]> = {};

        // Update cell
        function updatePosition(x: number, y: number, i: number, j: number) {
            if (x !== i) {
                worksheet.records[j][i].element.setAttribute('data-x', i);
                worksheet.records[j][i].element.setAttribute('aria-colindex', i + 2);
            }
            if (y !== j) {
                worksheet.records[j][i].element.setAttribute('data-y', j);
            }

            // Other updates
            if (isInserted(i, j)) {
                return;
            }
            if (references && (x !== i || y !== j)) {
                const columnIdFrom = getColumnNameFromId([x, y]);
                const columnIdTo = getColumnNameFromId([i, j]);
                affectedTokens[columnIdFrom] = columnIdTo;
            }
        }

        for (let j = firstRow; j < worksheet.records.length; j++) {
            for (let i = firstColumn; i < worksheet.records[0].length; i++) {
                if (worksheet.records[j][i]) {
                    // Current values
                    const x = parseInt(worksheet.records[j][i].element.getAttribute('data-x'));
                    const y = parseInt(worksheet.records[j][i].element.getAttribute('data-y'));

                    // Update column
                    if (worksheet.records[j][i].element.getAttribute('data-merged')) {
                        const columnIdFrom = getColumnNameFromId([x, y]);
                        const columnIdTo = getColumnNameFromId([i, j]);
                        if (mergeCellUpdates[columnIdFrom] == null) {
                            if (columnIdFrom === columnIdTo) {
                                mergeCellUpdates[columnIdFrom] = false;
                            } else {
                                const totalX = i - x;
                                const totalY = j - y;
                                mergeCellUpdates[columnIdFrom] = [columnIdTo, totalX, totalY];
                            }
                        }
                    } else {
                        updatePosition(x, y, i, j);
                    }
                }
            }
        }

        // Update the merged ranges. A range can move onto the old name of another range that did not move yet, so all
        // the old names go before the new names are written.
        const moved: [string, MergeInfo][] = [];
        for (const [key, update] of Object.entries(mergeCellUpdates)) {
            // A cell with data-merged that has no range in the options is skipped. Before, the update threw a TypeError.
            if (update && worksheet.options.mergeCells?.[key]) {
                const info = getIdFromColumnName(key);
                let x = info[0];
                let y = info[1];
                updatePosition(x, y, x + update[1], y + update[2]);

                const entry = worksheet.options.mergeCells[key] as MergeInfo;
                for (const cell of entry[2]!) {
                    x = parseInt(cell.getAttribute('data-x'));
                    y = parseInt(cell.getAttribute('data-y'));
                    cell.setAttribute('data-x', x + update[1]);
                    cell.setAttribute('aria-colindex', x + update[1] + 2);
                    cell.setAttribute('data-y', y + update[2]);
                }

                moved.push([update[0], entry]);
                delete worksheet.options.mergeCells[key];
            }
        }
        for (const [cellName, entry] of moved) {
            (worksheet.options.mergeCells as Record<string, MergeInfo>)[cellName] = entry;
        }

        if (references) {
            // A formula cell that a delete removed leaves the chain. Before, a change of a cell that it named threw, or
            // calculated the cell that then had its name.
            const removed = from.removed;
            if (removed) {
                // The old names of the removed cells
                const gone = new Set<string>();
                const other = removed.axis === 1 ? worksheet.headers.length : worksheet.options.data.length;
                for (let k = removed.index; k < removed.index + removed.count; k++) {
                    for (let o = 0; o < other; o++) {
                        gone.add(getColumnNameFromId(removed.axis === 1 ? [o, k] : [k, o]));
                    }
                }
                const chain = this.chain;
                // A removed cell that formulas named leaves the chain when the formulas name #REF! in its place
                if (removed.refErrors && worksheet.parent.config.updateReferences !== false) {
                    for (const name of gone) {
                        delete chain[name];
                    }
                }
                for (const name in chain) {
                    const dependents = chain[name];
                    for (const dependent of dependents) {
                        if (gone.has(dependent)) {
                            chain[name] = dependents.filter((dependent) => !gone.has(dependent));
                            break;
                        }
                    }
                }
            }

            if (worksheet.parent.config.updateReferences === false) {
                // The formula texts stay, so the chain keeps the names that they refer to. The formula cells moved.
                this.rename(affectedTokens, false);
                this.showResults(firstColumn, firstRow);
            } else {
                // Update formulas
                changed = this.updateFormulas(affectedTokens, keepsText, this.shiftOf(inserted, removed), removed?.refErrors ? removed : undefined);
            }

            // The formulas of the new cells calculate again at their positions, and get their place in the chain
            if (inserted) {
                const rows = inserted.axis === 1 ? [inserted.index, inserted.index + inserted.count] : [0, worksheet.options.data.length];
                const columns = inserted.axis === 0 ? [inserted.index, inserted.index + inserted.count] : [0, worksheet.options.data[0]?.length || 0];
                for (let j = rows[0]; j < Math.min(rows[1], worksheet.options.data.length); j++) {
                    for (let i = columns[0]; i < columns[1]; i++) {
                        if (worksheet.records[j]?.[i] && isFormula(worksheet.options.data[j][i])) {
                            worksheet.cells.showValue(i, j);
                        }
                    }
                }
            }

            // Update meta data
            worksheet.meta.update(affectedTokens);
        }

        // Update the size of the grid
        worksheet.accessibility.updateGridSize();

        // The frozen rows are the first rows. This comes after the cells know their new rows, because a row that leaves
        // the table saves and closes its editor.
        worksheet.rowManager.updateFreeze();

        // The tab stop of the grid can be in a deleted row or column
        worksheet.selection.updateTabStop();

        // Refresh selection
        worksheet.selection.refresh();

        // Update table with custom configuration if applicable
        worksheet.data.updateTable();

        return changed;
    }

    /**
     * Set the footers, or show the footers of the options again.
     * @param data - The new footers.
     */
    setFooter(data?: (string | number)[][]) {
        const worksheet = this.worksheet;

        if (data) {
            worksheet.options.footers = data;
        }

        if (worksheet.options.footers) {
            if (!worksheet.tfoot) {
                worksheet.tfoot = document.createElement('tfoot');
                worksheet.table.appendChild(worksheet.tfoot);
            }

            for (let j = 0; j < worksheet.options.footers.length; j++) {
                let tr;

                if (worksheet.tfoot.children[j]) {
                    tr = worksheet.tfoot.children[j];
                } else {
                    tr = document.createElement('tr');
                    const td = document.createElement('td');
                    tr.appendChild(td);
                    worksheet.tfoot.appendChild(tr);
                }
                for (let i = 0; i < worksheet.headers.length; i++) {
                    if (!worksheet.options.footers[j][i]) {
                        worksheet.options.footers[j][i] = '';
                    }

                    let td;

                    if (worksheet.tfoot.children[j].children[i + 1]) {
                        td = worksheet.tfoot.children[j].children[i + 1] as HTMLTableCellElement;
                    } else {
                        td = document.createElement('td');
                        tr.appendChild(td);

                        // Text align
                        const colAlign = worksheet.options.columns[i].align || worksheet.options.defaultColAlign || 'center';
                        td.style.textAlign = colAlign;
                    }
                    // A footer is in no cell, so its formula runs without coordinates. The DOM converts the result to text.
                    const value = worksheet.options.footers[j][i];
                    td.textContent = (
                        `${value}`.substr(0, 1) === '=' && worksheet.parent.config.parseFormulas !== false ? worksheet.formulas.execute(value as string) : value
                    ) as string;

                    // Hide/Show with hideColumn()/showColumn()
                    td.style.display = worksheet.cols[i].colElement.style.display;
                }
                // A delete of columns leaves the cells of the last columns
                while (worksheet.tfoot.children[j].children.length > worksheet.headers.length + 1) {
                    worksheet.tfoot.children[j].lastElementChild!.remove();
                }
            }

            // Update the size of the grid
            worksheet.accessibility.updateGridSize();
        }
    }
}
