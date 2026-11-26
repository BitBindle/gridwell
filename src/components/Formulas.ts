import libraryBase from '../utils/libraryBase.js';
import { getColumnName } from '../utils/helpers.js';
import { getColumnNameFromId, getIdFromColumnName } from '../utils/internalHelpers.js';
import type { HistoryCellChange, MergeInfo, Worksheet } from '../types/internal';

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

    /** Keep that the formula of the cell `dependent` refers to the cell `cell`. Both are cell names, such as A1. */
    add(cell: string, dependent: string) {
        if (!this.chain[cell]) {
            this.chain[cell] = [];
        }

        if (!this.chain[cell].includes(dependent)) {
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
            const tokens = expression.replace(/"(?:[^"]|"")*"/g, '""').match(/([A-Z]+[0-9]+)/g);

            // Direct self-reference protection
            if (parentId && tokens?.includes(parentId)) {
                console.error('Self Reference detected');
                return '#ERROR';
            } else {
                // Expressions to be used in the parsing
                const formulaExpressions: Record<string, unknown> = {};

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

                const ret = worksheet.dispatch('onbeforeformula', worksheet, expression, x, y);
                if (ret === false) {
                    return expression;
                } else if (ret) {
                    expression = ret;
                    source = ret;
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
     * Update formulas
     */
    private updateFormulas(referencesToUpdate: Record<string, string>) {
        const worksheet = this.worksheet;

        // Update formulas
        for (const row of worksheet.options.data) {
            for (let i = 0; i < worksheet.options.data[0].length; i++) {
                const value = `${row[i]}`;
                // Is formula
                if (value.substr(0, 1) === '=') {
                    // Replace tokens
                    const newFormula = this.updateFormula(value, referencesToUpdate);
                    if (newFormula !== value) {
                        row[i] = newFormula;
                    }
                }
            }
        }

        // Update formula chain
        worksheet.formulas.rename(referencesToUpdate);
    }

    /**
     * Update the references of the formulas after rows or columns move.
     */
    updateTableReferences() {
        const worksheet = this.worksheet;
        if (worksheet.skipUpdateTableReferences) {
            return;
        }

        // Update headers
        for (let i = 0; i < worksheet.headers.length; i++) {
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
        for (let j = 0; j < worksheet.rows.length; j++) {
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
            if (x !== i || y !== j) {
                const columnIdFrom = getColumnNameFromId([x, y]);
                const columnIdTo = getColumnNameFromId([i, j]);
                affectedTokens[columnIdFrom] = columnIdTo;
            }
        }

        for (let j = 0; j < worksheet.records.length; j++) {
            for (let i = 0; i < worksheet.records[0].length; i++) {
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

        // Update merged if applicable
        const keys = Object.keys(mergeCellUpdates);
        if (keys.length) {
            for (const key of keys) {
                const update = mergeCellUpdates[key];
                if (update) {
                    const info = getIdFromColumnName(key);
                    let x = info[0];
                    let y = info[1];
                    updatePosition(x, y, x + update[1], y + update[2]);

                    const columnIdFrom = key;
                    const columnIdTo = update[0];
                    for (const cell of (worksheet.options.mergeCells![columnIdFrom] as MergeInfo)[2]!) {
                        x = parseInt(cell.getAttribute('data-x'));
                        y = parseInt(cell.getAttribute('data-y'));
                        cell.setAttribute('data-x', x + update[1]);
                        cell.setAttribute('aria-colindex', x + update[1] + 2);
                        cell.setAttribute('data-y', y + update[2]);
                    }

                    worksheet.options.mergeCells![columnIdTo] = worksheet.options.mergeCells![columnIdFrom];
                    delete worksheet.options.mergeCells![columnIdFrom];
                }
            }
        }

        // Update formulas
        this.updateFormulas(affectedTokens);

        // Update meta data
        worksheet.meta.update(affectedTokens);

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
            }

            // Update the size of the grid
            worksheet.accessibility.updateGridSize();
        }
    }
}
