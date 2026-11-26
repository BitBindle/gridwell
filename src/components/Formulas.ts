import formula from '@jspreadsheet/formula';
import { getColumnName } from '../utils/helpers.js';
import { getColumnNameFromId, getIdFromColumnName } from '../utils/internalHelpers.js';
import type { HistoryCellChange, MergeInfo, Worksheet } from '../types/internal';

// The cells in progress of updateChain, by cell name
let chainLoopProtection: Record<string, boolean> = {};

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

    /**
     * Trying to extract a number from a string
     */
    private parseNumber(value: unknown, columnNumber: number) {
        const worksheet = this.worksheet;

        // Decimal point. Column 0 has one too.
        const column = worksheet.options.columns[columnNumber];
        const decimal = column && column.decimal ? column.decimal : '.';

        // Parse both parts of the number
        // Text, then its parts, then the digits of the integer part, then the digits as text
        let number: string | (string | string[] | null)[] = '' + value;
        number = number.split(decimal);
        number[0] = (number[0] as string).match(/[+-]?[0-9]/g);
        if (number[0]) {
            number[0] = (number[0] as string[]).join('');
        }
        if (number[1]) {
            // The pattern also matches an empty text, so there is a match
            number[1] = (number[1] as string).match(/[0-9]*/g)!.join('');
        }

        // Is a valid number
        if (number[0] && Number.isInteger(Number(number[0]))) {
            if (!number[1]) {
                value = Number(number[0] + '.00');
            } else {
                value = Number(number[0] + '.' + number[1]);
            }
        } else {
            value = null;
        }

        return value;
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

            // Parent column identification. executeFormula without x and y runs a formula that is in no cell.
            const parentId = getColumnNameFromId([x!, y!]);

            // Code protection
            if (formulaLoopProtection[parentId]) {
                console.error('Reference loop detected');
                return '#ERROR';
            }

            formulaLoopProtection[parentId] = true;

            // Convert range tokens
            function tokensUpdate(tokens: string[]) {
                for (let index = 0; index < tokens.length; index++) {
                    const f = [];
                    const token = tokens[index].split(':');
                    const e1 = getIdFromColumnName(token[0], true);
                    const e2 = getIdFromColumnName(token[1], true);

                    let x1, x2;

                    if (e1[0] <= e2[0]) {
                        x1 = e1[0];
                        x2 = e2[0];
                    } else {
                        x1 = e2[0];
                        x2 = e1[0];
                    }

                    let y1, y2;

                    if (e1[1] <= e2[1]) {
                        y1 = e1[1];
                        y2 = e2[1];
                    } else {
                        y1 = e2[1];
                        y2 = e1[1];
                    }

                    for (let j = y1; j <= y2; j++) {
                        for (let i = x1; i <= x2; i++) {
                            f.push(getColumnNameFromId([i, j]));
                        }
                    }

                    expression = expression.replace(tokens[index], f.join(','));
                }
            }

            // Range with $ remove $ (string literals are kept untouched)
            expression = expression.replace(/("(?:[^"]|"")*")|\$?([A-Z]+)\$?([0-9]+)/g, (match, literal, column, row) => {
                return literal ? literal : column + row;
            });

            let tokens = expression.match(/([A-Z]+[0-9]+):([A-Z]+[0-9]+)/g);
            if (tokens && tokens.length) {
                tokensUpdate(tokens);
            }

            // Get tokens
            tokens = expression.match(/([A-Z]+[0-9]+)/g);

            // Direct self-reference protection
            if (tokens && tokens.indexOf(parentId) > -1) {
                console.error('Self Reference detected');
                return '#ERROR';
            } else {
                // Expressions to be used in the parsing
                const formulaExpressions: Record<string, unknown> = {};

                if (tokens) {
                    for (let i = 0; i < tokens.length; i++) {
                        // Keep chain
                        worksheet.formulas.add(tokens[i], parentId);

                        // Do not calculate again. The formula engine reads the values of the references itself.
                        if (!formulaEngine && typeof (globalThis as unknown as Record<string, unknown>)[tokens[i]] === 'undefined') {
                            // Coords
                            const position = getIdFromColumnName(tokens[i], true);
                            // Get value
                            let value: unknown;

                            if (
                                typeof worksheet.options.data[position[1]] != 'undefined' &&
                                typeof worksheet.options.data[position[1]][position[0]] != 'undefined'
                            ) {
                                value = worksheet.options.data[position[1]][position[0]];
                            } else {
                                value = '';
                            }
                            // Get column data
                            if (('' + value).substr(0, 1) == '=') {
                                if (typeof formulaResults[tokens[i]] !== 'undefined') {
                                    value = formulaResults[tokens[i]];
                                } else {
                                    value = execute(value as string, position[0], position[1]);
                                    formulaResults[tokens[i]] = value;
                                }
                            }
                            // Type!
                            if (('' + value).trim() == '') {
                                // Null
                                formulaExpressions[tokens[i]] = null;
                            } else {
                                if (value == Number(value) && worksheet.parent.config.autoCasting != false) {
                                    // Number
                                    formulaExpressions[tokens[i]] = Number(value);
                                } else {
                                    // Trying any formatted number
                                    const number = worksheet.formulas.parseNumber(value, position[0]);
                                    if (worksheet.parent.config.autoCasting != false && number) {
                                        formulaExpressions[tokens[i]] = number;
                                    } else {
                                        // Text as a string literal, with its quotes and backslashes escaped
                                        formulaExpressions[tokens[i]] = JSON.stringify('' + value);
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
                    const name = '__JSSSTR' + literalIndex++ + '__';
                    formulaExpressions[name] = '"' + literal.slice(1, -1).replace(/""/g, '\\"') + '"';
                    return name;
                });

                // Convert formula to javascript
                let res;

                try {
                    if (formulaEngine) {
                        res = formulaEngine(worksheet, source, x === undefined ? x : Number(x), y === undefined ? y : Number(y));
                    } else {
                        res = formula(parsedExpression, formulaExpressions, x, y, worksheet);
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
                    const cell = getIdFromColumnName(worksheet.formulas.chain[cellId][i], true);
                    // Update cell
                    const value = '' + worksheet.options.data[cell[1]][cell[0]];
                    if (value.substr(0, 1) == '=') {
                        records.push(worksheet.cells.update(cell[0], cell[1], value, true));
                    } else {
                        // No longer a formula, remove from the chain
                        // This changes nothing: Object.keys returns a new array (issue #43)
                        Object.keys(worksheet.formulas.chain)[i] = null!;
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
        for (let j = 0; j < worksheet.options.data.length; j++) {
            for (let i = 0; i < worksheet.options.data[0].length; i++) {
                const value = '' + worksheet.options.data[j][i];
                // Is formula
                if (value.substr(0, 1) == '=') {
                    // Replace tokens
                    const newFormula = this.updateFormula(value, referencesToUpdate);
                    if (newFormula != value) {
                        worksheet.options.data[j][i] = newFormula;
                    }
                }
            }
        }

        // Update formula chain
        worksheet.formulas.rename(referencesToUpdate);
    }

    /**
     * Update cell references
     *
     * @return void
     */
    updateTableReferences() {
        const worksheet = this.worksheet;
        if (worksheet.skipUpdateTableReferences) {
            return;
        }

        // Update headers
        for (let i = 0; i < worksheet.headers.length; i++) {
            const x = parseInt(worksheet.headers[i].getAttribute('data-x'));

            if (x != i) {
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

                if (y != j) {
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
            if (x != i) {
                worksheet.records[j][i].element.setAttribute('data-x', i);
                worksheet.records[j][i].element.setAttribute('aria-colindex', i + 2);
            }
            if (y != j) {
                worksheet.records[j][i].element.setAttribute('data-y', j);
            }

            // Other updates
            if (x != i || y != j) {
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
                            if (columnIdFrom == columnIdTo) {
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
            for (let i = 0; i < keys.length; i++) {
                const update = mergeCellUpdates[keys[i]];
                if (update) {
                    const info = getIdFromColumnName(keys[i], true);
                    let x = info[0];
                    let y = info[1];
                    updatePosition(x, y, x + update[1], y + update[2]);

                    const columnIdFrom = keys[i];
                    const columnIdTo = update[0];
                    for (let j = 0; j < (worksheet.options.mergeCells![columnIdFrom] as MergeInfo)[2]!.length; j++) {
                        x = parseInt((worksheet.options.mergeCells![columnIdFrom] as MergeInfo)[2]![j].getAttribute('data-x'));
                        y = parseInt((worksheet.options.mergeCells![columnIdFrom] as MergeInfo)[2]![j].getAttribute('data-y'));
                        (worksheet.options.mergeCells![columnIdFrom] as MergeInfo)[2]![j].setAttribute('data-x', x + update[1]);
                        (worksheet.options.mergeCells![columnIdFrom] as MergeInfo)[2]![j].setAttribute('aria-colindex', x + update[1] + 2);
                        (worksheet.options.mergeCells![columnIdFrom] as MergeInfo)[2]![j].setAttribute('data-y', y + update[2]);
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
    setFooter(data?: string[][]) {
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
                    // The DOM converts a number or a boolean to text
                    td.textContent = worksheet.cells.parseValue(+worksheet.records.length + i, j, worksheet.options.footers[j][i]) as string;

                    // Hide/Show with hideColumn()/showColumn()
                    td.style.display = worksheet.cols[i].colElement.style.display;
                }
            }

            // Update the size of the grid
            worksheet.accessibility.updateGridSize();
        }
    }
}
