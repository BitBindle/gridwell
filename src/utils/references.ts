import { refreshSelection } from './selection.js';
import { getColumnName } from './helpers.js';
import { updateMeta } from './meta.js';
import { getColumnNameFromId, getIdFromColumnName } from './internalHelpers.js';
import { updateGridSize, updateRowIndex, updateTabStop } from './accessibility.js';
import { updateCell } from './render.js';
import { updateTable } from './table.js';
import type { HistoryCellChange, MergeInfo, Worksheet } from '../types/internal';

/**
 * The formula references: the cells to calculate again after a cell changes, and the references to rename after
 * columns or rows move, are inserted or are deleted.
 */
/**
 * Update all related cells in the chain
 */
// The cells in progress, by cell name
let chainLoopProtection: Record<string, boolean> = {};

export const updateFormulaChain = function (this: Worksheet, x: number, y: number, records: HistoryCellChange[]) {
    const obj = this;

    const cellId = getColumnNameFromId([x, y]);
    if (obj.formulas.chain[cellId] && obj.formulas.chain[cellId].length > 0) {
        if (chainLoopProtection[cellId]) {
            obj.records[y][x].element.innerHTML = '#ERROR';
            // The chain ends here
            obj.formulas.chain[cellId] = [];
        } else {
            // Protection
            chainLoopProtection[cellId] = true;

            for (let i = 0; i < obj.formulas.chain[cellId].length; i++) {
                const cell = getIdFromColumnName(obj.formulas.chain[cellId][i], true);
                // Update cell
                const value = '' + obj.options.data[cell[1]][cell[0]];
                if (value.substr(0, 1) == '=') {
                    records.push(updateCell.call(obj, cell[0], cell[1], value, true));
                } else {
                    // No longer a formula, remove from the chain
                    // This changes nothing: Object.keys returns a new array (issue #43)
                    Object.keys(obj.formulas.chain)[i] = null!;
                }
                updateFormulaChain.call(obj, cell[0], cell[1], records);
            }
        }
    }

    chainLoopProtection = {};
};

/**
 * Update formula
 */
export const updateFormula = function (formula: string, referencesToUpdate: Record<string, string>) {
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
};

/**
 * Update formulas
 */
const updateFormulas = function (this: Worksheet, referencesToUpdate: Record<string, string>) {
    const obj = this;

    // Update formulas
    for (let j = 0; j < obj.options.data.length; j++) {
        for (let i = 0; i < obj.options.data[0].length; i++) {
            const value = '' + obj.options.data[j][i];
            // Is formula
            if (value.substr(0, 1) == '=') {
                // Replace tokens
                const newFormula = updateFormula(value, referencesToUpdate);
                if (newFormula != value) {
                    obj.options.data[j][i] = newFormula;
                }
            }
        }
    }

    // Update formula chain
    obj.formulas.rename(referencesToUpdate);
};

/**
 * Update cell references
 *
 * @return void
 */
export const updateTableReferences = function (this: Worksheet) {
    const obj = this;
    if (obj.skipUpdateTableReferences) {
        return;
    }

    // Update headers
    for (let i = 0; i < obj.headers.length; i++) {
        const x = parseInt(obj.headers[i].getAttribute('data-x'));

        if (x != i) {
            // Update coords
            obj.headers[i].setAttribute('data-x', i);
            obj.headers[i].setAttribute('aria-colindex', i + 2);
            // Title
            if (!obj.headers[i].getAttribute('title')) {
                obj.headers[i].innerHTML = getColumnName(i);
            }
        }
    }

    // Update all rows
    for (let j = 0; j < obj.rows.length; j++) {
        if (obj.rows[j]) {
            const y = parseInt(obj.rows[j].element.getAttribute('data-y'));

            if (y != j) {
                // Update coords
                obj.rows[j].element.setAttribute('data-y', j);
                obj.rows[j].element.children[0].setAttribute('data-y', j);
                updateRowIndex.call(obj, obj.rows[j].element, j);
                // Row number
                obj.rows[j].element.children[0].innerHTML = String(j + 1);
            }
        }
    }

    // Regular cells affected by this change: the new cell name by the old one
    const affectedTokens: Record<string, string> = {};
    // The merged cells affected by this change: false, or the new cell name and the column and row offsets
    const mergeCellUpdates: Record<string, false | [string, number, number]> = {};

    // Update cell
    const updatePosition = function (x: number, y: number, i: number, j: number) {
        if (x != i) {
            obj.records[j][i].element.setAttribute('data-x', i);
            obj.records[j][i].element.setAttribute('aria-colindex', i + 2);
        }
        if (y != j) {
            obj.records[j][i].element.setAttribute('data-y', j);
        }

        // Other updates
        if (x != i || y != j) {
            const columnIdFrom = getColumnNameFromId([x, y]);
            const columnIdTo = getColumnNameFromId([i, j]);
            affectedTokens[columnIdFrom] = columnIdTo;
        }
    };

    for (let j = 0; j < obj.records.length; j++) {
        for (let i = 0; i < obj.records[0].length; i++) {
            if (obj.records[j][i]) {
                // Current values
                const x = parseInt(obj.records[j][i].element.getAttribute('data-x'));
                const y = parseInt(obj.records[j][i].element.getAttribute('data-y'));

                // Update column
                if (obj.records[j][i].element.getAttribute('data-merged')) {
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
                for (let j = 0; j < (obj.options.mergeCells![columnIdFrom] as MergeInfo)[2]!.length; j++) {
                    x = parseInt((obj.options.mergeCells![columnIdFrom] as MergeInfo)[2]![j].getAttribute('data-x'));
                    y = parseInt((obj.options.mergeCells![columnIdFrom] as MergeInfo)[2]![j].getAttribute('data-y'));
                    (obj.options.mergeCells![columnIdFrom] as MergeInfo)[2]![j].setAttribute('data-x', x + update[1]);
                    (obj.options.mergeCells![columnIdFrom] as MergeInfo)[2]![j].setAttribute('aria-colindex', x + update[1] + 2);
                    (obj.options.mergeCells![columnIdFrom] as MergeInfo)[2]![j].setAttribute('data-y', y + update[2]);
                }

                obj.options.mergeCells![columnIdTo] = obj.options.mergeCells![columnIdFrom];
                delete obj.options.mergeCells![columnIdFrom];
            }
        }
    }

    // Update formulas
    updateFormulas.call(obj, affectedTokens);

    // Update meta data
    updateMeta.call(obj, affectedTokens);

    // Update the size of the grid
    updateGridSize.call(obj);

    // The tab stop of the grid can be in a deleted row or column
    updateTabStop.call(obj);

    // Refresh selection
    refreshSelection.call(obj);

    // Update table with custom configuration if applicable
    updateTable.call(obj);
};
