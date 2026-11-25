import type { WorksheetHistoryRecord } from '../types/internal';

/**
 * The undo history of a worksheet: the records of the changes, the position of undo and redo in them, and the flag that
 * stops new records while a change must not be recorded.
 *
 * Each worksheet has one, as worksheet.history. Undoing and redoing a record changes the worksheet, so the worksheet
 * methods undo and redo (src/utils/history.ts) do that, and ask the history for the record.
 */
export class History {
    /** The records, oldest first. */
    records: WorksheetHistoryRecord[] = [];

    /** The index of the record that undo reverts next, or -1 when there is nothing to undo. */
    index = -1;

    /** While true, add keeps no records. Undo and redo set it while they apply a record. */
    ignore = false;

    /**
     * Keep a record of a change, unless the history ignores changes. The records that were undone are dropped, so they
     * can no longer be redone.
     */
    add(record: WorksheetHistoryRecord) {
        if (this.ignore != true) {
            // Increment and get the current history index
            const index = ++this.index;

            // Slice the array to discard undone changes
            this.records = this.records.slice(0, index + 1);

            // Keep history
            this.records[index] = record;
        }
    }

    /** Drop all the records. */
    clear() {
        this.records = [];
        this.index = -1;
    }

    /** Whether there is a record to undo. */
    canUndo() {
        return this.index >= 0;
    }

    /** Whether there is a record to redo. */
    canRedo() {
        return this.index < this.records.length - 1;
    }

    /** The record to undo, or undefined when there is none. The next record to undo is the one before it. */
    back(): WorksheetHistoryRecord | undefined {
        if (this.canUndo()) {
            return this.records[this.index--];
        }
        return undefined;
    }

    /** The record to redo, or undefined when there is none. The next record to redo is the one after it. */
    forward(): WorksheetHistoryRecord | undefined {
        if (this.canRedo()) {
            return this.records[++this.index];
        }
        return undefined;
    }
}
