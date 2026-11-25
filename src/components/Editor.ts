/**
 * The cell editor of a worksheet: which cell is being edited, where it is, and its HTML before the editor opened.
 *
 * Each worksheet has one, as worksheet.editor. Opening and closing an editor creates the controls of the column type
 * and writes the value to the cell, so the worksheet methods openEditor and closeEditor (src/utils/editor.ts) do that,
 * and keep the state here.
 */
export class Editor {
    /** The cell that is being edited, or null when no editor is open. */
    cell: HTMLTableCellElement | null = null;

    /** The HTML of the cell before the editor opened. Closing restores it when the value does not change. */
    html = '';

    /** The column of the cell, or -1 when no editor is open. */
    x = -1;

    /** The row of the cell, or -1 when no editor is open. */
    y = -1;

    /** Whether an editor is open. */
    isOpen(): this is Editor & { cell: HTMLTableCellElement } {
        return this.cell !== null;
    }

    /** Keep the cell that an editor opens on. */
    start(cell: HTMLTableCellElement, x: number, y: number) {
        this.cell = cell;
        this.html = cell.innerHTML;
        this.x = x;
        this.y = y;
    }

    /** Forget the cell when its editor closes, or when the column type has no editor to keep open. */
    end() {
        this.cell = null;
        this.html = '';
        this.x = -1;
        this.y = -1;
    }
}
