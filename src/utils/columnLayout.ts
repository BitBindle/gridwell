/**
 * The widths of the columns of a worksheet as prefix sums. A hidden column has the width 0. Columns are from 0, widths
 * are in grid pixels.
 */
export class ColumnLayout {
    readonly count: number;

    /** The width of all columns. */
    readonly total: number;

    private readonly offsets: Float64Array;

    constructor(widths: number[]) {
        this.count = widths.length;
        this.offsets = new Float64Array(this.count + 1);
        for (let x = 0; x < this.count; x++) {
            this.offsets[x + 1] = this.offsets[x] + Math.max(0, widths[x] || 0);
        }
        this.total = this.offsets[this.count];
    }

    /** The left edge of column x. offsetOf(count) is the width of all columns. */
    offsetOf(x: number) {
        return this.offsets[Math.max(0, Math.min(this.count, x))];
    }

    /** The column whose box holds the pixel at offset. A hidden column holds no pixel. */
    columnAt(offset: number) {
        // The last column whose left edge is at or before the offset. A hidden column before it has the same edge. An
        // offset after the last column gives the last column, and a worksheet without columns gives 0.
        let low = 0;
        let high = this.count - 1;
        let found = 0;
        while (low <= high) {
            const middle = (low + high) >> 1;
            if (this.offsets[middle] <= offset) {
                found = middle;
                low = middle + 1;
            } else {
                high = middle - 1;
            }
        }
        return found;
    }
}
