/** A run of rows from `first` to `last` (from 0) with one height in pixels, or hidden. */
export interface RowRun {
    first: number;
    last: number;
    height?: number;
    hidden?: boolean;
}

/** A run of rows in the layout: its height is undefined for the default height. */
interface LayoutRun {
    first: number;
    last: number;
    height: number | undefined;
    hidden: boolean;
}

/**
 * Sort the runs by their first row, remove the rows outside the sheet and the runs of rows with the default height that
 * show, and join runs that are next to each other and equal. Of two runs that overlap, the run later in that order wins
 * for its own rows.
 */
function normalRuns(runs: RowRun[], count: number, defaultHeight: number) {
    const sorted: LayoutRun[] = runs
        .map((run) => {
            const height = run.height! > 0 && run.height !== defaultHeight ? run.height : undefined;
            return { first: Math.max(0, Math.floor(run.first) || 0), last: Math.min(count - 1, Math.floor(run.last) || 0), height, hidden: !!run.hidden };
        })
        .filter((run) => run.first <= run.last && (run.height !== undefined || run.hidden))
        .sort((a, b) => a.first - b.first);

    // The result is sorted and has no overlaps. A run after the last run of the result keeps the order. A run that
    // overlaps runs of the result cuts each of them, and wins for its own rows. A sort after each run made 50,000 runs
    // of hidden rows take seconds, so only an overlap sorts.
    const result: LayoutRun[] = [];
    for (const run of sorted) {
        const last = result[result.length - 1];
        if (!last || run.first > last.last) {
            if (last && last.last + 1 === run.first && last.height === run.height && last.hidden === run.hidden) {
                last.last = run.last;
            } else {
                result.push(run);
            }
            continue;
        }
        const pieces: LayoutRun[] = [];
        for (const item of result) {
            if (item.last < run.first || item.first > run.last) {
                pieces.push(item);
                continue;
            }
            if (item.first < run.first) {
                pieces.push({ first: item.first, last: run.first - 1, height: item.height, hidden: item.hidden });
            }
            if (item.last > run.last) {
                pieces.push({ first: run.last + 1, last: item.last, height: item.height, hidden: item.hidden });
            }
        }
        pieces.push(run);
        pieces.sort((x, y) => x.first - y.first);
        result.length = 0;
        for (const piece of pieces) {
            const previous = result[result.length - 1];
            if (previous && previous.last + 1 === piece.first && previous.height === piece.height && previous.hidden === piece.hidden) {
                previous.last = piece.last;
            } else {
                result.push(piece);
            }
        }
    }
    return result;
}

/**
 * The heights and the hidden rows of a sheet, as sorted runs of equal rows. It gives the top of a row and the row at a
 * pixel without a loop over all rows. Rows are from 0, heights are in pixels.
 */
export class RowLayout {
    readonly count: number;

    /** The default height that the application gives. */
    readonly givenHeight: number;

    /** The smallest height of a row in the browser, or 0. A lower height shows as this height. */
    minHeight: number;

    /** The height of a row without a run. */
    defaultHeight: number;

    /** The rows with another height than the default height, or hidden. */
    runs: LayoutRun[];

    /** The top of the first row of each run. */
    private tops: number[] = [];

    /** The height of all rows. */
    total = 0;

    constructor(count: number, defaultHeight: number, runs: RowRun[] = [], minHeight = 0) {
        this.count = Math.max(0, Math.floor(count) || 0);
        this.givenHeight = defaultHeight > 0 ? defaultHeight : 20;
        this.minHeight = minHeight > 0 ? minHeight : 0;
        this.defaultHeight = Math.max(this.givenHeight, this.minHeight);
        this.runs = normalRuns(runs, this.count, this.givenHeight);
        this.index();
    }

    /** Set the smallest height of a row in the browser, for example from the padding of the cells. */
    setMinHeight(height: number) {
        this.minHeight = height > 0 ? height : 0;
        this.defaultHeight = Math.max(this.givenHeight, this.minHeight);
        this.index();
    }

    private index() {
        this.tops = new Array(this.runs.length);
        let top = 0;
        let next = 0;
        for (let i = 0; i < this.runs.length; i++) {
            const run = this.runs[i];
            top += (run.first - next) * this.defaultHeight;
            this.tops[i] = top;
            top += (run.last - run.first + 1) * this.runHeight(run);
            next = run.last + 1;
        }
        this.total = top + (this.count - next) * this.defaultHeight;
    }

    /** The index of the last run whose first row is at most `y`, or -1. */
    private runIndex(y: number) {
        let low = 0;
        let high = this.runs.length - 1;
        let found = -1;
        while (low <= high) {
            const middle = (low + high) >> 1;
            if (this.runs[middle].first <= y) {
                found = middle;
                low = middle + 1;
            } else {
                high = middle - 1;
            }
        }
        return found;
    }

    /** The height of each row of a run: 0 for hidden rows. */
    private runHeight(run: LayoutRun) {
        if (run.hidden) {
            return 0;
        }
        return Math.max(this.minHeight, run.height! > 0 ? run.height! : this.givenHeight);
    }

    /** The run of row `y`, or null for a row with the default height that shows. */
    runOf(y: number) {
        const i = this.runIndex(y);
        return i >= 0 && y <= this.runs[i].last ? this.runs[i] : null;
    }

    /** The height of row `y`: 0 for a hidden row. */
    heightOf(y: number) {
        const run = this.runOf(y);
        return run ? this.runHeight(run) : this.defaultHeight;
    }

    /** The height that row `y` has when it shows, also when it is hidden. */
    shownHeight(y: number) {
        const run = this.runOf(y);
        return run && run.height! > 0 ? Math.max(this.minHeight, run.height!) : this.defaultHeight;
    }

    isHidden(y: number) {
        return !!this.runOf(y)?.hidden;
    }

    /** The top of row `y`. `offsetOf(count)` is the height of all rows. */
    offsetOf(y: number) {
        if (y <= 0) {
            return 0;
        }
        if (y >= this.count) {
            return this.total;
        }
        const i = this.runIndex(y);
        if (i < 0) {
            return y * this.defaultHeight;
        }
        const run = this.runs[i];
        if (y <= run.last) {
            return this.tops[i] + (y - run.first) * this.runHeight(run);
        }
        return this.tops[i] + (run.last - run.first + 1) * this.runHeight(run) + (y - run.last - 1) * this.defaultHeight;
    }

    /** The row whose box holds the pixel `offset`. A hidden row holds no pixel. */
    rowAt(offset: number) {
        if (this.count === 0) {
            return 0;
        }
        if (offset <= 0) {
            return this.nextShown(0, 1);
        }
        if (offset >= this.total) {
            return this.nextShown(this.count - 1, -1);
        }
        // The last run whose top is at most the offset
        let low = 0;
        let high = this.runs.length - 1;
        let i = -1;
        while (low <= high) {
            const middle = (low + high) >> 1;
            if (this.tops[middle] <= offset) {
                i = middle;
                low = middle + 1;
            } else {
                high = middle - 1;
            }
        }
        if (i < 0) {
            return Math.min(this.count - 1, Math.floor(offset / this.defaultHeight));
        }
        const run = this.runs[i];
        const height = this.runHeight(run);
        const end = this.tops[i] + (run.last - run.first + 1) * height;
        if (offset < end && height > 0) {
            return run.first + Math.floor((offset - this.tops[i]) / height);
        }
        return Math.min(this.count - 1, run.last + 1 + Math.floor((offset - end) / this.defaultHeight));
    }

    /** The first row from `y` in the direction `step` (1 or -1) that shows, or `y` when none does. */
    nextShown(y: number, step: 1 | -1) {
        let row = y;
        while (row >= 0 && row < this.count) {
            const run = this.runOf(row);
            if (!run?.hidden) {
                return row;
            }
            row = step > 0 ? run.last + 1 : run.first - 1;
        }
        return y;
    }

    /**
     * Set the height (pixels, or null for the default height) or the hidden state of the rows `first` to `last`. A
     * property that `values` does not have stays as it is.
     */
    set(first: number, last: number, values: { height?: number | null; hidden?: boolean }) {
        first = Math.max(0, first);
        last = Math.min(this.count - 1, last);
        if (first > last) {
            return;
        }
        const pieces: RowRun[] = [];
        let row = first;
        while (row <= last) {
            const run = this.runOf(row);
            const end = run ? Math.min(run.last, last) : Math.min(last, this.nextRunFirst(row) - 1);
            const height = 'height' in values ? (values.height ?? undefined) : run?.height;
            const hidden = 'hidden' in values ? !!values.hidden : !!run?.hidden;
            pieces.push({ first: row, last: end, height, hidden });
            row = end + 1;
        }
        const kept: RowRun[] = [];
        for (const run of this.runs) {
            if (run.last < first || run.first > last) {
                kept.push(run);
                continue;
            }
            if (run.first < first) {
                kept.push({ first: run.first, last: first - 1, height: run.height, hidden: run.hidden });
            }
            if (run.last > last) {
                kept.push({ first: last + 1, last: run.last, height: run.height, hidden: run.hidden });
            }
        }
        this.runs = normalRuns(kept.concat(pieces), this.count, this.givenHeight);
        this.index();
    }

    /** The first row of the first run after row `y`, or the row count. */
    private nextRunFirst(y: number) {
        const next = this.runs[this.runIndex(y) + 1];
        return next ? next.first : this.count;
    }
}
