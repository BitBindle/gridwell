import { expect } from 'chai';

import jspreadsheet, { type Worksheet } from './support/jspreadsheet';

function rows(count: number) {
    const data: string[][] = [];
    for (let j = 0; j < count; j++) {
        data.push([`${j}`, `b${j}`]);
    }
    return data;
}

function rect(left: number, top: number, width: number, height: number) {
    // The rectangle has no toJSON, which the code does not call
    return () => ({ left, top, width, height, right: left + width, bottom: top + height, x: left, y: top }) as DOMRect;
}

// The rows are placed at the end of the task that changed them
function placed() {
    return new Promise((resolve) => {
        queueMicrotask(() => resolve(undefined));
    });
}

// The spreadsheet joins jspreadsheet.spreadsheet, which the scroll of the page reads, after it builds its worksheets
function built() {
    return new Promise((resolve) => {
        setTimeout(resolve, 10);
    });
}

// jsdom has no layout. Each row, also each header row, is 20 pixels high, and a hidden row has no height.
function withRowHeight(test: () => void | Promise<void>) {
    return async () => {
        const descriptor = Object.getOwnPropertyDescriptor(window.HTMLElement.prototype, 'offsetHeight')!;
        Object.defineProperty(window.HTMLTableRowElement.prototype, 'offsetHeight', {
            configurable: true,
            get() {
                return (this as HTMLElement).style.display === 'none' ? 0 : 20;
            },
        });
        try {
            await test();
        } finally {
            delete (window.HTMLTableRowElement.prototype as { offsetHeight?: number }).offsetHeight;
            Object.defineProperty(window.HTMLElement.prototype, 'offsetHeight', descriptor);
        }
    };
}

function frozen(worksheet: Worksheet) {
    return worksheet.rows
        .map((row, y) => ({ y, row: row.element }))
        .filter(({ row }) => row.classList.contains('jss_frozen_row'))
        .map(({ y, row }) => [y, row.style.getPropertyValue('--jss-frozen-top'), row.classList.contains('jss_frozen_last')]);
}

describe('Frozen rows', () => {
    it(
        'puts the first rows below the header row, and draws the line under the last one',
        withRowHeight(async () => {
            const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(5), freezeRows: 2 }] });
            await placed();

            expect(frozen(worksheet)).to.eql([
                [0, '20px', false],
                [1, '40px', true],
            ]);
        })
    );

    it(
        'freezes the first rows again after an insert, a move, a delete, a sort and an undo',
        withRowHeight(() => {
            const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(5), freezeRows: 2 }] });
            const first = worksheet.rows[0].element;
            const second = worksheet.rows[1].element;

            worksheet.insertRow(1, 0, true);
            expect(frozen(worksheet).map(([y]) => y)).to.eql([0, 1]);
            expect(second.classList.contains('jss_frozen_row')).to.equal(false);
            expect(second.style.getPropertyValue('--jss-frozen-top')).to.equal('');

            worksheet.undo();
            expect(worksheet.rows[1].element).to.equal(second);
            expect(frozen(worksheet).map(([y]) => y)).to.eql([0, 1]);

            worksheet.moveRow(0, 3);
            expect(first.classList.contains('jss_frozen_row')).to.equal(false);
            expect(frozen(worksheet).map(([y]) => y)).to.eql([0, 1]);

            worksheet.deleteRow(0, 1);
            expect(frozen(worksheet).map(([y]) => y)).to.eql([0, 1]);

            worksheet.orderBy(1, 1);
            expect(frozen(worksheet).map(([y]) => y)).to.eql([0, 1]);
            expect(worksheet.rows.filter((row) => row.element.classList.contains('jss_frozen_row')).length).to.equal(2);
        })
    );

    it(
        'moves the line and the rows after a frozen row that is hidden',
        withRowHeight(async () => {
            const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(5), freezeRows: 3 }] });

            worksheet.hideRow(1);
            await placed();
            expect(frozen(worksheet)).to.eql([
                [0, '20px', false],
                [1, '40px', false],
                [2, '40px', true],
            ]);

            worksheet.hideRow(2);
            await placed();
            expect(frozen(worksheet).map(([y, , last]) => [y, last])).to.eql([
                [0, true],
                [1, false],
                [2, false],
            ]);

            worksheet.showRow([1, 2]);
            await placed();
            expect(frozen(worksheet)).to.eql([
                [0, '20px', false],
                [1, '40px', false],
                [2, '60px', true],
            ]);
        })
    );

    it('freezes all rows when freezeRows is more than the rows, and the new rows of setData', () => {
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(2), freezeRows: 4 }] });
        expect(frozen(worksheet).map(([y]) => y)).to.eql([0, 1]);

        worksheet.setData(rows(6));
        expect(frozen(worksheet).map(([y]) => y)).to.eql([0, 1, 2, 3]);
    });

    it('places the rows again when the size observer sees a change, and destroy stops the observer', async () => {
        const observed: Element[] = [];
        let disconnects = 0;
        let callback: (() => void) | null = null;
        const original = (globalThis as { ResizeObserver?: unknown }).ResizeObserver;
        (globalThis as { ResizeObserver?: unknown }).ResizeObserver = class {
            constructor(cb: () => void) {
                callback = cb;
            }
            observe(element: Element) {
                observed.push(element);
            }
            disconnect() {
                disconnects++;
                observed.length = 0;
            }
        };

        const element = document.body.appendChild(document.createElement('div'));
        try {
            const [worksheet] = jspreadsheet(element, { worksheets: [{ data: rows(5), freezeRows: 2 }] });
            await new Promise((resolve) => setTimeout(resolve));

            expect(observed).to.eql([worksheet.thead, worksheet.rows[0].element, worksheet.rows[1].element]);

            // A web font makes the rows 30 pixels high
            Object.defineProperty(window.HTMLTableRowElement.prototype, 'offsetHeight', { configurable: true, get: () => 30 });
            try {
                callback!();
            } finally {
                delete (window.HTMLTableRowElement.prototype as { offsetHeight?: number }).offsetHeight;
            }
            expect(frozen(worksheet).map(([, top]) => top)).to.eql(['30px', '60px']);

            const before = disconnects;
            jspreadsheet.destroy(element as never);
            expect(disconnects).to.equal(before + 1);
            expect(observed).to.eql([]);
        } finally {
            (globalThis as { ResizeObserver?: unknown }).ResizeObserver = original;
            element.remove();
        }
    });
});

describe('Frozen rows with lazy loading', () => {
    it('keeps the frozen rows first in the table when the keyboard loads other pages', () => {
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(300), freezeRows: 2, lazyLoading: true, tableOverflow: true }] });
        const attached = () => Array.from(worksheet.tbody.children).map((row) => Number(row.getAttribute('data-y')));

        // A jump loads the last page
        worksheet.updateSelectionFromCoords(0, 298, 0, 298);
        worksheet.down();
        expect(attached().slice(0, 3)).to.eql([0, 1, 202]);
        expect(attached().length).to.equal(100);
        expect(attached()[99]).to.equal(299);

        // The rows above load before the rows after the frozen rows
        for (let k = 0; k < 120; k++) {
            worksheet.up();
        }
        expect(attached().slice(0, 2)).to.eql([0, 1]);
        expect(attached().length).to.equal(100);
        expect(new Set(attached()).size).to.equal(100);
        expect(attached().includes(worksheet.selection.range![3] as number)).to.equal(true);
    });

    it('keeps the frozen rows first and the other rows in order after a delete and a move of a frozen row', () => {
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(300), freezeRows: 2, lazyLoading: true, tableOverflow: true }] });
        const attached = () => Array.from(worksheet.tbody.children).map((row) => Number(row.getAttribute('data-y')));
        const ordered = (list: number[]) => list.every((y, k) => k === 0 || y > list[k - 1]);
        worksheet.updateSelectionFromCoords(0, 298, 0, 298);
        worksheet.down();

        worksheet.deleteRow(0);
        expect(attached().slice(0, 2)).to.eql([0, 1]);
        expect(worksheet.rows[1].element.parentNode).to.equal(worksheet.tbody);
        expect(ordered(attached())).to.equal(true);

        worksheet.moveRow(0, 250);
        expect(attached().slice(0, 2)).to.eql([0, 1]);
        expect(ordered(attached())).to.equal(true);

        // The keyboard loads the rows above in order
        worksheet.updateSelectionFromCoords(0, 215, 0, 215);
        for (let k = 0; k < 10; k++) {
            worksheet.up();
        }
        expect(attached().slice(0, 2)).to.eql([0, 1]);
        expect(ordered(attached())).to.equal(true);
    });

    it('freezes at most 40 rows with lazy loading, and a whole number of rows', () => {
        const [lazy] = jspreadsheet(root, { worksheets: [{ data: rows(300), freezeRows: 150, lazyLoading: true, tableOverflow: true }] });
        expect(frozen(lazy).length).to.equal(40);

        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(5), freezeRows: 1.5 }] });
        expect(frozen(worksheet).map(([y]) => y)).to.eql([0]);
    });

    it('removes the row that leaves the frozen rows after an insert at the top or a move into the frozen rows', () => {
        for (const change of [(w: Worksheet) => w.insertRow(1, 0, true), (w: Worksheet) => w.moveRow(250, 0)]) {
            const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(300), freezeRows: 2, lazyLoading: true, tableOverflow: true }] });
            const attached = () => Array.from(worksheet.tbody.children).map((row) => Number(row.getAttribute('data-y')));
            worksheet.updateSelectionFromCoords(0, 298, 0, 298);
            worksheet.down();

            change(worksheet);

            // The old row 1 is row 2 now, and it is not next to the rows that scroll
            expect(attached().slice(0, 2)).to.eql([0, 1]);
            expect(attached()[2]).to.be.greaterThan(150);
            const rest = attached().slice(2);
            expect(rest.every((y, k) => k === 0 || y === rest[k - 1] + 1)).to.equal(true);

            // The keyboard loads the rows above in order up to the frozen rows
            worksheet.updateSelectionFromCoords(0, rest[0], 0, rest[0]);
            for (let k = 0; k < rest[0]; k++) {
                worksheet.up();
            }
            expect(attached().slice(0, 5)).to.eql([0, 1, 2, 3, 4]);
        }
    });

    it('saves an open editor in its own cell when a change in the frozen rows removes its row', () => {
        for (const change of [(w: Worksheet) => w.insertRow(1, 0, true), (w: Worksheet) => w.moveRow(250, 0)]) {
            const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(300), freezeRows: 2, lazyLoading: true, tableOverflow: true }] });
            worksheet.updateSelectionFromCoords(0, 298, 0, 298);
            worksheet.down();
            const cell = worksheet.records[1][1].element;
            worksheet.openEditor(cell);
            (cell.children[0] as HTMLInputElement).value = 'EDITED';

            change(worksheet);

            // The edited row is row 2 now
            expect(worksheet.editor.isOpen()).to.equal(false);
            expect(worksheet.getValueFromCoords(1, 2)).to.equal('EDITED');
            expect([0, 1, 3].map((y) => worksheet.getValueFromCoords(1, y))).to.not.include('EDITED');
        }
    });

    it('keeps the rows that scroll after an insert of many rows at the top', () => {
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(300), freezeRows: 2, lazyLoading: true, tableOverflow: true }] });
        worksheet.updateSelectionFromCoords(0, 298, 0, 298);
        worksheet.down();

        worksheet.insertRow(100, 0, true);

        // The rows that scrolled before the insert, now 300 to 399, stay in the table. The inserted rows and the rows that
        // left the frozen rows leave it.
        const attached = Array.from(worksheet.tbody.children).map((row) => Number(row.getAttribute('data-y')));
        expect(worksheet.rows[399].element.parentNode).to.equal(worksheet.tbody);
        expect(attached.slice(0, 2)).to.eql([0, 1]);
        expect(attached.slice(2).every((y) => y >= 300)).to.equal(true);
    });

    it('keeps the selected row in the table while the keyboard moves through 400 rows with 40 frozen rows', () => {
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(400), freezeRows: 40, lazyLoading: true, tableOverflow: true }] });
        const missing: number[] = [];
        const check = () => {
            const y = worksheet.selection.range![3] as number;
            if (worksheet.rows[y].element.parentNode !== worksheet.tbody) {
                missing.push(y);
            }
        };

        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        for (let k = 0; k < 390; k++) {
            worksheet.down();
            check();
        }
        worksheet.updateSelectionFromCoords(0, 399, 0, 399);
        for (let k = 0; k < 390; k++) {
            worksheet.up();
            check();
        }

        expect(missing).to.eql([]);
    });

    it('keeps the frozen rows first when the content scrolls down', async () => {
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(300), freezeRows: 2, lazyLoading: true, tableOverflow: true }] });

        // The view is near the end of the attached rows, and the wheel turns down
        Object.defineProperty(worksheet.content, 'scrollHeight', { value: 2000 });
        Object.defineProperty(worksheet.content, 'clientHeight', { value: 300 });
        worksheet.content.scrollTop = 1700;
        worksheet.content.dispatchEvent(new window.WheelEvent('wheel', { deltaY: 100 }));
        // fixed wait: the lazy loading loads 100 ms after the wheel event
        await new Promise((resolve) => setTimeout(resolve, 150));

        const attached = Array.from(worksheet.tbody.children).map((row) => Number(row.getAttribute('data-y')));
        // The load adds rows 100 to 129 and removes the 30 rows after the frozen rows
        expect(attached.slice(0, 3)).to.eql([0, 1, 32]);
        expect(attached[99]).to.equal(129);
        expect(attached.length).to.equal(100);
    });
});

describe('The scroll to the active cell below the frozen rows', () => {
    // The content is 300 pixels high and scrolls 500 pixels. The header row and each row are 20 pixels high.
    function scrolledWorksheet() {
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(60), freezeRows: 2, tableOverflow: true }] });
        Object.defineProperty(worksheet.content, 'scrollTop', { value: 500, writable: true });
        worksheet.content.getBoundingClientRect = rect(0, 0, 400, 300);
        worksheet.tbody.getBoundingClientRect = rect(0, 21 - 500, 400, 1200);
        return worksheet;
    }

    it(
        'puts the top edge of a cell that a frozen row covers below the frozen rows',
        withRowHeight(() => {
            const worksheet = scrolledWorksheet();

            // The header and two frozen rows cover 60 pixels
            worksheet.records[30][0].element.getBoundingClientRect = rect(50, 45, 100, 20);
            worksheet.updateSelectionFromCoords(0, 31, 0, 31);
            worksheet.up();

            expect(worksheet.content.scrollTop).to.equal(500 + 45 - 60);
        })
    );

    it(
        'keeps the scroll position when the keyboard selects a cell of a frozen row',
        withRowHeight(() => {
            const worksheet = scrolledWorksheet();

            // Without frozen rows, the scroll would move this cell below the header row
            worksheet.records[0][0].element.getBoundingClientRect = rect(50, 15, 100, 20);
            worksheet.updateSelectionFromCoords(0, 1, 0, 1);
            worksheet.up();

            expect(worksheet.selection.range!.map(Number)).to.eql([0, 0, 0, 0]);
            expect(worksheet.content.scrollTop).to.equal(500);
        })
    );

    it(
        'scrolls to the top for the first row after the frozen rows',
        withRowHeight(() => {
            const worksheet = scrolledWorksheet();

            worksheet.records[2][0].element.getBoundingClientRect = rect(50, 21 - 500 + 40, 100, 20);
            worksheet.updateSelectionFromCoords(0, 3, 0, 3);
            worksheet.up();

            expect(worksheet.content.scrollTop).to.equal(0);
        })
    );
});

describe('Frozen rows and frozen columns', () => {
    it('gives a cell of a frozen row and a frozen column the sticky distance of its header', async () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [
                {
                    data: rows(5),
                    freezeRows: 1,
                    freezeColumns: 1,
                    tableOverflow: true,
                    columns: [{ width: 100 }, { width: 300 }],
                },
            ],
        });

        worksheet.content.scrollLeft = 0;
        worksheet.content.dispatchEvent(new Event('scroll'));
        Object.defineProperty(worksheet.content, 'scrollLeft', { value: 400, writable: true });
        worksheet.content.dispatchEvent(new Event('scroll'));

        const shifted = worksheet.records[1][0].element.style.left;
        expect(worksheet.records[0][0].element.style.left).to.equal(worksheet.headers[0].style.left);
        expect(shifted).to.not.equal(worksheet.headers[0].style.left);

        // After a move, the new first row is the frozen row
        worksheet.moveRow(0, 2);
        await placed();
        expect(worksheet.records[0][0].element.style.left).to.equal(worksheet.headers[0].style.left);
        expect(worksheet.records[2][0].element.style.left).to.equal(shifted);
    });
});

describe('setFreeze', () => {
    it(
        'freezes more rows, fewer rows and no rows, and keeps a number that is undefined',
        withRowHeight(async () => {
            const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(10), freezeRows: 1 }] });

            worksheet.setFreeze(3);
            await placed();
            expect(frozen(worksheet)).to.eql([
                [0, '20px', false],
                [1, '40px', false],
                [2, '60px', true],
            ]);

            worksheet.setFreeze(undefined, 2);
            expect(worksheet.options.freezeRows).to.equal(3);
            expect(worksheet.options.freezeColumns).to.equal(2);

            worksheet.setFreeze(1);
            await placed();
            expect(frozen(worksheet)).to.eql([[0, '20px', true]]);
            expect(worksheet.rows[2].element.style.getPropertyValue('--jss-frozen-top')).to.equal('');

            worksheet.setFreeze(0);
            await placed();
            expect(frozen(worksheet)).to.eql([]);
        })
    );

    it('freezes and unfreezes columns at the current scroll position', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: rows(5), tableOverflow: true, columns: [{ width: 100 }, { width: 300 }, { width: 100 }] }],
        });
        Object.defineProperty(worksheet.content, 'scrollLeft', { value: 400, writable: true });
        const classes = (cell: HTMLElement) => ['jss_freezed', 'jss_freezed_last'].filter((name) => cell.classList.contains(name));

        worksheet.setFreeze(undefined, 1);
        expect(classes(worksheet.headers[0])).to.eql(['jss_freezed', 'jss_freezed_last']);
        expect(classes(worksheet.records[2][0].element)).to.eql(['jss_freezed', 'jss_freezed_last']);
        expect(worksheet.headers[0].style.left).to.equal('0px');

        worksheet.setFreeze(undefined, 2);
        expect(classes(worksheet.headers[0])).to.eql(['jss_freezed']);
        expect(classes(worksheet.records[2][0].element)).to.eql(['jss_freezed']);
        expect(classes(worksheet.headers[1])).to.eql(['jss_freezed', 'jss_freezed_last']);
        expect(classes(worksheet.records[2][1].element)).to.eql(['jss_freezed', 'jss_freezed_last']);
        expect(worksheet.headers[1].style.left).to.equal('100px');

        worksheet.setFreeze(undefined, 0);
        for (const cell of [worksheet.headers[0], worksheet.headers[1], worksheet.records[2][0].element, worksheet.records[2][1].element]) {
            expect(classes(cell)).to.eql([]);
            expect(cell.style.left).to.equal('');
        }
    });

    it('gives a crossing cell the distance of its header when the number of frozen rows changes', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: rows(5), tableOverflow: true, columns: [{ width: 100 }, { width: 300 }] }],
        });
        Object.defineProperty(worksheet.content, 'scrollLeft', { value: 400, writable: true });
        worksheet.setFreeze(undefined, 2);
        const shifted = worksheet.records[3][1].element.style.left;
        expect(shifted).to.not.equal('100px');

        worksheet.setFreeze(2);
        expect(worksheet.records[1][1].element.style.left).to.equal('100px');
        expect(worksheet.records[3][1].element.style.left).to.equal(shifted);

        worksheet.setFreeze(0);
        expect(worksheet.records[1][1].element.style.left).to.equal(shifted);
    });

    it('freezes the columns again after a scroll that occurred while no column was frozen', () => {
        const [worksheet] = jspreadsheet(root, {
            worksheets: [{ data: rows(5), tableOverflow: true, columns: [{ width: 100 }, { width: 300 }] }],
        });
        const content = worksheet.content;
        function scrollTo(left: number) {
            Object.defineProperty(content, 'scrollLeft', { value: left, writable: true, configurable: true });
            content.dispatchEvent(new window.Event('scroll'));
        }
        worksheet.setFreeze(undefined, 1);
        scrollTo(400);
        worksheet.setFreeze(undefined, 0);
        scrollTo(0);
        worksheet.setFreeze(undefined, 1);

        scrollTo(400);

        expect(worksheet.headers[0].classList.contains('jss_freezed')).to.equal(true);
    });

    it('gives 0 for a value that is not a finite number of 0 or more', () => {
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(5), freezeRows: 2, freezeColumns: 1 }] });

        for (const value of [Infinity, NaN, -2, 'abc', null]) {
            worksheet.setFreeze(2, 1);
            worksheet.setFreeze(value as number, value as number);
            expect([worksheet.options.freezeRows, worksheet.options.freezeColumns]).to.eql([0, 0]);
        }
        worksheet.setFreeze(2.7, '3' as unknown as number);
        expect([worksheet.options.freezeRows, worksheet.options.freezeColumns]).to.eql([2, 3]);
    });

    it('puts the new frozen rows first in the table with lazy loading, and removes the rows that leave', () => {
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(300), lazyLoading: true, tableOverflow: true }] });
        const attached = () => Array.from(worksheet.tbody.children).map((row) => Number(row.getAttribute('data-y')));
        worksheet.updateSelectionFromCoords(0, 298, 0, 298);
        worksheet.down();

        worksheet.setFreeze(3);
        expect(attached().slice(0, 3)).to.eql([0, 1, 2]);
        expect(attached()[3]).to.be.greaterThan(150);

        worksheet.setFreeze(1);
        expect(attached()[0]).to.equal(0);
        expect(attached()[1]).to.be.greaterThan(150);
        const rest = attached().slice(1);
        expect(rest.every((y, k) => k === 0 || y === rest[k - 1] + 1)).to.equal(true);
    });

    it('removes the last frozen rows from the table with lazy loading, so that the rows above can load again', () => {
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(300), lazyLoading: true, tableOverflow: true, freezeRows: 3 }] });
        const attached = () => Array.from(worksheet.tbody.children).map((row) => Number(row.getAttribute('data-y')));
        worksheet.updateSelectionFromCoords(0, 298, 0, 298);
        worksheet.down();
        expect(attached().slice(0, 3)).to.eql([0, 1, 2]);

        worksheet.setFreeze(0);

        const first = attached()[0];
        expect(first).to.be.greaterThan(150);
        expect(attached().every((y, k) => k === 0 || y === attached()[k - 1] + 1)).to.equal(true);
        worksheet.view.loadUp();
        expect(attached()[0]).to.be.lessThan(first);
    });

    it('keeps a tab stop in the table when the selected row leaves it with lazy loading', () => {
        const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(300), lazyLoading: true, tableOverflow: true, freezeRows: 3 }] });
        worksheet.updateSelectionFromCoords(0, 298, 0, 298);
        worksheet.down();
        worksheet.updateSelectionFromCoords(1, 2, 1, 2);
        expect(worksheet.tbody.querySelector('td[tabindex="0"]')).to.equal(worksheet.records[2][1].element);

        worksheet.setFreeze(1);

        expect(worksheet.tbody.contains(worksheet.records[2][1].element)).to.equal(false);
        expect(worksheet.tbody.querySelector('td[tabindex="0"]')).to.not.equal(null);
    });
});

describe('Header rows stick below each other (issue #184)', () => {
    function tops(worksheet: Worksheet) {
        return Array.from(worksheet.thead.children).map((row) => (row as HTMLElement).style.getPropertyValue('--jss-header-top'));
    }

    it(
        'puts each nested header row and the filter row below the rows above it',
        withRowHeight(async () => {
            const [worksheet] = jspreadsheet(root, {
                worksheets: [{ data: rows(5), filters: true, nestedHeaders: [[{ title: 'Group', colspan: 2 }]] }],
            });
            await placed();

            expect(tops(worksheet)).to.eql(['0px', '20px', '40px']);
        })
    );

    it(
        'puts the frozen rows below all the header rows',
        withRowHeight(async () => {
            const [worksheet] = jspreadsheet(root, {
                worksheets: [{ data: rows(5), freezeRows: 1, nestedHeaders: [[{ title: 'Group', colspan: 2 }]] }],
            });
            await placed();

            expect(frozen(worksheet)).to.eql([[0, '40px', true]]);
        })
    );

    it(
        'gives a header of one row no distance',
        withRowHeight(async () => {
            const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(5) }] });
            await placed();

            expect(tops(worksheet)).to.eql(['']);
        })
    );

    it('places the header rows again when the size observer sees a change, and destroy stops the observer', async () => {
        const observed: Element[] = [];
        let callback: (() => void) | null = null;
        const original = (globalThis as { ResizeObserver?: unknown }).ResizeObserver;
        (globalThis as { ResizeObserver?: unknown }).ResizeObserver = class {
            constructor(cb: () => void) {
                callback = cb;
            }
            observe(element: Element) {
                observed.push(element);
            }
            disconnect() {
                observed.length = 0;
            }
        };

        const element = document.body.appendChild(document.createElement('div'));
        try {
            const [worksheet] = jspreadsheet(element, { worksheets: [{ data: rows(5), nestedHeaders: [[{ title: 'Group', colspan: 2 }]] }] });
            await new Promise((resolve) => setTimeout(resolve));

            expect(observed).to.eql(Array.from(worksheet.thead.children));

            // A web font makes the header rows 30 pixels high
            Object.defineProperty(window.HTMLTableRowElement.prototype, 'offsetHeight', { configurable: true, get: () => 30 });
            try {
                callback!();
            } finally {
                delete (window.HTMLTableRowElement.prototype as { offsetHeight?: number }).offsetHeight;
            }
            expect(tops(worksheet)).to.eql(['0px', '30px']);

            jspreadsheet.destroy(element as never);
            expect(observed).to.eql([]);
        } finally {
            (globalThis as { ResizeObserver?: unknown }).ResizeObserver = original;
            element.remove();
        }
    });
});

describe('The fill handle under the header and the frozen rows (issue #185)', () => {
    // jsdom has no layout. The cells of the header rows and of the frozen rows are sticky, so the test gives the first
    // cell of the last header row, and of the last frozen row, the place where it stops on the screen.
    function stick(row: Element, bottom: number) {
        (row.children[0] as HTMLElement).getBoundingClientRect = rect(0, bottom - 20, 50, 20);
    }

    // The content is at the top left of the screen. The bottom of the cell is `bottom` pixels below the top of the screen.
    function select(worksheet: Worksheet, x: number, y: number, bottom: number) {
        worksheet.content.getBoundingClientRect = rect(0, 0, 400, 300);
        worksheet.records[y][x].element.getBoundingClientRect = rect(150, bottom - 20, 100, 20);
        worksheet.updateSelectionFromCoords(x, y, x, y);
    }

    function scroll(worksheet: Worksheet, x: number, y: number, top: number, bottom: number) {
        Object.defineProperty(worksheet.content, 'scrollTop', { value: top, writable: true, configurable: true });
        worksheet.records[y][x].element.getBoundingClientRect = rect(150, bottom - 20, 100, 20);
        worksheet.content.dispatchEvent(new Event('scroll'));
    }

    it(
        'hides the fill handle while its cell is under the nested header rows, and shows it again after the scroll back',
        withRowHeight(async () => {
            const [worksheet] = jspreadsheet(root, {
                worksheets: [{ data: rows(40), tableOverflow: true, nestedHeaders: [[{ title: 'Group', colspan: 2 }]] }],
            });
            await placed();
            stick(worksheet.thead.lastElementChild!, 40);

            select(worksheet, 1, 5, 50);
            expect(worksheet.corner.style.display).to.equal('');

            // The header rows cover 40 pixels, and the fill handle starts 4 pixels above the bottom of the cell
            scroll(worksheet, 1, 5, 20, 30);
            expect(worksheet.corner.style.display).to.equal('none');

            scroll(worksheet, 1, 5, 0, 50);
            expect(worksheet.corner.style.display).to.equal('');
        })
    );

    it(
        'hides the fill handle of a cell under the frozen rows, but not of a cell in a frozen row',
        withRowHeight(async () => {
            const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(40), tableOverflow: true, freezeRows: 2 }] });
            await placed();
            // The header row and the two frozen rows cover 60 pixels
            stick(worksheet.thead.lastElementChild!, 20);
            stick(worksheet.rows[1].element, 60);

            select(worksheet, 1, 5, 55);
            expect(worksheet.corner.style.display).to.equal('none');

            select(worksheet, 1, 1, 55);
            expect(worksheet.corner.style.display).to.equal('');
        })
    );

    it(
        'hides the fill handle while a scroll of the page moves its cell under the header rows',
        withRowHeight(async () => {
            const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(40), filters: true }] });
            await built();
            // The header row and the filter row stick at the top of the window, and cover 40 pixels
            stick(worksheet.thead.lastElementChild!, 40);

            select(worksheet, 1, 10, 300);
            expect(worksheet.corner.style.display).to.equal('');

            worksheet.records[10][1].element.getBoundingClientRect = rect(150, 10, 100, 20);
            document.dispatchEvent(new Event('scroll'));
            expect(worksheet.corner.style.display).to.equal('none');

            worksheet.records[10][1].element.getBoundingClientRect = rect(150, 280, 100, 20);
            document.dispatchEvent(new Event('scroll'));
            expect(worksheet.corner.style.display).to.equal('');
        })
    );

    it(
        'hides the fill handle under the last frozen row that shows when the last frozen row is hidden',
        withRowHeight(async () => {
            const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(40), tableOverflow: true, freezeRows: 3 }] });
            await placed();
            worksheet.hideRow(2);
            await placed();
            // The header row and the two frozen rows that show cover 60 pixels. The hidden row has no size on the screen.
            stick(worksheet.thead.lastElementChild!, 20);
            stick(worksheet.rows[1].element, 60);

            select(worksheet, 1, 5, 55);
            expect(worksheet.corner.style.display).to.equal('none');
        })
    );

    it(
        'checks the fill handle of each worksheet in the scrolled element, also of a worksheet that is not current',
        withRowHeight(async () => {
            const first = document.createElement('div');
            const second = document.createElement('div');
            root.appendChild(first);
            root.appendChild(second);
            const a = jspreadsheet(first, { worksheets: [{ data: rows(40) }] })[0];
            const b = jspreadsheet(second, { worksheets: [{ data: rows(40) }] })[0];
            await built();
            stick(a.thead.lastElementChild!, 40);
            stick(b.thead.lastElementChild!, 40);
            select(a, 1, 10, 300);
            select(b, 1, 10, 300);
            jspreadsheet.current = b;

            // A scroll of the element around the first worksheet reads no layout of the second worksheet
            let reads = 0;
            b.records[10][1].element.getBoundingClientRect = () => {
                reads++;
                return rect(150, 10, 100, 20)();
            };
            a.records[10][1].element.getBoundingClientRect = rect(150, 10, 100, 20);
            first.dispatchEvent(new Event('scroll'));
            expect(a.corner.style.display).to.equal('none');
            expect(b.corner.style.display).to.equal('');
            expect(reads).to.equal(0);

            a.records[10][1].element.getBoundingClientRect = rect(150, 280, 100, 20);
            document.dispatchEvent(new Event('scroll'));
            expect(a.corner.style.display).to.equal('');
            expect(b.corner.style.display).to.equal('none');
        })
    );

    it(
        'reads no layout on a scroll when the worksheet has no fill handle',
        withRowHeight(async () => {
            const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(40), tableOverflow: true, selectionCopy: false }] });
            await placed();
            worksheet.updateSelectionFromCoords(1, 10, 1, 10);

            let reads = 0;
            worksheet.records[10][1].element.getBoundingClientRect = () => {
                reads++;
                return rect(150, 10, 100, 20)();
            };
            document.dispatchEvent(new Event('scroll'));
            scroll(worksheet, 0, 0, 20, 10);
            expect(reads).to.equal(0);
        })
    );

    it(
        'keeps the fill handle of a cell under the frozen columns hidden after a scroll down',
        withRowHeight(async () => {
            const [worksheet] = jspreadsheet(root, { worksheets: [{ data: rows(40), tableOverflow: true, freezeColumns: 1 }] });
            await placed();

            // The frozen column is 100 pixels wide, and the cell ends 50 pixels from the left of the content
            worksheet.content.getBoundingClientRect = rect(0, 0, 400, 300);
            worksheet.records[10][1].element.getBoundingClientRect = rect(0, 180, 50, 20);
            worksheet.updateSelectionFromCoords(1, 10, 1, 10);
            expect(worksheet.corner.style.display).to.equal('none');

            Object.defineProperty(worksheet.content, 'scrollTop', { value: 20, writable: true, configurable: true });
            worksheet.records[10][1].element.getBoundingClientRect = rect(0, 160, 50, 20);
            worksheet.content.dispatchEvent(new Event('scroll'));
            expect(worksheet.corner.style.display).to.equal('none');
        })
    );
});
