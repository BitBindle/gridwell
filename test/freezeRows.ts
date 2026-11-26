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
    it('gives a cell of a frozen row and a frozen column the sticky distance of its header', () => {
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
        expect(worksheet.records[0][0].element.style.left).to.equal(worksheet.headers[0].style.left);
        expect(worksheet.records[2][0].element.style.left).to.equal(shifted);
    });
});
