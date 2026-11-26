import { expect } from 'chai';

import jspreadsheet, { type Worksheet } from './support/jspreadsheet';

function rows(count: number) {
    const data: string[][] = [];
    for (let j = 0; j < count; j++) {
        data.push([`${j}`]);
    }
    return data;
}

describe('Lazy loading', () => {
    it('loads the page of the selected row when the keyboard moves from a row that is not attached', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: rows(300), lazyLoading: true, tableOverflow: true }],
        });

        const worksheet = instance[0];

        // Row 251 is not attached
        worksheet.updateSelectionFromCoords(0, 250, 0, 250);
        worksheet.down();

        expect(worksheet.selection.range!.map(Number)).to.eql([0, 251, 0, 251]);
        expect(worksheet.tbody.contains(worksheet.records[251][0].element)).to.equal(true);

        worksheet.updateSelectionFromCoords(0, 250, 0, 250);
        worksheet.up();

        expect(worksheet.selection.range!.map(Number)).to.eql([0, 249, 0, 249]);
        expect(worksheet.tbody.contains(worksheet.records[249][0].element)).to.equal(true);
    });

    it('moves up to the first attached row without an error', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: rows(300), lazyLoading: true, tableOverflow: true }],
        });

        const worksheet = instance[0];

        // A move near the last attached row attaches the next rows and removes the first rows
        worksheet.updateSelectionFromCoords(0, 80, 0, 80);
        worksheet.down();

        const first = parseInt((worksheet.tbody.firstChild as HTMLElement).getAttribute('data-y'));
        expect(first).to.be.above(0);

        worksheet.updateSelectionFromCoords(0, first + 1, 0, first + 1);
        worksheet.up();

        expect(worksheet.selection.range!.map(Number)).to.eql([0, first, 0, first]);
    });

    it('attaches only the rows of the page when more than 100 rows are attached', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: rows(300), lazyLoading: true, tableOverflow: true }],
        });

        const worksheet = instance[0];

        // The keyboard attaches the last page, and an insert on the last page attaches all the new rows
        worksheet.updateSelectionFromCoords(0, 298, 0, 298);
        worksheet.down();
        worksheet.insertRow(150);
        expect(worksheet.tbody.children.length).to.equal(250);

        // The keyboard attaches the first page
        worksheet.updateSelectionFromCoords(0, 1, 0, 1);
        worksheet.up();

        const attached = Array.from(worksheet.tbody.children).map((row) => Number(row.getAttribute('data-y')));
        expect(attached).to.eql(rows(100).map((row, y) => y));
    });

    it('attaches only the rows of the page after an insert in the middle of the attached rows', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ data: rows(300), lazyLoading: true, tableOverflow: true }],
        });

        const worksheet = instance[0];

        // The first page is attached, and an insert next to an attached row attaches all the new rows
        worksheet.insertRow(150, 50);
        expect(worksheet.tbody.children.length).to.equal(250);

        // The keyboard attaches the last page
        worksheet.updateSelectionFromCoords(0, 448, 0, 448);
        worksheet.down();

        const attached = Array.from(worksheet.tbody.children).map((row) => Number(row.getAttribute('data-y')));
        expect(attached).to.eql(rows(100).map((row, y) => 350 + y));
    });

    it('loads the page of the selected result when the keyboard moves with a search', () => {
        const data: string[][] = [];
        for (let j = 0; j < 600; j++) {
            data.push([(j % 2 ? 'b' : 'a') + j]);
        }

        const instance = jspreadsheet(root, {
            worksheets: [{ data, search: true, lazyLoading: true, tableOverflow: true }],
        });

        const worksheet = instance[0];

        // The search keeps the even rows. Rows 302 and 452 are not attached.
        worksheet.search('a');

        worksheet.updateSelectionFromCoords(0, 300, 0, 300);
        worksheet.down();

        expect(worksheet.selection.range!.map(Number)).to.eql([0, 302, 0, 302]);
        expect(worksheet.tbody.contains(worksheet.records[302][0].element)).to.equal(true);

        worksheet.updateSelectionFromCoords(0, 450, 0, 450);
        worksheet.down();

        expect(worksheet.selection.range!.map(Number)).to.eql([0, 452, 0, 452]);
        expect(worksheet.tbody.contains(worksheet.records[452][0].element)).to.equal(true);
    });
});

describe('Scroll position after a lazy load', () => {
    const wait = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));
    const rowPrototype = window.HTMLTableRowElement.prototype as unknown as Record<string, unknown>;

    // jsdom has no layout. The stubs give each row that shows a height of 20 pixels, and the view a height of 400 pixels
    // or the height that build gets.
    // With shown false, each row has no offsetParent, as in a worksheet that does not show.
    let shown: boolean;
    let scrollTop: number;
    const isShown = (row: HTMLElement) => row.parentNode != null && row.style.display !== 'none';

    beforeEach(() => {
        shown = true;
        scrollTop = 0;
        Object.defineProperty(rowPrototype, 'offsetParent', {
            configurable: true,
            get(this: HTMLElement) {
                return shown && isShown(this) ? this.parentNode!.parentNode : null;
            },
        });
        Object.defineProperty(rowPrototype, 'offsetTop', {
            configurable: true,
            get(this: HTMLElement) {
                if (!isShown(this)) {
                    return 0;
                }
                const siblings = Array.from(this.parentNode!.children) as HTMLElement[];
                return siblings.slice(0, siblings.indexOf(this)).filter(isShown).length * 20;
            },
        });
    });

    afterEach(() => {
        delete rowPrototype.offsetParent;
        delete rowPrototype.offsetTop;
    });

    async function build(viewHeight = 400) {
        const worksheet = jspreadsheet(root, { worksheets: [{ data: rows(300), lazyLoading: true, tableOverflow: true }] })[0];
        const content = worksheet.content;
        const scrollHeight = () => (Array.from(worksheet.tbody.children) as HTMLElement[]).filter(isShown).length * 20;
        Object.defineProperty(content, 'clientHeight', { get: () => viewHeight });
        Object.defineProperty(content, 'scrollHeight', { get: scrollHeight });
        Object.defineProperty(content, 'scrollTop', {
            get: () => scrollTop,
            set: (value: number) => {
                scrollTop = Math.max(0, Math.min(value, scrollHeight() - viewHeight));
            },
        });
        // The earlier tests can leave a timer of the lazy loading. The worksheet ignores a scroll until that timer ends.
        await wait(150);
        return worksheet;
    }

    async function wheel(worksheet: Worksheet, deltaY: number) {
        worksheet.content.dispatchEvent(new window.WheelEvent('wheel', { deltaY }));
        await wait(150);
    }

    function toBottom(worksheet: Worksheet) {
        worksheet.content.scrollTop = worksheet.content.scrollHeight;
    }

    const first = (worksheet: Worksheet) => Number((worksheet.tbody.firstChild as HTMLElement).getAttribute('data-y'));

    it('keeps the last attached row in its place in the view after a load at the bottom', async () => {
        const worksheet = await build();
        toBottom(worksheet);
        const anchor = worksheet.tbody.lastChild as HTMLTableRowElement;
        const place = anchor.offsetTop - scrollTop;

        await wheel(worksheet, 100);

        expect(first(worksheet)).to.equal(30);
        expect(worksheet.tbody.contains(anchor)).to.equal(true);
        expect(anchor.offsetTop - scrollTop).to.equal(place);
        expect(scrollTop).to.equal(1000);
    });

    it('keeps the last row that shows in its place when the last attached row is hidden', async () => {
        const worksheet = await build();
        (worksheet.tbody.lastChild as HTMLElement).style.display = 'none';
        toBottom(worksheet);
        const anchor = (worksheet.tbody.lastChild as HTMLElement).previousSibling as HTMLTableRowElement;
        const place = anchor.offsetTop - scrollTop;

        await wheel(worksheet, 100);

        expect(first(worksheet)).to.equal(30);
        expect(anchor.offsetTop - scrollTop).to.equal(place);
    });

    it('keeps the first attached row in its place in the view after a load at the top', async () => {
        const worksheet = await build();
        toBottom(worksheet);
        await wheel(worksheet, 100);
        worksheet.content.scrollTop = 0;
        const anchor = worksheet.tbody.firstChild as HTMLTableRowElement;

        await wheel(worksheet, -100);

        expect(first(worksheet)).to.equal(0);
        expect(anchor.offsetTop - scrollTop).to.equal(0);
        expect(scrollTop).to.equal(600);
    });

    it('loads no rows at the top while a view higher than 30 rows scrolls down', async () => {
        const worksheet = await build(1000);
        const firsts = [];
        for (let k = 0; k < 8; k++) {
            // The wheel scrolls the content, as in a browser
            worksheet.content.scrollTop += 200;
            await wheel(worksheet, 200);
            firsts.push(first(worksheet));
        }

        expect(firsts).to.eql([0, 0, 0, 0, 30, 30, 30, 60]);
    });

    it('keeps the scroll position when no attached row shows', async () => {
        const worksheet = await build();
        toBottom(worksheet);
        shown = false;

        await wheel(worksheet, 100);

        expect(first(worksheet)).to.equal(30);
        expect(scrollTop).to.equal(1600);
    });

    it('loads nothing for a scroll event without a change of the scroll position', async () => {
        const worksheet = await build();
        toBottom(worksheet);
        await wheel(worksheet, 100);

        worksheet.content.dispatchEvent(new window.Event('scroll'));
        await wait(150);

        expect(first(worksheet)).to.equal(30);
    });

    it('loads nothing at the top for a scroll down near the top', async () => {
        const worksheet = await build();
        toBottom(worksheet);
        await wheel(worksheet, 100);
        worksheet.content.scrollTop = 5;

        await wheel(worksheet, 100);

        expect(first(worksheet)).to.equal(30);
        expect(scrollTop).to.equal(5);
    });
});
