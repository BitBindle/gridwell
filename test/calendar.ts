import { expect } from 'chai';

import jspreadsheet, { type Jspreadsheet } from './support/jspreadsheet';

// The key codes, which the keys of the editor and of the grid read
const codes: Record<string, number> = {
    Tab: 9,
    Enter: 13,
    Escape: 27,
    PageUp: 33,
    PageDown: 34,
    End: 35,
    Home: 36,
    ArrowLeft: 37,
    ArrowUp: 38,
    ArrowRight: 39,
    ArrowDown: 40,
};

const key = (element: Element, keyName: string, options: KeyboardEventInit = {}) => {
    const keyCode = codes[keyName] ?? keyName.toUpperCase().charCodeAt(0);
    const event = new KeyboardEvent('keydown', { key: keyName, keyCode, which: keyCode, bubbles: true, cancelable: true, ...options });
    element.dispatchEvent(event);
    return event;
};

const press = (element: Element) => {
    element.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
    element.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
};

const waitForListeners = () => new Promise((resolve) => setTimeout(resolve));

const pad = (n: number) => `${n}`.padStart(2, '0');

/** Today as the text "YYYY-MM-DD". */
const today = () => {
    const now = new Date();
    return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
};

/** A worksheet with a calendar column, with the editor of A1 open. */
async function open(
    options: Jspreadsheet.CalendarOptions = {},
    value: Jspreadsheet.CellValue = '2024-03-05',
    worksheetOptions: Partial<Jspreadsheet.WorksheetOptions> = {}
) {
    const instance = jspreadsheet(root, {
        worksheets: [
            {
                data: [
                    [value, 'x'],
                    ['', 'y'],
                ],
                columns: [{ type: 'calendar', options }, {}],
                ...worksheetOptions,
            },
        ],
    });
    await waitForListeners();
    const worksheet = instance[0];
    worksheet.updateSelectionFromCoords(0, 0, 0, 0);
    const cell = worksheet.getCell('A1');
    worksheet.openEditor(cell);

    const input = cell.querySelector<HTMLInputElement & { calendar?: unknown }>('input')!;
    const calendar = cell.querySelector<HTMLElement>('.jcalendar')!;
    const grid = calendar.querySelector<HTMLTableElement>('[role="grid"]')!;
    const active = () => document.getElementById(input.getAttribute('aria-activedescendant') || '');
    const day = (text: string) => Array.from(grid.querySelectorAll<HTMLElement>('.jcalendar-set-day')).find((td) => td.textContent === text)!;
    const button = (className: string) => calendar.querySelector<HTMLElement>(`.${className}`)!;
    return { worksheet, cell, input, calendar, grid, active, day, button };
}

describe('The calendar', () => {
    it('is a combobox with a grid of the days of the month, and the day of the cell is active', async () => {
        const { input, calendar, grid, active, day } = await open();

        expect(calendar.classList.contains('jcalendar-focus')).to.equal(true);
        expect(document.activeElement).to.equal(input);
        expect(input.value).to.equal('2024-03-05');
        expect(input.readOnly).to.equal(true);
        expect(input.getAttribute('role')).to.equal('combobox');
        expect(input.getAttribute('aria-haspopup')).to.equal('grid');
        expect(input.getAttribute('aria-expanded')).to.equal('true');
        expect(input.getAttribute('aria-controls')).to.equal(grid.id);
        expect(input.getAttribute('aria-label')).to.equal('Edit cell A1');

        expect(grid.getAttribute('aria-label')).to.equal('March 2024');
        expect(calendar.querySelector('.jcalendar-month')!.textContent).to.equal('Mar');
        expect(calendar.querySelector('.jcalendar-year')!.textContent).to.equal('2024');

        const weekdays = Array.from(grid.querySelectorAll('[role="columnheader"]'));
        expect(weekdays.map((th) => th.textContent).join('')).to.equal('SMTWTFS');
        expect(weekdays[1].getAttribute('aria-label')).to.equal('Monday');

        // The 1st of March 2024 is a Friday
        const days = Array.from(grid.querySelectorAll('.jcalendar-set-day'));
        expect(days.slice(0, 7).map((td) => td.textContent)).to.eql(['', '', '', '', '', '1', '2']);
        expect(days.filter((td) => td.textContent).length).to.equal(31);

        expect(active()).to.equal(day('5'));
        expect(active()!.getAttribute('role')).to.equal('gridcell');
        expect(active()!.getAttribute('aria-label')).to.equal('Tuesday 5 March 2024');
        expect(active()!.getAttribute('aria-selected')).to.equal('true');
        expect(active()!.classList.contains('jcalendar-selected')).to.equal(true);
        expect(day('6').getAttribute('aria-selected')).to.equal('false');

        // The buttons of the previous and the next month have names
        expect(calendar.querySelector('.jcalendar-prev button')!.getAttribute('aria-label')).to.equal('Previous');
        expect(calendar.querySelector('.jcalendar-next button')!.getAttribute('aria-label')).to.equal('Next');
    });

    it('moves the active day with the arrows, Page Up, Page Down, Home and End', async () => {
        const { input, grid, active } = await open({}, '2024-01-31');

        const event = key(input, 'ArrowRight');
        expect(event.defaultPrevented).to.equal(true);
        expect(active()!.getAttribute('aria-label')).to.equal('Thursday 1 February 2024');
        expect(grid.getAttribute('aria-label')).to.equal('February 2024');

        key(input, 'ArrowLeft');
        expect(active()!.getAttribute('aria-label')).to.equal('Wednesday 31 January 2024');
        key(input, 'ArrowUp');
        expect(active()!.getAttribute('aria-label')).to.equal('Wednesday 24 January 2024');
        key(input, 'ArrowDown');
        key(input, 'ArrowDown');
        expect(active()!.getAttribute('aria-label')).to.equal('Wednesday 7 February 2024');

        // The start and the end of the week, which starts on Sunday
        key(input, 'Home');
        expect(active()!.getAttribute('aria-label')).to.equal('Sunday 4 February 2024');
        key(input, 'End');
        expect(active()!.getAttribute('aria-label')).to.equal('Saturday 10 February 2024');

        // A month, on the last day of a shorter month, and a year with Shift
        key(input, 'PageUp');
        key(input, 'PageUp');
        expect(active()!.getAttribute('aria-label')).to.equal('Sunday 10 December 2023');
        key(input, 'PageDown', { shiftKey: true });
        expect(active()!.getAttribute('aria-label')).to.equal('Tuesday 10 December 2024');
        key(input, 'PageUp', { shiftKey: true });
        key(input, 'PageDown');
        expect(active()!.getAttribute('aria-label')).to.equal('Wednesday 10 January 2024');
    });

    it('starts the week on the startingDay of the column', async () => {
        const { input, grid, active } = await open({ startingDay: 1 }, '2024-03-05');

        expect(Array.from(grid.querySelectorAll('[role="columnheader"]')).map((th) => th.getAttribute('aria-label'))).to.eql([
            'Monday',
            'Tuesday',
            'Wednesday',
            'Thursday',
            'Friday',
            'Saturday',
            'Sunday',
        ]);
        expect(
            Array.from(grid.querySelectorAll('.jcalendar-set-day'))
                .slice(0, 7)
                .map((td) => td.textContent)
        ).to.eql(['', '', '', '', '1', '2', '3']);

        key(input, 'Home');
        expect(active()!.getAttribute('aria-label')).to.equal('Monday 4 March 2024');
        key(input, 'End');
        expect(active()!.getAttribute('aria-label')).to.equal('Sunday 10 March 2024');
    });

    it('saves the active day with Enter, and the selection moves down', async () => {
        const { worksheet, cell, input } = await open({ format: 'DD/MM/YYYY' });
        const ends: [number, number, boolean][] = [];
        worksheet.parent.config.oneditionend = (_worksheet, _cell, x, y, _value, save) => {
            ends.push([x, y, save]);
        };

        expect(input.value).to.equal('05/03/2024');
        key(input, 'ArrowRight');
        key(document.activeElement!, 'Enter');

        expect(worksheet.editor.isOpen()).to.equal(false);
        expect(worksheet.getValueFromCoords(0, 0)).to.equal('2024-03-06 00:00:00');
        expect(cell.textContent).to.equal('06/03/2024');
        expect(ends).to.eql([[0, 0, true]]);
        expect(worksheet.selection.range).to.eql([0, 1, 0, 1]);
    });

    it('saves the active day with Tab, and the selection moves right', async () => {
        const { worksheet, input } = await open();

        key(input, 'ArrowDown');
        key(input, 'Tab');

        expect(worksheet.editor.isOpen()).to.equal(false);
        expect(worksheet.getValueFromCoords(0, 0)).to.equal('2024-03-12 00:00:00');
        expect(worksheet.selection.range).to.eql([1, 0, 1, 0]);
    });

    it('closes without saving with Escape', async () => {
        const { worksheet, input } = await open();
        const ends: boolean[] = [];
        worksheet.parent.config.oneditionend = (_worksheet, _cell, _x, _y, _value, save) => {
            ends.push(save);
        };

        key(input, 'ArrowRight');
        key(input, 'Escape');

        expect(worksheet.editor.isOpen()).to.equal(false);
        expect(worksheet.getValueFromCoords(0, 0)).to.equal('2024-03-05');
        expect(ends).to.eql([false]);
    });

    it('saves the day that is pressed', async () => {
        const { worksheet, day } = await open();

        press(day('20'));

        expect(worksheet.editor.isOpen()).to.equal(false);
        expect(worksheet.getValueFromCoords(0, 0)).to.equal('2024-03-20 00:00:00');
    });

    it('removes the date with the reset button, and saves the active day with the done button', async () => {
        const first = await open();
        press(first.button('jcalendar-reset'));
        expect(first.worksheet.editor.isOpen()).to.equal(false);
        expect(first.worksheet.getValueFromCoords(0, 0)).to.equal('');

        root.innerHTML = '';
        const second = await open();
        key(second.input, 'ArrowLeft');
        press(second.button('jcalendar-confirm'));
        expect(second.worksheet.getValueFromCoords(0, 0)).to.equal('2024-03-04 00:00:00');

        root.innerHTML = '';
        const third = await open({ resetButton: false });
        expect(third.button('jcalendar-reset').style.display).to.equal('none');
    });

    it('saves today with Enter on an empty cell, and starts with today with the today option', async () => {
        const first = await open({}, '');
        expect(first.input.value).to.equal('');
        expect(first.active()!.getAttribute('data-date')).to.equal(today());
        expect(first.active()!.classList.contains('jcalendar-today')).to.equal(true);
        key(first.input, 'Enter');
        expect(first.worksheet.getValueFromCoords(0, 0)).to.equal(`${today()} 00:00:00`);

        root.innerHTML = '';
        const second = await open({ today: true }, '');
        expect(second.input.value).to.equal(today());
    });

    it('opens on the date of a number of days, as Excel counts them', async () => {
        const { input, active } = await open({}, 45356);

        expect(input.value).to.equal('2024-03-05');
        expect(active()!.getAttribute('data-date')).to.equal('2024-03-05');
    });

    it('chooses the time with the selects, which Tab reaches before it saves', async () => {
        const { worksheet, input, calendar, day, button } = await open({ time: true, format: 'YYYY-MM-DD HH24:MI' }, '2024-03-05 10:20:00');
        const [hour, minute] = Array.from(calendar.querySelectorAll<HTMLSelectElement>('.jcalendar-select'));

        expect(input.value).to.equal('2024-03-05 10:20');
        expect(hour.value).to.equal('10');
        expect(minute.value).to.equal('20');
        expect(hour.getAttribute('aria-label')).to.equal('Hour');
        expect(minute.getAttribute('aria-label')).to.equal('Minute');
        expect((hour.parentElement as HTMLElement).style.display).to.equal('');

        // A press on a day does not close the calendar with the time
        press(day('12'));
        expect(worksheet.editor.isOpen()).to.equal(true);

        key(input, 'Tab');
        expect(document.activeElement).to.equal(hour);
        hour.value = '15';
        hour.dispatchEvent(new Event('change'));
        key(hour, 'Tab');
        expect(document.activeElement).to.equal(minute);
        key(minute, 'Tab', { shiftKey: true });
        expect(document.activeElement).to.equal(hour);
        key(hour, 'Tab', { shiftKey: true });
        expect(document.activeElement).to.equal(input);
        expect(worksheet.editor.isOpen()).to.equal(true);

        key(input, 'Tab');
        key(hour, 'Tab');
        minute.value = '45';
        minute.dispatchEvent(new Event('change'));
        key(minute, 'Tab');

        expect(worksheet.editor.isOpen()).to.equal(false);
        expect(worksheet.getValueFromCoords(0, 0)).to.equal('2024-03-12 15:45:00');
        expect(worksheet.selection.range).to.eql([1, 0, 1, 0]);

        // The update button saves the day with the time
        worksheet.openEditor(worksheet.getCell('A1'));
        const again = worksheet.getCell('A1');
        key(again.querySelector('input')!, 'ArrowRight');
        press(again.querySelector('.jcalendar-update')!);
        expect(worksheet.getValueFromCoords(0, 0)).to.equal('2024-03-13 15:45:00');
        expect(button('jcalendar-update')).to.not.equal(null);
    });

    it('reads the date that is typed in the format of the column when it is not readonly', async () => {
        const { worksheet, input, active } = await open({ readonly: false, format: 'DD/MM/YYYY' });

        expect(input.readOnly).to.equal(false);

        // The left and the right arrows, Home and End move in the text
        expect(key(input, 'ArrowLeft').defaultPrevented).to.equal(false);
        expect(key(input, 'Home').defaultPrevented).to.equal(false);

        input.value = '20/04/2024';
        input.dispatchEvent(new Event('input', { bubbles: true }));
        expect(active()!.getAttribute('data-date')).to.equal('2024-04-20');

        // A text that is not a date changes nothing
        input.value = '20/04';
        input.dispatchEvent(new Event('input', { bubbles: true }));
        expect(active()!.getAttribute('data-date')).to.equal('2024-04-20');

        key(input, 'ArrowDown');
        key(input, 'Enter');
        expect(worksheet.getValueFromCoords(0, 0)).to.equal('2024-04-27 00:00:00');

        // An empty text removes the date
        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        worksheet.openEditor(worksheet.getCell('A1'));
        const again = worksheet.getCell('A1').querySelector('input')!;
        again.value = '';
        again.dispatchEvent(new Event('input', { bubbles: true }));
        key(again, 'Enter');
        expect(worksheet.getValueFromCoords(0, 0)).to.equal('');
    });

    it('shows the months and the years when their names in the header are pressed', async () => {
        const { worksheet, input, calendar, grid, active, day } = await open();

        press(calendar.querySelector('.jcalendar-year')!);
        const years = Array.from(grid.querySelectorAll('.jcalendar-set-year'));
        expect(years.length).to.equal(25);
        expect(years[0].textContent).to.equal('2012');
        expect(grid.getAttribute('aria-label')).to.equal('2012 – 2036');
        expect(active()!.textContent).to.equal('2024');
        expect(active()!.getAttribute('aria-selected')).to.equal('true');

        // The keys move between the years, and Page Up and Page Down by a page
        key(input, 'ArrowDown');
        expect(active()!.textContent).to.equal('2029');
        key(input, 'ArrowLeft');
        expect(active()!.textContent).to.equal('2028');
        key(input, 'PageDown');
        expect(active()!.textContent).to.equal('2053');
        key(input, 'PageUp');
        press(calendar.querySelector('.jcalendar-prev button')!);
        expect(active()!.textContent).to.equal('2003');
        press(calendar.querySelector('.jcalendar-next button')!);
        expect(active()!.textContent).to.equal('2028');

        // Enter shows the days of the year
        key(input, 'Enter');
        expect(worksheet.editor.isOpen()).to.equal(true);
        expect(grid.getAttribute('aria-label')).to.equal('March 2028');

        // The months
        press(calendar.querySelector('.jcalendar-month')!);
        const months = Array.from(grid.querySelectorAll('.jcalendar-set-month'));
        expect(months.map((td) => td.textContent)).to.eql(['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']);
        expect(grid.getAttribute('aria-label')).to.equal('2028');
        expect(active()!.getAttribute('aria-label')).to.equal('March 2028');
        key(input, 'ArrowDown');
        expect(active()!.getAttribute('aria-label')).to.equal('July 2028');
        key(input, 'ArrowRight');
        key(input, 'PageDown');
        expect(active()!.getAttribute('aria-label')).to.equal('August 2029');
        press(calendar.querySelector('.jcalendar-prev button')!);
        expect(active()!.getAttribute('aria-label')).to.equal('August 2028');

        // A press on a month shows its days, and a press on a year shows the days of its month
        press(grid.querySelectorAll('.jcalendar-set-month')[1]);
        expect(grid.getAttribute('aria-label')).to.equal('February 2028');
        press(calendar.querySelector('.jcalendar-year')!);
        press(Array.from(grid.querySelectorAll('.jcalendar-set-year')).find((td) => td.textContent === '2030')!);
        expect(grid.getAttribute('aria-label')).to.equal('February 2030');

        // The previous and the next month
        press(calendar.querySelector('.jcalendar-next button')!);
        expect(grid.getAttribute('aria-label')).to.equal('March 2030');
        press(day('31'));
        expect(worksheet.getValueFromCoords(0, 0)).to.equal('2030-03-31 00:00:00');
    });

    it('disables the days out of the valid range, which cannot be chosen', async () => {
        const { worksheet, input, day } = await open({ validRange: ['2024-03-03', '2024-03-20'] });

        expect(day('2').classList.contains('jcalendar-disabled')).to.equal(true);
        expect(day('2').getAttribute('aria-disabled')).to.equal('true');
        expect(day('3').hasAttribute('aria-disabled')).to.equal(false);
        expect(day('21').getAttribute('aria-disabled')).to.equal('true');

        press(day('21'));
        expect(worksheet.editor.isOpen()).to.equal(true);

        key(input, 'End');
        key(input, 'ArrowDown');
        key(input, 'ArrowDown');
        key(input, 'Enter');
        expect(worksheet.editor.isOpen()).to.equal(false);
        expect(worksheet.getValueFromCoords(0, 0)).to.equal('2024-03-05');
    });

    it('opens again with the arrows or a press on the input', async () => {
        const { input, calendar } = await open();
        const control = (input as HTMLInputElement & { calendar: { close(ignoreEvents?: boolean): string } }).calendar;

        expect(control.close(true)).to.equal('2024-03-05');
        expect(calendar.classList.contains('jcalendar-focus')).to.equal(false);
        expect(input.getAttribute('aria-expanded')).to.equal('false');
        expect(input.hasAttribute('aria-activedescendant')).to.equal(false);

        // Other keys do nothing while it is closed
        expect(key(input, 'ArrowLeft').defaultPrevented).to.equal(false);
        key(input, 'ArrowDown');
        expect(calendar.classList.contains('jcalendar-focus')).to.equal(true);

        control.close(true);
        input.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
        expect(calendar.classList.contains('jcalendar-focus')).to.equal(true);
    });

    it('saves the date after a press outside it', async () => {
        const { worksheet, input } = await open();

        key(input, 'ArrowRight');
        document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));

        expect(worksheet.editor.isOpen()).to.equal(false);
        // The active day is not chosen, and the value does not change
        expect(worksheet.getValueFromCoords(0, 0)).to.equal('2024-03-05');
        expect(worksheet.history.records.length).to.equal(0);
    });

    it('keeps the focus in the input after a press in the calendar', async () => {
        const { calendar, day } = await open();

        const down = new MouseEvent('mousedown', { bubbles: true, cancelable: true });
        day('12').dispatchEvent(down);
        expect(down.defaultPrevented).to.equal(true);

        const select = new MouseEvent('mousedown', { bubbles: true, cancelable: true });
        calendar.querySelector('select')!.dispatchEvent(select);
        expect(select.defaultPrevented).to.equal(false);
    });

    it('has a fixed position with tableOverflow, under or over the input', async () => {
        const { calendar } = await open({}, '2024-03-05', { tableOverflow: true });
        const wrapper = calendar.firstElementChild as HTMLElement;

        expect(wrapper.style.position).to.equal('fixed');
        expect(wrapper.style.top).to.equal('2px');
    });

    it('opens over the input when there is no room under it', async () => {
        const instance = jspreadsheet(root, { worksheets: [{ data: [['2024-03-05']], columns: [{ type: 'calendar' }] }] });
        await waitForListeners();
        const worksheet = instance[0];
        const cell = worksheet.getCell('A1');
        const rect = { top: 700, bottom: 720, left: 900, right: 1000, height: 20, width: 100, x: 900, y: 700 } as DOMRect;
        const content = { top: 0, bottom: 300, left: 0, right: 300, height: 300, width: 300, x: 0, y: 0 } as DOMRect;
        const original = HTMLElement.prototype.getBoundingClientRect;
        HTMLElement.prototype.getBoundingClientRect = function (this: HTMLElement) {
            return this.classList.contains('jcalendar-content') ? content : this.tagName === 'INPUT' ? rect : original.call(this);
        };
        try {
            worksheet.openEditor(cell);
        } finally {
            HTMLElement.prototype.getBoundingClientRect = original;
        }
        const wrapper = cell.querySelector<HTMLElement>('.jcalendar > .jcalendar-container')!;

        expect(wrapper.style.top).to.equal('-322px');
        expect(parseInt(wrapper.style.left)).to.be.below(0);
        worksheet.closeEditor(cell, false);
    });

    it('opens over the input when the element that scrolls has no room under it, and stops at its edge', async () => {
        const instance = jspreadsheet(root, { worksheets: [{ data: [['2024-03-05']], columns: [{ type: 'calendar' }] }] });
        await waitForListeners();
        const worksheet = instance[0];
        const cell = worksheet.getCell('A1');
        root.style.cssText = 'position: relative; overflow: auto';
        const rect = { top: 250, bottom: 270, left: 10, right: 110, height: 20, width: 100, x: 10, y: 250 } as DOMRect;
        const content = { top: 0, bottom: 300, left: 0, right: 300, height: 300, width: 300, x: 0, y: 0 } as DOMRect;
        const scroller = { top: 0, bottom: 300 } as DOMRect;
        const original = HTMLElement.prototype.getBoundingClientRect;
        HTMLElement.prototype.getBoundingClientRect = function (this: HTMLElement) {
            if (this === root) {
                return scroller;
            }
            return this.classList.contains('jcalendar-content') ? content : this.tagName === 'INPUT' ? rect : original.call(this);
        };
        try {
            worksheet.openEditor(cell);
        } finally {
            HTMLElement.prototype.getBoundingClientRect = original;
            root.removeAttribute('style');
        }
        const wrapper = cell.querySelector<HTMLElement>('.jcalendar > .jcalendar-container')!;

        // 250 pixels over the input, where the element that scrolls starts
        expect(wrapper.style.top).to.equal('-270px');
        worksheet.closeEditor(cell, false);
    });
});
