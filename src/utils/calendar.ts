import { translate } from './dictionary.js';
import { formatDate, parseDate, toDateParts } from './dates.js';
import { createId } from './ids.js';
import { getPopupSide } from './position.js';

/** The input of a calendar, which keeps its calendar. */
export type CalendarElement = HTMLInputElement & { calendar?: Calendar };

/** The options of a calendar. */
export type CalendarOptions = {
    /** The date that the calendar starts with, as the text "YYYY-MM-DD HH:MI:SS", or as a number of days as Excel counts them. */
    value?: string | number | null;
    /** The format of the date in the input, such as "DD/MM/YYYY". */
    format?: string;
    /** The placeholder of the input. */
    placeholder?: string;
    /** Do not let the date be typed in the input. */
    readonly?: boolean;
    /** Start with today when there is no date. */
    today?: boolean;
    /** Show the selects of the hour and the minute. */
    time?: boolean;
    /** Show the button that removes the date. */
    resetButton?: boolean;
    /** The first and the last day that can be chosen, as the text "YYYY-MM-DD". Either can be empty. */
    validRange?: [string | null | undefined, string | null | undefined] | null;
    /** The first day of the week: 0 for Sunday, 1 for Monday, … */
    startingDay?: number | null;
    /** Give the calendar a fixed position, for a calendar in an element that scrolls. */
    position?: boolean | null;
    /** Open the calendar when it is created. */
    opened?: boolean;
    onopen?: ((element: CalendarElement, calendar: Calendar) => void) | null;
    /** Called when the calendar closes, unless close is called with ignoreEvents. */
    onclose?: ((element: CalendarElement, calendar: Calendar) => void) | null;
    /**
     * Called when the date changes.
     * @param element - The input of the calendar.
     * @param value - The date, as the text "YYYY-MM-DD HH:MI:SS", or an empty text when it is removed.
     * @param calendar - The calendar.
     */
    onchange?: ((element: CalendarElement, value: string, calendar: Calendar) => void) | null;
};

/** A day that the calendar shows. The month is 1 to 12. */
type Day = { year: number; month: number; day: number; hours: number; minutes: number };

/** What the grid of the calendar shows: the days of a month, the months of a year or a page of years. */
type View = 'days' | 'months' | 'years';

/** The number of years that the grid of the years shows. */
const yearsPerPage = 25;

/**
 * Write the number with two digits, such as "05".
 * @param n - The number.
 */
function pad(n: number) {
    return `${n}`.padStart(2, '0');
}

/**
 * The number of days of a month.
 * @param year - The year.
 * @param month - The month, 1 to 12.
 */
function daysInMonth(year: number, month: number) {
    return new Date(Date.UTC(year, month, 0)).getUTCDate();
}

/**
 * The date of a day as the text "YYYY-MM-DD".
 * @param day - The day.
 */
function toIso(day: Pick<Day, 'year' | 'month' | 'day'>) {
    return `${day.year}-${pad(day.month)}-${pad(day.day)}`;
}

/**
 * A calendar under an input: the days of a month in a grid, and buttons that remove the date and close the calendar.
 * The input is an ARIA combobox whose popup is the grid, and the focus stays in the input: aria-activedescendant
 * names the day that the keys are on. The arrows move between the days, Page Up and Page Down to the previous and the
 * next month (with Shift, year), Home and End to the start and the end of the week, and Enter and Tab choose the day
 * and close the calendar. When the date can be typed, the left and the right arrows, Home and End move in the text.
 * Escape closes it without changing the date, and so does a press outside it.
 *
 * The editor of a calendar column uses it for the cell. The input keeps the calendar as its calendar property, and the
 * calendar has the classes of the jSuites calendar.
 */
export class Calendar {
    /** The type, as jSuites names its components. */
    readonly type = 'calendar';

    /** The input of the calendar. */
    readonly element: CalendarElement;

    /** The options of the calendar, with their defaults. */
    readonly options: Required<Omit<CalendarOptions, 'value'>> & { value: string };

    /** The element of the calendar, which holds the grid and the buttons. Its position can be set. */
    readonly content: HTMLElement;

    // The element that shows the calendar while it has the class jcalendar-focus
    private readonly container: HTMLElement & { calendar?: Calendar };

    // The element of content that is placed under the input
    private readonly wrapper: HTMLElement;

    // The month and the year in the header, which open the grid of the months and of the years
    private readonly monthLabel: HTMLElement;
    private readonly yearLabel: HTMLElement;

    // The grid
    private readonly grid: HTMLTableElement;
    private readonly body: HTMLTableSectionElement;

    // The selects of the time
    private readonly hour: HTMLSelectElement;
    private readonly minute: HTMLSelectElement;

    // The day that the keys are on, and what the grid shows
    private cursor: Day;
    private view: View = 'days';

    constructor(element: CalendarElement, options: CalendarOptions = {}) {
        this.element = element;
        this.options = {
            value: '',
            format: options.format || 'YYYY-MM-DD',
            placeholder: options.placeholder ?? '',
            readonly: options.readonly ?? true,
            today: options.today ?? false,
            time: options.time ?? false,
            resetButton: options.resetButton ?? true,
            validRange: options.validRange ?? null,
            startingDay: options.startingDay ?? null,
            position: options.position ?? null,
            opened: options.opened ?? false,
            onopen: options.onopen ?? null,
            onclose: options.onclose ?? null,
            onchange: options.onchange ?? null,
        };

        const listId = createId('jss-calendar');

        element.classList.add('jcalendar-input');
        element.readOnly = this.options.readonly;
        element.autocomplete = 'off';
        if (this.options.placeholder) {
            element.placeholder = this.options.placeholder;
        }
        element.setAttribute('role', 'combobox');
        element.setAttribute('aria-haspopup', 'grid');
        element.setAttribute('aria-expanded', 'false');
        element.setAttribute('aria-controls', listId);

        this.container = document.createElement('div');
        this.container.className = 'jcalendar jcalendar-container';

        this.wrapper = document.createElement('div');
        this.wrapper.className = 'jcalendar-container';
        this.container.appendChild(this.wrapper);

        this.content = document.createElement('div');
        this.content.className = 'jcalendar-content';
        this.wrapper.appendChild(this.content);

        // The buttons that remove the date and that close the calendar
        const controls = document.createElement('div');
        controls.className = 'jcalendar-controls';
        controls.style.borderBottom = '1px solid #ddd';
        const reset = this.createButton('jcalendar-reset', translate('Reset'), () => this.reset());
        if (!this.options.resetButton) {
            reset.style.display = 'none';
        }
        controls.appendChild(reset);
        controls.appendChild(this.createButton('jcalendar-confirm', translate('Done'), () => this.confirm()));
        this.content.appendChild(controls);

        const tables = document.createElement('div');
        tables.className = 'jcalendar-table';
        this.content.appendChild(tables);

        // The header, with the buttons of the previous and the next page, and the month and the year
        const header = document.createElement('table');
        header.className = 'jcalendar-nav';
        header.setAttribute('role', 'presentation');
        header.setAttribute('cellpadding', '0');
        header.setAttribute('cellspacing', '0');
        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        const prev = document.createElement('td');
        prev.className = 'jcalendar-prev';
        prev.appendChild(this.createButton('', translate('Previous'), () => this.page(-1), true));
        const title = document.createElement('td');
        title.className = 'jcalendar-header';
        this.monthLabel = document.createElement('span');
        this.monthLabel.className = 'jcalendar-month';
        this.monthLabel.addEventListener('click', () => this.show('months'));
        this.yearLabel = document.createElement('span');
        this.yearLabel.className = 'jcalendar-year';
        this.yearLabel.addEventListener('click', () => this.show('years'));
        title.appendChild(this.monthLabel);
        title.appendChild(this.yearLabel);
        const next = document.createElement('td');
        next.className = 'jcalendar-next';
        next.appendChild(this.createButton('', translate('Next'), () => this.page(1), true));
        headerRow.appendChild(prev);
        headerRow.appendChild(title);
        headerRow.appendChild(next);
        thead.appendChild(headerRow);
        header.appendChild(thead);
        tables.appendChild(header);

        // The grid of the days, the months or the years
        this.grid = document.createElement('table');
        this.grid.id = listId;
        this.grid.setAttribute('role', 'grid');
        this.grid.setAttribute('cellpadding', '0');
        this.grid.setAttribute('cellspacing', '0');
        this.body = document.createElement('tbody');
        this.grid.appendChild(this.body);
        this.grid.addEventListener('click', (e) => {
            const td = (e.target as Element).closest('td');
            if (td?.id) {
                this.pressCell(td);
            }
        });
        tables.appendChild(this.grid);

        // The time, and the button that chooses the day with the time
        const footer = document.createElement('div');
        footer.className = 'jcalendar-controls';
        const time = document.createElement('div');
        time.className = 'jcalendar-time';
        time.style.maxWidth = '140px';
        this.hour = this.createSelect(24, translate('Hour'), (value) => (this.cursor.hours = value));
        this.minute = this.createSelect(60, translate('Minute'), (value) => (this.cursor.minutes = value));
        time.appendChild(this.hour);
        time.appendChild(this.minute);
        if (!this.options.time) {
            time.style.display = 'none';
        }
        footer.appendChild(time);
        const update = document.createElement('div');
        update.style.flexGrow = '10';
        update.appendChild(this.createButton('jcalendar-update', translate('Update'), () => this.confirm()));
        footer.appendChild(update);
        this.content.appendChild(footer);

        // A press in the calendar keeps the focus in the input, but the selects of the time take it
        this.content.addEventListener('mousedown', (e) => {
            if ((e.target as Element).tagName !== 'SELECT') {
                e.preventDefault();
            }
        });
        this.content.addEventListener('keydown', this.keydown);
        element.addEventListener('keydown', this.keydown);
        element.addEventListener('input', () => this.typeDate());
        element.addEventListener('mouseup', () => this.open());

        element.parentNode!.insertBefore(this.container, element.nextSibling);
        element.calendar = this;
        this.container.calendar = this;

        this.cursor = this.toDay(null);
        this.setValue(options.value || (this.options.today ? this.now() : ''), true);

        if (this.options.opened) {
            this.open();
        }
    }

    /** Tell whether the calendar is open. */
    isOpened() {
        return this.container.classList.contains('jcalendar-focus');
    }

    /** Open the calendar under the input, or over it when there is no room under it, on the month of the date. */
    open() {
        if (this.isOpened()) {
            return;
        }

        this.cursor = this.toDay(this.options.value);
        this.view = 'days';
        this.render();

        this.container.classList.add('jcalendar-focus');
        this.element.setAttribute('aria-expanded', 'true');
        document.addEventListener('mousedown', this.pressOutside, true);
        document.addEventListener('touchstart', this.pressOutside, true);

        // Under the input, or over it when there is no room under it in the window or in the element that scrolls. The
        // calendar is two pixels away from the input, and it stops at the edge of the room when it does not fit.
        const rect = this.element.getBoundingClientRect();
        const rectContent = this.content.getBoundingClientRect();
        const { above, room } = getPopupSide(this.element, this.wrapper, this.options.position === true, rectContent.height + 2);
        const offset = Math.min(rectContent.height + 2, room);
        if (this.options.position) {
            this.wrapper.style.position = 'fixed';
            this.wrapper.style.top = `${above ? rect.top - offset : rect.bottom + 2}px`;
            this.wrapper.style.left = `${rect.left}px`;
        } else {
            this.wrapper.style.top = above ? `${-(rect.height + offset)}px` : '2px';

            if (window.innerWidth < rect.left + rectContent.width) {
                this.wrapper.style.left = `${window.innerWidth - (rect.left + rectContent.width + 20)}px`;
            } else {
                this.wrapper.style.left = '0px';
            }
        }

        this.options.onopen?.(this.element, this);

        if (this.element.isConnected && document.activeElement !== this.element) {
            this.element.focus({ preventScroll: true });
        }
    }

    /**
     * Close the calendar.
     * @param ignoreEvents - Do not call onclose.
     * @returns The date, as the text "YYYY-MM-DD HH:MI:SS" or as it was given, or an empty text.
     */
    close(ignoreEvents?: boolean) {
        document.removeEventListener('mousedown', this.pressOutside, true);
        document.removeEventListener('touchstart', this.pressOutside, true);

        if (this.isOpened()) {
            this.container.classList.remove('jcalendar-focus');
            this.element.setAttribute('aria-expanded', 'false');
            this.element.removeAttribute('aria-activedescendant');
            // The text that was typed shows as the date
            this.element.value = formatDate(this.options.value, this.options.format);

            if (!ignoreEvents) {
                this.options.onclose?.(this.element, this);
            }
        }

        return this.options.value;
    }

    /** The date, as the text "YYYY-MM-DD HH:MI:SS" or as it was given, or an empty text when there is none. */
    getValue() {
        return this.options.value;
    }

    /**
     * Change the date, show it in the input in the format, and call onchange.
     * @param value - The date, as the text "YYYY-MM-DD HH:MI:SS" or "YYYY-MM-DD", which it keeps, or as a number of days
     * as Excel counts them. A text in the format of the calendar is read too. An empty value removes the date.
     * @param ignoreEvents - Do not call onchange.
     */
    setValue(value: string | number | null | undefined, ignoreEvents?: boolean) {
        // A date in the ISO format stays as it is, so that closing the calendar without a choice changes nothing
        const parts = toDateParts(value) || toDateParts(parseDate(value, this.options.format));
        let date = '';
        if (typeof value === 'string' && toDateParts(value)) {
            date = value;
        } else if (parts) {
            date = `${toIso(parts)} ${pad(parts.hours)}:${pad(parts.minutes)}:${pad(parts.seconds)}`;
        }

        this.element.value = formatDate(date, this.options.format);

        if (date !== this.options.value) {
            this.options.value = date;
            if (!ignoreEvents) {
                this.options.onchange?.(this.element, date, this);
            }
        }

        if (this.isOpened()) {
            this.cursor = this.toDay(date);
            this.render();
        }
    }

    /** Remove the date and close the calendar. */
    reset() {
        this.setValue('');
        this.close();
    }

    /**
     * Tell whether a day can be chosen: it is in the valid range.
     * @param day - The day.
     */
    isValid(day: Pick<Day, 'year' | 'month' | 'day'>) {
        const range = this.options.validRange;
        const iso = toIso(day);
        return !range || ((!range[0] || iso >= range[0].substring(0, 10)) && (!range[1] || iso <= range[1].substring(0, 10)));
    }

    /** Choose the day of the keys, with the time of the selects, and close the calendar. A day out of the range is not chosen. */
    private confirm() {
        if (this.options.readonly || this.element.value.trim() !== '') {
            this.choose();
        } else {
            // The date that was typed is removed
            this.setValue('');
        }
        this.close();
    }

    /** Choose the day of the keys, with the time of the selects. A day out of the range is not chosen. */
    private choose() {
        const c = this.cursor;
        if (this.isValid(c)) {
            const time = this.options.time ? `${pad(c.hours)}:${pad(c.minutes)}:00` : '00:00:00';
            this.setValue(`${toIso(c)} ${time}`);
        }
    }

    /** Read the date that is typed in the input, and show its month. */
    private typeDate() {
        const parts = toDateParts(parseDate(this.element.value, this.options.format));
        if (parts && this.isOpened()) {
            this.cursor = {
                ...parts,
                hours: this.options.time ? parts.hours : this.cursor.hours,
                minutes: this.options.time ? parts.minutes : this.cursor.minutes,
            };
            this.view = 'days';
            this.render();
        }
    }

    /**
     * Show the grid of the days, the months or the years.
     * @param view - What the grid shows.
     */
    private show(view: View) {
        this.view = view;
        this.render();
    }

    /**
     * Go to the previous or the next month, year or page of years.
     * @param step - -1 or 1.
     */
    private page(step: number) {
        if (this.view === 'days') {
            this.moveMonths(step);
        } else if (this.view === 'months') {
            this.moveMonths(step * 12);
        } else {
            this.cursor.year += step * yearsPerPage;
        }
        this.render();
    }

    /**
     * Choose the cell of the grid that is pressed: a day closes the calendar unless it has the time, and a month or a
     * year shows its days.
     * @param td - The cell.
     */
    private pressCell(td: HTMLTableCellElement) {
        const [year, month, day] = td.getAttribute('data-date')!.split('-').map(Number);
        if (this.view === 'days') {
            if (!this.isValid({ year, month, day })) {
                return;
            }
            this.cursor = { ...this.cursor, year, month, day };
            if (this.options.time) {
                this.render();
            } else {
                this.choose();
                this.close();
            }
        } else {
            this.cursor.year = year;
            if (this.view === 'months') {
                this.cursor.month = month;
            }
            this.cursor.day = Math.min(this.cursor.day, daysInMonth(this.cursor.year, this.cursor.month));
            this.show('days');
        }
    }

    /**
     * Move the day of the keys by a number of months, on the same day or the last day of the month.
     * @param months - The number of months, negative to go back.
     */
    private moveMonths(months: number) {
        const index = this.cursor.year * 12 + this.cursor.month - 1 + months;
        this.cursor.year = Math.floor(index / 12);
        this.cursor.month = (index % 12) + 1;
        this.cursor.day = Math.min(this.cursor.day, daysInMonth(this.cursor.year, this.cursor.month));
    }

    /**
     * Move the day of the keys by a number of days.
     * @param days - The number of days, negative to go back.
     */
    private moveDays(days: number) {
        const date = new Date(Date.UTC(this.cursor.year, this.cursor.month - 1, this.cursor.day + days));
        this.cursor.year = date.getUTCFullYear();
        this.cursor.month = date.getUTCMonth() + 1;
        this.cursor.day = date.getUTCDate();
    }

    /** The day of the week of the day of the keys, counted from the first day of the week. */
    private weekday() {
        const day = new Date(Date.UTC(this.cursor.year, this.cursor.month - 1, this.cursor.day)).getUTCDay();
        return (day - (this.options.startingDay || 0) + 7) % 7;
    }

    /**
     * The day of a date, or today with the time of now when there is no date.
     * @param value - The date.
     */
    private toDay(value: string | null): Day {
        const parts = toDateParts(value);
        if (parts) {
            return { year: parts.year, month: parts.month, day: parts.day, hours: parts.hours, minutes: parts.minutes };
        }
        const now = new Date();
        return { year: now.getFullYear(), month: now.getMonth() + 1, day: now.getDate(), hours: now.getHours(), minutes: now.getMinutes() };
    }

    /** Today, as the text "YYYY-MM-DD HH:MI:SS". */
    private now() {
        const now = this.toDay(null);
        return `${toIso(now)} ${pad(now.hours)}:${pad(now.minutes)}:00`;
    }

    /** Show the header and the grid of the view, with the day of the keys active. */
    private render() {
        const c = this.cursor;
        const firstOfMonth = toIso({ year: c.year, month: c.month, day: 1 });
        this.monthLabel.textContent = formatDate(firstOfMonth, 'MMM');
        this.yearLabel.textContent = `${c.year}`;
        this.body.textContent = '';

        const today = toIso(this.toDay(null));
        const value = this.options.value.substring(0, 10);
        let active: HTMLTableCellElement | null = null;

        const addCell = (row: HTMLTableRowElement, className: string, date: string, text: string, label: string, isActive: boolean) => {
            const td = document.createElement('td');
            td.className = className;
            td.id = `${this.grid.id}-${date}`;
            td.setAttribute('role', 'gridcell');
            td.setAttribute('data-date', date);
            td.setAttribute('aria-label', label);
            td.textContent = text;
            if (isActive) {
                td.classList.add('jcalendar-selected');
                active = td;
            }
            row.appendChild(td);
            return td;
        };

        if (this.view === 'days') {
            this.grid.setAttribute('aria-label', formatDate(firstOfMonth, 'MMMM YYYY'));

            // The short names of the days of the week, named by their full names
            const weekdays = document.createElement('tr');
            weekdays.setAttribute('align', 'center');
            for (let i = 0; i < 7; i++) {
                // 2023-01-01 is a Sunday
                const date = `2023-01-${pad(1 + ((i + (this.options.startingDay || 0)) % 7))}`;
                const th = document.createElement('td');
                th.className = 'jcalendar-weekday';
                th.setAttribute('role', 'columnheader');
                th.setAttribute('aria-label', formatDate(date, 'DDDD'));
                th.textContent = formatDate(date, 'DDD');
                weekdays.appendChild(th);
            }
            this.body.appendChild(weekdays);

            const first = (new Date(Date.UTC(c.year, c.month - 1, 1)).getUTCDay() - (this.options.startingDay || 0) + 7) % 7;
            const days = daysInMonth(c.year, c.month);
            for (let week = 0; week < 6; week++) {
                const row = document.createElement('tr');
                row.setAttribute('align', 'center');
                for (let i = 0; i < 7; i++) {
                    const day = week * 7 + i - first + 1;
                    if (day < 1 || day > days) {
                        const td = document.createElement('td');
                        td.className = 'jcalendar-set-day';
                        td.setAttribute('role', 'gridcell');
                        row.appendChild(td);
                        continue;
                    }
                    const date = toIso({ year: c.year, month: c.month, day });
                    const td = addCell(row, 'jcalendar-set-day', date, `${day}`, formatDate(date, 'DDDD D MMMM YYYY'), day === c.day);
                    td.setAttribute('aria-selected', date === value ? 'true' : 'false');
                    if (date === today) {
                        td.classList.add('jcalendar-today');
                    }
                    if (!this.isValid({ year: c.year, month: c.month, day })) {
                        td.classList.add('jcalendar-disabled');
                        td.setAttribute('aria-disabled', 'true');
                    }
                }
                this.body.appendChild(row);
            }
        } else if (this.view === 'months') {
            this.grid.setAttribute('aria-label', `${c.year}`);
            let row = document.createElement('tr');
            for (let month = 1; month <= 12; month++) {
                const date = toIso({ year: c.year, month, day: 1 });
                const td = addCell(row, 'jcalendar-set-month', date, formatDate(date, 'MMM'), formatDate(date, 'MMMM YYYY'), month === c.month);
                td.setAttribute('aria-selected', date.substring(0, 7) === value.substring(0, 7) ? 'true' : 'false');
                if (month % 4 === 0) {
                    this.body.appendChild(row);
                    row = document.createElement('tr');
                }
            }
        } else {
            const start = c.year - Math.floor(yearsPerPage / 2);
            this.grid.setAttribute('aria-label', `${start} – ${start + yearsPerPage - 1}`);
            let row = document.createElement('tr');
            for (let i = 0; i < yearsPerPage; i++) {
                const year = start + i;
                const date = toIso({ year, month: 1, day: 1 });
                const td = addCell(row, 'jcalendar-set-year', date, `${year}`, `${year}`, year === c.year);
                td.setAttribute('aria-selected', `${year}` === value.substring(0, 4) ? 'true' : 'false');
                if ((i + 1) % 5 === 0) {
                    this.body.appendChild(row);
                    row = document.createElement('tr');
                }
            }
        }

        this.hour.value = `${c.hours}`;
        this.minute.value = `${c.minutes}`;

        if (active) {
            this.element.setAttribute('aria-activedescendant', (active as HTMLTableCellElement).id);
        }
    }

    /**
     * A button of the calendar.
     * @param className - The class of the button.
     * @param text - Its text, or its name when it has no text.
     * @param onclick - What it does.
     * @param named - The text is the name of the button, which shows no text.
     */
    private createButton(className: string, text: string, onclick: () => void, named?: boolean) {
        const button = document.createElement('button');
        button.type = 'button';
        if (className) {
            button.className = className;
        }
        if (named) {
            button.setAttribute('aria-label', text);
            button.tabIndex = -1;
        } else {
            button.textContent = text;
        }
        button.addEventListener('click', onclick);
        return button;
    }

    /**
     * A select of the hour or of the minute.
     * @param count - The number of options, from 0.
     * @param label - Its name.
     * @param onchange - What a new option changes.
     */
    private createSelect(count: number, label: string, onchange: (value: number) => void) {
        const select = document.createElement('select');
        select.className = 'jcalendar-select';
        select.setAttribute('aria-label', label);
        for (let i = 0; i < count; i++) {
            const option = document.createElement('option');
            option.value = `${i}`;
            option.textContent = pad(i);
            select.appendChild(option);
        }
        select.addEventListener('change', () => onchange(Number(select.value)));
        return select;
    }

    /** The keys on the input and in the calendar. */
    private keydown = (e: KeyboardEvent) => {
        const target = e.target as HTMLElement;
        const onInput = target === this.element;

        if (!this.isOpened()) {
            if (onInput && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) {
                this.open();
                e.preventDefault();
            }
            return;
        }

        // The keys of the grid are on the input
        const editable = !this.options.readonly;
        let handled = true;

        switch (e.key) {
            case 'Escape':
                // The editor of a cell closes without saving on Escape, so onclose is not called
                this.close(true);
                return;
            case 'Enter':
                if (onInput && this.view !== 'days') {
                    this.show('days');
                } else {
                    // The editor of a cell saves the date, and the key goes on to the grid
                    this.confirm();
                    return;
                }
                break;
            case 'Tab':
                // With the time, Tab moves between the input and the selects of the time before it chooses the day
                if (this.options.time && !e.shiftKey && target !== this.minute) {
                    (target === this.hour ? this.minute : this.hour).focus();
                } else if (this.options.time && e.shiftKey && target !== this.element) {
                    (target === this.minute ? this.hour : this.element).focus();
                } else {
                    this.confirm();
                    return;
                }
                break;
            default:
                handled = onInput && this.moveCursor(e, editable);
        }

        if (handled) {
            e.preventDefault();
            e.stopPropagation();
        }
    };

    /**
     * Move the day, the month or the year of the keys with an arrow, Page Up, Page Down, Home or End.
     * @param e - The key.
     * @param editable - The date can be typed, so the left and the right arrows, Home and End move in the text.
     * @returns Whether the key moved it.
     */
    private moveCursor(e: KeyboardEvent, editable: boolean) {
        const inText = editable && this.view === 'days';
        const columns = this.view === 'days' ? 7 : this.view === 'months' ? 4 : 5;
        const steps: Record<string, number | undefined> = {
            ArrowLeft: inText ? undefined : -1,
            ArrowRight: inText ? undefined : 1,
            ArrowUp: -columns,
            ArrowDown: columns,
        };

        if (steps[e.key] !== undefined) {
            const step = steps[e.key]!;
            if (this.view === 'days') {
                this.moveDays(step);
            } else if (this.view === 'months') {
                this.moveMonths(step);
            } else {
                this.cursor.year += step;
            }
        } else if (e.key === 'PageUp' || e.key === 'PageDown') {
            const step = e.key === 'PageUp' ? -1 : 1;
            if (this.view === 'days' && e.shiftKey) {
                this.moveMonths(step * 12);
            } else {
                this.page(step);
                return true;
            }
        } else if ((e.key === 'Home' || e.key === 'End') && this.view === 'days' && !inText) {
            this.moveDays(e.key === 'Home' ? -this.weekday() : 6 - this.weekday());
        } else {
            return false;
        }

        this.render();
        return true;
    }

    /** A press outside the calendar and its input closes it. */
    private pressOutside = (e: Event) => {
        const path = e.composedPath();
        if (!path.includes(this.container) && !path.includes(this.element)) {
            this.close();
        }
    };
}
