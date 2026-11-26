import { translate } from './dictionary.js';

// Dates in the formats of the calendar columns and the masks, such as "DD/MM/YYYY HH24:MI". A date is kept as the text
// "YYYY-MM-DD HH:MI:SS", or as a number of days since 1899-12-30, as in Excel, whose decimal part is the time.

const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const weekdays = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

// The tokens of a format, the longest first, so that "MMMM" is not read as "MM" twice
const tokens = [
    'AM/PM',
    'MONTH',
    'HH24',
    'HH12',
    'YYYY',
    'MMMM',
    'DDDD',
    'MMM',
    'MON',
    'DDD',
    'DAY',
    'YY',
    'MM',
    'DD',
    'DY',
    'WD',
    'HH',
    'MI',
    'SS',
    'M',
    'D',
    'H',
];

/** The parts of a date. The month is 1 to 12. */
interface DateParts {
    year: number;
    month: number;
    day: number;
    hours: number;
    minutes: number;
    seconds: number;
}

/**
 * Split a format into its tokens and the text between them. The format is read in capitals, as "dd/mm/yyyy" and
 * "DD/MM/YYYY" are the same format.
 * @param format - The format, such as "DD/MM/YYYY".
 */
function tokenize(format: string) {
    const parts: { token?: string; text?: string }[] = [];
    const upper = format.toUpperCase();
    let i = 0;
    while (i < upper.length) {
        const token = tokens.find((t) => upper.startsWith(t, i));
        if (token) {
            parts.push({ token });
            i += token.length;
        } else {
            const last = parts[parts.length - 1];
            if (last?.text !== undefined) {
                last.text += format[i];
            } else {
                parts.push({ text: format[i] });
            }
            i++;
        }
    }
    return parts;
}

/**
 * Tell whether a format has the tokens of a date or a time, such as "DD/MM/YYYY" or "HH:MI".
 * @param format - The format.
 */
export function isDateFormat(format: string) {
    return tokenize(format).some((part) => part.token !== undefined && part.token.length > 1);
}

/**
 * The number of days since 1899-12-30 of a date, as Excel counts them. Excel counts from the 31st of December 1899, and
 * also counts the 29th of February 1900, which did not exist, so the dates before March 1900 are one day less than the
 * days since 1899-12-30: 1 is the 1st of January 1900.
 * @param parts - The date.
 */
export function toExcelNumber(parts: DateParts) {
    const days = (Date.UTC(parts.year, parts.month - 1, parts.day) - Date.UTC(1899, 11, 30)) / 86400000;
    const time = (parts.hours * 3600 + parts.minutes * 60 + parts.seconds) / 86400;
    return (days < 61 ? days - 1 : days) + time;
}

/**
 * Read the parts of a date from the text "YYYY-MM-DD HH:MI:SS", whose time may be missing, or from a number of days as
 * Excel counts them.
 * @param value - The date.
 * @returns The parts, or null when the value is not a date.
 */
export function toDateParts(value: unknown): DateParts | null {
    if (typeof value === 'number') {
        if (!Number.isFinite(value) || value < 0) {
            return null;
        }
        const days = Math.floor(value);
        const date = new Date(Date.UTC(1899, 11, 30) + (days < 61 ? days + 1 : days) * 86400000);
        // The seconds of the time, rounded so that 0.5 is 12:00:00 and not 11:59:59
        const seconds = Math.round((value - days) * 86400);
        return {
            year: date.getUTCFullYear(),
            month: date.getUTCMonth() + 1,
            day: date.getUTCDate(),
            hours: Math.floor(seconds / 3600) % 24,
            minutes: Math.floor(seconds / 60) % 60,
            seconds: seconds % 60,
        };
    }

    const match = /^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?/.exec(`${value ?? ''}`);
    if (!match) {
        return null;
    }
    const parts = {
        year: Number(match[1]),
        month: Number(match[2]),
        day: Number(match[3]),
        hours: Number(match[4] ?? 0),
        minutes: Number(match[5] ?? 0),
        seconds: Number(match[6] ?? 0),
    };
    return isValid(parts) ? parts : null;
}

/**
 * Tell whether the parts are a date that exists, such as not the 31st of April, and a time of the day.
 * @param parts - The date.
 */
function isValid(parts: DateParts) {
    const date = new Date(Date.UTC(parts.year, parts.month - 1, parts.day));
    return date.getUTCMonth() === parts.month - 1 && date.getUTCDate() === parts.day && parts.hours < 24 && parts.minutes < 60 && parts.seconds < 60;
}

/**
 * Write the number with two digits, such as "05".
 * @param n - The number.
 */
function pad(n: number) {
    return `${n}`.padStart(2, '0');
}

/**
 * Show a date in a format. The names of the months and of the days are translated with the dictionary of
 * jspreadsheet.setDictionary.
 * @param value - The date, as the text "YYYY-MM-DD HH:MI:SS" or as a number of days as Excel counts them.
 * @param format - The format, such as "DD/MM/YYYY". "YYYY-MM-DD" when it is missing.
 * @returns The date in the format, or an empty text when the value is not a date.
 */
export function formatDate(value: unknown, format?: string) {
    const parts = toDateParts(value);
    if (!parts) {
        return '';
    }

    const weekday = new Date(Date.UTC(parts.year, parts.month - 1, parts.day)).getUTCDay();
    const hours12 = parts.hours % 12 || 12;
    const month = translate(months[parts.month - 1]);
    const day = translate(weekdays[weekday]);

    const values: Record<string, string> = {
        YYYY: `${parts.year}`,
        YY: pad(parts.year % 100),
        MONTH: month.toUpperCase(),
        MMMM: month,
        MMM: month.substring(0, 3),
        MON: month.substring(0, 3).toUpperCase(),
        MM: pad(parts.month),
        M: `${parts.month}`,
        DDDD: day,
        DAY: day.toUpperCase(),
        WD: day,
        DDD: day.substring(0, 3),
        DY: day.substring(0, 3),
        DD: pad(parts.day),
        D: `${parts.day}`,
        HH24: pad(parts.hours),
        HH: pad(parts.hours),
        H: `${parts.hours}`,
        HH12: pad(hours12),
        MI: pad(parts.minutes),
        SS: pad(parts.seconds),
        'AM/PM': parts.hours < 12 ? 'AM' : 'PM',
    };

    return tokenize(format || 'YYYY-MM-DD')
        .map((part) => (part.token ? values[part.token] : part.text))
        .join('');
}

/**
 * Escape the characters of a text that a regular expression reads as a pattern.
 * @param text - The text.
 */
function escapeRegExp(text: string) {
    return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Read a date that is written in a format, such as "05/03/2024" in "DD/MM/YYYY". The names of the months are read in
 * English or as the dictionary translates them, in any case.
 * @param text - The date as it is written.
 * @param format - The format, such as "DD/MM/YYYY".
 * @returns The date as the text "YYYY-MM-DD HH:MI:SS", or an empty text when the text is not a date in the format.
 */
export function parseDate(text: unknown, format: string) {
    // The names in English and as the dictionary translates them, in pairs, cut to a length for the short names
    function names(list: string[], length?: number) {
        return list.flatMap((name) => [name, translate(name)]).map((name) => (length ? name.substring(0, length) : name));
    }
    const longMonths = names(months);
    const shortMonths = names(months, 3);
    const days = [...names(weekdays), ...names(weekdays, 3)].map(escapeRegExp).join('|');

    const parts: DateParts = { year: new Date().getFullYear(), month: 1, day: 1, hours: 0, minutes: 0, seconds: 0 };
    let pm: boolean | null = null;

    // The month of a name, which is in a pair of the list
    function month(list: string[]): [string, (match: string) => void] {
        return [list.map(escapeRegExp).join('|'), (m) => (parts.month = (list.findIndex((name) => name.toLowerCase() === m.toLowerCase()) >> 1) + 1)];
    }

    // A number of one or two digits, or of the given pattern, that sets a part of the date
    function digits(key: keyof DateParts, pattern = '\\d{1,2}'): [string, (match: string) => void] {
        return [pattern, (m) => (parts[key] = Number(m))];
    }
    // The name of a day tells nothing that the date does not tell
    const weekday: [string, (match: string) => void] = [days, () => {}];

    // The pattern of each token, and how its match changes the date
    const readers: Record<string, [string, (match: string) => void]> = {
        YYYY: digits('year', '\\d{4}'),
        // Two digits are a year of this century up to 29, as in Excel, and of the last century from 30
        YY: ['\\d{2}', (m) => (parts.year = Number(m) + (Number(m) < 30 ? 2000 : 1900))],
        MONTH: month(longMonths),
        MMMM: month(longMonths),
        MMM: month(shortMonths),
        MON: month(shortMonths),
        MM: digits('month'),
        M: digits('month'),
        DD: digits('day'),
        D: digits('day'),
        DDDD: weekday,
        DAY: weekday,
        WD: weekday,
        DDD: weekday,
        DY: weekday,
        HH24: digits('hours'),
        HH: digits('hours'),
        H: digits('hours'),
        HH12: ['\\d{1,2}', (m) => (parts.hours = Number(m) % 12)],
        MI: digits('minutes', '\\d{2}'),
        SS: digits('seconds', '\\d{2}'),
        'AM/PM': ['AM|PM', (m) => (pm = m.toUpperCase() === 'PM')],
    };

    const list = tokenize(format);
    // The text between the tokens matches itself, and any space matches any number of spaces
    const pattern = list.map((part) => (part.token ? `(${readers[part.token][0]})` : escapeRegExp(part.text!).replace(/\s+/g, '\\s+'))).join('');
    const match = new RegExp(`^\\s*${pattern}\\s*$`, 'i').exec(`${text ?? ''}`);
    if (!match) {
        return '';
    }

    let group = 1;
    for (const part of list) {
        if (part.token) {
            readers[part.token][1](match[group++]);
        }
    }
    if (pm !== null && parts.hours < 12 && pm) {
        parts.hours += 12;
    }
    if (!isValid(parts)) {
        return '';
    }

    return `${parts.year}-${pad(parts.month)}-${pad(parts.day)} ${pad(parts.hours)}:${pad(parts.minutes)}:${pad(parts.seconds)}`;
}
