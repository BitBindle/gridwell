import { formatDate, isDateFormat, parseDate, toDateParts, toExcelNumber } from './dates.js';

// The masks of the columns: a number format as in Excel, such as "#,##0.00" or "0.00%;[Red]-0.00%", a date format,
// such as "DD/MM/YYYY", or a locale with the options of Intl.NumberFormat.

/** The mask of a column. */
export interface Mask {
    /** A number format, such as "#,##0.00", or a date format, such as "DD/MM/YYYY". */
    mask?: string;
    /** A locale for Intl.NumberFormat, when there is no mask. */
    locale?: string;
    /** The options of Intl.NumberFormat for the locale. */
    options?: Intl.NumberFormatOptions;
    /** The decimal separator of the mask, when the mask does not tell it, such as "," for "#.##0,00". */
    decimal?: string;
}

/** A part of a number format. */
type Token =
    | { type: 'digit'; char: '0' | '#' | '?' }
    | { type: 'separator'; char: string }
    | { type: 'decimal' }
    | { type: 'group' }
    | { type: 'exponent'; sign: '+' | '-'; digits: number }
    | { type: 'text' }
    | { type: 'literal'; text: string };

/** A section of a number format, for the positive numbers, the negative numbers, zero or a text. */
interface Section {
    tokens: Token[];
    /** The color in brackets, such as "red" for [Red]. */
    color: string | null;
    /** The format of a date, when the section is one. */
    date: string | null;
    /** The number of percent signs, each of which multiplies the number by 100. */
    percent: number;
}

const colors = ['black', 'blue', 'cyan', 'green', 'magenta', 'red', 'white', 'yellow'];

// The masks that were read, by text
const cache = new Map<string, { sections: Section[]; decimal: string; group: string }>();

/**
 * Split a number format into its sections, at the semicolons outside quotes and brackets.
 * @param format - The format.
 */
function splitSections(format: string) {
    const sections = [''];
    let quoted = false;
    let bracket = false;
    for (let i = 0; i < format.length; i++) {
        const c = format[i];
        if (c === '\\' && !quoted) {
            sections[sections.length - 1] += c + (format[i + 1] ?? '');
            i++;
            continue;
        }
        if (c === '"' && !bracket) {
            quoted = !quoted;
        } else if (c === '[' && !quoted) {
            bracket = true;
        } else if (c === ']' && !quoted) {
            bracket = false;
        } else if (c === ';' && !quoted && !bracket) {
            sections.push('');
            continue;
        }
        sections[sections.length - 1] += c;
    }
    return sections;
}

/**
 * Read the tokens of a section of a number format. The points and the commas are separators until the decimal
 * separator of the mask is known.
 * @param text - The section.
 */
function readSection(text: string): Section {
    const tokens: Token[] = [];
    let color: string | null = null;
    let percent = 0;
    // The format without its quoted texts and brackets, to tell whether it is a date
    let plain = '';

    function literal(value: string) {
        const last = tokens[tokens.length - 1];
        if (last?.type === 'literal') {
            last.text += value;
        } else {
            tokens.push({ type: 'literal', text: value });
        }
    }

    for (let i = 0; i < text.length; i++) {
        const c = text[i];
        if (c === '[') {
            const end = text.indexOf(']', i);
            const content = text.substring(i + 1, end < 0 ? text.length : end);
            if (colors.includes(content.toLowerCase())) {
                color = content.toLowerCase();
            } else if (content.startsWith('$')) {
                // A currency of Excel, such as [$€-407], shows its symbol
                literal(content.substring(1).split('-')[0]);
            }
            i = end < 0 ? text.length : end;
        } else if (c === '"') {
            const end = text.indexOf('"', i + 1);
            literal(text.substring(i + 1, end < 0 ? text.length : end));
            i = end < 0 ? text.length : end;
        } else if (c === '\\') {
            literal(text[i + 1] ?? '');
            i++;
        } else if (c === '_') {
            // The width of the next character, as a space
            literal(' ');
            i++;
        } else if (c === '*') {
            // A character that fills the cell, which a cell does not repeat
            i++;
        } else if (c === '0' || c === '#' || c === '?') {
            tokens.push({ type: 'digit', char: c });
        } else if (c === '.' || c === ',') {
            tokens.push({ type: 'separator', char: c });
        } else if (c === '%') {
            percent++;
            literal('%');
        } else if ((c === 'E' || c === 'e') && (text[i + 1] === '+' || text[i + 1] === '-') && text[i + 2] === '0') {
            let digits = 0;
            while (text[i + 2 + digits] === '0') {
                digits++;
            }
            tokens.push({ type: 'exponent', sign: text[i + 1] as '+' | '-', digits });
            i += 1 + digits;
        } else if (c === '@') {
            tokens.push({ type: 'text' });
        } else {
            literal(c);
            plain += c;
        }
    }

    const hasDigits = tokens.some((token) => token.type === 'digit');
    const date = !hasDigits && isDateFormat(plain) ? text.replace(/\[[^\]]*\]/g, '') : null;

    return { tokens, color, date, percent };
}

/**
 * Tell which separator of a format is the decimal separator, when the mask does not say it. When the format has a point
 * and a comma, the last one is. One separator is a thousands separator when three digits follow it, and it is a comma
 * or follows a #, as in "#,##0" or "#.##0".
 * @param tokens - The tokens of the first section.
 */
function findDecimal(tokens: Token[]) {
    const separators = tokens.filter((token): token is { type: 'separator'; char: string } => token.type === 'separator');
    const chars = [...new Set(separators.map((separator) => separator.char))];
    if (chars.length > 1) {
        return separators[separators.length - 1].char;
    }
    if (chars.length === 1) {
        const index = tokens.indexOf(separators[0]);
        const after = tokens.slice(index + 1);
        const digits = after.findIndex((token) => token.type !== 'digit');
        const following = digits < 0 ? after.length : digits;
        const before = tokens[index - 1];
        const group = separators.length > 1 || (following === 3 && (chars[0] === ',' || (before?.type === 'digit' && before.char === '#')));
        return group ? (chars[0] === ',' ? '.' : ',') : chars[0];
    }
    return '.';
}

/**
 * Read a number format, and resolve its separators into the decimal separator and the thousands separator.
 * @param format - The format.
 * @param decimal - The decimal separator, when the mask gives it.
 */
function readFormat(format: string, decimal?: string) {
    const key = `${decimal ?? ''}\u0000${format}`;
    const cached = cache.get(key);
    if (cached) {
        return cached;
    }

    const sections = splitSections(format).map(readSection);
    const dec = decimal || findDecimal(sections[0].tokens);
    const group = dec === ',' ? '.' : ',';

    for (const section of sections) {
        let seenDecimal = false;
        section.tokens = section.tokens.map((token, index, list): Token => {
            if (token.type !== 'separator') {
                return token;
            }
            const digitBefore = list.slice(0, index).some((t) => t.type === 'digit');
            const digitAfter = list.slice(index + 1).some((t) => t.type === 'digit');
            if (token.char === dec && !seenDecimal && digitBefore) {
                seenDecimal = true;
                return { type: 'decimal' };
            }
            if (token.char === group && !seenDecimal && digitBefore && digitAfter) {
                return { type: 'group' };
            }
            return { type: 'literal', text: token.char };
        });
    }

    const result = { sections, decimal: dec, group };
    cache.set(key, result);
    return result;
}

/**
 * Read a value as a number: a number, or a text that is a number in JavaScript, such as "1234.5".
 * @param value - The value.
 * @returns The number, or null.
 */
function toNumber(value: unknown) {
    if (typeof value === 'number') {
        return Number.isFinite(value) ? value : null;
    }
    if (typeof value === 'string' && value.trim() !== '' && Number.isFinite(Number(value))) {
        return Number(value);
    }
    return null;
}

/**
 * Round a positive number to a number of decimals, as it is written in decimal: 1.005 gives 1.01, where toFixed gives 1.00.
 * @param value - The number.
 * @param decimals - The number of decimals.
 */
function round(value: number, decimals: number) {
    const rounded = Number(`${Math.round(Number(`${value}e${decimals}`))}e-${decimals}`);
    // A number that is too large or too small for the exponent notation keeps the rounding of toFixed
    return (Number.isFinite(rounded) ? rounded : value).toFixed(decimals);
}

/**
 * Write a number with a section of a number format.
 * @param n - The number, positive unless the section is the only one.
 * @param section - The section.
 * @param decimal - The decimal separator.
 * @param group - The thousands separator.
 */
function formatNumber(n: number, section: Section, decimal: string, group: string) {
    const tokens = section.tokens;
    const decimalIndex = tokens.findIndex((token) => token.type === 'decimal');
    const exponentIndex = tokens.findIndex((token) => token.type === 'exponent');
    const integerEnd = decimalIndex >= 0 ? decimalIndex : exponentIndex >= 0 ? exponentIndex : tokens.length;

    const integerDigits = tokens.slice(0, integerEnd).filter((token) => token.type === 'digit');
    const fractionDigits = tokens.slice(integerEnd, exponentIndex >= 0 ? exponentIndex : tokens.length).filter((token) => token.type === 'digit');
    const grouping = tokens.slice(0, integerEnd).some((token) => token.type === 'group');

    let value = Math.abs(n) * 100 ** section.percent;
    let exponent = 0;
    if (exponentIndex >= 0 && value !== 0) {
        const zeros = Math.max(1, integerDigits.length);
        exponent = Math.floor(Math.log10(value)) - zeros + 1;
        value = value / 10 ** exponent;
        // Rounding can give one digit more, such as 9.999 to 10.00
        if (Number(round(value, fractionDigits.length)) >= 10 ** zeros) {
            value /= 10;
            exponent++;
        }
    }

    const [integerText, fractionText = ''] = round(value, fractionDigits.length).split('.');

    // The decimals: a # or a ? leaves out a zero at the end, and a ? shows a space in its place
    const fraction = fractionDigits.map((token, i) => ({ char: (token as { char: string }).char, digit: fractionText[i] }));
    for (let i = fraction.length - 1; i >= 0 && fraction[i].digit === '0' && fraction[i].char !== '0'; i--) {
        fraction[i].digit = fraction[i].char === '?' ? ' ' : '';
    }

    // The integer digits, from the right. The first placeholder gets the digits that are left.
    let digits = integerText === '0' && !integerDigits.some((token) => (token as { char: string }).char === '0') ? '' : integerText;
    const integer: string[] = [];
    if (grouping) {
        const zeros = integerDigits.filter((token) => (token as { char: string }).char === '0').length;
        integer[0] = digits.padStart(zeros, '0').replace(/\B(?=(\d{3})+(?!\d))/g, group);
    } else {
        for (let i = integerDigits.length - 1; i >= 0; i--) {
            const char = (integerDigits[i] as { char: string }).char;
            if (i === 0) {
                integer[i] = digits || (char === '0' ? '0' : char === '?' ? ' ' : '');
            } else if (digits) {
                integer[i] = digits[digits.length - 1];
                digits = digits.substring(0, digits.length - 1);
            } else {
                integer[i] = char === '0' ? '0' : char === '?' ? ' ' : '';
            }
        }
    }

    let integerIndex = 0;
    let fractionIndex = 0;
    let result = '';
    for (const [index, token] of tokens.entries()) {
        if (token.type === 'digit') {
            if (index < integerEnd) {
                result += integer[integerIndex++] ?? '';
            } else {
                result += fraction[fractionIndex++]?.digit ?? '';
            }
        } else if (token.type === 'decimal') {
            result += decimal;
        } else if (token.type === 'exponent') {
            const sign = exponent < 0 ? '-' : token.sign === '+' ? '+' : '';
            result += `E${sign}${`${Math.abs(exponent)}`.padStart(token.digits, '0')}`;
        } else if (token.type === 'literal') {
            result += token.text;
        } else if (token.type === 'text') {
            result += `${n}`;
        }
    }
    return result;
}

/**
 * Write a text with a section of a number format: the section shows the text in place of its @.
 * @param text - The text.
 * @param section - The section.
 */
function formatText(text: string, section: Section) {
    return section.tokens.map((token) => (token.type === 'text' ? text : token.type === 'literal' ? token.text : '')).join('');
}

/**
 * The decimal separator and the thousands separator of a locale.
 * @param locale - The locale.
 * @param options - The options of Intl.NumberFormat.
 */
function localeSeparators(locale: string, options?: Intl.NumberFormatOptions) {
    const parts = new Intl.NumberFormat(locale, options).formatToParts(1234567.5);
    return {
        decimal: parts.find((part) => part.type === 'decimal')?.value ?? '.',
        group: parts.find((part) => part.type === 'group')?.value ?? '',
    };
}

/**
 * Show a value with a mask.
 * @param value - The value of a cell.
 * @param mask - The mask of its column.
 * @returns The text of the cell, and the color that the mask gives it, such as "red" for [Red].
 */
export function render(value: unknown, mask: Mask): { text: string; color: string | null } {
    if (value === '' || value == null) {
        return { text: '', color: null };
    }

    const n = toNumber(value);

    if (!mask.mask) {
        if (n === null || !mask.locale) {
            return { text: `${value}`, color: null };
        }
        return { text: new Intl.NumberFormat(mask.locale, mask.options).format(n), color: null };
    }

    const { sections, decimal, group } = readFormat(mask.mask, mask.decimal);

    if (n === null) {
        // A text shows in the fourth section, or in a section that is only for texts
        const section = sections[3] ?? sections.find((s) => s.tokens.some((token) => token.type === 'text'));
        if (sections[0].date) {
            const date = formatDate(value, sections[0].date);
            return { text: date || `${value}`, color: sections[0].color };
        }
        return section ? { text: formatText(`${value}`, section), color: section.color } : { text: `${value}`, color: null };
    }

    // The section of the number: the second for a negative number and the third for zero, when the format has them
    let section = sections[0];
    let number = n;
    if (n < 0 && sections[1]) {
        section = sections[1];
        number = -n;
    } else if (n === 0 && sections[2]) {
        section = sections[2];
    }

    if (section.date) {
        return { text: formatDate(number, section.date) || `${value}`, color: section.color };
    }

    let text = formatNumber(number, section, decimal, group);
    // The first section shows the sign of a negative number, unless the number shows as zero
    if (number < 0 && /[1-9]/.test(text)) {
        text = `-${text}`;
    }
    return { text, color: section.color };
}

/**
 * The value that the editor of a cell shows: a number with the decimal separator of the mask, and without thousands
 * separators or symbols, so that no digit is lost. A percentage shows as such, as in "12.5%", and a date in the format of
 * the mask.
 * @param value - The value of the cell.
 * @param mask - The mask of its column.
 */
export function toEditor(value: unknown, mask: Mask) {
    const n = toNumber(value);
    let decimal = '.';
    let percent = false;

    if (mask.mask) {
        const format = readFormat(mask.mask, mask.decimal);
        if (format.sections[0].date) {
            return formatDate(value, format.sections[0].date) || `${value ?? ''}`;
        }
        decimal = format.decimal;
        percent = format.sections[0].percent > 0;
    } else if (mask.locale) {
        decimal = localeSeparators(mask.locale, mask.options).decimal;
        percent = mask.options?.style === 'percent';
    }

    if (n === null) {
        return `${value ?? ''}`;
    }

    // 15 significant digits, so that 0.07 shows as 7% and not 7.000000000000001%
    const text = `${percent ? Number((n * 100).toPrecision(15)) : n}`.replace('.', decimal);
    return percent ? `${text}%` : text;
}

/**
 * Escape the characters of a text that a regular expression reads as a pattern.
 * @param text - The text.
 */
function escapeRegExp(text: string) {
    return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Read the value that is typed in a cell with a mask: a number, with or without the symbols and the thousands
 * separators of the mask, such as "$ 1,234.50", "(5)" or "12%", or a date in the format of a date mask, which gives the
 * number of days as Excel counts them.
 * @param text - The text that is typed.
 * @param mask - The mask of the column.
 * @returns The number, or the text when it is not a number.
 */
export function extract(text: string, mask: Mask): string | number {
    const value = `${text}`.trim();
    if (value === '') {
        return text;
    }

    let decimal = '.';
    let group = ',';
    let percent = value.includes('%');

    if (mask.mask) {
        const format = readFormat(mask.mask, mask.decimal);
        if (format.sections[0].date) {
            const parts = toDateParts(parseDate(value, format.sections[0].date));
            return parts ? toExcelNumber(parts) : text;
        }
        decimal = format.decimal;
        group = format.group;
        percent = percent || format.sections[0].percent > 0;
    } else if (mask.locale) {
        ({ decimal, group } = localeSeparators(mask.locale, mask.options));
        percent = percent || mask.options?.style === 'percent';
    }

    // Without the thousands separators and the spaces, the digits are one number, with symbols before and after it
    const plain = value.split(group).join('').replace(/\s/g, '');
    const match = new RegExp(`^([^\\d]*?)(\\d*)(?:${escapeRegExp(decimal)}(\\d*))?([^\\d]*)$`).exec(plain);
    if (!match || (match[2] === '' && !match[3])) {
        return text;
    }

    let n = Number(`${match[2] || '0'}.${match[3] || '0'}`);
    if (match[1].includes('-') || match[4].includes('-') || (match[1].includes('(') && match[4].includes(')'))) {
        n = -n;
    }
    return percent ? Number((n / 100).toPrecision(15)) : n;
}
