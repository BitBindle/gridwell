/**
 * jsuites has these methods, but its typings do not declare them.
 */

// This file must be a module, so that the declaration below adds to the JSuites interface instead of replacing the module
export {};

/** The mask options that getMask makes from a column, in the jSuites.mask format. */
export interface MaskOptions {
    mask?: string;
    locale?: string;
    options?: { decimal?: string } & Record<string, unknown>;
    /** The input that the mask applies to. */
    input?: HTMLElement;
}

/** jSuites.calendar has these functions, but its typings miss them. */
export interface CalendarFunctions {
    /** Read a date from a text in the format. Returns 'YYYY-MM-DD hh:mm:ss', or '' when the text is not a date. */
    extractDateFromString(date: unknown, format: string): string;

    /** Format a date. With a format that has [h], it returns the number of hours, a number when the value is empty. */
    getDateString(value: unknown, options?: string | { format?: string; mask?: string }): string | number;
}

/** jSuites.mask has these functions, but its typings miss them. */
export interface MaskFunctions {
    /** Format a value with the mask. A blank value (null, undefined or '') is returned as it is. */
    render(value: unknown, options: MaskOptions, fullMask?: boolean): string | null | undefined;

    /** Read the raw value from a masked text. The code calls it with values that are not blank, for which it returns the options with the value. */
    extract(value: unknown, options: MaskOptions, returnObject: true): MaskOptions & { value: string | number };
}

declare module 'jsuites/dist/types' {
    interface JSuites {
        /** Translate a text with the dictionary from setDictionary. */
        translate(text: string): string;

        /** Make a random unique id. */
        guid(): string;

        /** Make an image upload control in the element. */
        image(element: HTMLElement, options?: object): unknown;

        /** Make one request, or several requests and call complete when all have finished. */
        ajax(options: object | object[], complete?: () => void): unknown;
    }
}
