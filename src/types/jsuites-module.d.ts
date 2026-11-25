/**
 * The jsuites module for the type check, with the typings of jsuites corrected where they do not match what jsuites does.
 *
 * tsconfig.json maps the module 'jsuites' to this file. The build still bundles nothing from jsuites: it gets jsuites from
 * a global or require (see build.cjs).
 */
import type { JSuites as Base } from 'jsuites/dist/types';
import type { CalendarFunctions, MaskFunctions } from './jsuites';

/** The options of a jsuites component, the second parameter of the function that creates it */
type ComponentOptions<F> = F extends (el: HTMLElement, options: infer O) => unknown ? O : never;

/** A dropdown, as jSuites.dropdown returns it */
export type Dropdown = ReturnType<Base['dropdown']>;

/**
 * The options of jSuites.dropdown. jsuites takes any CSS width, the id of an item as the value, a list or null as the
 * value of a dropdown with multiple, and gives onclose the element of the dropdown, which holds the dropdown.
 */
export type DropdownOptions = Omit<ComponentOptions<Base['dropdown']>, 'width' | 'value' | 'onclose'> & {
    width?: number | string;
    value?: string | number | (string | number)[] | null;
    onclose?: (el: HTMLElement & { dropdown: Dropdown }) => void;
};

/** The options of jSuites.contextmenu. jsuites has defaults for each of them, and onclick is a function. */
export type ContextmenuOptions = Partial<Omit<ComponentOptions<Base['contextmenu']>, 'onclick'>> & {
    onclick?: (instance: ReturnType<Base['contextmenu']>, e: MouseEvent) => void;
};

/**
 * The options of jSuites.calendar. jsuites has defaults for each of them, takes a truthy position for a fixed position,
 * and gives onclose the input.
 */
export type CalendarOptions = Partial<Omit<ComponentOptions<Base['calendar']>, 'onclose' | 'position'>> & {
    position?: boolean;
    onclose?: (el: HTMLElement) => void;
};

/** The options of jSuites.color. jsuites has defaults for each of them, and takes a truthy position for a fixed position. */
export type ColorOptions = Partial<Omit<ComponentOptions<Base['color']>, 'position'>> & {
    position?: boolean;
};

export interface JSuites extends Omit<Base, 'calendar' | 'color' | 'contextmenu' | 'dropdown' | 'mask'> {
    calendar: ((el: HTMLElement, options: CalendarOptions) => ReturnType<Base['calendar']>) & CalendarFunctions;
    color(el: HTMLElement, options: ColorOptions): ReturnType<Base['color']>;
    contextmenu(el: HTMLElement, options: ContextmenuOptions): ReturnType<Base['contextmenu']>;
    dropdown(el: HTMLElement, options: DropdownOptions): Dropdown;
    mask: Base['mask'] & MaskFunctions;

    /** Translate a text with the dictionary from setDictionary. */
    translate(text: string): string;

    /** Make a random unique id. */
    guid(): string;

    /** Make an image upload control in the element. */
    image(element: HTMLElement, options?: object): unknown;

    /** Make one request, or several requests and call complete when all have finished. */
    ajax(options: object | object[], complete?: () => void): unknown;
}

declare const jSuites: JSuites;

export default jSuites;
