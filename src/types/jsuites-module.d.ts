/**
 * The jsuites module for the type check, with the typings of jsuites corrected where they do not match what jsuites does.
 *
 * tsconfig.json maps the module 'jsuites' to this file. The build still bundles nothing from jsuites: it gets jsuites from
 * a global or require (see build.cjs).
 */
import type { JSuites as Base } from 'jsuites/dist/types';

/** The options of a jsuites component, the second parameter of the function that creates it */
type ComponentOptions<F> = F extends (el: HTMLElement, options: infer O) => unknown ? O : never;

/** A dropdown, as jSuites.dropdown returns it. The typings miss its close method, which ignoreEvents keeps from calling onclose. */
export type Dropdown = ReturnType<Base['dropdown']> & { close(ignoreEvents?: boolean): void };

/**
 * The options of jSuites.dropdown. jsuites takes any CSS width, the id of an item as the value, a list or null as the
 * value of a dropdown with multiple, and gives onclose the element of the dropdown, which holds the dropdown.
 */
export type DropdownOptions = Omit<ComponentOptions<Base['dropdown']>, 'width' | 'value' | 'onclose'> & {
    width?: number | string;
    value?: string | number | (string | number)[] | null;
    onclose?: (el: HTMLElement & { dropdown: Dropdown }) => void;
};

/**
 * The options of jSuites.calendar. jsuites has defaults for each of them, takes a truthy position for a fixed position,
 * and gives onclose the input.
 */
export type CalendarOptions = Partial<Omit<ComponentOptions<Base['calendar']>, 'onclose' | 'position'>> & {
    position?: boolean;
    onclose?: (el: HTMLElement) => void;
};

/** A color picker, as jSuites.color returns it. jsuites also keeps it on its element as color. */
export type Color = ReturnType<Base['color']>;

/** The options of jSuites.color. jsuites has defaults for each of them, and takes a truthy position for a fixed position. */
export type ColorOptions = Partial<Omit<ComponentOptions<Base['color']>, 'position' | 'onopen'>> & {
    position?: boolean;
    onopen?: (el: HTMLElement & { color: Color }, instance: Color) => void;
};

/**
 * The options of jSuites.picker with items of type T. jsuites takes a list as the data and any CSS width, gives render
 * the item, and gives onchange the item twice and its index as text.
 */
export type PickerOptions<T> = Omit<ComponentOptions<Base['picker']>, 'data' | 'render' | 'onchange' | 'width'> & {
    data?: T[];
    width?: number | string;
    render?: (item: T, element: HTMLElement) => string;
    onchange?: (el: HTMLElement, instance: ReturnType<Base['picker']>, reservedValue: T, value: T, index: string, e: MouseEvent) => void;
};

/**
 * The options of jSuites.toolbar. jsuites reads more properties of an item than its typings name, for example title, the
 * type 'i' and the picker options of a select, and gives updateState the arguments of update.
 */
export type ToolbarOptions = Omit<ComponentOptions<Base['toolbar']>, 'items'> & {
    items: object[];
};

export interface JSuites extends Omit<Base, 'calendar' | 'color' | 'dropdown' | 'picker' | 'toolbar'> {
    calendar(el: HTMLElement, options: CalendarOptions): ReturnType<Base['calendar']>;
    color(el: HTMLElement, options: ColorOptions): Color;
    dropdown(el: HTMLElement, options: DropdownOptions): Dropdown;
    picker<T>(el: HTMLElement, options: PickerOptions<T>): ReturnType<Base['picker']>;
    toolbar(el: HTMLElement, options: ToolbarOptions): ReturnType<Base['toolbar']>;

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
