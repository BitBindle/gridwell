export = jspreadsheet;
export as namespace jspreadsheet;

declare const jspreadsheet: jspreadsheet.JSpreadsheet;

declare namespace jspreadsheet {
    type CellValue = string | number | boolean;

    type DropdownSourceItem =
        | string
        | number
        | {
              id: string | number;
              name: string;
              title?: string;
              image?: string;
              group?: string;
          };

    interface CalendarOptions {
        /**
         * @default "YYYY-MM-DD"
         */
        format?: string;

        /** Placeholder. */
        placeholder?: string;

        /**
         * Do not let the date be typed in the editor. When it is false, the date can be typed in the format.
         * @default true
         */
        readonly?: boolean;

        /**
         * Show the reset button.
         * @default true
         */
        resetButton?: boolean;

        /**
         * Show timepicker.
         * @default false
         */
        time?: boolean;

        /** Start with today when the cell has no date. */
        today?: boolean;

        /**
         * The first and the last day that can be chosen, as the text "YYYY-MM-DD". Either can be empty, for no limit.
         * The other days are shown disabled.
         */
        validRange?: [string | null | undefined, string | null | undefined];

        /**
         * The first day of the week in the calendar: 0 for Sunday, 1 for Monday, …
         * @default 0
         */
        startingDay?: number;
    }

    interface CustomEditor {
        /**
         * Event responsible for closing the editor of a cell with a custom editor. Without it, closing the editor keeps the value of the cell.
         * @param cell - Td tag whose editor should close.
         * @param save - If true, the value returned by this event will be the cell's new value. Otherwise, the value returned by this event is ignored.
         * @param x - Cell column index.
         * @param y - Cell row index.
         * @param instance - Worksheet instance.
         * @param options - Column configuration object.
         * @returns New cell value.
         */
        closeEditor?: (cell: HTMLTableCellElement, save: boolean, x: number, y: number, instance: WorksheetInstance, options: Column) => CellValue | undefined;

        /**
         * Event called when creating new cells.
         * @param cell - HTML element prepared to be the new cell.
         * @param value - Cell value.
         * @param x - Cell column index.
         * @param y - Cell row index.
         * @param instance - Worksheet instance.
         * @param options - Column configuration object.
         * @returns HTML element that will be the new cell.
         */
        createCell?: (cell: HTMLTableCellElement, value: CellValue, x: number, y: number, instance: WorksheetInstance, options: Column) => HTMLTableCellElement;

        /**
         * Event responsible for opening the editor of a cell with a custom editor. Without it, the cells of the column cannot be edited.
         * @param cell - Td tag whose editor should open.
         * @param value - Cell value.
         * @param x - Cell column index.
         * @param y - Cell row index.
         * @param instance - Worksheet instance.
         * @param options - Column configuration object.
         * @param e - Event that called this method.
         */
        openEditor?: (
            cell: HTMLTableCellElement,
            value: CellValue,
            x: number,
            y: number,
            instance: WorksheetInstance,
            options: Column,
            e: KeyboardEvent | MouseEvent | TouchEvent | undefined
        ) => void;

        /**
         * Event called before changing the value of a cell.
         *
         * The returned value will be the cell's new value.
         * @param cell - Cell whose value has changed.
         * @param value - New value.
         * @param x - Cell column index.
         * @param y - Cell row index.
         * @param instance - Worksheet instance.
         * @param options - Column configuration object.
         */
        updateCell?: (
            cell: HTMLTableCellElement,
            value: CellValue | undefined,
            x: number,
            y: number,
            instance: WorksheetInstance,
            options: Column
        ) => CellValue | undefined;
    }

    type HorizontalAlign = 'center' | 'left' | 'right' | 'justify';

    interface BaseColumn {
        /** Cell alignment. */
        align?: HorizontalAlign;

        decimal?: string;

        /** The mask of the cells when the column has no {@link mask}. */
        format?: string;

        /** Cell mask. */
        mask?: string;

        /** Name used to refer to this column if the data is arranged as an array of objects. */
        name?: string;

        /**
         * Prevent user from changing cell values.
         * @default false
         */
        readOnly?: boolean;

        /**
         * Defines a modification that must be made to a cell value before it is displayed in the spreadsheet.
         * @param cell - Cell whose value has changed.
         * @param value - New value.
         * @param x - Cell column index.
         * @param y - Cell row index.
         * @param instance - Worksheet instance.
         * @param options - Column configuration object.
         */
        render?: (cell: HTMLTableCellElement, value: CellValue | undefined, x: number, y: number, instance: WorksheetInstance, options: Column) => void;

        /** Custom column title. */
        title?: string;

        /**
         * The type of cells in this column.
         * @default "text"
         */
        type?: 'text' | 'numeric' | 'hidden' | 'dropdown' | 'autocomplete' | 'checkbox' | 'radio' | 'calendar' | 'image' | 'color' | 'html' | CustomEditor;

        /** Column width. */
        width?: string | number;

        /**
         * Enable automatic word wrapping in cells in this column.
         * @default false
         */
        wordWrap?: boolean;
    }

    interface DropdownColumn extends BaseColumn {
        autocomplete?: boolean;

        /**
         * The text between the ids of the items in the value of a cell with {@link multiple}, and between their names in
         * the cell. Without it, the value is split on `;`, the editor joins the ids with `;`, and the names are joined
         * with `; `.
         */
        delimiter?: string;

        /**
         * Gives the items that the dropdown of a cell shows, instead of {@link source}.
         * @param element - The element of the spreadsheet.
         * @param cell - The td of the cell being edited.
         * @param x - The column of the cell, as text.
         * @param y - The row of the cell, as text.
         * @param source - The source of the column.
         */
        filter?: (element: HTMLElement, cell: HTMLTableCellElement, x: string, y: string, source: DropdownSourceItem[] | undefined) => DropdownSourceItem[];

        /**
         * Allow selection of more than one item.
         * @default false
         */
        multiple?: boolean;

        /** Options available in the dropdown. */
        source?: DropdownSourceItem[];

        /** Url to fetch options from an external source. */
        url?: string;
    }

    interface CalendarColumn extends BaseColumn {
        /** Calendar options. */
        options?: CalendarOptions;
    }

    interface NumericColumn extends BaseColumn {
        /**
         * Save an empty text when the editor of a cell is emptied, instead of 0.
         * @default false
         */
        allowEmpty?: boolean;

        /** The locale of the numbers, when the column has no {@link BaseColumn.mask} or {@link BaseColumn.format}. */
        locale?: string;

        /** The options of the numbers with a {@link locale}, as Intl.NumberFormat takes them. */
        options?: Intl.NumberFormatOptions;
    }

    interface ColorColumn extends Omit<BaseColumn, 'render'> {
        /** If undefined, the cell shows the hex code of the color, if "square", it shows a square filled with the color. */
        render?: 'square';
    }

    type Column = DropdownColumn | CalendarColumn | NumericColumn | ColorColumn | BaseColumn;

    interface Row {
        /** Row height. */
        height?: string | number;

        /** Text that replaces the row number. */
        title?: string;
    }

    /**
     * The toolbar of a spreadsheet, which the element of the toolbar keeps as its toolbar property.
     */
    interface ToolbarInstance {
        /** The settings of the toolbar, with its items. */
        options: Record<string, unknown>;
        /** Replace the items of the toolbar. */
        create(items: ToolbarItem[]): void;
        /** Call updateState on each item, with the two arguments. */
        update(a?: unknown, b?: unknown): void;
        /** Mark an item element as selected, with the class jtoolbar-selected. */
        selectItem(element: HTMLElement): void;
        /** Disable or enable the toolbar. */
        setReadonly(state: boolean): void;
        /** The element of the toolbar. */
        get(): HTMLElement;
        show(): void;
        hide(): void;
        /** Remove the items of the toolbar. */
        destroy(): void;
    }

    /**
     * The picker of a select item of the toolbar, which the element of the item keeps as its picker property.
     */
    interface PickerInstance {
        /** The options of the picker, with their defaults. */
        options: Record<string, unknown>;
        /** The key of the selected option: the first value, or the key as text after the user chose one. */
        getValue(): number | string;
        /** Select the option with a key. With an event, the picker calls its onchange. */
        setValue(index: number | string, event?: Event): void;
        /** Tell whether the list of options is open. */
        isOpened(): boolean;
        open(): void;
        close(): void;
    }

    interface ToolbarItemBase {
        /** Tooltip shown when hovering over this option. */
        tooltip?: string;

        /**
         * Method called when the toolbar state should be updated.
         * @param toolbarElement - Toolbar HTML element.
         * @param toolbarInstance - Toolbar instance.
         * @param itemElement - Toolbar item HTML element.
         * @param worksheetInstance - Worksheet instance.
         */
        updateState?: (
            toolbarElement: HTMLDivElement,
            toolbarInstance: ToolbarInstance,
            itemElement: HTMLDivElement,
            worksheetInstance: WorksheetInstance
        ) => void;

        /**
         * The toolbar also reads id, width, title, active, disabled, class (the classes of the icon in place of
         * material-icons) and render(itemElement, item) of an item, and k, the style property of a color item.
         */
        [property: string]: unknown;
    }

    interface ToolbarIconItem extends ToolbarItemBase {
        /** Defines the icon (from material icons). */
        content: string;

        /**
         * Event fired when clicking on the html item referring to that item.
         * @param toolbarElement - Toolbar HTML element.
         * @param toolbarInstance - Toolbar instance.
         * @param itemElement - Toolbar item HTML element.
         * @param event - Pointer event that triggered the onclick.
         */
        onclick?: (toolbarElement: HTMLDivElement, toolbarInstance: ToolbarInstance, itemElement: HTMLDivElement, event: PointerEvent) => void;
    }

    interface ToolbarSelectItem extends ToolbarItemBase {
        type: 'select';

        /** Defines the icon (from material icons). Without it, the item shows the selected option. */
        content?: string;

        /**
         * Event called when the item picker value is changed.
         * @param itemElement - Toolbar item HTML element.
         * @param pickerInstance - Picker instance.
         * @param value - New picker value.
         * @param value2 - New picker value.
         * @param valueIndex - Index of the new picker value.
         * @param event - Pointer event that triggered this event.
         */
        onchange?: (
            itemElement: HTMLDivElement,
            pickerInstance: PickerInstance,
            value: string,
            value2: string,
            valueIndex: string,
            event: PointerEvent
        ) => void;

        /**
         * Options available in the picker.
         */
        options?: string[];

        /**
         * Creates picker items based on the picker options.
         * @param option - One of the items in the {@link ToolbarSelectItem.options} array.
         * @param element - The element of the option in the list, or null for the option that the picker shows.
         * @returns The HTML of the option, or an element.
         */
        render?: (option: string, element: HTMLElement | null) => string | Element;

        /** The index in {@link ToolbarSelectItem.options} of the option that the item shows first. */
        value?: number | string;

        /** Item width. */
        width?: string;
    }

    interface ToolbarColorItem extends ToolbarItemBase {
        type: 'color';

        /** Defines the icon (from material icons). */
        content: string;

        /**
         * Style that should be changed when input value changes. If this property is set, the onclick event is overridden.
         */
        k: string;
    }

    interface ToolbarDivisorItem {
        type: 'divisor';
    }

    /**
     * Item that makes up the toolbar configuration array. This item may have properties not described here.
     */
    type ToolbarItem = ToolbarIconItem | ToolbarSelectItem | ToolbarColorItem | ToolbarDivisorItem;

    /**
     * The settings of the toolbar: the items, and other properties, which the toolbar keeps in its options.
     */
    interface ToolbarOptions {
        items: ToolbarItem[];

        [property: string]: unknown;
    }

    interface NestedHeaderCell {
        id?: string;
        colspan?: number;
        title?: string;
        align?: string;
    }

    /** A cell that changed, in onafterchanges. */
    interface CellChange {
        /** The new value. It is undefined when setValue was called without a value. */
        value: CellValue | undefined;
        oldValue: CellValue;
        /** Column index of the cell. */
        x: number;
        /** Row index of the cell. */
        y: number;
    }

    /** A change that undo reverts. The other properties depend on the action. */
    interface HistoryRecord {
        action: string;
        [key: string]: unknown;
    }

    /**
     * The selection of a worksheet, {@link WorksheetInstance.selection}. {@link WorksheetInstance.updateSelectionFromCoords}
     * and the other selection methods change it.
     */
    interface WorksheetSelection {
        /**
         * The selected range, or null when nothing is selected. The column [0] and row [1] of the cell where the
         * selection starts, and the column [2] and row [3] of the cell where it ends. The end can be left of or above the
         * start. {@link WorksheetInstance.getSelection} gives it from left to right and top to bottom.
         */
        range: [number, number, number, number] | null;

        /** The selected range from left to right and top to bottom, or null before the first selection. */
        bounds: [number, number, number, number] | null;

        /** The selected cells. */
        highlighted: { element: HTMLTableCellElement; x: number; y: number }[];

        /** The cells that a drag of the selection corner fills, while the drag goes on. */
        fill: HTMLTableCellElement[];

        /** The cell that has the roving tabindex, the tab stop of the grid, or null before the grid has one. */
        activeCell: HTMLTableCellElement | null;

        /** Whether the keyboard focus is in the grid or in its editor. */
        focusInGrid: boolean;

        /** The column whose header a press selected first, while the selection is of whole columns, or null. */
        column: number | null;

        /** The row whose number a press selected first, while the selection is of whole rows, or null. */
        row: number | null;

        /** Remove the fill border from the cells that a drag of the corner was going to fill, and forget them. */
        clearFill(): void;
    }

    /**
     * Which rows a worksheet shows, {@link WorksheetInstance.view}. {@link WorksheetInstance.search},
     * the filters ({@link WorksheetInstance.openFilter},
     * {@link WorksheetInstance.resetFilters}), {@link WorksheetInstance.page} and the lazy loading change it.
     */
    interface WorksheetView {
        /** The indexes of the rows that a search or the filters found, or null when the worksheet shows every row. */
        results: number[] | null;

        /** The values that each column is filtered by, by column index. null or a missing entry is a column without filter. */
        filters: (string[] | null)[];

        /** The page that the worksheet shows, from 0, when it has pagination. */
        pageNumber: number;
    }

    /**
     * The state of a mouse gesture on a worksheet, {@link WorksheetInstance.pointer}. The mouse handlers set it when a
     * button goes down and clear it when it goes up.
     */
    interface WorksheetPointer {
        /** The column or row that the mouse resizes, or null. */
        resizing: ResizeColumnInfo | ResizeRowInfo | null;

        /** The column or row that the mouse moves, or null. */
        dragging: DragColumnInfo | DragRowInfo | null;

        /** The header or row element that shows a resize cursor, or null. */
        cursor: HTMLElement | null;

        /** True while the mouse drags the selection corner to fill cells. */
        corner: boolean;
    }

    /**
     * The formula chain of a worksheet, {@link WorksheetInstance.formulas}. The worksheet adds to it when it calculates
     * a formula.
     */
    interface WorksheetFormulas {
        /**
         * For each cell that a formula refers to, by its name (such as A1), the names of the cells whose formulas refer
         * to it.
         */
        chain: Record<string, string[]>;

        /** Keep that the formula of the cell `dependent` refers to the cell `cell`. Both are cell names. */
        add(cell: string, dependent: string): void;

        /** Rename the cells in the chain, with the new name by the old name, after rows or columns moved. */
        rename(references: Record<string, string>): void;
    }

    /**
     * The cell editor of a worksheet, {@link WorksheetInstance.editor}. {@link WorksheetInstance.openEditor} and
     * {@link WorksheetInstance.closeEditor} open and close it.
     */
    interface WorksheetEditor {
        /** The cell that is being edited, or null when no editor is open. */
        cell: HTMLTableCellElement | null;

        /** The HTML of the cell before the editor opened. */
        html: string;

        /** The column of the cell, or -1 when no editor is open. */
        x: number;

        /** The row of the cell, or -1 when no editor is open. */
        y: number;

        /** Whether an editor is open. */
        isOpen(): this is WorksheetEditor & { cell: HTMLTableCellElement };

        /** Keep the cell that an editor opens on, without creating the editor. */
        start(cell: HTMLTableCellElement, x: number, y: number): void;

        /** Forget the cell, without closing its editor. */
        end(): void;
    }

    /**
     * What a worksheet copied last, {@link WorksheetInstance.clipboard}.
     */
    interface WorksheetClipboard {
        /** The text of the last copy, or an empty string before the first copy. */
        data: string;

        /** The hash of {@link WorksheetClipboard.data}, or null before the first copy. */
        hash: number | null;

        /** The styles of the copied cells, one entry for each cell, by row. A paste of the copied text applies them. */
        style: string[];

        /** The hidden textarea that a copy selects, to put the text on the system clipboard. */
        textarea: HTMLTextAreaElement;

        /** Keep the text of a copy and the styles of its cells, without putting the text on the system clipboard. */
        keep(data: string, style: string[]): void;

        /** Whether text is the text of the last copy. */
        matches(text: string): boolean;

        /** Put text on the system clipboard. */
        write(text: string): void;

        /** Draw the dashed border around copied cells. */
        showBorder(cells: { element: HTMLTableCellElement }[]): void;

        /** Remove the dashed border from the cells of every worksheet in the document. */
        clearBorder(): void;
    }

    /**
     * The undo history of a worksheet, {@link WorksheetInstance.history}.
     */
    interface WorksheetHistory {
        /** The records of the changes, oldest first. */
        records: HistoryRecord[];

        /** The index in {@link WorksheetHistory.records} of the record that undo reverts next, or -1 when there is nothing to undo. */
        index: number;

        /** While true, the worksheet keeps no records of its changes. */
        ignore: boolean;

        /** Keep a record of a change. The records that were undone are dropped. Does nothing while {@link WorksheetHistory.ignore} is true. */
        add(record: HistoryRecord): void;

        /** Drop all the records. */
        clear(): void;

        /** Whether {@link WorksheetInstance.undo} has a change to undo. */
        canUndo(): boolean;

        /** Whether {@link WorksheetInstance.redo} has a change to redo. */
        canRedo(): boolean;

        /** Move back one record and return it, without changing the worksheet. Returns undefined when there is nothing to undo. */
        back(): HistoryRecord | undefined;

        /** Move forward one record and return it, without changing the worksheet. Returns undefined when there is nothing to redo. */
        forward(): HistoryRecord | undefined;
    }

    /**
     * Item compared in column ordering.
     *
     * It is composed respectively by the index of the row it represents and by the value of that row in the sorting column.
     */
    type SortingItem = [number, CellValue];

    /** The meta information of a cell: any properties that the application keeps on it. */
    /** The handlers of the events in {@link SpreadsheetOptions}, by event name. onevent gets every event, so it is not an event. */
    type EventHandlers = {
        [K in Exclude<keyof SpreadsheetOptions, 'onevent'> as K extends `on${string}` ? K : never]-?: Extract<
            SpreadsheetOptions[K],
            (...args: never[]) => unknown
        >;
    };

    /** The name of an event, such as 'onchange'. */
    type EventName = keyof EventHandlers;

    /** The arguments of onevent: the name of an event, then the arguments of its handler. */
    type EventArguments = { [E in EventName]: [event: E, ...args: Parameters<EventHandlers[E]>] }[EventName];

    /**
     * The type of onevent, which gets every event: one signature for each event, with the name of the event, then the
     * arguments of its handler. A handler can name fewer parameters than an event gives, such as `(event) => …` or
     * `(event, instance) => …`. To narrow the arguments with a test of the name, take them as
     * `(...args: jspreadsheet.EventArguments) => …`.
     */
    type EventListener = {
        [E in EventName]: (listener: (event: E, ...args: Parameters<EventHandlers[E]>) => unknown) => void;
    }[EventName] extends (listener: infer L) => void
        ? L
        : never;

    type MetaInformation = Record<string, unknown>;

    type ContextMenuItem =
        | {
              /** A line between the items. It needs no title. */
              type: 'line' | 'divisor';
              title?: string;
          }
        | ContextMenuOption;

    /** An item of the context menu that the user can choose. */
    type ContextMenuOption = {
        type?: 'default';
        title: string;
        icon?: string;
        id?: string;
        disabled?: boolean;
        /** element is the element of the item in the menu, and e the release of the mouse or the key (Enter or Space) that chose it. */
        onclick?(element: HTMLDivElement, e: MouseEvent | KeyboardEvent): void;
        shortcut?: string;
        tooltip?: string;
        submenu?: Array<ContextMenuItem>;
    };

    type ContextMenuRole =
        'select-all' | 'fill-handle' | 'row' | 'nested' | 'tabs' | 'toolbar' | 'pagination' | 'cell' | 'grid' | 'footer' | 'header' | 'applications';

    interface SpreadsheetOptions {
        /**
         * Show or not the "about" item in the context menu. A text is shown instead of the version when the item is clicked.
         * @default true
         */
        about?: boolean | string;

        /**
         * Spreadsheet plugins. Each function is called with the spreadsheet, an empty object and the spreadsheet options, and returns the plugin.
         */
        plugins?: Record<string, PluginFactory>;

        /**
         * Node for binding the javascript events and reading the text selection. This property is normally used when JSS is running as a web component, with its shadow root.
         */
        root?: HTMLElement | ShadowRoot;

        /**
         * Styles that the worksheets refer to by index: a number in {@link WorksheetOptions.style} is the index of a style in this list.
         */
        style?: string[];

        /**
         * Allow table export as csv.
         * @default true
         */
        allowExport?: boolean;

        /**
         * If true, Jss will try to convert cell contents used in formulas to numbers
         * @default true
         */
        autoCasting?: boolean;

        /**
         * Auto increment actions when using the dragging corner.
         * @default true
         */
        autoIncrement?: boolean;

        /**
         * Give a note that the accessible name of the active cell adds after the address and the text of the cell, for
         * example "cached, not recalculated". The name of A5 is then "A5 12, cached, not recalculated". Call
         * `worksheet.labelCell` after the note of a cell changes.
         * @param instance - Worksheet of the cell.
         * @param x - Column index of the cell.
         * @param y - Row index of the cell.
         * @returns Text of the note. An empty text, null or undefined adds nothing.
         */
        cellNote?: (instance: WorksheetInstance, x: number, y: number) => string | null | undefined;

        /**
         * Creates context menu when user clicks with secondary mouse button.
         * @param instance - Instance of the worksheet on which the click was made.
         * @param colIndex - Horizontal index of the element that was clicked. The meaning of this value depends on the {@link role} argument.
         * @param rowIndex - Vertical index of the element that was clicked. The meaning of this value depends on the {@link role} argument.
         * @param event - pointer event that triggered this event.
         * @param items - jss default context menu.
         * @param role - indicates in which part of the spreadsheet the click occurred.
         * @param x - Horizontal index of the element that was clicked. The meaning of this value depends on the {@link role} argument.
         * @param y - Vertical index of the element that was clicked. The meaning of this value depends on the {@link role} argument.
         * @returns Context menu configuration that should be created, or false to open no menu of the spreadsheet.
         */
        contextMenu?: (
            instance: WorksheetInstance,
            colIndex: string | number | null,
            rowIndex: string | number | null,
            event: PointerEvent,
            items: ContextMenuItem[],
            role: ContextMenuRole,
            x: string | number | null,
            y: string | number | null
        ) => ContextMenuItem[] | false | null | undefined;

        /**
         * Enable the formula debug notices.
         * @default false
         */
        debugFormulas?: boolean;

        /**
         * Give the text that the text editor shows when it opens on a cell, in place of the value of the cell. For example, show
         * a date serial number as a date. The hook does not run when the editor opens empty, for example on a typed character.
         * If the user leaves the text as it is, the cell keeps its value. A changed text goes through the mask of the column
         * and `onbeforechange`, where the application can convert it back to a value.
         * @param instance - Worksheet of the cell.
         * @param x - Column index of the cell.
         * @param y - Row index of the cell.
         * @param value - Value of the cell. For a formula cell, the formula.
         * @returns Text of the editor, or undefined to show the value as the column shows it.
         */
        editValue?: (instance: WorksheetInstance, x: number, y: number, value: CellValue) => CellValue | undefined;

        /**
         * Give the text that a cell shows, for example a value in a number format of the application. It runs for the columns
         * that show their value as text: the columns of the types text, numeric and html, and the columns without a type.
         * The sort, the search, the filters, the processed values of `getValue` and `getData`, and the copy use the shown text.
         * The footers do not use the hook. In an html column, and with `parseHTML`, the text is HTML: escape the text of users.
         * @param instance - Worksheet of the cell.
         * @param x - Column index of the cell.
         * @param y - Row index of the cell.
         * @param value - Value of the cell, or the result of its formula: an array for a range, and undefined for a cell
         * without a value.
         * @param cell - Element of the cell, when the value shows in a cell. The hook can change its style, for example its
         * `textAlign`. The column `align` applies first. A new value of the cell does not reset a style that the hook set.
         * @returns Text of the cell, or undefined to show the value with the mask of the column.
         */
        formatValue?: (
            instance: WorksheetInstance,
            x: number,
            y: number,
            value: CellValue | CellValue[] | undefined,
            cell?: HTMLElement
        ) => CellValue | undefined;

        /**
         * Calculate formulas with an engine of the application in place of the engine of `jspreadsheet.setFormulaEngine`.
         * The worksheet does not calculate the referenced cells first. The engine reads their raw values and calculates referenced formulas itself.
         * The worksheet still calculates the dependent cells again when a referenced cell changes.
         * A formula that refers to its own cell shows "#ERROR", and the engine does not get it.
         * @param instance - Worksheet of the formula.
         * @param expression - Formula text as typed, with the "=", or the text that `onbeforeformula` returned.
         * @param x - Column index of the cell. Undefined for a footer, and when `executeFormula` gets no coordinates.
         * @param y - Row index of the cell. Undefined for a footer, and when `executeFormula` gets no coordinates.
         * @returns Value that the cell shows. If the function throws, the cell shows "#ERROR".
         */
        formulaEngine?: (instance: WorksheetInstance, expression: string, x?: number, y?: number) => CellValue;

        /**
         * Fullscreen mode.
         * @default false
         */
        fullscreen?: boolean;

        /**
         * Include header titles on download.
         * @default false
         */
        includeHeadersOnDownload?: boolean;

        /**
         * Texts of the accessible names that assistive technology reads. A missing text is the English default, which
         * `jspreadsheet.setDictionary` can translate. `{0}` in a text stands for a value. Set `labels` when you create the
         * spreadsheet.
         */
        labels?: {
            /**
             * Name of the corner cell that selects all cells.
             * @default 'Select all'
             */
            selectAll?: string;

            /**
             * Name of the cell editor. `{0}` is the address of the cell, for example B4.
             * @default 'Edit cell {0}'
             */
            editCell?: string;

            /**
             * Name of the hidden text area that the copy action uses.
             * @default 'Clipboard helper'
             */
            clipboardHelper?: string;

            /**
             * Name of the filter of a column. `{0}` is the title of the column.
             * @default 'Filter column {0}'
             */
            filterColumn?: string;

            /**
             * Name of the list of worksheet tabs.
             * @default 'Worksheets'
             */
            worksheets?: string;

            /**
             * Name of the region that holds the worksheets.
             * @default 'Tab Panels'
             */
            tabPanels?: string;

            /**
             * Name of the button that adds a worksheet.
             * @default 'Add a new worksheet'
             */
            addWorksheet?: string;
        };

        /** Spreadsheet namespace */
        namespace?: string;

        /**
         * Occurs after all changes are applied in the tables.
         * @param instance - Instance of the worksheet where the change occurred.
         * @param changes - list of changes.
         */
        onafterchanges?: (instance: WorksheetInstance, changes: CellChange[]) => void;

        /**
         * Occurs before the value of a cell changes. Return false to cancel the change, or another value to write it
         * instead of the new value. To write the value false, return nothing: newValue is written.
         * @param instance - Instance of the worksheet where the changes will occur.
         * @param cell - HTML element that represents the cell being changed.
         * @param colIndex - Cell column index being changed.
         * @param rowIndex - Cell row index being changed.
         * @param newValue - Value being applied to the cell. It is undefined when setValue was called without a value.
         */
        onbeforechange?: (
            instance: WorksheetInstance,
            cell: HTMLTableCellElement,
            colIndex: number,
            rowIndex: number,
            newValue: CellValue | undefined
        ) => undefined | CellValue;

        /**
         * Occurs before a column is excluded. If this method returns false, the removal will be canceled.
         * @param instance - Instance of the worksheet where columns will be removed.
         * @param removedColumns - Indexes of the columns to be removed.
         */
        onbeforedeletecolumn?: (instance: WorksheetInstance, removedColumns: number[]) => undefined | boolean;

        /**
         * Occurs before a row is deleted. If this method returns false, the removal will be canceled.
         * @param instance - Instance of the worksheet where rows will be removed.
         * @param removedRows - Indexes of the rows to be removed.
         */
        onbeforedeleterow?: (instance: WorksheetInstance, removedRows: number[]) => undefined | boolean;

        /**
         * Intercept and parse a formula just before the execution.
         * @param instance - Instance of the worksheet.
         * @param expression - Formula that triggered the event.
         * @param x - Column index of the cell whose formula triggered the event. Undefined for a footer, and when
         * `executeFormula` gets no coordinates.
         * @param y - Row index of the cell whose formula triggered the event. Undefined for a footer, and when
         * `executeFormula` gets no coordinates.
         */
        onbeforeformula?: (instance: WorksheetInstance, expression: string, x?: number, y?: number) => false | string | undefined;

        /**
         * Occurs before a new column is inserted. If this method returns false, the insertion will be canceled.
         * @param instance - Instance of the worksheet where columns will be added.
         * @param columns - Settings for columns to be added.
         */
        onbeforeinsertcolumn?: (
            instance: WorksheetInstance,
            columns: {
                column: number;
                options: Column;
                data?: CellValue[];
            }[]
        ) => undefined | boolean;

        /**
         * Occurs before a new row is inserted. If this method returns false, the insertion will be canceled.
         * @param instance - Instance of the worksheet where rows will be added.
         * @param rows - Settings for rows to be added.
         */
        onbeforeinsertrow?: (
            instance: WorksheetInstance,
            rows: {
                row: number;
                data: CellValue[];
            }[]
        ) => undefined | boolean;

        /**
         * Occurs before the paste action is performed.
         *
         * If it returns false, the jss cancels the paste.
         * If it returns a string, it will be the content pasted into the worksheet.
         * If it returns rows, of `{ value }` objects as it got them or of values, they are pasted instead.
         *
         * @param instance - Instance of the worksheet where data will be pasted.
         * @param copiedText - The rows of the pasted cells, with the text of each cell.
         * @param colIndex - Column index where it will start the paste.
         * @param rowIndex - Row index where it will start the paste.
         */
        onbeforepaste?: (
            instance: WorksheetInstance,
            copiedText: { value: string }[][],
            colIndex: number | string,
            rowIndex: number | string
        ) => undefined | boolean | string | ({ value: CellValue } | CellValue)[][];

        /**
         * Occurs before the worksheet sends a request to a server: to load its {@link WorksheetOptions.csv} or
         * {@link WorksheetOptions.url}, the source of a dropdown column, or to post its changes with
         * {@link WorksheetOptions.persistence}. It occurs while the events are ignored too.
         *
         * The handler can change the options of the request, such as its headers, or return other options for fetch.
         * If it returns false, the request is not sent: the worksheet is built without the data of the request, or the
         * changes are not posted.
         * @param instance - Instance of the worksheet that sends the request.
         * @param url - The address of the request.
         * @param options - The options of fetch: the method, the headers, the body and the cache.
         */
        onbeforerequest?: (instance: WorksheetInstance, url: string, options: RequestInit) => RequestInit | false | undefined | void;

        /**
         * Occurs before persisting any changes to the server.
         *
         * This event is only called when the spreadsheet has the {@link WorksheetOptions.persistence} property set.
         *
         * If this event returns false, the change is not persisted on the server.
         * If it returns a truthy value, that value is persisted instead of the initial value.
         * @param spreadsheetInstance - Spreadsheet instance.
         * @param worksheetInstance - Instance of the worksheet to be saved.
         * @param data - The changed rows. The values of a row are by column name, or by column index when the column has no name.
         */
        onbeforesave?: (
            spreadsheetInstance: SpreadsheetInstance,
            worksheetInstance: WorksheetInstance,
            data: { row: number; data: Record<string, CellValue> }[]
        ) => unknown;

        /**
         * Occurs before the selection is changed.
         * @param instance - Worksheet instance where the selection will occur.
         * @param borderLeftIndex - Index of the first column contained by the selection.
         * @param borderTopIndex - Index of the first row contained by the selection.
         * @param borderRightIndex - Index of the last column contained by the selection.
         * @param borderBottomIndex - Index of the last row contained by the selection.
         * @param origin - Javascript event that triggered this jss event.
         */
        onbeforeselection?: (
            instance: WorksheetInstance,
            borderLeftIndex: number,
            borderTopIndex: number,
            borderRightIndex: number,
            borderBottomIndex: number,
            origin: Event | undefined
        ) => false | undefined;

        /**
         * Occurs when the table is blurred.
         * @param instance - Instance of the worksheet that was blurred.
         */
        onblur?: (instance: WorksheetInstance) => void;

        /**
         * Occurs after a column value is changed.
         * @param instance - Instance of the worksheet where the change occurred.
         * @param cell - HTML element that represents the cell being changed.
         * @param colIndex - Cell column index being changed.
         * @param rowIndex - Cell row index being changed.
         * @param newValue - New cell value. It is undefined when setValue was called without a value.
         * @param oldValue - Old cell value.
         */
        onchange?: (
            instance: WorksheetInstance,
            cell: HTMLTableCellElement,
            colIndex: number,
            rowIndex: number,
            newValue: CellValue | undefined,
            oldValue: CellValue
        ) => void;

        /**
         * Occurs when a column heading is changed.
         * @param instance - Instance of the worksheet where the change occurred.
         * @param colIndex - Index of the column that was renamed.
         * @param newValue - New column title.
         * @param oldValue - Old column title.
         */
        onchangeheader?: (instance: WorksheetInstance, colIndex: number, newValue: string, oldValue: string) => void;

        /**
         * Occurs when a "setMeta" is called.
         *
         * @param instance - Instance of the worksheet where the change occurred.
         * @param cellName - The changes of the meta information, by cell name.
         */
        onchangemeta?: (instance: WorksheetInstance, cellName: Record<string, MetaInformation>) => void;

        /**
         * Occurs when the page is changed.
         * @param instance - Instance of the worksheet where the change occurred.
         * @param newPageNumber - Page the worksheet is on.
         * @param oldPageNumber - Page the worksheet was on.
         * @param quantityPerPage - Maximum number of lines on pages.
         */
        onchangepage?: (instance: WorksheetInstance, newPageNumber: number, oldPageNumber: number, quantityPerPage: number) => void;

        /**
         * Occurs when a "setStyle" is called.
         *
         * @param instance - Instance of the worksheet where the change occurred.
         * @param cellName - An object with the changes.
         */
        onchangestyle?: (instance: WorksheetInstance, changes: Record<string, string>) => void;

        /**
         * Occurs when a comment is changed.
         * @param instance - Instance of the worksheet where the change occurred.
         * @param newComments - New comments, by cell name, as setComments got them. A falsy value removes the comment.
         * @param oldComments - Old comments, by cell name.
         */
        oncomments?: (instance: WorksheetInstance, newComments: Record<string, string | null | undefined>, oldComments: Record<string, string | null>) => void;

        /**
         * Occurs when the contents of one or more cells are copied. Returning false cancels the copy or the cut, and text replaces the copied data.
         * @param instance - Instance of the worksheet where the change occurred.
         * @param selectedRange - Copied cell range.
         * @param copiedData - Data from copied cell range.
         * @param cut - If true, the action was cut. Otherwise, the action was copy.
         */
        oncopy?: (
            instance: WorksheetInstance,
            selectedRange: [number, number, number, number],
            copiedData: string,
            cut: boolean | undefined
        ) => string | false | undefined;

        /**
         * Occurs when a cell is created.
         * @param instance - Instance of the worksheet where the cell was created.
         * @param cell - Cell HTML element.
         * @param colIndex - Cell column index.
         * @param rowIndex - Cell row index.
         * @param newValue - Cell value.
         */
        oncreatecell?: (instance: WorksheetInstance, cell: HTMLTableCellElement, colIndex: number, rowIndex: number, newValue: CellValue) => void;

        /**
         * Occurs when an editor is opened.
         * @param instance - Instance of the worksheet where the change occurred.
         * @param td - Td tag of the cell whose editor was opened.
         * @param colIndex - Column index of the cell whose editor was opened.
         * @param rowIndex - Row index of the cell whose editor was opened.
         * @param input - Input of the editor that was opened.
         * @param options - Column settings.
         */
        oncreateeditor?: (instance: WorksheetInstance, td: HTMLTableCellElement, colIndex: number, rowIndex: number, input: null, options: Column) => void;

        /**
         * Occurs after a column is excluded.
         * @param instance - Instance of the worksheet where the change occurred.
         * @param removedColumns - Indexes of the columns that were removed.
         */
        ondeletecolumn?: (instance: WorksheetInstance, removedColumns: number[]) => void;

        /**
         * Occurs after a row is excluded.
         * @param instance - Instance of the worksheet where the change occurred.
         * @param removedRows - Indexes of the rows that were removed.
         */
        ondeleterow?: (instance: WorksheetInstance, removedRows: number[]) => void;

        /**
         * Occurs when a worksheet is created.
         * @param worksheet - Instance of the new worksheet.
         * @param worksheetOptions - Options of the new worksheet.
         * @param index - Index of the new worksheet.
         */
        oncreateworksheet?: (worksheet: WorksheetInstance, worksheetOptions: WorksheetOptions, index: number) => void;

        /**
         * Occurs when a worksheet is deleted.
         * @param worksheet - Instance of the deleted worksheet.
         * @param index - Index of the deleted worksheet.
         */
        ondeleteworksheet?: (worksheet: WorksheetInstance, index: number) => void;

        /**
         * Occurs when a closeEditor is called.
         * @param instance - Instance of the worksheet where the change occurred.
         * @param td - Td tag of the cell whose editor was opened.
         * @param colIndex - Column index of the cell whose editor was opened.
         * @param rowIndex - Row index of the cell whose editor was opened.
         * @param editorValue - Value that was in the editor. It is undefined when the value was not saved.
         * @param wasSaved - Whether the value which was in the editor was saved in the cell or not.
         */
        oneditionend?: (
            instance: WorksheetInstance,
            td: HTMLTableCellElement,
            colIndex: number,
            rowIndex: number,
            editorValue: CellValue | undefined,
            wasSaved: boolean
        ) => void;

        /**
         * Occurs when a openEditor is called.
         * @param instance - Instance of the worksheet where the change occurred.
         * @param td - Td tag of the cell whose editor was opened.
         * @param colIndex - Column index of the cell whose editor was opened.
         * @param rowIndex - Row index of the cell whose editor was opened.
         */
        oneditionstart?: (instance: WorksheetInstance, td: HTMLTableCellElement, colIndex: number, rowIndex: number) => void;

        /**
         * Event fired when any other event fires. It runs before the event that was called.
         *
         * Its result counts as a result of the called event. When onevent, the handler of the event or the onevent of a
         * plugin returns false, the result of the event is false, so any of them can cancel an `onbefore…` action. Otherwise
         * the result is the last one that is not undefined.
         *
         * It gets the name of the event that was called, then the arguments of the event. See {@link EventListener}.
         */
        onevent?: EventListener;

        /**
         * Occurs when the table is focused.
         * @param instance - Instance of the worksheet that was focused on.
         */
        onfocus?: (instance: WorksheetInstance) => void;

        /**
         * Occurs after a new column is inserted.
         * @param instance - Instance of the worksheet where the change occurred.
         * @param columns - Columns that have been added.
         */
        oninsertcolumn?: (
            instance: WorksheetInstance,
            columns: {
                column: number;
                options: Column;
                data?: CellValue[];
            }[]
        ) => void;

        /**
         * Occurs after a new row is inserted.
         * @param instance - Instance of the worksheet where the change occurred.
         * @param rows - Rows that have been added.
         */
        oninsertrow?: (
            instance: WorksheetInstance,
            rows: {
                row: number;
                data: CellValue[];
            }[]
        ) => void;

        /**
         * Event fired when a spreadsheet is created.
         * @param instance - Jspreadsheet instance.
         */
        onload?: (instance: SpreadsheetInstance) => void;

        /**
         * Occurs when a group of cells is merged.
         * @param instance - Instance of the worksheet where the change occurred.
         * @param merges - Merges that were created.
         */
        onmerge?: (instance: WorksheetInstance, merges: Record<string, [number, number]>) => void;

        /**
         * Occurs when a merge is removed.
         * @param instance - Instance of the worksheet where the change occurred.
         * @param cellName - Address of the cell that merge was removed.
         * @param beforeMerges - Merges that were before the merge was removed.
         */
        onunmerge?: (instance: WorksheetInstance, cellName: string, beforeMerges: Record<string, [number, number]>) => void;

        /**
         * Occurs after a column is moved to a new position.
         * @param instance - Instance of the worksheet where the change occurred.
         * @param oldPosition - Column index before movement.
         * @param newPosition - Column index after movement.
         * @param quantity - Number of columns that were moved.
         */
        onmovecolumn?: (instance: WorksheetInstance, oldPosition: number, newPosition: number, quantity: number) => void;

        /**
         * Occurs after a row is moved to a new position.
         * @param instance - Instance of the worksheet where the change occurred.
         * @param oldPosition - Row index before movement.
         * @param newPosition - Row index after movement.
         * @param quantity - Number of rows that were moved.
         */
        onmoverow?: (instance: WorksheetInstance, oldPosition: number, newPosition: number, quantity: number) => void;

        /**
         * Occurs after a paste action is performed in the javascript table.
         * @param instance - Instance of the worksheet where the change occurred.
         * @param pastedInfo - The pasted cells, one item for each cell.
         */
        onpaste?: (
            instance: WorksheetInstance,
            pastedInfo: {
                x: number;
                y: number;
                value: CellValue;
            }[]
        ) => void;

        /**
         * Occurs when a change is redone.
         * @param instance - Instance of the worksheet where the change occurred.
         * @param historyRecord - History item that was redone. If there are no more actions to redo, it takes the value undefined.
         */
        onredo?: (instance: WorksheetInstance, historyRecord: HistoryRecord | undefined) => void;

        /**
         * Occurs after a change in column width.
         * @param instance - Instance of the worksheet where the change occurred.
         * @param colIndex - Index of columns that were resized.
         * @param newWidth - New column widths.
         * @param oldWidth - Old column widths.
         */
        onresizecolumn?: (instance: WorksheetInstance, colIndex: number | number[], newWidth: number | number[], oldWidth: number | number[]) => void;

        /**
         * Occurs after a change in row height.
         * @param instance - Instance of the worksheet where the change occurred.
         * @param rowIndex - Index of row being resized.
         * @param newHeight - New row height.
         * @param oldHeight - Old row height.
         */
        onresizerow?: (instance: WorksheetInstance, rowIndex: number, newHeight: number, oldHeight: number) => void;

        /**
         * Occurs when persistence on the server succeeds.
         * @param spreadsheetInstance - Spreadsheet instance.
         * @param worksheetInstance - Instance of the worksheet to be saved.
         * @param data - Data that has been sent to the server: the changed rows, or what {@link SpreadsheetOptions.onbeforesave} returned.
         */
        onsave?: (spreadsheetInstance: SpreadsheetInstance, worksheetInstance: WorksheetInstance, data: unknown) => void;

        /**
         * Occurs when selection is changed.
         * @param instance - Instance of the worksheet where the change occurred.
         * @param borderLeftIndex - Index of the first column contained by the selection.
         * @param borderTopIndex - Index of the first row contained by the selection.
         * @param borderRightIndex - Index of the last column contained by the selection.
         * @param borderBottomIndex - Index of the last row contained by the selection.
         * @param origin - Javascript event that triggered this jss event.
         */
        onselection?: (
            instance: WorksheetInstance,
            borderLeftIndex: number,
            borderTopIndex: number,
            borderRightIndex: number,
            borderBottomIndex: number,
            origin: Event | undefined
        ) => void;

        /**
         * Occurs after a colum is sorted.
         * @param instance - Instance of the worksheet where the change occurred.
         * @param colIndex - Index of the column that was sorted.
         * @param order - Sorting direction. 0 for ascending and 1 for descending.
         * @param newOrderValues - The new order of the rows.
         */
        onsort?: (instance: WorksheetInstance, colIndex: number, order: 0 | 1, newOrderValues: number[]) => void;

        /**
         * Occurs when the rows that a search or a filter shows have changed.
         * @param instance - Instance of the worksheet.
         * @param results - Indexes of the rows that are shown, or null when every row is shown.
         */
        onupdateresult?: (instance: WorksheetInstance, results: number[] | null) => void;

        /**
         * Occurs when a change is undone.
         * @param instance - Instance of the worksheet where the change occurred.
         * @param historyRecord - History item that was undone. If there are no more actions to undo, it takes the value undefined.
         */
        onundo?: (instance: WorksheetInstance, historyRecord: HistoryRecord | undefined) => void;

        /**
         * Enable execution of formulas inside the table.
         * @default true
         */
        parseFormulas?: boolean;

        /**
         * Function used in sorting columns. If not specified, the default function will be used.
         * @param order - Sorting direction. 0 for ascending and 1 for descending.
         */
        sorting?: (order: 0 | 1) => (itemA: SortingItem, itemB: SortingItem) => number;

        /**
         * If false, HTML inside cell values will be treated as regular text. If true, the HTML will be treated as HTML.
         * @default false
         */
        parseHTML?: boolean;

        /**
         * If true, the button to create new worksheets is shown.
         * @default false
         */
        tabs?: boolean;

        /**
         * If true, the worksheet tabs are hidden with any number of worksheets. The worksheets stay and openWorksheet still opens them.
         * @default false
         */
        hideTabs?: boolean;

        /**
         * Add custom toolbars: true for the default toolbar, the items, or the toolbar settings. A function gets the
         * settings of the default toolbar and returns the settings to use.
         */
        toolbar?: boolean | ToolbarItem[] | ToolbarOptions | ((defaultToolbar: ToolbarOptions) => ToolbarOptions);

        /**
         * Worksheet settings.
         */
        worksheets: WorksheetOptions[];
    }

    interface WorksheetOptions {
        /**
         * Allow comments over the cells.
         * @default true
         */
        allowComments?: boolean;

        /**
         * Allow delete a column.
         * @default true
         */
        allowDeleteColumn?: boolean;

        /**
         * Allow delete a row.
         * @default true
         */
        allowDeleteRow?: boolean;

        /**
         * Allow remove all rows. Otherwise, at least one row will be kept.
         * @default false
         */
        allowDeletingAllRows?: boolean;

        /**
         * Allow insert a new column.
         * @default true
         */
        allowInsertColumn?: boolean;

        /**
         * Allow insert a new row.
         * @default true
         */
        allowInsertRow?: boolean;

        /**
         * It has no effect. Version 5 let Tab on the last column insert a column; Tab there now moves the focus out of
         * the grid.
         * @deprecated
         * @default true
         */
        allowManualInsertColumn?: boolean;

        /**
         * Allow user to insert a new row using the Enter key on the last row.
         * @default true
         */
        allowManualInsertRow?: boolean;

        /**
         * Allow rename a column.
         * @default true
         */
        allowRenameColumn?: boolean;

        /**
         * Css classes to apply to cells. Only one class per cell is accepted.
         * @example
         * {
         *    A1: "some-class",
         *    B3: "another-class"
         * }
         */
        classes?: Record<string, string>;

        /**
         * Allow column dragging.
         * @default true
         */
        columnDrag?: boolean;

        /**
         * Allow column resizing.
         * @default true
         */
        columnResize?: boolean;

        /** Column settings. */
        columns?: Column[];

        /**
         * Allow column sorting.
         * @default true
         */
        columnSorting?: boolean;

        /**
         * Worksheet comments. Each object key is a cell name and its value is its comment.
         */
        comments?: Record<string, string>;

        /**
         * Load a external CSV file from this URL.
         */
        csv?: string;

        /**
         * Default delimiter for the CSV file. This value is used for both import and export.
         * @default ","
         */
        csvDelimiter?: string;

        /**
         * Default filename for a download method.
         * @default "jspreadsheet"
         */
        csvFileName?: string;

        /**
         * Load header titles from the CSV file.
         * @default false
         */
        csvHeaders?: boolean;

        /** Data loaded into the spreadsheet. */
        data?: CellValue[][] | Record<string, CellValue>[];

        /**
         * Default horizontal alignment used when a column does not have its alignment specified.
         * @default "center"
         */
        defaultColAlign?: HorizontalAlign;

        /**
         * Default width for columns with no specified width.
         * @default 100
         */
        defaultColWidth?: number;

        /**
         * Default row height.
         */
        defaultRowHeight?: number;

        /**
         * Allow table edition.
         * @default true
         */
        editable?: boolean;

        /**
         * Enable column filters.
         * @default false
         */
        filters?: boolean;

        /**
         * Set the initial footer of the spreadsheet: the rows of the footer, with a text, a number or a formula for each column.
         */
        footers?: (string | number)[][];

        /**
         * Number of columns frozen at the top of the spreadsheet.
         */
        freezeColumns?: number;

        /** Activate the table lazyloading. */
        lazyLoading?: boolean;

        /** Cells to be merged in the table innitialization. */
        mergeCells?: Record<string, [number, number]>;

        /**
         * Meta information.
         */
        meta?: Record<string, MetaInformation>;

        /**
         * Minimum table dimensions: [cols, rows].
         * @default [0, 0]
         */
        minDimensions?: [number, number];

        /**
         * Minimum number of spare cols.
         * @default 0
         */
        minSpareCols?: number;

        /**
         * Minimum number of spare rows.
         * @default 0
         */
        minSpareRows?: number;

        /** Define the nested headers. */
        nestedHeaders?: NestedHeaderCell[][];

        /** Number of rows per page. */
        pagination?: number;

        /**
         * Values available in the dropdown for choosing the number of rows per page.
         *
         * This dropdown is only visible when the {@link WorksheetOptions.search} option is true and the {@link WorksheetOptions.pagination} option is greater than 0.
         */
        paginationOptions?: number[];

        /**
         * Try to identify the column type when the instance is created with a table tag.
         * @default false
         */
        parseTableAutoCellType?: boolean;

        /**
         * If creating the instance with a table tag, if this tag has no header, transform the first line into the header.
         * @default false
         */
        parseTableFirstRowAsHeader?: boolean;

        /**
         * Route where requests for data persistence will be sent. If true, the {@link WorksheetOptions.url} property value will be used instead.
         */
        persistence?: boolean | string;

        /**
         * Spreadsheet plugins.
         * @deprecated The worksheet options do not load plugins. Use {@link SpreadsheetOptions.plugins}.
         */
        plugins?: Record<string, PluginFactory>;

        /**
         * DOM element for binding the javascript events.
         * @deprecated The worksheet options do not bind events. Use {@link SpreadsheetOptions.root}.
         */
        root?: HTMLElement;

        /**
         * Allow row dragging.
         * @default true
         */
        rowDrag?: boolean;

        /**
         * Allow row resizing.
         * @default true
         */
        rowResize?: boolean;

        /** Row settings. */
        rows?: Row[] | Record<number, Row>;

        /**
         * Allow search in the table.
         * @default false
         */
        search?: boolean;

        /**
         * If true, Jss will capitalize characters that are part of formulas. This does not apply to characters enclosed in double quotes, which represent text within formulas.
         * @default true
         */
        secureFormulas?: boolean;

        /**
         * Display the copy icon in the lower right corner of the selection.
         * @default true
         */
        selectionCopy?: boolean;

        /**
         * Cell styles, by cell name. A number is the index of a style in {@link SpreadsheetOptions.style}.
         */
        style?: Record<string, string | number>;

        /**
         * Set the max height of the table.
         * This property is only used when {@link WorksheetOptions.tableOverflow} is allowed.
         */
        tableHeight?: string | number;

        /**
         * Allow table overflow.
         * @default false
         */
        tableOverflow?: boolean;

        /**
         * Set the max width of the table.
         * This property is only used when {@link WorksheetOptions.tableOverflow} is allowed.
         */
        tableWidth?: string | number;

        /**
         * If true, cell contents may overflow over empty cells.
         * @default false
         */
        textOverflow?: boolean;

        /** Load a external json file from this URL. */
        url?: string;

        /**
         * Global text wrapping.
         * @default false
         */
        wordWrap?: boolean;

        /**
         * Worksheet name.
         */
        worksheetName?: string;
    }

    interface JspreadsheetInstanceElement extends HTMLDivElement {
        /**
         * Jss instance this element belongs to
         */
        spreadsheet: SpreadsheetInstance;
    }

    interface JworksheetInstanceElement extends HTMLDivElement {
        /**
         * Jss worksheet instance this element belongs to
         */
        jspreadsheet: WorksheetInstance;
    }

    interface DragInfo {
        /**
         * Index where the row or column will be moved.
         */
        destination: number;
    }

    interface DragColumnInfo extends DragInfo {
        /**
         * Index of the column being moved.
         */
        column: number;

        /**
         * HTML element that corresponds to the column being moved.
         */
        element: HTMLTableCellElement;
    }

    interface DragRowInfo extends DragInfo {
        /**
         * Index of the row being moved.
         */
        row: number;

        /**
         * HTML element that corresponds to the row being moved.
         */
        element: HTMLTableRowElement;
    }

    interface ResizeInfo {
        /**
         * Mouse position when resizing started. This position refers to the vertical axis, when resizing a row, or the horizontal axis, when resizing a column.
         */
        mousePosition: number;
    }

    interface ResizeRowInfo extends ResizeInfo {
        /**
         * HTML element that represents the row.
         */
        element: HTMLTableRowElement;

        /**
         * Old row height.
         */
        height: number;

        /**
         * Index of the row being resized.
         */
        row: number;
    }

    interface ResizeColumnInfo extends ResizeInfo {
        /**
         * Index of the column being resized.
         */
        column: number;

        /**
         * Old column width.
         */
        width: number;
    }

    /**
     * Creates a plugin. The spreadsheet calls it with itself, an empty object and the spreadsheet options.
     */
    type PluginFactory = (spreadsheet: SpreadsheetInstance, options: object, config: SpreadsheetOptions) => Plugin;

    interface Plugin {
        /**
         * This method is called before a worksheet is created.
         * @param instance - New worksheet instance.
         */
        beforeinit?: (instance: WorksheetInstance) => void;

        /**
         * Get spreadsheet config information. The library does not call it.
         */
        getConfig?: () => SpreadsheetOptions;

        /**
         * This method is called when a worksheet is created.
         * @param instance - New worksheet instance.
         */
        init?: (instance: WorksheetInstance) => void;

        /**
         * Event fired when any other event fires, after the handlers of the spreadsheet options.
         *
         * Returning false cancels an `onbefore…` action, as the handler in the options does. Returning nothing keeps the
         * result of the handlers before it, and returning another value replaces it.
         *
         * It gets the name of the event that was called, then the arguments of the event. See {@link EventListener}.
         */
        onevent?: EventListener;

        /**
         * This method is called before the spreadsheet sends data to the server.
         * @param instance - Worksheet instance.
         * @param method - Name of the method whose execution needs to be saved on the server.
         * @param data - Arguments of the method whose execution needs to be saved on the server: the changes, as onafterchanges gets them.
         */
        persistence?: (instance: WorksheetInstance, method: string, data: { data: CellChange[] }) => void;

        /**
         * Method called before the context menu is displayed. If this method returns anything other than a falsy value, that value overrides the context menu settings.
         * @param instance - Instance of the worksheet on which the click was made.
         * @param colIndex - Horizontal index of the element that was clicked. The meaning of this value depends on the {@link role} argument.
         * @param rowIndex - Vertical index of the element that was clicked. The meaning of this value depends on the {@link role} argument.
         * @param event - pointer event that triggered this method.
         * @param items - jss default context menu.
         * @param role - indicates in which part of the spreadsheet the click occurred.
         * @param x - Horizontal index of the element that was clicked. The meaning of this value depends on the {@link role} argument.
         * @param y - Vertical index of the element that was clicked. The meaning of this value depends on the {@link role} argument.
         * @returns Context menu configuration that should be created.
         */
        contextMenu?: (
            instance: WorksheetInstance,
            colIndex: number | null,
            rowIndex: number | null,
            event: PointerEvent,
            items: ContextMenuItem[],
            role: ContextMenuRole,
            x: number | null,
            y: number | null
        ) => ContextMenuItem[] | null | undefined;

        /**
         * Method called before the toolbar is displayed. If this method returns anything other than a falsy value, that value overrides the toolbar settings.
         * @param toolbar - The toolbar settings.
         */
        toolbar?: (toolbar: ToolbarOptions) => ToolbarOptions | null | undefined;
    }

    /**
     * A spreadsheet: an instance of the Spreadsheet class, which holds the worksheets, the toolbar and the settings. The
     * methods are on the prototype, so they need the spreadsheet as `this`. The fields that only the library uses are
     * not declared here.
     */
    interface SpreadsheetInstance {
        /**
         * Read a text to the user of assistive technology, for example "B4 changed by another user". The spreadsheet has a
         * polite live region, so the text does not stop what the screen reader reads. The region is hidden on the screen.
         * The spreadsheet adds the region when it is built, so call `announce` after `onload`. Each call adds the text as a
         * new element, so the same text two times is read two times.
         * @param text - Text to read.
         */
        announce: (text: string) => void;

        /**
         * Spreadsheet settings.
         */
        config: SpreadsheetOptions;

        /**
         * The element of the context menu of the spreadsheet.
         */
        contextMenu: HTMLDivElement;

        /**
         * Root HTML element of this jss instance.
         */
        el: JspreadsheetInstanceElement;

        /**
         * Alias for el.
         */
        element: JspreadsheetInstanceElement;

        /**
         * Toogle table fullscreen mode.
         * @param activate - Desired mode. Default: The opposite of the current mode.
         */
        fullscreen: (activate?: boolean) => void;

        /**
         * Get the spreadsheet settings.
         */
        getConfig: () => SpreadsheetOptions;

        /**
         * Get the index of the currently active worksheet
         */
        getWorksheetActive: () => number;

        /**
         * Hide the toolbar.
         */
        hideToolbar: () => void;

        /**
         * If true, the spreadsheet does not emit events, except onbeforerequest, onbeforesave and onsave: the persistence
         * still saves the changes, for example of an undo or a redo.
         */
        ignoreEvents?: boolean;

        /**
         * Spreadsheet plugins.
         */
        plugins: Record<string, Plugin>;

        /**
         * Change the spreadsheet settings.
         * @param config - The settings to change.
         */
        setConfig: (config: Partial<SpreadsheetOptions>) => void;

        /**
         * Add new plugins to the spreadsheet.
         * @param plugins - New plugins.
         */
        setPlugins: (plugins: Record<string, PluginFactory>) => void;

        /**
         * Show the toolbar using the current settings.
         */
        showToolbar: () => void;

        /**
         * HTML div tag used as the toolbar of this jss instance.
         */
        toolbar: HTMLDivElement;

        /**
         * Instances of the worksheets that make up this spreadsheet.
         */
        worksheets: WorksheetInstance[];
    }

    /**
     * A worksheet: an instance of the Worksheet class, one table of a spreadsheet. `jspreadsheet()` returns the
     * worksheets. The methods are on the prototype, so they need the worksheet as `this`. The fields that only the
     * library uses, such as the col tags and the selection corner, are not declared here.
     */
    interface WorksheetInstance {
        /**
         * Close a cell editor. Does nothing when the cell is not the cell being edited.
         * @param cell - HTML td tag whose editor must be closed.
         * @param save - Whether or not to save editor content in cell.
         */
        closeEditor: (cell: HTMLTableCellElement, save: boolean) => void;

        content: HTMLDivElement;

        /**
         * Copies or cuts the contents of selected cells in the worksheet. Does nothing when no cell is selected.
         * @param cut - If true, the operation is cut, if not, it is copy.
         */
        copy: (cut?: boolean) => void;

        /**
         * Create a new worksheet.
         * @param options - Worksheet options.
         */
        createWorksheet: (options: WorksheetOptions) => void;

        /**
         * Remove columns.
         *
         * This method returns false if the {@link SpreadsheetOptions.onbeforedeletecolumn} event returns false or if the "This action will destroy any existing merged cells. Are you sure?" dialog receives a negative response.
         * @param columnNumber - Column index from which removal starts.
         * @param numOfColumns - Number of columns to be removed.
         */
        deleteColumn: (columnNumber?: number, numOfColumns?: number) => false | undefined;

        /**
         * Remove rows.
         *
         * This method returns false if the {@link SpreadsheetOptions.onbeforedeleterow} event returns false or if the dialog cases "This action will destroy any existing merged cells. Are you sure?" or "This action will clear your search results. Are you sure?" receive a negative response.
         * @param rowNumber - Row index from which removal starts.
         * @param numOfRows - Number of rows to be removed.
         */
        deleteRow: (rowNumber?: number, numOfRows?: number) => false | undefined;

        /**
         * Delete a worksheet.
         * @param position - Worksheet index.
         */
        deleteWorksheet: (position: number) => void;

        /**
         * Remove all merged cells.
         */
        destroyMerge: () => void;

        /**
         * Emit an event.
         * @param event - Event name.
         * @param args - Arguments that should be passed to the event.
         * @returns The result of the last handler that ran, or null when none ran.
         */
        dispatch: <E extends EventName>(event: E, ...args: Parameters<EventHandlers[E]>) => unknown;

        /**
         * Simulates the action of the "arrow down" key. Does nothing when no cell is selected.
         * @param shiftKey - If true, the method simulates the action of the "arrow down" key while the Shift key is pressed.
         * @param ctrlKey - If true, the method simulates the action of the "arrow down" key while the Ctrl key is pressed.
         */
        down: (shiftKey?: boolean, ctrlKey?: boolean) => void;

        /**
         * Get the current data as a CSV file.
         * @param includeHeaders - If true, include the header regardless of the {@link SpreadsheetOptions.includeHeadersOnDownload} property value.
         * @param processed - If true, the result will contain the displayed cell values. Otherwise, the result will contain the actual cell values.
         */
        download: (includeHeaders?: boolean, processed?: boolean) => void;

        /**
         * Root HTML element of this worksheet instance.
         */
        element: JworksheetInstanceElement;

        /**
         * Execute a formula.
         * @param expression - Formula to be executed.
         * @param x - Column index of the cell where the formula is.
         * @param y - Row index of the cell where the formula is.
         * @returns The result of the formula, as the formula engine gives it, or '#ERROR'. A formula that refers to a range
         * can give an array. Without a formula engine, '#ENGINE'.
         */
        executeFormula: (expression: string, x?: number, y?: number) => unknown;

        /**
         * Table row containing filter inputs.
         */
        filter: null | HTMLTableRowElement;

        /**
         * Simulates the action of the Home key. Does nothing when no cell is selected.
         * @param shiftKey - If true, the method simulates the action of the Home key while the Shift key is pressed.
         * @param ctrlKey - If true, the method simulates the action of the Home key while the Ctrl key is pressed.
         */
        first: (shiftKey?: boolean, ctrlKey?: boolean) => void;

        /** The formula chain: which formulas the worksheet calculates again when a cell changes. */
        formulas: WorksheetFormulas;

        /**
         * Get cell DOM element by cell name.
         * @param cell - Cell name.
         */
        getCell(cell: string): HTMLTableCellElement;

        /**
         * Get cell DOM element by cell coords.
         * @param x - Cell column index.
         * @param y - Cell row index.
         */
        getCell(x: number, y: number): HTMLTableCellElement;

        /**
         * Get cell DOM element by cell coordinates.
         * @param x - Column index of the cell.
         * @param y - Row index of the cell.
         */
        getCellFromCoords: (x: number, y: number) => HTMLTableCellElement;

        /**
         * Get the data from one column by its index.
         * @param columnNumber - Column index.
         * @param processed - If true, the return is constructed using the innerHTML of the cells. Otherwise, it is constructed using the {@link WorksheetOptions.data} property. Default: false.
         */
        getColumnData: (columnNumber: number, processed?: boolean) => CellValue[];

        /**
         * Get comments from one or all cells.
         * @param cell - Cell name. If it is a falsy value, the comments of all cells are returned.
         */
        getComments: (cell?: string) => Record<string, string> | string;

        /**
         * Get worksheet config information.
         */
        getConfig: () => WorksheetOptions;

        /**
         * Get the full or partial table data.
         * @param highlighted - If true, get only data from highlighted cells. If false, get data from all cells. Default: false.
         * @param processed - If true, the return is constructed using the innerHTML of the cells. Otherwise, it is constructed using the {@link WorksheetOptions.data} property. Default: false.
         * @param delimiter - Column delimiter. If this property is specified, the result is text formatted like a csv, with a line break after each row. An empty delimiter counts as none.
         * @param asJson - If this property is true and there is no delimiter, each row is an object with the column indexes as keys.
         */
        getData: {
            (highlighted?: boolean, processed?: boolean, delimiter?: undefined | null | '', asJson?: false): CellValue[][];
            (highlighted: boolean | undefined, processed: boolean | undefined, delimiter: '', asJson: true): Record<number, CellValue>[];
            (highlighted: boolean | undefined, processed: boolean | undefined, delimiter: string, asJson?: boolean): string;
            (highlighted: boolean | undefined, processed: boolean | undefined, delimiter: undefined | null, asJson: true): Record<number, CellValue>[];
            (highlighted?: boolean, processed?: boolean, delimiter?: string | null, asJson?: boolean): CellValue[][] | Record<number, CellValue>[] | string;
        };

        /**
         * Get data from a range.
         * @param range - Range of cells whose values are to be returned.
         * @param processed - If true, the method returns the values of the HTML elements of the cells. Otherwise, the method returns the values of the cells in the options.data array.
         */
        getDataFromRange: (range: string, processed?: boolean) => CellValue[][];

        /**
         * Get the column title.
         * @param column - Column index.
         */
        getHeader: (column: number) => string;

        /**
         * Get all header titles.
         * @param asArray - If true, returns the items in an array, if false, returns them separated by ";" within a single string.
         */
        getHeaders: (asArray?: boolean) => string | string[];

        /**
         * Get height of all rows.
         */
        getHeight(row?: undefined): string[];

        /**
         * Get height of one row.
         * @param row - Column index.
         */
        getHeight(row: number): string;

        /**
         * Get the coordinates of the highlighted selections.
         */
        getHighlighted: () => [number, number, number, number][];

        /**
         * Get the innerHTML of a cell.
         * @param cell - Cell name.
         */
        getLabel(cell: string): string;

        /**
         * Get the innerHTML of a cell.
         * @param x - Cell column index.
         * @param y - Cell row index.
         */
        getLabel(x: number, y: number): string;

        /**
         * Get information from one or all merged cells
         * @param cellName - Cell name. If it is a falsy value, it returns the information of all merges. If the given cell is not the anchor of a merge, it returns null.
         */
        getMerge(cellName?: string): Record<string, [number, number]> | [number, number] | null;

        /**
         * Get the meta information of all cells.
         * @returns The meta information by cell name, or undefined when no cell has meta information.
         */
        getMeta(cell?: undefined): Record<string, MetaInformation> | undefined;

        /**
         * Get the meta information of a cell.
         * @param cell - Cell name. An empty name gives the meta information of all cells.
         * @returns The meta information of the cell, or null when it has none.
         */
        getMeta(cell: string): MetaInformation | null;

        /**
         * Get the range description of the selected cells.
         */
        getRange: () => string;

        /**
         * Get data from a row by its index.
         * @param rowNumber - Row index.
         * @param processed - If true, the return is constructed using the innerHTML of the cells. Otherwise, it is constructed using the {@link WorksheetOptions.data} property. Default: false.
         */
        getRowData: (rowNumber: number, processed?: boolean) => CellValue[] | undefined;

        /**
         * Get information from selected cells in the worksheet.
         * @param columnNameOnly - If true, the method returns the names of the selected cells. Otherwise, the method returns the records of the selected cells.
         */
        getSelected: {
            (columnNameOnly?: false): { element: HTMLTableCellElement; x: number; y: number }[];
            (columnNameOnly: true): string[];
            (columnNameOnly?: boolean): { element: HTMLTableCellElement; x: number; y: number }[] | string[];
        };

        /**
         * Get indexes of the columns that have highlighted cells.
         * @param visibleOnly - If true, the method returns only visible columns.
         */
        getSelectedColumns: (visibleOnly?: boolean) => number[];

        /**
         * Get indexes of the rows that have highlighted cells.
         * @param visibleOnly - If true, the method returns only visible rows.
         */
        getSelectedRows: (visibleOnly?: boolean) => number[];

        /**
         * Get the coordinates of the range that is selected in the worksheet.
         * @returns The first column, the first row, the last column and the last row, or null when no cell is selected.
         */
        getSelection: () => [number, number, number, number] | null;

        /**
         * Get styles from one or all cells.
         * @param cell - Name or coordinate of a cell. If omitted, returns styles for all cells.
         * @param key - Style property. if specified, returns only that property. Otherwise, it returns all the cell's style properties.
         */
        getStyle: (cell?: string | [number, number], key?: string) => string | Record<string, string>;

        /**
         * Get the value of a cell.
         * @param cell - Cell name.
         * @param processedValue - If true, it returns the cell's innerHTML. Otherwise, it returns the value of the cell in the {@link WorksheetOptions.data} property.
         */
        getValue: (cell: string, processedValue?: boolean) => CellValue | null;

        /**
         * Get the value of a cell by its coordinates.
         * @param x - Column index.
         * @param y - Row index.
         * @param processedValue - If true, it returns the cell's innerHTML. Otherwise, it returns the value of the cell in the {@link WorksheetOptions.data} property.
         */
        getValueFromCoords: (x: number, y: number, processedValue?: boolean) => CellValue | null;

        /**
         * Get the width of one or all columns.
         * @param column - Index of the column. If omitted, returns the widths of all columns.
         */
        getWidth: (column?: number) => number | (number | string)[];

        /**
         * Get the index of the currently active worksheet
         */
        getWorksheetActive: () => number;

        /**
         * List of cells that make up the header.
         */
        headers: HTMLTableCellElement[];

        /**
         * Hide a column.
         * @param colNumber - Column indexes.
         */
        hideColumn: (colNumber: number | number[]) => void;

        /**
         * Hide row count column.
         */
        hideIndex: () => void;

        /**
         * Hide Row.
         * @param rowNumber - Row indexes.
         */
        hideRow: (rowNumber: number | number[]) => void;

        /**
         * The undo history of the worksheet.
         */
        history: WorksheetHistory;

        /**
         * What the worksheet copied last.
         */
        clipboard: WorksheetClipboard;

        /**
         * The cell editor of the worksheet.
         */
        editor: WorksheetEditor;

        /**
         * Check if a cell is within the current selection.
         * @param x - Cell column index.
         * @param y - Cell row index.
         */
        isSelected: (x: number, y: number) => boolean;

        /**
         * Insert one or more columns.
         *
         * This method returns false if the {@link SpreadsheetOptions.onbeforeinsertcolumn} event returns false or if the "This action will destroy any existing merged cells. Are you sure?" dialog receives a negative response.
         * @param mixed - Number of columns to insert. It can also be an array of values, but in this case, only one column is inserted, whose data is based on the array items. Default: 1.
         * @param columnNumber - Index of the column used as reference for the insertion. Default: last column.
         * @param insertBefore - Insert new columns before (a truthy value) or after the reference column. Default: false.
         * @param properties - New column properties.
         */
        insertColumn: (mixed?: number | CellValue[], columnNumber?: number, insertBefore?: boolean | number, properties?: Column[]) => false | undefined;

        /**
         * Insert one or more rows.
         *
         * This method returns false if the {@link SpreadsheetOptions.onbeforeinsertrow} event returns false or if the "This action will destroy any existing merged cells. Are you sure?" or "This action will clear your search results. Are you sure?" dialogs receive a negative response.
         * @param mixed - Number of rows to insert. It can also be an array of values, but in this case, only one row is inserted, whose data is based on the array items. Default: 1.
         * @param rowNumber - Index of the row used as reference for the insertion. Default: last row.
         * @param insertBefore - Insert new rows before (a truthy value) or after the reference row. Default: false.
         */
        insertRow: (mixed?: number | CellValue[], rowNumber?: number, insertBefore?: boolean | number) => false | undefined;

        /**
         * Check if a cell is read only.
         * @param x - Cell column index.
         * @param y - Cell row index.
         */
        isReadOnly(x: number, y: number): boolean;

        /**
         * Check if a cell is read only.
         * @param cellName - Cell name.
         */
        isReadOnly(cellName: string): boolean;

        /**
         * Set the accessible name of a cell again, from its address, its text and the `cellNote` option. Call it after the
         * note of a cell changes. Only the active cell has a name, so the call does nothing for another cell.
         * @param x - Column index of the cell.
         * @param y - Row index of the cell.
         */
        labelCell: (x: number, y: number) => void;

        /**
         * Simulates the action of the End key. Does nothing when no cell is selected.
         * @param shiftKey - If true, the method simulates the action of the End key while the Shift key is pressed.
         * @param ctrlKey - If true, the method simulates the action of the End key while the Ctrl key is pressed.
         */
        last: (shiftKey?: boolean, ctrlKey?: boolean) => void;

        /**
         * Simulates the action of the "arrow left" key. Does nothing when no cell is selected.
         * @param shiftKey - If true, the method simulates the action of the "arrow left" key while the Shift key is pressed.
         * @param ctrlKey - If true, the method simulates the action of the "arrow left" key while the Ctrl key is pressed.
         */
        left: (shiftKey?: boolean, ctrlKey?: boolean) => void;

        /**
         * Move a column.
         *
         * This method returns false if the "This action will destroy any existing merged cells. Are you sure?" dialog receives a negative response.
         * @param o - Column index.
         * @param d - New column index.
         */
        moveColumn: (o: number, d: number) => false | undefined;

        /**
         * Move a row.
         *
         * This method returns false if the "This action will destroy any existing merged cells. Are you sure?" or "This action will clear your search results. Are you sure?" dialogs receive a negative response.
         * @param o - Row index.
         * @param d - New row index.
         */
        moveRow: (o: number, d: number) => false | undefined;

        /**
         * Start the edition for one cell.
         * @param cell - Cell HTML Element.
         * @param empty - If true, the editor opens without content even if the cell had content.
         * @param event - Js event that triggered the editor opening. This argument is passed to the "openEditor" method of custom editors.
         */
        openEditor: (cell: HTMLTableCellElement, empty?: boolean, event?: KeyboardEvent | MouseEvent | TouchEvent) => void;

        /**
         * Open the column filter.
         *
         * This method only runs if the {@link WorksheetOptions.filters} property is true.
         * @param columnId - Column index.
         */
        openFilter: (columnId: number) => void;

        /**
         * Open the worksheet by index.
         * @param position - Worksheet index.
         */
        openWorksheet: (position: number) => void;

        /**
         * Spreadsheet settings.
         */
        options: WorksheetOptions;

        /**
         * Reorder rows based on values in a column.
         *
         * This method returns false if the "This action will destroy any existing merged cells. Are you sure?" dialog receives a negative response, or returns true if the sort is successful.
         * @param column - Column index. If the value of this parameter is less than 0, the method returns false and does not perform sorting.
         * @param order - Sorting direction. 0 for ascending and 1 for descending. Without it, the direction changes: descending when the column is sorted ascending, else ascending.
         */
        orderBy: (column: number, order?: 0 | 1) => boolean | undefined;

        /**
         * Go to page. Valid only when {@link WorksheetOptions.pagination} is true.
         * @param pageNumber - Page number (starting at 0).
         */
        page: (pageNumber: number) => void;

        /**
         * Div with pagination controls.
         */
        pagination: HTMLDivElement;

        /**
         * Spreadsheet of which this worksheet is part.
         */
        parent: SpreadsheetInstance;

        /**
         * Pastes content into one or more cells. When the selection is larger than the content, the content is repeated over the selection.
         * @param x - Column index of the cell from which the content will be pasted.
         * @param y - Row index of the cell from which the content will be pasted.
         * @param data - Content to be pasted.
         */
        paste: (x: number, y: number, data: string) => false | undefined;

        /**
         * The state of a mouse gesture: the column or row that the mouse resizes or moves, and the drag of the selection corner.
         */
        pointer: WorksheetPointer;

        /**
         * Get the number of pages of the worksheet.
         */
        quantiyOfPages: () => number;

        /**
         * List of HTML elements representing table cells.
         */
        records: {
            element: HTMLTableCellElement;
            x: number;
            y: number;
        }[][];

        /**
         * Redo previously undone action
         */
        redo: () => void;

        /**
         * Remove a merge.
         * @param cellName - Merge anchor cell.
         * @param data - Data to be placed in cells released from the merge.
         */
        removeMerge: (cellName: string, data?: CellValue[]) => void;

        /**
         * Reset all filters.
         */
        resetFilters: () => void;

        /**
         * Reset search
         */
        resetSearch: () => void;

        /**
         * Reset highlighted cell selection.
         * @returns If there were highlighted cells, it returns 1, otherwise it returns 0.
         */
        resetSelection: () => 0 | 1;

        /**
         * Reset styles of one or more cells.
         * @param o - Object whose keys are the names of the cells that must have their styles reset. Each value is the new style of the cell, as in {@link setStyle}.
         * @param ignoreHistoryAndEvents - If true, do not add this action to history.
         */
        resetStyle: (o: Record<string, string | string[]>, ignoreHistoryAndEvents?: boolean) => void;

        /**
         * Simulates the action of the "arrow right" key. Does nothing when no cell is selected.
         * @param shiftKey - If true, the method simulates the action of the "arrow right" key while the Shift key is pressed.
         * @param ctrlKey - If true, the method simulates the action of the "arrow right" key while the Ctrl key is pressed.
         */
        right: (shiftKey?: boolean, ctrlKey?: boolean) => void;

        /**
         * List of rows that make up the table.
         */
        rows: {
            element: HTMLTableRowElement;
            y: number;
        }[];

        /**
         * Search for some text.
         * @param query - Text to be searched.
         */
        search: (query: string) => void;

        /**
         * Text field used to perform searches.
         */
        searchInput: HTMLInputElement;

        /**
         * Select all table cells.
         */
        selectAll: () => void;

        /**
         * The selection of the worksheet.
         */
        selection: WorksheetSelection;

        /**
         * Set the data from one column by index.
         * @param colNumber - Column index.
         * @param data - New data. Positions with the null value are not changed in the table.
         * @param force - If true, the method also changes the contents of readonly columns.
         */
        setColumnData: (colNumber: number, data: (CellValue | null)[], force?: boolean) => void;

        /**
         * Set or remove a comment.
         * @param cellId - Name of the cell.
         * @param comments - New comment. If it is a falsy value, the method just uncomments the cell.
         */
        setComments(cellId: string, comments: string | null | undefined): void;

        /**
         * Set or remove comments.
         * @param cellId - Object whose keys are cell names and values are cell comments. If the value of a key is a falsy value, the cell comment is removed.
         */
        setComments(cellId: Record<string, string | null | undefined>): void;

        /**
         * Change the worksheet settings.
         * @param config - The settings of the worksheet to change.
         * @param spreadsheetLevel - False or missing: the settings are applied to the worksheet.
         */
        setConfig(config: Partial<WorksheetOptions>, spreadsheetLevel?: false): void;

        /**
         * Change the spreadsheet settings.
         * @param config - The settings of the spreadsheet to change.
         * @param spreadsheetLevel - True: the settings are applied to the spreadsheet.
         */
        setConfig(config: Partial<SpreadsheetOptions>, spreadsheetLevel: true): void;

        /**
         * Set data.
         * @param data - New data. It can be an array of cell values or an array of objects whose values are cell values.
         */
        setData: (data?: CellValue[][] | Record<string, CellValue>[]) => void;

        /**
         * Set a column title.
         * @param column - Column index.
         * @param newValue - New title. Empty string or undefined to reset the header title.
         */
        setHeader: (column: number, newValue?: string) => void;

        /**
         * Change row height.
         * @param row - Row index.
         * @param height - New height. An integer greater than zero.
         */
        setHeight: (row: number, height: number) => void;

        /**
         * Merge cells.
         * @param cellName - Name of a cell. If it is a falsy value, this method merges the selected cells in the table and ignores all parameters of this method.
         * @param colspan - Number of columns this merge occupies.
         * @param rowspan - Number of rows this merge occupies.
         * @returns If the "cellName" parameter is a falsy value, and there are no cells selected in the table, this method returns null.
         */
        setMerge: (cellName?: string, colspan?: number, rowspan?: number) => null | undefined;

        /**
         * Set a property on a cell's meta information.
         * @param o - Cell name.
         * @param k - Property name.
         * @param v - Property value.
         */
        setMeta(o: string, k: string, v: string): void;

        /**
         * Remove current and define new meta information for one or more cells.
         * @param o - Object with the new meta information.
         */
        setMeta(o: Record<string, MetaInformation>): void;

        /**
         * Change the read only state of a cell.
         * @param cell - Cell HTML element or its name.
         * @param state - New read only state.
         */
        setReadOnly: (cell: string | HTMLTableCellElement, state: boolean) => void;

        /**
         * Set a row data by index.
         * @param rowNumber - Row index.
         * @param data - New data. Positions with the null value are not changed in the table.
         * @param force - If true, the method also changes the contents of readonly columns.
         */
        setRowData: (rowNumber: number, data: (CellValue | null)[], force?: boolean) => void;

        /**
         * Change a single style of one or more cells.
         * @param o - Name of a cell.
         * @param k - property to be changed.
         * @param v - New property value. If equal to the property's current value and the "force" parameter is false, removes that property from the style.
         * @param force - If true, changes the value of the property even if the cell is read-only. Also, if true, even if the new value of the property is the same as the current one, the property is not removed.
         * @param ignoreHistoryAndEvents - If true, do not add this action to history. The onchangestyle event still occurs.
         */
        setStyle(o: string, k: string, v: string, force?: boolean, ignoreHistoryAndEvents?: boolean): void;

        /**
         * Change cell styles.
         * @param o - Object where each key is the name of a cell and each value is the style changes for that cell. Each value can be a string with css styles separated by semicolons or an array where each item is a string with a css style.
         * @param k - It is not used.
         * @param v - It is not used.
         * @param force - If true, changes the value of the property even if the cell is read-only. Also, if true, even if the new value of the property is the same as the current one, the property is not removed.
         * @param ignoreHistoryAndEvents - If true, do not add this action to history. The onchangestyle event still occurs.
         */
        setStyle(o: Record<string, string | string[]>, k?: null | undefined, v?: null | undefined, force?: boolean, ignoreHistoryAndEvents?: boolean): void;

        /**
         * Change the value of one or more cells.
         * @param cell - Name of a cell, HTML element that represents a cell or an array whose items can be any of the previous alternatives or objects. When an array item is an object, it must have the cell coordinates ("x" and "y") and can have the cell's new value ("value"), but if does not have it, the "value" parameter is used instead.
         * @param value - New cell value.
         * @param force - If true, changes the value of even read-only cells.
         */
        setValue: (
            cell: string | HTMLTableCellElement | (string | { x: number; y: number; value?: CellValue } | HTMLTableCellElement)[],
            value?: CellValue,
            force?: boolean
        ) => void;

        /**
         * Set a cell value based on its coordinates.
         * @param x - Cell column index.
         * @param y - Cell row index.
         * @param value - New value.
         * @param force - If true, changes the value of even read-only cells.
         */
        setValueFromCoords: (x: number, y: number, value: CellValue, force?: boolean) => void;

        /**
         * Set the width of a column.
         * @param column - Column index.
         * @param width - New width.
         */
        setWidth(column: number, width: number): void;

        /**
         * Set the width of one or more columns.
         * @param column - Column indexes.
         * @param width - New widths.
         */
        setWidth(column: number[], width: number | number[]): void;

        /**
         * Show hidden column.
         * @param colNumber - Column index.
         */
        showColumn: (colNumber: number | number[]) => void;

        /**
         * Show row count column.
         */
        showIndex: () => void;

        /**
         * Show hidden row.
         * @param rowNumber - Row index.
         */
        showRow: (rowNumber: number | number[]) => void;

        /**
         * HTML table tag of this jss instance.
         */
        table: HTMLTableElement;

        /**
         * HTML tbody tag of this jss instance.
         */
        tbody: HTMLTableSectionElement;

        /**
         * HTML thead tag of this jss instance.
         */
        thead: HTMLTableSectionElement;

        /**
         * Undo last action.
         */
        undo: () => void;

        /**
         * Simulates the action of the "arrow up" key. Does nothing when no cell is selected.
         * @param shiftKey - If true, the method simulates the action of the "arrow up" key while the Shift key is pressed.
         * @param ctrlKey - If true, the method simulates the action of the "arrow up" key while the Ctrl key is pressed.
         */
        up: (shiftKey?: boolean, ctrlKey?: boolean) => void;

        /**
         * Select cells.
         * @param x1 - Column index of the first cell of the selection. If omitted or null, rows "y1" through "y2" are selected.
         * @param y1 - Row index of the first cell of the selection. If omitted or null, columns "x1" through "x2" are selected.
         * @param x2 - Column index of the last cell of the selection. Default: Parameter "x1".
         * @param y2 - Row index of the last cell of the selection. Default: Parameter "y1".
         */
        updateSelectionFromCoords: (x1: number | null, y1: number | null, x2?: number | null, y2?: number | null) => false | undefined;

        /** Which rows the worksheet shows: the results of a search or the filters, the filter values and the page. */
        view: WorksheetView;

        /**
         * Get the page index of a row.
         * @param cell - Row index.
         */
        whichPage: (cell: number) => number;
    }

    type Version = () => {
        host: string;
        license: string;
        print: () => [string];
        version: string;
    };

    interface JssHelpers {
        /**
         * Extract the configuration to create a new spreadsheet from a static HTML element.
         * @param element - Table element.
         * @param options - Worksheet options, which the options of the table are added to.
         */
        createFromTable: (element: HTMLTableElement, options?: WorksheetOptions) => WorksheetOptions;

        /**
         * Internal method. Get the position of the caret in an element.
         */
        getCaretIndex: (e: Node) => number;

        /**
         * Get the column letter(s) based on its index.
         * @param i - Column index.
         */
        getColumnName: (i: number) => string;

        /**
         * Get "A1" style coordinates based on column and row indices.
         * @param x - Column index.
         * @param y - Row index.
         */
        getCellNameFromCoords: (x: number, y: number) => string;

        /**
         * Get column and row indices based on coordinate in "A1" style.
         * @param columnName - Coordinate in "A1" style.
         */
        getCoordsFromCellName: (columnName: string) => [number, number | null] | undefined;

        /**
         * Get coordinates from a range.
         * @param range - Range in "A1:B2" style.
         * @returns Array filled with the x and y coordinates of the first and last cells in the range.
         */
        getCoordsFromRange: (range: string) => [number, number, number, number];

        /**
         * Internal method. Swap the keys and the values of an object. A number value becomes an index of the list.
         */
        invert: (o: Record<string, string | number>) => string[] & Record<string, string>;

        /**
         * Parse CSV string to JS array.
         * @param str - Text in csv format.
         * @param delimiter - Csv delimiter.
         */
        parseCSV: (str: string, delimiter?: string) => string[][];
    }

    /**
     * The options of a spreadsheet created from a table. The first worksheet is created from the table, so the worksheets
     * can be missing.
     */
    type TableSpreadsheetOptions = Omit<SpreadsheetOptions, 'worksheets'> & { worksheets?: WorksheetOptions[] };

    /**
     * An engine that calculates formulas, such as the default export of `@jspreadsheet/formula`. See
     * `jspreadsheet.setFormulaEngine`.
     * @param expression - The formula without the "=". Its references, and its text literals, are names of `variables`.
     * @param variables - The values of the references and the text literals, by name. A text is a JavaScript string literal.
     * @param x - Column index of the cell. Undefined for a footer, and when `executeFormula` gets no coordinates.
     * @param y - Row index of the cell. Undefined for a footer, and when `executeFormula` gets no coordinates.
     * @param instance - Worksheet of the formula.
     * @returns Value that the cell shows. If the function throws or returns a function, the cell shows "#ERROR".
     */
    type FormulaEngine = (
        expression: string,
        variables: Record<string, unknown>,
        x: number | undefined,
        y: number | undefined,
        instance: WorksheetInstance
    ) => unknown;

    interface JSpreadsheet {
        /**
         * Create a spreadsheet from a table, which is replaced by a div in its parent, so it must have a parent: without
         * one, this throws. The first worksheet loads the data and the options of the table. It returns the worksheets,
         * which are added when the spreadsheet is built, after it returns.
         * @param element - The table whose data the spreadsheet loads.
         * @param options - The options of the spreadsheet. The options of the first worksheet are added to the options of the table.
         */
        (element: HTMLTableElement, options?: TableSpreadsheetOptions): WorksheetInstance[];

        /**
         * Create a spreadsheet in an element. A table is replaced by a div in its parent, so it must have a parent: without
         * one, this throws. It returns the worksheets, which are added when the spreadsheet is built, after it returns.
         * @param element - The div, or the table whose data the spreadsheet loads.
         * @param options - The options of the spreadsheet.
         */
        (element: HTMLDivElement | HTMLTableElement, options: SpreadsheetOptions): WorksheetInstance[];

        /**
         * Current instance of jss.
         */
        current: null | WorksheetInstance;

        /**
         * Destroy an instance of jss.
         * @param element - Root element of jss instance.
         * @param destroyEventHandlers - Remove event listeners. Default: false.
         */
        destroy: (element: JspreadsheetInstanceElement, destroyEventHandlers?: boolean) => void;

        /**
         * Destroy all instances of jss.
         */
        destroyAll: () => void;

        /**
         * Get a worksheet instance by name and namespace.
         *
         * Returns undefined when the namespace has no worksheet with this name, and an empty object when no spreadsheet has the namespace.
         * @param worksheetName - Name of the searched worksheet. If null or undefined, the method returns the worksheets of the namespace by name.
         * @param namespace - Namespace name.
         */
        getWorksheetInstanceByName: (
            worksheetName: string | null | undefined,
            namespace: string
        ) => WorksheetInstance | Record<string, WorksheetInstance> | undefined;

        helpers: JssHelpers;

        /**
         * Register the engine that calculates the formulas of all spreadsheets, such as the default export of
         * `@jspreadsheet/formula`. Without an engine, a formula shows "#ENGINE", unless the spreadsheet has the
         * `formulaEngine` option. Register it before you create the spreadsheets: a formula that was calculated without an
         * engine shows "#ENGINE" until it is calculated again.
         * @param engine - The engine, or null to remove it.
         */
        setFormulaEngine: (engine: FormulaEngine | null) => void;

        /**
         * Defines translations.
         * @param o - Translations.
         */
        setDictionary: (o: Record<string, string>) => void;

        spreadsheet: SpreadsheetInstance[];

        /**
         * Basic version information.
         */
        version: Version;
    }
}
