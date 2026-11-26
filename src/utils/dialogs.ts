import type jspreadsheet from '../jspreadsheet';

/** The options of a spreadsheet that answer its dialogs. */
type DialogOptions = Pick<jspreadsheet.SpreadsheetOptions, 'alert' | 'confirm'>;

/**
 * Ask the user to confirm an action. The confirm option answers in place of the dialog of the browser, and false confirms
 * without a dialog.
 * @param config - The options of the spreadsheet.
 * @param text - The translated question.
 */
export function askConfirm(config: DialogOptions, text: string) {
    if (config.confirm === false) {
        return true;
    }
    if (typeof config.confirm === 'function') {
        return config.confirm(text) === true;
    }
    return confirm(text);
}

/**
 * Show a message to the user. The alert option shows it in place of the dialog of the browser, and false shows nothing.
 * @param config - The options of the spreadsheet.
 * @param text - The translated message.
 */
export function showAlert(config: DialogOptions, text: string) {
    if (config.alert === false) {
        return;
    }
    if (typeof config.alert === 'function') {
        config.alert(text);
        return;
    }
    alert(text);
}
