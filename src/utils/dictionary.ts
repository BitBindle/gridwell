// The translations of gridwell.setDictionary, by English text. A Map, so that a text such as "constructor" is not
// read from the prototype of an object.
const dictionary = new Map<string, string>();

/**
 * Add translations to the dictionary. A text that is already in it gets the new translation.
 * @param translations - The translations, by English text. An empty translation shows the English text again.
 */
export function setDictionary(translations: Record<string, string>) {
    for (const [text, translation] of Object.entries(translations)) {
        dictionary.set(text, translation);
    }
}

/**
 * Translate a text with the dictionary. A text without a translation stays as it is.
 * @param text - The English text.
 * @param context - Tells apart two texts that are the same in English, such as the month May and its short name. The
 * dictionary has the translation of the text by "text (context)", such as "May (short)".
 */
export function translate(text: string, context?: string) {
    return dictionary.get(context ? `${text} (${context})` : text) || text;
}
