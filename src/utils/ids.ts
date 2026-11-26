// A random part, so that two copies of the library on one page do not give the same ids
const prefix = Math.random().toString(36).slice(2, 8);

// The number of the last id
let last = 0;

/**
 * Give an id that no other element of the page has.
 * @param name - The start of the id, which tells what the element is.
 */
export function createId(name: string) {
    last++;
    return `${name}-${prefix}-${last}`;
}
