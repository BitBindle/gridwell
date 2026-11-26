import type { Worksheet } from '../types/internal';

/**
 * Send a request of a worksheet to a server with fetch, and read the body of its response.
 * It sends the headers that jSuites.ajax sent: servers can tell from X-Requested-With that a script sent the request.
 * The onbeforerequest event can change the options of the request, or cancel it.
 * @param worksheet - The worksheet that sends the request.
 * @param url - The address of the request.
 * @param type - How to read the body of the response: as JSON or as text.
 * @param body - The form to post. Without it, the request is a GET.
 * @returns The body of the response, or null when onbeforerequest cancels the request, when it fails, or when the
 * server answers with a status that is not 2xx. A failure is logged with console.error.
 */
export async function request(worksheet: Worksheet, url: string, type: 'json' | 'text', body?: FormData): Promise<{ result: unknown } | null> {
    let options: RequestInit = {
        method: body ? 'POST' : 'GET',
        headers: body ? { 'X-Requested-With': 'XMLHttpRequest', Accept: 'application/json' } : { 'X-Requested-With': 'XMLHttpRequest' },
        body,
        cache: 'no-cache',
    };

    const ret = worksheet.dispatch('onbeforerequest', worksheet, url, options);
    if (ret === false) {
        return null;
    }
    if (ret && typeof ret === 'object') {
        options = ret;
    }

    try {
        const response = await fetch(url, options);
        if (!response.ok) {
            throw new Error(`Jspreadsheet: the ${options.method ?? 'GET'} request to ${url} failed with the status ${response.status}`);
        }

        return { result: type === 'json' ? await response.json() : await response.text() };
    } catch (error) {
        console.error(error);
        return null;
    }
}
