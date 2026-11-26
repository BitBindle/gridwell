/**
 * Send a request to a server with fetch and read the body of its response.
 * It sends the headers that jSuites.ajax sent: servers can tell from X-Requested-With that a script sent the request.
 * The promise is rejected when the request fails, or when the server answers with a status that is not 2xx.
 * @param url - The address of the request.
 * @param type - How to read the body of the response: as JSON or as text.
 * @param body - The form to post. Without it, the request is a GET.
 */
export async function request(url: string, type: 'json' | 'text', body?: FormData): Promise<unknown> {
    const method = body ? 'POST' : 'GET';
    const headers: Record<string, string> = { 'X-Requested-With': 'XMLHttpRequest' };
    if (body) {
        headers.Accept = 'application/json';
    }

    const response = await fetch(url, { method, headers, body, cache: 'no-cache' });
    if (!response.ok) {
        throw new Error(`Jspreadsheet: the ${method} request to ${url} failed with the status ${response.status}`);
    }

    return type === 'json' ? response.json() : response.text();
}
