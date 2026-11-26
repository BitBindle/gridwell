/**
 * A stub of fetch for the tests, which answers the requests of the library with the responses that a test gives.
 */

/** A request that the library sent. */
export interface StubbedRequest {
    url: string;
    init: RequestInit;
}

/** The response to a url: its body, and its status, 200 by default. */
export interface StubbedResponse {
    body?: unknown;
    status?: number;
}

/**
 * Replace fetch with a stub until restore is called. The stub answers each request with the response of its url, or
 * with a 404 when there is none. A response is a body to read as JSON or as text.
 */
export function stubFetch() {
    const original = globalThis.fetch;
    const requests: StubbedRequest[] = [];
    const responses: Record<string, StubbedResponse> = {};

    globalThis.fetch = (async (url: string, init: RequestInit) => {
        requests.push({ url, init });
        const { body, status = 200 } = responses[url] ?? { status: 404 };
        return {
            ok: status >= 200 && status < 300,
            status,
            json: async () => (typeof body === 'string' ? JSON.parse(body) : body),
            text: async () => (typeof body === 'string' ? body : JSON.stringify(body)),
        };
    }) as unknown as typeof fetch;

    return {
        requests,
        responses,
        // Wait until the promises of the responses are resolved
        settle: () => new Promise((resolve) => setTimeout(resolve)),
        restore() {
            globalThis.fetch = original;
        },
    };
}
