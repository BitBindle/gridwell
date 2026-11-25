/**
 * Which rows a worksheet shows: the rows that a search or the filters found, the values of the filters, and the page.
 * The search, the filters, the pagination and the lazy loading read and change it.
 *
 * Each worksheet has one, as worksheet.view. Showing the rows changes the DOM of the worksheet, so search, filters,
 * page and the lazy loading functions stay worksheet functions (src/utils/search.ts, filter.ts, pagination.ts and
 * lazyLoading.ts), and keep their state here.
 */
export class View {
    /** The indexes of the rows that a search or the filters found, or null when the worksheet shows every row. */
    results: number[] | null = null;

    /** The values that each column is filtered by, by column index. null or a missing entry is a column without filter. */
    filters: (string[] | null)[] = [];

    /** The page that the worksheet shows, from 0, when it has pagination. */
    pageNumber = 0;
}
