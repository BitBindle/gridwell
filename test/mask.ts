import { expect } from 'chai';

import gridwell, { type Gridwell, type Worksheet } from './support/gridwell';

/**
 * The texts that the cells of the first row show.
 * @param columns - The columns.
 * @param data - The values of the first row.
 */
function shown(columns: Gridwell.Column[], data: Gridwell.CellValue[]) {
    const worksheet = gridwell(root, { worksheets: [{ data: [data], columns }] })[0];
    return worksheet.records[0].map((record) => record.element.textContent);
}

/**
 * Open the editor of a cell, type a text in it and close it.
 * @param worksheet - The worksheet.
 * @param name - The name of the cell, such as "A1".
 * @param text - The text to type, or undefined to keep the text of the editor.
 * @returns The text that the editor showed when it opened.
 */
function edit(worksheet: Worksheet, name: string, text?: string) {
    const cell = worksheet.getCell(name);
    worksheet.openEditor(cell);
    const input = cell.children[0] as HTMLInputElement;
    const opened = input.value;
    if (text !== undefined) {
        input.value = text;
    }
    worksheet.closeEditor(cell, true);
    return opened;
}

describe('Number formats', () => {
    it('shows the decimals, the thousands separators and the symbols of the mask', () => {
        const masks = ['0.00', '#.##', '#,##0.00', '#.##0,00', '$ #,##0.00', '0 "kg"', '000', '#', '0', '#,##0', '0.0#', '[$€-407] #,##0.00'];
        const values = [1.5, 1.5, 1234.5, 1234.5, 1234.5, 12, 7, 0.5, 0.4, 1234567, 1.5, 3];
        expect(
            shown(
                masks.map((mask) => ({ type: 'numeric', mask })),
                values
            )
        ).to.eql(['1.50', '1.5', '1,234.50', '1.234,50', '$ 1,234.50', '12 kg', '007', '1', '0', '1,234,567', '1.5', '€ 3.00']);
    });

    it('rounds the decimals as the number is written', () => {
        expect(shown([{ mask: '0.00' }, { mask: '0.0' }, { mask: '0' }], [1.005, 0.05, 2.5])).to.eql(['1.01', '0.1', '3']);
    });

    it('takes the decimal separator of the column, and reads the other separator as the thousands separator', () => {
        expect(
            shown(
                [
                    { mask: '$ #.##,00', decimal: ',' },
                    { mask: '#.##0', decimal: ',' },
                ],
                [2000, 1234567]
            )
        ).to.eql(['$ 2.000,00', '1.234.567']);
    });

    it('shows a percentage, a number in scientific notation and a phone number', () => {
        expect(
            shown(
                [
                    { mask: '0.0%' },
                    { mask: '0%' },
                    { mask: '0.00E+00' },
                    { mask: '0.00E+00' },
                    { mask: '0.00E+00' },
                    { mask: '(000) 000-0000' },
                    { mask: '00-00' },
                ],
                [0.125, 0.07, 1234, 0.00012, 9.999, 5551234567, 12345]
            )
        ).to.eql(['12.5%', '7%', '1.23E+03', '1.20E-04', '1.00E+01', '(555) 123-4567', '123-45']);
    });

    it('puts the minus before the symbols when the mask has no section for the negative numbers', () => {
        expect(shown([{ mask: '$ #,##0.00' }, { mask: '0.00' }, { mask: '0' }], [-1234.5, -0.001, -0.2])).to.eql(['-$ 1,234.50', '0.00', '0']);
    });

    it('shows a negative number, zero and a text with the sections of the mask', () => {
        const mask = '#,##0.00;[Red](#,##0.00);"zero";"text: "@';
        expect(shown([{ mask }, { mask }, { mask }, { mask }], [1234.5, -1234.5, 0, 'abc'])).to.eql(['1,234.50', '(1,234.50)', 'zero', 'text: abc']);
    });

    it('shows a text that is not a number as it is', () => {
        expect(shown([{ mask: '0.00' }, { mask: '0;-0' }, { mask: '0.00' }, { mask: '0.00' }], ['abc', 'x', '', '1.5'])).to.eql(['abc', 'x', '', '1.50']);
    });

    it('marks the cells of a section with the color [Red] red', () => {
        const worksheet = gridwell(root, {
            worksheets: [{ data: [[5, -5, 'a']], columns: [{ mask: '[Red]0;0' }, { mask: '0;[Red]0' }, { mask: '0;[Red]0' }] }],
        })[0];

        expect(worksheet.records[0].map((record) => record.element.classList.contains('red'))).to.eql([true, true, false]);
        expect(worksheet.getCell('B1').textContent).to.equal('5');

        worksheet.setValue('A1', -5);

        expect(worksheet.getCell('A1').classList.contains('red')).to.equal(false);
        expect(worksheet.getCell('A1').textContent).to.equal('5');
    });

    it('shows a number with the locale and the options of the column (issue #136)', () => {
        expect(
            shown(
                [
                    { type: 'numeric', locale: 'en-US', options: { style: 'currency', currency: 'USD' } },
                    { type: 'numeric', locale: 'de-DE', options: { style: 'currency', currency: 'EUR' }, decimal: ',' },
                    { type: 'numeric', locale: 'en-US', options: { style: 'percent' } },
                    { type: 'numeric', locale: 'en-US' },
                ],
                [3, 1234.5, 0.25, 'abc']
            )
        ).to.eql(['$3.00', '1.234,50 €', '25%', 'abc']);
    });

    it('shows a number with a date mask as a date, as Excel counts the days', () => {
        expect(
            shown(
                [{ mask: 'YYYY-MM-DD' }, { mask: 'DD/MM/YYYY HH24:MI' }, { mask: 'DD/MM/YYYY' }, { mask: 'DD/MM/YYYY' }, { mask: 'DD/MM/YYYY' }],
                [45000, 45000.75, '2024-03-05', 'soon', 3]
            )
        ).to.eql(['2023-03-15', '15/03/2023 18:00', '05/03/2024', 'soon', '03/01/1900']);
    });
});

describe('The editor of a cell with a mask', () => {
    it('shows the number without the symbols of the mask, and keeps a number typed with them as a number', () => {
        const worksheet = gridwell(root, {
            worksheets: [
                {
                    data: [[1234.5, 1234.5, 1234.5, -5]],
                    columns: [{ type: 'numeric', mask: '$ #,##0.00' }, { mask: '#.##0,00' }, { type: 'numeric', locale: 'de-DE' }, { mask: '0;(0)' }],
                },
            ],
        })[0];

        expect(edit(worksheet, 'A1', '$ 2,000.10')).to.equal('1234.5');
        expect(worksheet.getValue('A1')).to.equal(2000.1);
        expect(edit(worksheet, 'B1', '1.234.567,5')).to.equal('1234,5');
        expect(worksheet.getValue('B1')).to.equal(1234567.5);
        expect(edit(worksheet, 'C1', '-1,5')).to.equal('1234,5');
        expect(worksheet.getValue('C1')).to.equal(-1.5);
        expect(edit(worksheet, 'D1', '(7)')).to.equal('-5');
        expect(worksheet.getValue('D1')).to.equal(-7);
    });

    it('shows a percentage with its percent sign, and reads a number typed in it as a percentage', () => {
        const worksheet = gridwell(root, {
            worksheets: [
                {
                    data: [[0.125, 0.07, 0.5]],
                    columns: [{ mask: '0.0%' }, { mask: '0%' }, { type: 'numeric', locale: 'en-US', options: { style: 'percent' } }],
                },
            ],
        })[0];

        expect(edit(worksheet, 'A1', '50')).to.equal('12.5%');
        expect(worksheet.getValue('A1')).to.equal(0.5);
        expect(edit(worksheet, 'B1', '7.5%')).to.equal('7%');
        expect(worksheet.getValue('B1')).to.equal(0.075);
        expect(edit(worksheet, 'C1', '20%')).to.equal('50%');
        expect(worksheet.getValue('C1')).to.equal(0.2);
    });

    it('keeps a text that is not a number, a formula and an empty text as they are typed', () => {
        const worksheet = gridwell(root, {
            worksheets: [
                {
                    data: [[1, 2, 3, 4]],
                    columns: [{ mask: '0.00' }, { mask: '0.00' }, { type: 'numeric', mask: '0.00', allowEmpty: true }, { type: 'numeric', mask: '0.00' }],
                },
            ],
        })[0];

        edit(worksheet, 'A1', 'abc');
        expect(worksheet.getValue('A1')).to.equal('abc');
        expect(worksheet.getCell('A1').textContent).to.equal('abc');
        expect(edit(worksheet, 'A1', '=1+1')).to.equal('abc');
        expect(worksheet.getValue('A1')).to.equal('=1+1');
        // A formula opens as it is written
        expect(edit(worksheet, 'A1')).to.equal('=1+1');
        edit(worksheet, 'B1', '');
        expect(worksheet.getValue('B1')).to.equal('');
        edit(worksheet, 'C1', '');
        expect(worksheet.getValue('C1')).to.equal('');
        edit(worksheet, 'D1', '');
        expect(worksheet.getValue('D1')).to.equal(0);
    });

    it('shows the date of a date mask, and reads a date typed in its format as the number of days', () => {
        const worksheet = gridwell(root, {
            worksheets: [{ data: [[45000, 'later']], columns: [{ mask: 'DD/MM/YYYY' }, { mask: 'DD/MM/YYYY' }] }],
        })[0];

        expect(edit(worksheet, 'A1', '05/03/2024')).to.equal('15/03/2023');
        expect(worksheet.getValue('A1')).to.equal(45356);
        expect(worksheet.getCell('A1').textContent).to.equal('05/03/2024');
        expect(edit(worksheet, 'B1', '31/02/2024')).to.equal('later');
        expect(worksheet.getValue('B1')).to.equal('31/02/2024');
    });
});

describe('Date formats', () => {
    afterEach(() => {
        gridwell.setDictionary({ March: 'March', Tuesday: 'Tuesday' });
    });

    it('shows the date of a calendar cell in the format of its column', () => {
        const formats = ['DD/MM/YYYY', 'DD MMM YY', 'DDD, DD MMMM YYYY', 'MON D', 'DY WD', 'HH12:MI AM/PM', 'HH24:MI:SS', 'MM/DD/YYYY HH:MI', 'DAY MONTH'];
        const data = [
            '2024-03-05 00:00:00',
            '2024-03-05',
            '2024-03-05',
            '2024-03-05',
            '2024-03-05',
            '2024-03-05 00:05:00',
            '2024-03-05 07:08:09',
            '2024-03-05 13:30:00',
            '2024-03-05',
        ];
        expect(
            shown(
                formats.map((format) => ({ type: 'calendar', options: { format } })),
                data
            )
        ).to.eql(['05/03/2024', '05 Mar 24', 'Tue, 05 March 2024', 'MAR 5', 'Tue Tuesday', '12:05 AM', '07:08:09', '03/05/2024 13:30', 'TUESDAY MARCH']);
    });

    it('reads the date of a calendar cell in the format of its column, and shows a text that is not a date as it is', () => {
        expect(
            shown(
                [
                    { type: 'calendar', options: { format: 'DD/MM/YYYY' } },
                    { type: 'calendar', options: { format: 'DD MMMM YYYY' } },
                    { type: 'calendar', options: { format: 'MM/DD/YY HH12:MI AM/PM' } },
                    { type: 'calendar', options: { format: 'DD/MM/YYYY' } },
                    { type: 'calendar' },
                    { type: 'calendar', options: { format: 'DD/MM/YYYY' } },
                    { type: 'calendar', options: { format: 'DY, D MON YYYY HH24:MI:SS' } },
                    { type: 'calendar', options: { format: 'DDDD M/D/YY H:MI' } },
                ],
                ['05/03/2024', '5 march 2024', '03/05/99 01:30 PM', 'tomorrow', '2024-03-05', '', 'tue, 5 mar 2024 07:08:09', 'Tuesday 3/5/24 7:08']
            )
        ).to.eql(['05/03/2024', '05 March 2024', '03/05/99 01:30 PM', 'tomorrow', '2024-03-05', '', 'Tue, 5 MAR 2024 07:08:09', 'Tuesday 3/5/24 7:08']);
    });

    it('shows and reads the names of the months and the days in the language of the dictionary', () => {
        gridwell.setDictionary({ March: 'März', Tuesday: 'Dienstag' });

        const worksheet = gridwell(root, {
            worksheets: [
                {
                    data: [['2024-03-05', '']],
                    columns: [
                        { type: 'calendar', options: { format: 'DDDD, DD MMMM YYYY' } },
                        { type: 'calendar', options: { format: 'DD MMMM YYYY' } },
                    ],
                },
            ],
        })[0];

        expect(worksheet.getCell('A1').textContent).to.equal('Dienstag, 05 März 2024');

        worksheet.paste(1, 0, '05 März 2024');

        expect(worksheet.getValue('B1')).to.equal('2024-03-05 00:00:00');
    });

    it('shows and reads the short names of the months and the days with their own translations (issue #155)', () => {
        const translations = {
            May: 'Tháng Năm',
            'May (short)': 'Thg 5',
            June: 'Tháng Sáu',
            Jun: 'Thg 6',
            July: 'Tháng Bảy',
            Jul: 'Thg 7',
            Dec: 'Thg 12',
            Saturday: 'Thứ Bảy',
            Sat: 'T7',
            Mon: 'T2',
        };
        gridwell.setDictionary(translations);
        try {
            const worksheet = gridwell(root, {
                worksheets: [
                    {
                        data: [['2024-06-01', '2024-07-01', '2024-05-04', '2024-05-04', '', '', '']],
                        columns: Array.from({ length: 7 }, (v, i) => ({
                            type: 'calendar' as const,
                            options: { format: i === 3 ? 'DDDD D MMMM YYYY' : 'DDD, DD MMM YYYY' },
                        })),
                    },
                ],
            })[0];

            // The months have short names of their own, and the short name of May is not its full name
            expect(worksheet.records[0].slice(0, 4).map((record) => record.element.textContent)).to.eql([
                'T7, 01 Thg 6 2024',
                'T2, 01 Thg 7 2024',
                'T7, 04 Thg 5 2024',
                'Thứ Bảy 4 Tháng Năm 2024',
            ]);

            worksheet.paste(4, 0, 'T2, 01 Thg 7 2024\tt7, 04 THG 5 2024\tT2, 02 Thg 12 2024');
            expect(worksheet.getValueFromCoords(4, 0)).to.equal('2024-07-01 00:00:00');
            expect(worksheet.getValueFromCoords(5, 0)).to.equal('2024-05-04 00:00:00');
            expect(worksheet.getValueFromCoords(6, 0)).to.equal('2024-12-02 00:00:00');
        } finally {
            gridwell.setDictionary(Object.fromEntries(Object.keys(translations).map((text) => [text, ''])));
        }
    });

    it('keeps a pasted date in the ISO format, reads a date in the format of the column and empties a text that is not a date', () => {
        const worksheet = gridwell(root, {
            worksheets: [
                { data: [['', '', '', '']], columns: Array.from({ length: 4 }, () => ({ type: 'calendar' as const, options: { format: 'DD/MM/YYYY' } })) },
            ],
        })[0];

        worksheet.paste(0, 0, '2024-03-05\t05/03/2024\t31/02/2024\tsoon');

        expect(worksheet.getValueFromCoords(0, 0)).to.equal('2024-03-05');
        expect(worksheet.getValueFromCoords(1, 0)).to.equal('2024-03-05 00:00:00');
        expect(worksheet.getValueFromCoords(2, 0)).to.equal('');
        expect(worksheet.getValueFromCoords(3, 0)).to.equal('');
        expect(worksheet.records[0].map((record) => record.element.textContent)).to.eql(['05/03/2024', '05/03/2024', '', '']);
    });
});
