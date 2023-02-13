export type CellData = { [header: string]: string | number | boolean }[];

export async function getSelectedData(): Promise<CellData> {
    return await Excel.run(async context => {
        const range = context.workbook.getSelectedRange();

        return await getDataForRange(context, range);
    });
}

export async function getDataForAddress(address: string): Promise<CellData> {
    return await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);

        return await getDataForRange(context, range);
    });
}

export async function getDataForTable(tableName: string): Promise<CellData> {
    return await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const table = sheet.tables.getItemOrNullObject(tableName);

        try {
            await context.sync();

            if (!table.isNullObject) {
                const r = await getDataForRange(context, table.getRange());
                return r
            }
        } catch (e) {
            console.error(e);
        }

        return [];
    });
}

async function getDataForRange(context: Excel.RequestContext, range: Excel.Range): Promise<CellData> {
    range.load("values");

    await context.sync();

    const values = range.values as (string | number | boolean)[][];
    const headers = values[0].map(value => value.toString());

    // Remove the header row, filter out empty rows, and map each cell in each
    // row to its header.
    const r = values
        .slice(1)
        .filter(cells => cells.some(cell => cell !== ""))
        .map(cells => {
            const row = {};
            headers.forEach((header, index) => {
                row[header] = cells[index];
            });

            return row;
        });
    console.log(`data.range`, { r, range })
    return r
}
