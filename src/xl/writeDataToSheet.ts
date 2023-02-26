async function getOrCreateWorksheet(
    context: Excel.RequestContext,
    name: string
): Promise<Excel.Worksheet> {
    const workbook = context.workbook;
    const worksheets = workbook.worksheets;
    const sheet = worksheets.getItemOrNullObject(name);
    await context.sync();
    if (sheet.isNullObject) {
        const sheet = worksheets.add(name);
        return sheet;
    }
    return sheet;
}

/**
 * Tables are stored in a worksheet of the same name
 * @param sheetName
 */
export async function writeDataToSheet(
    sheetName: string,
    headers: string[],
    rows: (string | number)[][]
) {
    return Excel.run(async (context) => {
        //const worksheetWithTableName = `table ${sheetName}`;

        const sheet = await getOrCreateWorksheet(context, sheetName);
        sheet.activate();

        // clean out tables
        sheet.tables.load("items");
        await context.sync();
        sheet.tables.items.forEach((table) => table.delete());
        await context.sync();

        // clean out any older data
        sheet.getRange().clear();

        // write headers
        const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
        headerRange.values = [headers];
        await context.sync();

        // write rows
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const range = sheet.getRangeByIndexes(i + 1, 0, 1, row.length);
            range.values = [row];

            // sync every 50 to avoid issues where there is too much data
            if (i % 50 === 0) {
                await context.sync();
            }
        }

        // Create the new table
        const tableRange = sheet.getRangeByIndexes(0, 0, rows.length + 1, headers.length);
        const table = sheet.tables.add(tableRange, true);

        // have the table name match the sheet name
        table.name = `${sheetName}`;
    });
}
