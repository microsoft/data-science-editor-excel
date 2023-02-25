async function getOrCreateWorksheet(
    context: Excel.RequestContext,
    name: string
): Promise<Excel.Worksheet> {
    const workbook = context.workbook;
    const worksheets = workbook.worksheets;
    let sheet = worksheets.getItemOrNullObject(name);
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
    headers: string[]
    //data: (string|number)[][]
) {
    return Excel.run(async (context) => {
        //const worksheetWithTableName = `table ${sheetName}`;

        const sheet = await getOrCreateWorksheet(context, sheetName);

        const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
        headerRange.values = [headers];
        await context.sync();

        // just write headers for now.

        // const workbook = context.workbook;
        // const tables = workbook.tables;

        // const table = tables.getItemOrNullObject(tableName);
        // await context.sync();

        // // table does not exist make it.
        // if (table.isNullObject) {
        //     // table does not exist, create the table
        //     const worksheets = workbook.worksheets;
        //     worksheets.load("items/name");
        //     await context.sync();

        //     const names = worksheets.items.map(x => x.name);
        //     if (names.includes(worksheetWithTableName)) {
        //         throw new Error("table sheet already exists")
        //     }

        //     const sheet = worksheets.add(worksheetWithTableName);
        //     const table = sheet.tables.add("A1", false);
        //     table.
        //     // sheet.tables.add()
        //     // const sheet = worksheets.getItemOrNullObject(worksheetWithTableName);
        //     // sheet.

        // }
    });
}
