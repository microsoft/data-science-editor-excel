export const dataScienceExportSheet = "Data Science Editor Export";
export const dataScienceExportTable = "DataScienceEditorTable";

export async function createTableForAddress(name: string, address: string): Promise<void> {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const table = sheet.tables.add(address, true /* hasHeaders */);
        table.name = name;
        await context.sync();
    });
}

export async function createTableForSelection(name: string): Promise<void> {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const selection = context.workbook.getSelectedRange();
        const table = sheet.tables.add(selection, true /* hasHeaders */);
        table.name = name;
        await context.sync();
    });
}

export async function isExistingTable(
    context: Excel.RequestContext,
    name: string,
    sheet: Excel.Worksheet
): Promise<boolean> {
    const table = sheet.tables.getItemOrNullObject(name);
    await context.sync();

    return !table.isNullObject;
}

async function fillTableFromDataset(
    dataset,
    context: Excel.RequestContext,
    sheet: Excel.Worksheet
) {
    // TODO: Investigate why these calls lead to 'cannot read properties of undefined'
    // errors even though we can still grab their lengths.
    if (dataset.length === 0) {
        return;
    }

    let columnNamesCreated = false;
    const resultValues = [[]];

    for (const rowKey in dataset) {
        // Populate the column titles first
        if (!columnNamesCreated) {
            const columnNames = Object.keys(dataset[rowKey]);

            for (const key in columnNames) {
                resultValues[0].push(columnNames[key]);
            }

            columnNamesCreated = true;
        }

        resultValues.push([]);
        const columnValues = Object.values(dataset[rowKey]);

        for (const key in columnValues) {
            resultValues[Number(rowKey) + 1].push(String(columnValues[key]));
        }
    }

    const tableRange = sheet.getRangeByIndexes(
        0 /*startRow*/,
        0 /*startCol*/,
        resultValues.length,
        resultValues[0].length
    );
    tableRange.load("values");
    await context.sync();
    tableRange.values = resultValues;
    await context.sync();

    const table = sheet.tables.add(tableRange, true /*hasHeaders*/);
    table.load("name");
    await context.sync();
    table.name = dataScienceExportTable;
}

export async function insertTableFromDataset(dataset) {
    return await Excel.run(async (context) => {
        // Insert a new sheet if it doesn't exist
        let sheet = context.workbook.worksheets.getItemOrNullObject(dataScienceExportSheet);
        await context.sync();

        sheet.load("visibility");
        await context.sync();
        sheet.visibility = "Visible";

        if (!sheet.isNullObject) {
            sheet.delete();
            await context.sync();
        }

        sheet = context.workbook.worksheets.add(dataScienceExportSheet);
        await context.sync();

        sheet.load("visibility");
        await context.sync();
        sheet.visibility = "VeryHidden";

        await fillTableFromDataset(dataset, context, sheet);
        await context.sync();
    });
}

export async function getAllTables(context: Excel.RequestContext): Promise<Excel.Table[]> {
    const tables = context.workbook.tables;
    tables.load();
    await context.sync();
    return tables.items;
}
