import { writeDataToSheet } from "./xl/writeDataToSheet";
import { getDataForTable } from "./xl/data";

const colour = "#107C41";

// export type Block = any;
// export type DataTable = any;

// Need Typings for this protocol

export let currentWorkspace: unknown;

const argumentNameExportDataToSheetName = "name";

export const blocks = [
    {
        kind: "block",
        type: "excel_import_table",
        message0: "table %1",
        colour,
        args0: [
            {
                type: "ds_field_iframe_data_chooser",
                name: "table name",
                dataId: "table",
            },
        ],
        nextStatement: "DataScienceStatement",
        dataPreviewField: true,
        template: "meta",
    },

    // Export a table to a sheet.
    {
        kind: "block",
        type: "excelExportDataToSheet",
        message0: "export data to sheet %1",
        colour,
        args0: [
            {
                type: "string",
                name: argumentNameExportDataToSheetName,
                dataId: "table",
            },
        ],
        previousStatement: "DataScienceStatement",
        nextStatement: "DataScienceStatement",
        dataPreviewField: false,
        template: "meta",
    },
];

export const category = [
    {
        kind: "category",
        name: "Tables",
        colour,
        contents: blocks.map((block) => ({ kind: "block", type: block.type })),
        order: 100,
    },
];

export const transforms = {
    // don't rename these identifiers, they are used in the serialized blocky and will break existing files
    // eslint-disable-next-line @typescript-eslint/naming-convention
    excel_import_table: async (b) => {
        const tableName = b.inputs[0].fields["table name"].value;
        if (!tableName) {
            console.debug(`table.load no table selected`);
            return { dataset: [] };
        }

        const dataset = await getDataForTable(tableName);
        console.debug(`table.load`, { tableName, dataset });
        if (!dataset) return { warning: "table not found", dataset: [] };
        return { dataset };
    },

    excelExportData: async (b) => {
        console.log("excel_export_data");
        const sheetName = b.inputs[0].fields[argumentNameExportDataToSheetName].value;
        if (!sheetName) {
            console.debug(`table.write no sheet selected`);
            return { dataset: [] };
        }

        writeDataToSheet(sheetName, ["test"]);
    },
};

export const setCurrentWorkspace = (workspace) => {
    currentWorkspace = workspace;
};
