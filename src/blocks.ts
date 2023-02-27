import { writeDataToSheet } from "./xl/writeDataToSheet";
import { getDataForTable } from "./xl/data";
import { DATA_SCIENCE_STATEMENT_TYPE } from "./constants";

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
        nextStatement: DATA_SCIENCE_STATEMENT_TYPE,
        dataPreviewField: true,
        template: "meta",
    },

    // Export a table to a sheet.
    {
        kind: "block",
        type: "excelExportDataToSheet",
        message0: "export data to sheet %1",
        tooltip: "look a tooltip",
        colour,
        args0: [
            {
                // there are some restrictions to this field
                // must be a valid table or sheet name
                // to simplify - no spaces only [a-z][A_Z] characters
                type: "field_input",
                name: argumentNameExportDataToSheetName,
                spellcheck: false,
                //dataId: "table",
            },
        ],
        previousStatement: DATA_SCIENCE_STATEMENT_TYPE,
        nextStatement: DATA_SCIENCE_STATEMENT_TYPE,
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

    excelExportDataToSheet: async (b, data: { [key: string]: string | number }[]) => {
        console.log("excel_export_data");
        console.log(b);
        console.log(data);
         // block id stays fixed as long as the block exists
         // this allows tracking of the specific block across calls
         // track all block ids and associate a worksheetId with each of them
        const blockId = b.id;
        debugger;
        const sheetName = b.inputs[0].fields[argumentNameExportDataToSheetName].value;
        if (!sheetName) {
            console.debug(`table.write no sheet selected`);
            return { dataset: [] };
        }

        // Data is an array of objects that have properties
        // only consider the first object properties
        const headers: string[] = [];
        if (data.length > 0) {
            headers.push(...Object.getOwnPropertyNames(data[0]));
        }

        const rows = data.map((o) => {
            const row = headers.map((name) => o[name]);
            return row;
        });

        // when typing each character alters the sheet name.
        // would need to constantly rename sheet?
        writeDataToSheet(sheetName, headers, rows);
        return data;
    },
};

export const setCurrentWorkspace = (workspace) => {
    currentWorkspace = workspace;
};
