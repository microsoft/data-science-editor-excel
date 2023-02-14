import { getDataForTable } from "./xl/data";

const colour = "#107C41";

export type Block = any;
export type DataTable = any;

export let currentWorkspace;

export const blocks = [
    {
        kind: "block",
        type: "excel_import_table",
        message0: "table %1",
        colour,
        args0: [
            {
                type: "jacdac_field_iframe_data_chooser",
                name: "table name",
                dataId: "table",
            },
        ],
        nextStatement: "DataScienceStatement",
        dataPreviewField: true,
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
    excel_import_table: async (b) => {
        const tableName = b.inputs[0].fields["table name"].value;
        if (!tableName) {
            console.log(`table.load no table selected`);
            return { dataset: [] };
        }

        const dataset = await getDataForTable(tableName);
        console.log(`table.load`, { tableName, dataset });
        if (!dataset) return { warning: "table not found", dataset: [] };
        return { dataset };
    },
};

export const setCurrentWorkspace = (workspace) => {
    currentWorkspace = workspace;
};
