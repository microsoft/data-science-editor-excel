import { getDataForTable } from "./xl/data";

const colour = "#107C41";

// export type Block = any;
// export type DataTable = any;

export let currentWorkspace: unknown;

export const blocks = [
    {
        kind: "block",
        type: "excelImportTable",
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
    excelImportTable: async (b) => {
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
};

export const setCurrentWorkspace = (workspace) => {
    currentWorkspace = workspace;
};
