import { getDataForTable } from "./xl/data";

const colour = "#107C41";

export interface Block {
    kind: "block";
    type: string;
    message0: string;
    colour: string;
    args0: unknown[];
    nextStatement: string;
    dataPreviewField: boolean;
    template: string;
}
// export type DataTable = any;

export let currentWorkspace: unknown;

/**
 * Definition of a block
 */
export const blocks: Block[] = [
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
];

export const category = [
    {
        kind: "category",
        name: "Tables",
        colour,
        contents: blocks.map(({ kind, type }) => ({ kind, type })),
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
};

export const setCurrentWorkspace = (workspace) => {
    currentWorkspace = workspace;
};
