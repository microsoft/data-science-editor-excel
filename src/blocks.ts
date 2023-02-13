import { getDataForTable } from "./xl/data"
import {
    insertTableFromDataset,
    DATA_SCIENCE_EXPORT_TABLE,
} from "./xl/tables"
import { insertBarChart, clearAllChartsFromActiveSheet } from "./xl/charts"

const colour = "#107C41"

export type Block = any
export type DataTable = any

export let currentWorkspace
var savedInput = new Map()

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
    {
        kind: "block",
        type: "excel_chart_bar",
        tooltip: "Renders the block data in a bar chart",
        message0: "bar chart of index %1 value %2 %3 %4",
        args0: [
            {
                type: "jacdac_field_data_column_chooser",
                name: "index",
            },
            {
                type: "jacdac_field_data_column_chooser",
                name: "value",
                dataType: "number",
            },
            {
                type: "jacdac_field_json_settings",
                name: "settings",
                schema: {
                    type: "object",
                    properties: {
                        title: {
                            type: "string",
                            title: "Chart title",
                        },
                        encoding: {
                            index: {
                                title: "Index",
                                type: "object",
                                properties: {
                                    axis: {
                                        type: "object",
                                        properties: {
                                            title: {
                                                type: "string",
                                                title: "Title",
                                            },
                                        },
                                    },
                                },
                            },
                            value: {
                                title: "Value",
                                type: "object",
                                properties: {
                                    scale: {
                                        type: "object",
                                        properties: {
                                            domainMin: {
                                                type: "number",
                                                title: "minimum",
                                                description:
                                                    "Sets the minimum value in the scale domain",
                                            },
                                            domainMax: {
                                                type: "number",
                                                title: "maximum",
                                                description:
                                                    "Sets the maximum value in the scale domain",
                                            },
                                        },
                                    },
                                    axis: {
                                        type: "object",
                                        properties: {
                                            title: {
                                                type: "string",
                                                title: "Title",
                                            },
                                        },
                                    },
                                },
                            },
                        },
                    },
                },
            },
            {
                type: "input_dummy",
            },
        ],
        previousStatement: "DataScienceStatement",
        nextStatement: "DataScienceStatement",
        colour,
        template: "meta",
        inputsInline: false,
        dataPreviewField: false,
        dsl: "chart",
    },
]

export const category = [
    {
        kind: "category",
        name: "Tables",
        colour,
        contents: blocks.map((block) => ({ kind: "block", type: block.type })),
        order: 100,
    },
]

export const transforms = {
    excel_import_table: async (b) => {
        const tableName = b.inputs[0].fields["table name"].value
        if (!tableName) {
            console.log(`table.load no table selected`)
            return { dataset: [] }
        }

        const dataset = await getDataForTable(tableName)
        console.log(`table.load`, { tableName, dataset })
        if (!dataset) return { warning: "table not found", dataset: [] }
        return { dataset }
    },
    excel_chart_bar: async (b, dataset) => {
        const index = b.inputs[0].fields["index"].value
        const value = b.inputs[0].fields["value"].value

        if (index === undefined || value === undefined) {
            // Waiting for the user to select...
            return
        }

        // Clear the charts before we delete the existing chart
        // data source sheet to prevent chart errors.
        await clearAllChartsFromActiveSheet()
        await insertTableFromDataset(dataset)
        insertBarChart(DATA_SCIENCE_EXPORT_TABLE, index, value)
    },
}

export const setCurrentWorkspace = (workspace) => {
    currentWorkspace = workspace
}
