import { DATA_SCIENCE_EXPORT_SHEET } from "./tables"

export const clearAllChartsFromActiveSheet = async () => {
    Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getActiveWorksheet()
        await context.sync()

        // Clear the charts on the sheet
        sheet.charts.load("items")
        await context.sync()

        for (let i = 0; i < sheet.charts.items.length; i++) {
            sheet.charts.items[i].delete()
        }

        await context.sync()
    })
}

const clearAllSeriesFromChart = async (
    context: Excel.RequestContext,
    chart: Excel.Chart
) => {
    chart.series.load("items")
    await context.sync()

    for (let i = 0; i < chart.series.items.length; i++) {
        chart.series.items[i].delete()
    }

    await context.sync()
}

const getRangeForSelectedColumn = async (
    context: Excel.RequestContext,
    sheet: Excel.Worksheet,
    selectedTable: string,
    columnName: string
) => {
    let tableRange = sheet.tables.getItemOrNullObject(selectedTable).getRange()

    if (tableRange === undefined) {
        return
    }

    tableRange.load("columnCount, rowCount")
    await context.sync()

    for (let i = 0; i < tableRange.columnCount; i++) {
        let currentCol = tableRange.getColumn(i)
        currentCol.load("values, rowIndex, columnIndex")
        await context.sync()

        let rowIndex = currentCol.rowIndex
        let columnIndex = currentCol.columnIndex

        // Find the selected column ranges without the column titles
        if (currentCol.values[0][0] == columnName) {
            return sheet.getRangeByIndexes(
                rowIndex + 1,
                columnIndex,
                tableRange.rowCount - 1,
                1
            )
        }
    }

    return null
}

export const insertBarChart = async (
    selectedTable: string,
    selectedIndex: string,
    selectedValue: string
) => {
    Excel.run(async (context) => {
        if (
            selectedTable === undefined ||
            selectedIndex === undefined ||
            selectedValue === undefined
        ) {
            // Not enough information, so let's bail
            return
        }

        let exportSheet = context.workbook.worksheets.getItemOrNullObject(
            DATA_SCIENCE_EXPORT_SHEET
        )
        await context.sync()

        let sheet = context.workbook.worksheets.getActiveWorksheet()
        await context.sync()

        if (sheet.isNullObject) {
            // Up-to-date DSE table should have been exported
            return
        }

        // Create an empty chart
        let emptyRange = sheet.getRange("XFD1048576") // TODO: Actually find a blank cell
        let chart = sheet.charts.add("ColumnClustered", emptyRange, "Auto")

        await clearAllSeriesFromChart(context, chart)

        // Find the ranges for the series and category names
        let seriesRange = await getRangeForSelectedColumn(
            context,
            exportSheet,
            selectedTable,
            selectedValue
        )
        let categoryNamesRange = await getRangeForSelectedColumn(
            context,
            exportSheet,
            selectedTable,
            selectedIndex
        )

        // Fill in the series and category names
        chart.series.add(selectedValue).setValues(seriesRange)
        chart.axes.categoryAxis.title.text = selectedIndex
        chart.axes.categoryAxis.setCategoryNames(categoryNamesRange)
        chart.legend.position = "Bottom"
        chart.legend.format.fill.setSolidColor("white")

        await context.sync()
    })
}
