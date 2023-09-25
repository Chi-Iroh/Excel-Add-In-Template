import { CellLocation } from "../cell/CellLocation";
import { EXCEL_ROWS_MAX, EXCEL_COLUMNS_MAX } from "./limits";

/**
 * Create an Excel.Range object containing all the cells of a worksheet
 * @param worksheet current Excel worksheet
 * @returns the range
 */
export function getWholeWorksheetRange(worksheet : Excel.Worksheet) : Excel.Range {
    const worksheetRange = worksheet.getRange(`A1:${CellLocation.fromIndices(EXCEL_COLUMNS_MAX, EXCEL_ROWS_MAX).stringifyInstanceCoords()}`);
    worksheetRange.load("address");
    return worksheetRange;
}