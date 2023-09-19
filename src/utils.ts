import { EXCEL_COLUMNS_MAX, EXCEL_ROWS_MAX } from "./excel/limits";
import { CellLocation } from "./cell/CellLocation";

export function handleError(error : Error) {
    document.getElementById("error")!.innerText = `${error.stack}\n${error.name} : ${error.message}`;
}

export function resetError() : void {
    document.getElementById("error")!.innerText = "";
}