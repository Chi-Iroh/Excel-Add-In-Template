import { EXCEL_COLUMNS_MAX, EXCEL_ROWS_MAX } from "./excel/limits";
import { CellLocation } from "./cell/CellLocation";

export function handleError(error : Error) {
    document.getElementById("error")!.innerText = `${error.stack}\n${error.name} : ${error.message}`;
}

export function resetError() : void {
    document.getElementById("error")!.innerText = "";
}

export function stringIndexFindIf(str : string, predicate : (char : string) => boolean) : number {
    const charArray : string[] = str.split("");
    return charArray.findIndex(predicate);
}

export function stringFindIf(str : string, predicate : (char : string) => boolean) : string | undefined {
    const charArray : string[] = str.split("");
    return charArray.find(predicate);
}