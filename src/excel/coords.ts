import { strFindIfNot } from "../utils";
import { addAscii, isLetter, toAscii } from "../charType";

/**
 * Regex to check if a string represents a valid cell coordinate.
 */
export const EXCEL_COORD_REGEX = /^[A-Z]+[0-9]+$/i;

/**
 * Checks if a string represents valid coordinates.
 * @param str string to parse
 * @returns true if coordinates are valid, false otherwise
 */
export function areCoordsValid(str : string) : boolean {
    return EXCEL_COORD_REGEX.test(str);
}

/**
 * Check if a string starts with cell coordinates.
 * @param str string to parse
 * @returns [false, 0] if the string doesn't start with coordinates, [true, <number of chars concerned>] otherwise
 */
export function beginsWithValidCoords(str : string) : [boolean, number] {
    const matches = /^[A-Z]+[0-9]+/i.exec(str);

    if (matches == null) {
        return [false, 0];
    }
    return [true, matches[0].length]
}

/**
 * Converts a column to its index.
 * @param columnStr the column (A, B, C...), either in lowercase or uppercase
 * @returns the index (0, 1, 2...)
 * @throws Error if the column is empty or contains non-letter characters
 * @link https://stackoverflow.com/questions/21229180/convert-column-index-into-corresponding-column-letter
 */
export function columnToIndex(columnStr : string) : number {
    if (columnStr.length == 0) {
        throw new Error("Column mustn't be empty !");
    } else if (strFindIfNot(columnStr, isLetter) != -1) {
        throw new Error("Column must be only made of letters !");
    }
    columnStr = columnStr.toUpperCase();

    const length : number = columnStr.length;
    const asciiCodeRightBeforeCapitalA : number = toAscii('A') - 1;
    let index : number = 0;
    let pow : number = Math.pow(26, length - 1);

    for (let i = 0; i < length; i++) {
        index += (toAscii(columnStr[i]) - asciiCodeRightBeforeCapitalA) * pow;
        pow /= 26;
    }
    return index - 1;
}

/**
 * Converts an index to its column.
 * @param columnIndex the index (0, 1, 2...)
 * @returns the column (A, B, C...)
 * @throws Error if index is strictly negative
 * @link https://stackoverflow.com/questions/21229180/convert-column-index-into-corresponding-column-letter
 */
export function indexToColumn(columnIndex : number) : string {
    if (columnIndex < 0) {
        throw new Error(`Column index must be greater or equal than 0, got ${columnIndex} !`);
    }

    let columnStr : string = "";
    columnIndex++;

    while (columnIndex > 0) {
        const digit : number = (columnIndex - 1) % 26;
        columnStr = addAscii('A', digit) + columnStr;
        columnIndex = (columnIndex - digit - 1) / 26;
    }
    return columnStr;
}

/**
 * Converts coordinates to their indices.
 * @param coords cell coordinates
 * @returns [<row index>, <column index>]
 * @throws Error if coords aren't valid
 */
export function coordsToIndices(coords : string) : [number, number] {
    if (!areCoordsValid(coords)) {
        throw new Error(`"${coords}" isn't a proper Excel coordinate !`);
    }
    const firstDigitIndex : number = coords.search(/[0-9]/);
    coords = coords.toUpperCase();

    return [
        Number(coords.substring(firstDigitIndex)) - 1,      // row
        columnToIndex(coords.substring(0, firstDigitIndex)) // column
    ];
}

/**
 * Converts indices to their corresponfing coordinates.
 * @param rowIndex index of the row
 * @param columnIndex index of the column, must be greater or equal to 0
 * @returns the coord corresponding to the indices
 * @throws Error if either rowIndex or columnIndex is strictly negative
 */
export function indicesToCoord(rowIndex : number, columnIndex : number) : string {
    if (rowIndex < 0) {
        throw new Error(`Row index must be greater or equal than 0, got ${rowIndex} !`);
    } else if (columnIndex < 0) {
        throw new Error(`Column index must be greater or equal than 0, got ${columnIndex} !`);
    }
    return indexToColumn(columnIndex) + (rowIndex + 1).toString();
}