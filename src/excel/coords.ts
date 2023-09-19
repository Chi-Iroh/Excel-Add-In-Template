import { addAscii, toAscii } from "../charType";

export const EXCEL_COORD_REGEX = /^[A-Z]+[0-9]+$/i;

export function isCoordValid(coord : string) : boolean {
    return EXCEL_COORD_REGEX.test(coord);
}

export function beginsWithValidCoord(str : string) : [boolean, number] {
    const matches = /^[A-Z]+[0-9]+/i.exec(str);

    if (matches == null) {
        return [false, 0];
    }
    return [true, matches[0].length]
}

export function columnToIndex(column : string) : number {
    const length : number = column.length;
    const asciiCodeRightBeforeCapitalA : number = toAscii('A') - 1;
    let index : number = 0;
    let pow : number = Math.pow(26, length - 1);

    for (let i = 0; i < length; i++) {
        index += (toAscii(column[i]) - asciiCodeRightBeforeCapitalA) * pow;
        pow /= 26;
    }
    return index - 1;
}

// https://stackoverflow.com/questions/21229180/convert-column-index-into-corresponding-column-letter

export function indexToColumn(index : number) : string {
    if (index < 0) {
        throw new Error("Index mustn't be negative !");
    }

    let column : string = "";

    index++;
    while (index > 0) {
        const digit : number = (index - 1) % 26;
        column = addAscii('A', digit) + column;
        index = (index - digit - 1) / 26;
    }
    return column;
}

export function coordToIndices(cellLocation : string) : [number, number] {
    if (!isCoordValid(cellLocation)) {
        throw new Error(`"${cellLocation}" isn't a proper Excel coordinate !\nEx: A1, B12, AD42`);
    }
    const firstDigitIndex : number = cellLocation.search(/[0-9]/);
    cellLocation = cellLocation.toUpperCase();

    return [
        Number(cellLocation.substring(firstDigitIndex)) - 1,
        columnToIndex(cellLocation.substring(0, firstDigitIndex))
    ];
}

export function indicesToCoord(columnIndex : number, rowIndex : number) : string {
    return indexToColumn(columnIndex) + (rowIndex + 1).toString();
}