import { CellLocation, Direction } from "./CellLocation";

/**
 *  A class to manipulate a single cell of an Excel worksheet.
 * @note Excel API only provides a class to manipulate ranges of cells.
 */
export class Cell {
    private location : CellLocation;
    private range : Excel.Range;
    private worksheet : Excel.Worksheet;

    /**
     * @param worksheet worksheet to edit
     * @param location location of the cell, either string like "A1" or a CellLocation object
     */
    constructor(worksheet : Excel.Worksheet, location : string | CellLocation) {
        this.worksheet = worksheet;
        this.location = new CellLocation(location);
        const locationString : string = this.location.stringifyInstanceCoords();
        this.range = worksheet.getRange(locationString);
        this.range.load("address");
        if (!this.range.hasOwnProperty("values")) { // values property doesn't exist if range is empty
            this.range.values = [[ "" ]]
        }
    }

    /**
     * Writes a value in the cell
     * @param formulaOrValue formula or value to set
     */
    public setValue(formulaOrValue : string | number) : void {
        this.range.values[0][0] = formulaOrValue;
    }

    /**
     * Computes the value of the cell.m
     * @returns Computes the formula and returns its result as a number.
     * @returns undefined if empty cell or contains text not starting with '='
     */
    public computeValue() : number | undefined {
        const valueBeforeBeingCalulated : string | number = this.range.values[0][0];
        if (String(this.range.values[0][0]) == "") {
            return undefined;
        } else if (typeof valueBeforeBeingCalulated === "string" && valueBeforeBeingCalulated[0] != '=') { // String but not a formula
            return undefined;
        }

        this.range.calculate();
        return Number(this.range.values[0][0]);
    }

    /**
     * Reads the value of the formula.
     * @returns cell content as string
     * @returns undefined id the cell is empty
     * @note This function doesn't evaluate the formula, contrary to Cell.computeValue.
     */
    public getValueAsString() : string | undefined {
        if (String(this.range.values[0][0]) == "") {
            return undefined;
        }
        return String(this.range.values[0][0]);
    }

    /**
     * Creates a Cell instance from a distance and direction relative to this instance.
     * @param distance distance from this instance, can be 0 (useless) or negative (better change the direction)
     * @param direction either up, right, down or left
     * @returns the new Cell instance
     */
    public nextCell(distance : number, direction : Direction) : Cell {
        return new Cell(this.worksheet, this.location.nextCellLocation(distance, direction));
    }

    /**
     *  See Excel.Range.untrack()
     */
    public untrack() : void {
        this.range.untrack();
    }

    /**
     *  Gets the location of the cell
     * @returns string containing location, as "A1"
     */
    public getLocation() : string {
        return this.location.stringifyInstanceCoords();
    }

    /**
     *  Updates Excel worksheet global array
     * @note For testing purposes only, won't do anything in production
     */
    public updateAllValues() : void {
        if ((this.range as any)["updateAllValues"] != undefined) {
            (this.range as any).updateAllValues();
        }
    }
};