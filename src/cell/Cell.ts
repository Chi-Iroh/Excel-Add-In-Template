import { CellLocation, Direction } from "./CellLocation";

export class Cell {
    private location : CellLocation;
    private range : Excel.Range;
    private worksheet : Excel.Worksheet;

    /**
     * Checks if the location represents a single cell.
     * @param location is the string representing the cell location
     * @returns true if the location string represents a single cell (A1, BD85 etc..), false if not (A2:B7, 42, hello etc..)
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

    public setValue(formulaOrValue : string | number) : void {
        this.range.values[0][0] = formulaOrValue;
    }

    public getValue() : number | undefined {
        this.range.calculate();
        if (String(this.range.values[0][0]) == "") {
            return undefined;
        }
        return Number(this.range.values[0][0]);
    }

    public getValueAsString() : string | undefined {
        if (String(this.range.values[0][0]) == "") {
            return undefined;
        }
        return String(this.range.values[0][0]);
    }

    public nextCell(distance : number, direction : Direction) : Cell | null {
        let nextCellLocation : CellLocation = this.location.nextCellLocation(distance, direction);

        if (nextCellLocation == null) {
            return null;
        }
        return new Cell(this.worksheet, nextCellLocation);
    }

    public untrack() : void {
        this.range.untrack();
    }

    public getLocation() : string {
        return this.location.stringifyInstanceCoords();
    }

    public updateAllValues() : void {
        if ((this.range as any)["updateAllValues"] != undefined) {
            (this.range as any).updateAllValues();
        }
    }
};