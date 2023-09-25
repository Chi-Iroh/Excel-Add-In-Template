import { EXCEL_ROWS_MAX, EXCEL_COLUMNS_MAX } from "../excel/limits";
import { indicesToCoord, coordsToIndices } from "../excel/coords";

/**
 * Directions to help selecting cells.
 */
export enum Direction {
    Up,
    Down,
    Right,
    Left
};

/**
 * A class to represent and manipulate a cell's location.
 */
export class CellLocation {
    private rowIndex : number;
    private columnIndex : number;

    /**
     * Converts coordinates to Excel format.
     * @param columnIndex zero-based column index (>=0)
     * @param rowIndex zero-based row indew (>=0)
     * @returns Excel-formatted coordinates string
     * @throw Error if either rowIndex or columnIndex is strictly negative
     */
    public static stringifyCoords(rowIndex : number, columnIndex : number) : string {
        return indicesToCoord(rowIndex, columnIndex);
    }

    /**
     * Creates a new CellLocation instance from indices
     * @param rowIndex zero-based row index
     * @returns new CellLocation instance from these indices.
     * @note No constructor overloading in TypeScript.
     */
    public static fromIndices(rowIndex : number, columnIndex : number) : CellLocation {
        return new CellLocation(indicesToCoord(rowIndex, columnIndex));
    }

    /**
     * CellLocation instantiation from coords
     * @param location Excel-formatted coordinates string / another CellLocation instance
     */
    constructor(location : string | CellLocation) {
        if (typeof location === "string") {
            const indices : [number, number] = coordsToIndices(location);
            this.rowIndex = indices[0] as number;
            this.columnIndex = indices[1] as number;
        } else {
            this.rowIndex = location.rowIndex;
            this.columnIndex = location.columnIndex;
        }
    }

    /**
     * Converts instance coordinates to Excel-formatted coordinates string.
     * @returns coordinates string
     */
    public stringifyInstanceCoords() : string {
        return CellLocation.stringifyCoords(this.rowIndex, this.columnIndex);
    }

    /**
     * Checks if a location is valid, according a distance a direction from this instance.
     * @param distance distance in cells from this instance
     * @param direction direction in which to go
     * @returns false if location is out of bounds, true otherwise
     */
    public isLocationValid(distance : number, direction : Direction) : boolean {
        if (direction == Direction.Down) {
            return this.columnIndex < EXCEL_ROWS_MAX - distance;
        } else if (direction == Direction.Up) {
            return this.columnIndex >= distance;
        } else if (direction == Direction.Right) {
            return this.rowIndex < EXCEL_COLUMNS_MAX - distance;
        } else if (direction == Direction.Left) {
            return this.rowIndex >= distance;
        }
        return false;
    }

    /**
     * Finds a CellLocation according a distance a direction from this instance.
     * @param distance distance in cells from this instance
     * @param direction direction in which to go
     * @returns new CellLocation
     * @throws Error if out of bounds location (see CellLocation.isLocationValid) or invalid direction (bad value casted as enum).
     */
    public nextCellLocation(distance : number, direction : Direction) : CellLocation {
        if (!this.isLocationValid(distance, direction)) {
            throw new Error("Cannot step on this direction !");
        }
        if (direction == Direction.Down) {
            return new CellLocation(CellLocation.stringifyCoords(this.rowIndex + distance, this.columnIndex));
        } else if (direction == Direction.Up) {
            return new CellLocation(CellLocation.stringifyCoords(this.rowIndex - distance, this.columnIndex));
        } else if (direction == Direction.Right) {
            return new CellLocation(CellLocation.stringifyCoords(this.rowIndex, this.columnIndex + distance));
        } else if (direction == Direction.Left) {
            return new CellLocation(CellLocation.stringifyCoords(this.rowIndex, this.columnIndex - distance));
        }
        throw new Error(`Bad direction value, must be >=0 and <=3 (use Direction.XXX), got ${direction} !`);
    }
};