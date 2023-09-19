import { EXCEL_ROWS_MAX, EXCEL_COLUMNS_MAX } from "../excel/limits";
import { indicesToCoord, coordToIndices } from "../excel/coords";

export enum Direction {
    Up,
    Down,
    Right,
    Left
};

export class CellLocation {
    private rowIndex : number;
    private columnIndex : number;

    public static stringifyCoords(columnIndex : number, rowIndex : number) {
        return indicesToCoord(columnIndex, rowIndex);
    }

    public static fromIndices(columnIndex : number, rowIndex : number) : CellLocation {
        return new CellLocation(indicesToCoord(columnIndex, rowIndex));
    }

    constructor(location : string | CellLocation) {
        if (typeof location === "string") {
            const indices : [number, number] = coordToIndices(location);
            this.rowIndex = indices[0] as number;
            this.columnIndex = indices[1] as number;
        } else {
            this.rowIndex = location.rowIndex;
            this.columnIndex = location.columnIndex;
        }
    }

    public stringifyInstanceCoords() : string {
        return CellLocation.stringifyCoords(this.columnIndex, this.rowIndex);
    }

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

    public nextCellLocation(distance : number, direction : Direction) : CellLocation {
        if (!this.isLocationValid(distance, direction)) {
            throw new Error("Cannot step on this direction !");
        }
        if (direction == Direction.Down) {
            return new CellLocation(CellLocation.stringifyCoords(this.columnIndex, this.rowIndex + distance));
        } else if (direction == Direction.Up) {
            return new CellLocation(CellLocation.stringifyCoords(this.columnIndex, this.rowIndex - distance));
        } else if (direction == Direction.Right) {
            return new CellLocation(CellLocation.stringifyCoords(this.columnIndex + distance, this.rowIndex));
        } else if (direction == Direction.Left) {
            return new CellLocation(CellLocation.stringifyCoords(this.columnIndex - distance, this.rowIndex));
        }
        throw new Error("direction is not in enum !");
    }
};