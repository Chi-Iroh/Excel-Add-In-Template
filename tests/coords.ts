import "./setup/setup";
import * as coords from "../src/excel/coords";

test("Coords", () => {
    expect(coords.columnToIndex("A")).toBe(0);
    expect(coords.columnToIndex("AA")).toBe(26);
    expect(coords.indexToColumn(0)).toBe("A");
    expect(coords.indexToColumn(26)).toBe("AA");
})