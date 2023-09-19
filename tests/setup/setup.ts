import { OfficeMockObject } from "office-addin-mock";
import { coordToIndices } from "../../src/excel/coords";
import { isDigit, isLetter } from "../../src/charType";
import { stringIndexFindIf } from "../../src/utils";
import { EXCEL_COLUMNS_MAX, EXCEL_ROWS_MAX } from "../../src/excel/limits";
import * as math from "mathjs";

(EXCEL_COLUMNS_MAX as any) = 10;
(EXCEL_ROWS_MAX as any) = 10;

export let excelAllValues : (string | number)[][] = Array.from({length : EXCEL_ROWS_MAX}, row => Array(EXCEL_COLUMNS_MAX).fill(0));

export class ExcelRangeMock {
  private topLeftRowIndex : number = 0;
  private topLeftColumnIndex : number = 0;
  private bottomRightRowIndex : number = 0;
  private bottomRightColumnIndex : number = 0;
  private nRows : number = 0;
  private nColumns : number = 0;
  public values : (string | number)[][] = [];

  constructor(location : string) {
    const dotsIndex : number = location.indexOf(":");
    if (dotsIndex != -1) {
      const topLeftCell : string = location.substring(0, dotsIndex);
      const topLeftIndices = coordToIndices(topLeftCell);
      const bottomRightCell : string = location.substring(dotsIndex);
      const bottomRightIndices = coordToIndices(bottomRightCell);
      this.topLeftRowIndex = topLeftIndices[0] as number;
      this.topLeftColumnIndex = topLeftIndices[1] as number;
      this.bottomRightRowIndex = bottomRightIndices[0] as number;
      this.bottomRightColumnIndex = bottomRightIndices[1] as number;
    } else {
      const indices = coordToIndices(location);
      this.topLeftRowIndex = indices[0] as number;
      this.topLeftColumnIndex = indices[1] as number;
      this.bottomRightRowIndex = indices[0] as number;
      this.bottomRightColumnIndex = indices[1] as number;
    }
    this.nRows = 1 + this.bottomRightRowIndex - this.topLeftRowIndex;
    this.nColumns = 1 + this.bottomRightColumnIndex - this.topLeftColumnIndex;
    this.values = Array.from({length: this.nRows}, row => new Array(this.nColumns).fill(0))

    for (let rowIndex : number = 0; rowIndex < this.nRows; rowIndex++) {
      for (let columnIndex : number = 0; columnIndex < this.nColumns; columnIndex++) {
        this.values[rowIndex][columnIndex] = excelAllValues[this.topLeftRowIndex + rowIndex][this.topLeftColumnIndex + columnIndex];
      }
    }
  }

  private getOtherCellValueAsNumber(cell : string) : number {
    let cellRange = new ExcelRangeMock(cell);
    cellRange.calculate();
    return Number(cellRange.values[0][0]);
  }

  private extractCellLocation(location : string) : string | undefined {
    const digitIndex : number = stringIndexFindIf(location, isDigit);
    if (!isLetter(location[0]) || digitIndex == -1) {
      return undefined;
    }

    const digitsPart : string = location.substring(digitIndex);
    const isNotDigit = (char : string) => { return !isDigit(char); }
    const nextNonDigitIndex = stringIndexFindIf(digitsPart, isNotDigit);
    if (nextNonDigitIndex == -1) {
      return location;
    }
    return location.substring(0, digitIndex + nextNonDigitIndex);
  }

  private calculateSingleCell(rowIndex : number, columnIndex : number) : void {
    let formula : string = this.values[rowIndex][columnIndex] as string;
    let mathExpr : string = "";

    if (formula[0] == "=") {
      formula = formula.substring(1);
    }

    for (let i = 0; i < formula.length; i++) {
      const cellLocation : string | undefined = this.extractCellLocation(formula.substring(i));

      if (typeof cellLocation === "string") {
        console.log(cellLocation);
        mathExpr += this.getOtherCellValueAsNumber(cellLocation).toString();
        i += cellLocation.length - 1;
      } else {
        mathExpr += formula[i];
      }
      console.log(mathExpr)
    }
    this.values[rowIndex][columnIndex] = math.evaluate(mathExpr);
  }

  public calculate() : void {
    for (let rowIndex : number = 0; rowIndex < this.values.length; rowIndex++) {
      for (let columnIndex : number = 0; columnIndex < this.values[rowIndex].length; columnIndex++) {
        const cellCopy : string | number = this.values[rowIndex][columnIndex];
        if (typeof cellCopy === "string" && cellCopy[0] == "=") {
          this.calculateSingleCell(rowIndex, columnIndex);
        }
      }
    }
  }

  public untrack() {}

  public load(property : any) {}

  public updateAllValues() : void {
    for (let rowIndex : number = 0; rowIndex <= this.values?.length; rowIndex++) {
      for (let columnIndex : number = 0; columnIndex < this.values[rowIndex]?.length; columnIndex++) {
        excelAllValues[this.topLeftRowIndex + rowIndex][this.bottomRightColumnIndex + columnIndex] = this.values[rowIndex][columnIndex];
      }
    }
  }
};

export class ExcelRequestContextMock {
  public workbook = {
    worksheets: {
      getActiveWorksheet: function() {
        return {
          getRange : function(location : string) {
            return new ExcelRangeMock(location);
          }
        }
      }
    }
  }
};

const excelMockData = {
  RequestContext: ExcelRequestContextMock,
  Range: ExcelRangeMock
}

const officeMockData = {
  onReady: async function () {}
};

global.Office = new OfficeMockObject(officeMockData) as any;
global.Excel = new OfficeMockObject(excelMockData) as any;