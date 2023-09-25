import "./setup/setup";
import { Cell } from "../src/cell/Cell";

test("Sum of cells", () => {
  const worksheet = new Excel.RequestContext().workbook.worksheets.getActiveWorksheet();
  const A1 : Cell = new Cell(worksheet, "A1");
  const A2 : Cell = new Cell(worksheet, "A2");
  const A3 : Cell = new Cell(worksheet, "A3");
  const A4 : Cell = new Cell(worksheet, "A4");
  const A5 : Cell = new Cell(worksheet, "A5");

  A1.setValue(42);
  A1.updateAllValues();

  A2.setValue(42);
  A2.updateAllValues();

  A3.setValue("=A1+A2");
  A3.updateAllValues();

  expect(A1.computeValue()).toBe(42);
  expect(new Cell(worksheet, "A1").computeValue()).toBe(42)

  expect(A2.computeValue()).toBe(42);
  expect(A3.computeValue()).toBe(84);

  A4.setValue("=A3*A3");
  A4.updateAllValues();
  expect(A4.computeValue()).toBe(84 * 84);

  A5.setValue("=SQRT(A4)");
  A5.updateAllValues();
  expect(A5.computeValue()).toBe(A3.computeValue());
});