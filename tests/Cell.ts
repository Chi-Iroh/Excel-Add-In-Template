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

  expect(A1.getValue()).toBe(42);
  expect(new Cell(worksheet, "A1").getValue()).toBe(42)

  expect(A2.getValue()).toBe(42);
  expect(A3.getValue()).toBe(84);

  A4.setValue("=A3*A3");
  A4.updateAllValues();
  expect(A4.getValue()).toBe(84 * 84);

  A5.setValue("=SQRT(A4)");
  A5.updateAllValues();
  expect(A5.getValue()).toBe(A3.getValue());
});