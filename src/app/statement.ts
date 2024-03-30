import XLSX, { WorkBook } from "xlsx";

export class Statement {
  constructor(private statement: WorkBook) {}

  public getDimensions(sheet: XLSX.WorkSheet) {
    const dim = sheet["!ref"];

    if (!dim) {
      throw new Error("No dimensions found in the statement");
    }

    const [start, end] = dim.split(":");

    if (!start || !end) {
      throw new Error("Invalid dimensions found in the statement");
    }

    const startCol = start.match(/[A-Z]+/)?.[0];
    const startRow = start.match(/\d+/)?.[0];
    const endCol = end.match(/[A-Z]+/)?.[0];
    const endRow = end.match(/\d+/)?.[0];

    if (!startCol || !startRow || !endCol || !endRow) {
      throw new Error("Invalid dimensions found in the statement");
    }

    return {
      startCol,
      startRow: parseInt(startRow),
      endCol,
      endRow: parseInt(endRow),
    };
  }
}
