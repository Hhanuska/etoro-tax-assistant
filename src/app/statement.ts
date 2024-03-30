import XLSX, { WorkBook } from "xlsx";

export class Statement {
  private dimensions: { [sheetName: string]: Dimensions } = {};

  private colMap: { [sheetName: string]: { [col: string]: string } } = {};

  constructor(private statement: WorkBook) {
    this.dimensions["Closed Positions"] = this.getDimensions(
      statement.Sheets["Closed Positions"]
    );
    this.colMap["Closed Positions"] = this.getColMap("Closed Positions");
  }

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

  public getColMap(sheetName: string) {
    const sheet = this.statement.Sheets[sheetName];
    const dimensions = this.dimensions[sheetName];
    const colMap: { [col: string]: string } = {};

    for (
      let i = dimensions.startCol.charCodeAt(0);
      i <= dimensions.endCol.charCodeAt(0);
      i++
    ) {
      const col = String.fromCharCode(i);
      const cell = sheet[`${col}${dimensions.startRow}`];
      if (!cell || !cell.v) {
        continue;
      }
      colMap[cell.v] = col;
    }

    console.log(colMap);

    return colMap;
  }
}

interface Dimensions {
  startCol: string;
  startRow: number;
  endCol: string;
  endRow: number;
}
