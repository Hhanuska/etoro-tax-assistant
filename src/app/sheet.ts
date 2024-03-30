import XLSX, { WorkBook } from "xlsx";

export class Sheet {
  private dimensions: Dimensions;

  private colMap: { [col: string]: string } = {};

  constructor(private sheet: XLSX.WorkSheet) {
    this.dimensions = this.getDimensions();
    this.colMap = this.getColMap();
  }

  public getDimensions() {
    const dim = this.sheet["!ref"];

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

  public getColMap() {
    const colMap: { [col: string]: string } = {};

    for (
      let i = this.dimensions.startCol.charCodeAt(0);
      i <= this.dimensions.endCol.charCodeAt(0);
      i++
    ) {
      const col = String.fromCharCode(i);
      const cell = this.sheet[`${col}${this.dimensions.startRow}`];
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
