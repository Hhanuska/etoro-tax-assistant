import XLSX, { WorkBook } from "xlsx";

export class Sheet {
  public dimensions: XLSX.Range;

  public colMap: { [col: string]: number } = {};

  constructor(public sheet: XLSX.WorkSheet) {
    this.dimensions = this.getDimensions();
    this.colMap = this.getColMap();
  }

  public getDimensions() {
    const dim = this.sheet["!ref"];

    if (!dim) {
      throw new Error("No dimensions found in the statement");
    }
    const range = XLSX.utils.decode_range(dim);
    return range;
  }

  public getColMap() {
    const colMap: { [col: string]: number } = {};

    for (let col = this.dimensions.s.c; col <= this.dimensions.e.c; col++) {
      const cell =
        this.sheet[XLSX.utils.encode_cell({ c: col, r: this.dimensions.s.r })];
      if (!cell || !cell.v) {
        continue;
      }
      colMap[cell.v] = col;
    }

    return colMap;
  }

  public refreshColMap() {
    this.dimensions = this.getDimensions();
    this.colMap = this.getColMap();
  }
}

interface Dimensions {
  startCol: string;
  startRow: number;
  endCol: string;
  endRow: number;
}
