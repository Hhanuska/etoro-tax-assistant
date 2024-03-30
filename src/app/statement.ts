import { WorkBook } from "xlsx";
import { Sheet } from "./sheet";

export class Statement {
  public sheets: { [sheetName: string]: Sheet } = {};

  constructor(private statement: WorkBook) {
    this.statement.SheetNames.forEach((sheetName) => {
      this.sheets[sheetName] = new Sheet(this.statement.Sheets[sheetName]);
    });
  }
}
