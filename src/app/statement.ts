import XLSX, { WorkBook } from "xlsx";

export class Statement {
  constructor(private statement: WorkBook) {}
}
