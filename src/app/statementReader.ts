import fs from "fs";
import XLSX from "xlsx";

export class StatementReader {
  private static INPUT_PATH = "./files/input/";

  public static getInputFilePaths() {
    return fs
      .readdirSync(this.INPUT_PATH)
      .filter((file) => file.endsWith(".xlsx"))
      .map((file) => `${this.INPUT_PATH}${file}`);
  }

  public static readInputFile(filePath: string) {
    return XLSX.readFile(filePath);
  }
}
