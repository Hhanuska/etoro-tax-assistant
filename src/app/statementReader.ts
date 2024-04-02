import fs from "fs";
import XLSX from "xlsx";

export class StatementReader {
  private static INPUT_PATH = "./files/input/";

  private static OUTPUT_PATH = "./files/output/";

  public static getInputFilePaths() {
    return fs
      .readdirSync(this.INPUT_PATH)
      .filter((file) => file.endsWith(".xlsx"))
      .map((file) => `${this.INPUT_PATH}${file}`);
  }

  public static readInputFile(filePath: string) {
    return XLSX.readFile(filePath);
  }

  public static writeOutputFile(workbook: XLSX.WorkBook, name: string) {
    if (!fs.existsSync(this.OUTPUT_PATH)) {
      fs.mkdirSync(this.OUTPUT_PATH);
    }

    XLSX.writeFile(workbook, `${this.OUTPUT_PATH}${name}.xlsx`);
  }
}
