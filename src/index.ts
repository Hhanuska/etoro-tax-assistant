import { MNB } from "./currency/mnb";
import * as fs from "fs";
import * as XLSX from "xlsx";

// MNB.getExchangeRates(2022).then((response) => console.log(response));

const INPUT_PATH = "./files/input/";

fs.readdirSync(INPUT_PATH).forEach((file) => {
  if (!file.endsWith(".xlsx")) return;
  const workbook = XLSX.readFile(`${INPUT_PATH}${file}`);
  console.log(workbook.SheetNames);
});
