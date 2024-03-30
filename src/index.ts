import XLSX, { WorkBook } from "xlsx";
import { StatementReader } from "./app/statementReader";
import { MNB } from "./currency/mnb";
import { Statement } from "./app/statement";

const statements = StatementReader.getInputFilePaths().map((path) =>
  StatementReader.readInputFile(path)
);

statements.forEach(handleStatement);

async function handleStatement(s: WorkBook) {
  const statement = new Statement(s);

  const dates = getDates(statement);

  const oldestDate = new Date(dates[0]);

  console.log(oldestDate);
}

function getDates(statement: Statement) {
  const dates: number[] = [];

  const sheet = statement.sheets["Closed Positions"];

  const col = sheet.colMap["Open Date"];

  for (
    let i = sheet.dimensions.startRow + 1;
    i <= sheet.dimensions.endRow;
    i++
  ) {
    const cell: XLSX.CellObject = sheet.sheet[`${col}${i}`];
    if (!cell || !cell.v) {
      continue;
    }
    const [date, time] = cell.v.toString().split(" ");
    dates.push(new Date(reformatDate(date)).valueOf());
  }

  return dates.sort();
}

function reformatDate(date: string) {
  return date.split("/").reverse().join("-");
}
