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
  const newestDate = new Date(dates[dates.length - 1]);

  const rates = await MNB.getExchangeRates(oldestDate, newestDate);
}

function getDates(statement: Statement) {
  const dates: number[] = [];

  const closedPositionsSheet = statement.sheets["Closed Positions"];

  const col = closedPositionsSheet.colMap["Open Date"];

  for (
    let i = closedPositionsSheet.dimensions.startRow + 1;
    i <= closedPositionsSheet.dimensions.endRow;
    i++
  ) {
    const cell: XLSX.CellObject = closedPositionsSheet.sheet[`${col}${i}`];
    if (!cell || !cell.v) {
      continue;
    }
    const date = dateAndTimeToDate(cell.v.toString());
    dates.push(date.valueOf());
  }

  const activitySheet = statement.sheets["Account Activity"];

  const dateCol = activitySheet.colMap["Date"];

  for (
    let i = activitySheet.dimensions.startRow + 1;
    i <= activitySheet.dimensions.endRow;
    i++
  ) {
    const cell: XLSX.CellObject = activitySheet.sheet[`${dateCol}${i}`];
    if (!cell || !cell.v) {
      continue;
    }
    const date = dateAndTimeToDate(cell.v.toString());
    dates.push(date.valueOf());
  }

  return dates.sort();
}

function dateAndTimeToDate(dateAndTime: string): Date {
  const [date, time] = dateAndTime.split(" ");
  return new Date(reformatDate(date));
}

function reformatDate(date: string) {
  return date.split("/").reverse().join("-");
}
