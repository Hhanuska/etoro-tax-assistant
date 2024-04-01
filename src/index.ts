import XLSX, { WorkBook } from "xlsx";
import { StatementReader } from "./app/statementReader";
import { MNB, MNBRate } from "./currency/mnb";
import { Statement } from "./app/statement";

const statements = StatementReader.getInputFilePaths().map((path) => {
  const fileName = path.split("/")[path.split("/").length - 1];
  return {
    name: fileName.substring(0, fileName.indexOf(".xlsx")),
    wb: StatementReader.readInputFile(path),
  };
});

statements.forEach(handleStatement);

async function handleStatement(s: { name: string; wb: WorkBook }) {
  const statement = new Statement(s.wb);

  const dates = getDates(statement);

  const oldestDate = new Date(dates[0]);
  const newestDate = new Date(dates[dates.length - 1]);

  const rates = await MNB.getExchangeRates(oldestDate, newestDate);

  addExchangeRatesToStatement(statement, rates);

  StatementReader.writeOutputFile(
    statement.getWorkbook(),
    `./files/output/${s.name}_output.xlsx`
  );
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

function addExchangeRatesToStatement(statement: Statement, rates: MNBRate[]) {
  const activitySheet = statement.sheets["Account Activity"];
  const dateCol = activitySheet.colMap["Date"];

  const exchRateCol = String.fromCharCode(
    activitySheet.dimensions.endCol.charCodeAt(0) + 1
  );

  activitySheet.sheet[
    "!ref"
  ] = `${activitySheet.dimensions.startCol}${activitySheet.dimensions.startRow}:${exchRateCol}${activitySheet.dimensions.endRow}`;

  activitySheet.sheet[`${exchRateCol}${activitySheet.dimensions.startRow}`] = {
    t: "s",
    v: "Exchange Rate",
  };

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
    const rate = MNB.getExchangeRate(date, rates);

    activitySheet.sheet[`${exchRateCol}${i}`] = {
      t: "n",
      v: rate,
    };
  }

  return statement;
}
