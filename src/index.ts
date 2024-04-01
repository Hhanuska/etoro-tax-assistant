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

  addExchangeRatesToActivity(statement, rates);
  addExchangeRatesToClosedPositions(statement, rates);
  createSummary(statement);

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

function addExchangeRatesToActivity(statement: Statement, rates: MNBRate[]) {
  const activitySheet = statement.sheets["Account Activity"];
  const dateCol = activitySheet.colMap["Date"];

  const exchRateCol = String.fromCharCode(
    activitySheet.dimensions.endCol.charCodeAt(0) + 1
  );

  const convertedCol = String.fromCharCode(
    activitySheet.dimensions.endCol.charCodeAt(0) + 2
  );

  activitySheet.sheet[
    "!ref"
  ] = `${activitySheet.dimensions.startCol}${activitySheet.dimensions.startRow}:${convertedCol}${activitySheet.dimensions.endRow}`;

  activitySheet.sheet[`${exchRateCol}${activitySheet.dimensions.startRow}`] = {
    t: "s",
    v: "Exchange Rate",
  };

  activitySheet.sheet[`${convertedCol}${activitySheet.dimensions.startRow}`] = {
    t: "s",
    v: "Amount (HUF)",
  };

  for (
    let i = activitySheet.dimensions.startRow + 1;
    i <= activitySheet.dimensions.endRow;
    i++
  ) {
    const dateCell: XLSX.CellObject = activitySheet.sheet[`${dateCol}${i}`];
    if (!dateCell || !dateCell.v) {
      continue;
    }
    const date = dateAndTimeToDate(dateCell.v.toString());
    const rate = MNB.getExchangeRate(date, rates);

    activitySheet.sheet[`${exchRateCol}${i}`] = {
      t: "n",
      v: rate,
    };

    const amountCol = activitySheet.colMap["Amount"];

    const convertedAmountCell: XLSX.CellObject = {
      t: "n",
      f: `=${amountCol}${i} * ${exchRateCol}${i}`,
    };

    activitySheet.sheet[`${convertedCol}${i}`] = convertedAmountCell;
  }

  statement.sheets["Account Activity"].refreshColMap();

  return statement;
}

function addExchangeRatesToClosedPositions(
  statement: Statement,
  rates: MNBRate[]
) {
  const closedPositionsSheet = statement.sheets["Closed Positions"];
  const openDateCol = closedPositionsSheet.colMap["Open Date"];
  const closeDateCol = closedPositionsSheet.colMap["Close Date"];

  const exchRateOpenCol = String.fromCharCode(
    closedPositionsSheet.dimensions.endCol.charCodeAt(0) + 1
  );

  const convertedOpenCol = String.fromCharCode(
    closedPositionsSheet.dimensions.endCol.charCodeAt(0) + 2
  );

  const exchRateCloseCol = String.fromCharCode(
    closedPositionsSheet.dimensions.endCol.charCodeAt(0) + 3
  );

  const convertedCloseCol = String.fromCharCode(
    closedPositionsSheet.dimensions.endCol.charCodeAt(0) + 4
  );

  const convertedProfitCol = String.fromCharCode(
    closedPositionsSheet.dimensions.endCol.charCodeAt(0) + 5
  );

  closedPositionsSheet.sheet[
    "!ref"
  ] = `${closedPositionsSheet.dimensions.startCol}${closedPositionsSheet.dimensions.startRow}:${convertedProfitCol}${closedPositionsSheet.dimensions.endRow}`;

  closedPositionsSheet.sheet[
    `${exchRateOpenCol}${closedPositionsSheet.dimensions.startRow}`
  ] = {
    t: "s",
    v: "Exchange Rate at open date",
  };

  closedPositionsSheet.sheet[
    `${exchRateCloseCol}${closedPositionsSheet.dimensions.startRow}`
  ] = {
    t: "s",
    v: "Exchange Rate at close date",
  };

  closedPositionsSheet.sheet[
    `${convertedOpenCol}${closedPositionsSheet.dimensions.startRow}`
  ] = {
    t: "s",
    v: "Amount at open (HUF)",
  };

  closedPositionsSheet.sheet[
    `${convertedCloseCol}${closedPositionsSheet.dimensions.startRow}`
  ] = {
    t: "s",
    v: "Amount at close (HUF)",
  };

  closedPositionsSheet.sheet[
    `${convertedProfitCol}${closedPositionsSheet.dimensions.startRow}`
  ] = {
    t: "s",
    v: "Profit (HUF)",
  };

  for (
    let i = closedPositionsSheet.dimensions.startRow + 1;
    i <= closedPositionsSheet.dimensions.endRow;
    i++
  ) {
    const openDateCell: XLSX.CellObject =
      closedPositionsSheet.sheet[`${openDateCol}${i}`];
    if (!openDateCell || !openDateCell.v) {
      continue;
    }
    const openDate = dateAndTimeToDate(openDateCell.v.toString());
    const openRate = MNB.getExchangeRate(openDate, rates);

    closedPositionsSheet.sheet[`${exchRateOpenCol}${i}`] = {
      t: "n",
      v: openRate,
    };

    const amountCol = closedPositionsSheet.colMap["Amount"];

    const convertedOpenAmountCell: XLSX.CellObject = {
      t: "n",
      f: `=${amountCol}${i} * ${exchRateOpenCol}${i}`,
    };

    closedPositionsSheet.sheet[`${convertedOpenCol}${i}`] =
      convertedOpenAmountCell;

    const closeDateCell: XLSX.CellObject =
      closedPositionsSheet.sheet[`${closeDateCol}${i}`];
    if (!closeDateCell || !closeDateCell.v) {
      continue;
    }
    const closeDate = dateAndTimeToDate(closeDateCell.v.toString());
    const closeRate = MNB.getExchangeRate(closeDate, rates);

    closedPositionsSheet.sheet[`${exchRateCloseCol}${i}`] = {
      t: "n",
      v: closeRate,
    };

    const profitCol =
      closedPositionsSheet.colMap["Profit"] ??
      closedPositionsSheet.colMap["Profit(USD)"];

    const convertedCloseAmountCell: XLSX.CellObject = {
      t: "n",
      f: `=(${amountCol}${i} - ${profitCol}${i}) * ${exchRateCloseCol}${i}`,
    };

    closedPositionsSheet.sheet[`${convertedCloseCol}${i}`] =
      convertedCloseAmountCell;

    const convertedProfitCell: XLSX.CellObject = {
      t: "n",
      f: `=${convertedOpenCol}${i} - ${convertedCloseCol}${i}`,
    };

    closedPositionsSheet.sheet[`${convertedProfitCol}${i}`] =
      convertedProfitCell;
  }

  statement.sheets["Closed Positions"].refreshColMap();

  return statement;
}
