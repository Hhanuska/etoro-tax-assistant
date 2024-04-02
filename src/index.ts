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

function addExchangeRatesToDividends(statement: Statement, rates: MNBRate[]) {
  const dividendSheet = statement.sheets["Dividends"];
  const dateCol = dividendSheet.colMap["Date of Payment"];

  const exchRateCol = String.fromCharCode(
    dividendSheet.dimensions.endCol.charCodeAt(0) + 1
  );
  const convertedReceivedCol = String.fromCharCode(
    dividendSheet.dimensions.endCol.charCodeAt(0) + 2
  );
  const convertedWithheldCol = String.fromCharCode(
    dividendSheet.dimensions.endCol.charCodeAt(0) + 3
  );

  dividendSheet.sheet[
    "!ref"
  ] = `${dividendSheet.dimensions.startCol}${dividendSheet.dimensions.startRow}:${convertedWithheldCol}${dividendSheet.dimensions.endRow}`;

  dividendSheet.sheet[`${exchRateCol}${dividendSheet.dimensions.startRow}`] = {
    t: "s",
    v: "Exchange Rate",
  };
  dividendSheet.sheet[
    `${convertedReceivedCol}${dividendSheet.dimensions.startRow}`
  ] = {
    t: "s",
    v: "Amount received (HUF)",
  };
  dividendSheet.sheet[
    `${convertedWithheldCol}${dividendSheet.dimensions.startRow}`
  ] = {
    t: "s",
    v: "Amount withheld (HUF)",
  };

  for (
    let i = dividendSheet.dimensions.startRow + 1;
    i <= dividendSheet.dimensions.endRow;
    i++
  ) {
    const dateCell: XLSX.CellObject = dividendSheet.sheet[`${dateCol}${i}`];
    if (!dateCell || !dateCell.v) {
      continue;
    }
    const date = dateAndTimeToDate(dateCell.v.toString());
    const rate = MNB.getExchangeRate(date, rates);

    dividendSheet.sheet[`${exchRateCol}${i}`] = {
      t: "n",
      v: rate,
    };

    const amountCol = dividendSheet.colMap["Net Dividend Received (USD)"];

    const convertedAmountCell: XLSX.CellObject = {
      t: "n",
      f: `=${amountCol}${i} * ${exchRateCol}${i}`,
    };

    dividendSheet.sheet[`${convertedReceivedCol}${i}`] = convertedAmountCell;

    const withheldCol = dividendSheet.colMap["Withholding Tax Amount (USD)"];

    const convertedWithheldCell: XLSX.CellObject = {
      t: "n",
      f: `=${withheldCol}${i} * ${exchRateCol}${i}`,
    };

    dividendSheet.sheet[`${convertedWithheldCol}${i}`] = convertedWithheldCell;
  }

  statement.sheets["Dividends"].refreshColMap();

  return statement;
}

function createSummary(statement: Statement) {
  const sheet = XLSX.utils.json_to_sheet([]);

  sheet["!ref"] = "A1:X99";

  sheet["A1"] = {
    t: "s",
    v: "Summary generated by https://github.com/Hhanuska/etoro-tax-assistant",
  };

  sheet["B2"] = {
    t: "s",
    v: "USD",
  };
  sheet["C2"] = {
    t: "s",
    v: "HUF",
  };

  const activityValueUsdCol =
    statement.sheets["Account Activity"].colMap["Amount"];
  const activityValueHufCol =
    statement.sheets["Account Activity"].colMap["Amount (HUF)"];
  const activityTypeCol = statement.sheets["Account Activity"].colMap["Type"];

  sheet["A3"] = {
    t: "s",
    v: "Deposits",
  };
  sheet["B3"] = {
    t: "n",
    f: `SUMIFS('Account Activity'!${activityValueUsdCol}:${activityValueUsdCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Deposit")`,
  };
  sheet["C3"] = {
    t: "n",
    f: `SUMIFS('Account Activity'!${activityValueHufCol}:${activityValueHufCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Deposit")`,
  };

  sheet["A4"] = {
    t: "s",
    v: "Withdrawals",
  };
  sheet["B4"] = {
    t: "n",
    f: `ABS(SUMIFS('Account Activity'!${activityValueUsdCol}:${activityValueUsdCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Withdraw Request")+SUMIFS('Account Activity'!${activityValueUsdCol}:${activityValueUsdCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Withdraw Request Cancelled"))`,
  };
  sheet["C4"] = {
    t: "n",
    f: `ABS(SUMIFS('Account Activity'!${activityValueHufCol}:${activityValueHufCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Withdraw Request")+SUMIFS('Account Activity'!${activityValueHufCol}:${activityValueHufCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Withdraw Request Cancelled"))`,
  };

  sheet["A5"] = {
    t: "s",
    v: "Withdrawal Fees",
  };
  sheet["B5"] = {
    t: "n",
    f: `ABS(SUMIFS('Account Activity'!${activityValueUsdCol}:${activityValueUsdCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Withdraw Fee")+SUMIFS('Account Activity'!${activityValueUsdCol}:${activityValueUsdCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Withdraw Fee Cancelled"))`,
  };
  sheet["C5"] = {
    t: "n",
    f: `ABS(SUMIFS('Account Activity'!${activityValueHufCol}:${activityValueHufCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Withdraw Fee")+SUMIFS('Account Activity'!${activityValueHufCol}:${activityValueHufCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Withdraw Fee Cancelled"))`,
  };

  sheet["A6"] = {
    t: "s",
    v: "Net Deposits",
  };
  sheet["B6"] = {
    t: "n",
    f: `B3-B4`,
  };
  sheet["C6"] = {
    t: "n",
    f: `C3-C4`,
  };

  sheet["A7"] = {
    t: "s",
    v: "Total Cost",
  };
  sheet["B7"] = {
    t: "n",
    f: `B5+B6`,
  };
  sheet["C7"] = {
    t: "n",
    f: `C5+C6`,
  };

  const activityAssetTypeCol =
    statement.sheets["Account Activity"].colMap["Asset type"];

  sheet["A9"] = {
    t: "s",
    v: "Open Positions (Total)",
  };
  sheet["B9"] = {
    t: "n",
    f: `SUMIFS('Account Activity'!${activityValueUsdCol}:${activityValueUsdCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Open Position")`,
  };
  sheet["C9"] = {
    t: "n",
    f: `SUMIFS('Account Activity'!${activityValueHufCol}:${activityValueHufCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Open Position")`,
  };

  sheet["A10"] = {
    t: "s",
    v: "Open Positions (Stocks)",
  };
  sheet["B10"] = {
    t: "n",
    f: `SUMIFS('Account Activity'!${activityValueUsdCol}:${activityValueUsdCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Open Position", 'Account Activity'!${activityAssetTypeCol}:${activityAssetTypeCol}, "Stocks")`,
  };
  sheet["C10"] = {
    t: "n",
    f: `SUMIFS('Account Activity'!${activityValueHufCol}:${activityValueHufCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Open Position", 'Account Activity'!${activityAssetTypeCol}:${activityAssetTypeCol}, "Stocks")`,
  };

  sheet["A11"] = {
    t: "s",
    v: "Open Positions (CFD)",
  };
  sheet["B11"] = {
    t: "n",
    f: `SUMIFS('Account Activity'!${activityValueUsdCol}:${activityValueUsdCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Open Position", 'Account Activity'!${activityAssetTypeCol}:${activityAssetTypeCol}, "CFD")`,
  };
  sheet["C11"] = {
    t: "n",
    f: `SUMIFS('Account Activity'!${activityValueHufCol}:${activityValueHufCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Open Position", 'Account Activity'!${activityAssetTypeCol}:${activityAssetTypeCol}, "CFD")`,
  };

  sheet["A12"] = {
    t: "s",
    v: "Open Positions (Crypto)",
  };
  sheet["B12"] = {
    t: "n",
    f: `SUMIFS('Account Activity'!${activityValueUsdCol}:${activityValueUsdCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Open Position", 'Account Activity'!${activityAssetTypeCol}:${activityAssetTypeCol}, "Crypto")`,
  };
  sheet["C12"] = {
    t: "n",
    f: `SUMIFS('Account Activity'!${activityValueHufCol}:${activityValueHufCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Open Position", 'Account Activity'!${activityAssetTypeCol}:${activityAssetTypeCol}, "Crypto")`,
  };

  sheet["A14"] = {
    t: "s",
    v: "Closed Positions (Total)",
  };
  sheet["B14"] = {
    t: "n",
    f: `SUMIFS('Account Activity'!${activityValueUsdCol}:${activityValueUsdCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Position closed")`,
  };
  sheet["C14"] = {
    t: "n",
    f: `SUMIFS('Account Activity'!${activityValueHufCol}:${activityValueHufCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Position closed")`,
  };

  sheet["A15"] = {
    t: "s",
    v: "Closed Positions (Stocks)",
  };
  sheet["B15"] = {
    t: "n",
    f: `SUMIFS('Account Activity'!${activityValueUsdCol}:${activityValueUsdCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Position closed", 'Account Activity'!${activityAssetTypeCol}:${activityAssetTypeCol}, "Stocks")`,
  };
  sheet["C15"] = {
    t: "n",
    f: `SUMIFS('Account Activity'!${activityValueHufCol}:${activityValueHufCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Position closed", 'Account Activity'!${activityAssetTypeCol}:${activityAssetTypeCol}, "Stocks")`,
  };

  sheet["A16"] = {
    t: "s",
    v: "Closed Positions (CFD)",
  };
  sheet["B16"] = {
    t: "n",
    f: `SUMIFS('Account Activity'!${activityValueUsdCol}:${activityValueUsdCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Position closed", 'Account Activity'!${activityAssetTypeCol}:${activityAssetTypeCol}, "CFD")`,
  };
  sheet["C16"] = {
    t: "n",
    f: `SUMIFS('Account Activity'!${activityValueHufCol}:${activityValueHufCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Position closed", 'Account Activity'!${activityAssetTypeCol}:${activityAssetTypeCol}, "CFD")`,
  };

  sheet["A17"] = {
    t: "s",
    v: "Closed Positions (Crypto)",
  };
  sheet["B17"] = {
    t: "n",
    f: `SUMIFS('Account Activity'!${activityValueUsdCol}:${activityValueUsdCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Position closed", 'Account Activity'!${activityAssetTypeCol}:${activityAssetTypeCol}, "Crypto")`,
  };
  sheet["C17"] = {
    t: "n",
    f: `SUMIFS('Account Activity'!${activityValueHufCol}:${activityValueHufCol}, 'Account Activity'!${activityTypeCol}:${activityTypeCol}, "Position closed", 'Account Activity'!${activityAssetTypeCol}:${activityAssetTypeCol}, "Crypto")`,
  };

  const closedPositionsProfitUsdCol =
    statement.sheets["Closed Positions"].colMap["Profit"] ??
    statement.sheets["Closed Positions"].colMap["Profit(USD)"];
  const closedPositionsProfitHufCol =
    statement.sheets["Closed Positions"].colMap["Profit (HUF)"];
  const closedPositionsAssetTypeCol =
    statement.sheets["Closed Positions"].colMap["Type"];

  sheet["A19"] = {
    t: "s",
    v: "Profit from closed positions (Total)",
  };
  sheet["B19"] = {
    t: "n",
    f: `SUM('Closed Positions'!${closedPositionsProfitUsdCol}:${closedPositionsProfitUsdCol})`,
  };
  sheet["C19"] = {
    t: "n",
    f: `SUM('Closed Positions'!${closedPositionsProfitHufCol}:${closedPositionsProfitHufCol})`,
  };

  sheet["A20"] = {
    t: "s",
    v: "Profit from closed positions (Stocks)",
  };
  sheet["B20"] = {
    t: "n",
    f: `SUMIFS('Closed Positions'!${closedPositionsProfitUsdCol}:${closedPositionsProfitUsdCol}, 'Closed Positions'!${closedPositionsAssetTypeCol}:${closedPositionsAssetTypeCol}, "Stocks")`,
  };
  sheet["C20"] = {
    t: "n",
    f: `SUMIFS('Closed Positions'!${closedPositionsProfitHufCol}:${closedPositionsProfitHufCol}, 'Closed Positions'!${closedPositionsAssetTypeCol}:${closedPositionsAssetTypeCol}, "Stocks")`,
  };

  sheet["A21"] = {
    t: "s",
    v: "Profit from closed positions (CFD)",
  };
  sheet["B21"] = {
    t: "n",
    f: `SUMIFS('Closed Positions'!${closedPositionsProfitUsdCol}:${closedPositionsProfitUsdCol}, 'Closed Positions'!${closedPositionsAssetTypeCol}:${closedPositionsAssetTypeCol}, "CFD")`,
  };
  sheet["C21"] = {
    t: "n",
    f: `SUMIFS('Closed Positions'!${closedPositionsProfitHufCol}:${closedPositionsProfitHufCol}, 'Closed Positions'!${closedPositionsAssetTypeCol}:${closedPositionsAssetTypeCol}, "CFD")`,
  };

  sheet["A22"] = {
    t: "s",
    v: "Profit from closed positions (Crypto)",
  };
  sheet["B22"] = {
    t: "n",
    f: `SUMIFS('Closed Positions'!${closedPositionsProfitUsdCol}:${closedPositionsProfitUsdCol}, 'Closed Positions'!${closedPositionsAssetTypeCol}:${closedPositionsAssetTypeCol}, "Crypto")`,
  };
  sheet["C22"] = {
    t: "n",
    f: `SUMIFS('Closed Positions'!${closedPositionsProfitHufCol}:${closedPositionsProfitHufCol}, 'Closed Positions'!${closedPositionsAssetTypeCol}:${closedPositionsAssetTypeCol}, "Crypto")`,
  };

  XLSX.utils.book_append_sheet(statement.getWorkbook(), sheet, "Summary");
}
