import { WorkBook } from "xlsx";
import { StatementReader } from "./app/statementReader";
import { MNB } from "./currency/mnb";
import { Statement } from "./app/statement";

const statements = StatementReader.getInputFilePaths().map((path) =>
  StatementReader.readInputFile(path)
);

statements.forEach(handleStatement);

async function handleStatement(s: WorkBook) {
  const statement = new Statement(s);
}
