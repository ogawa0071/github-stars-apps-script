function onOpen(): void {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("GitHub Stars")
    .addItem("Get", "main")
    .addToUi();
}

const getSpreadsheetValues = (): {
  columns: GoogleAppsScript.Spreadsheet.Range[];
  nextColumns: GoogleAppsScript.Spreadsheet.Range[];
} => {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getActiveSheet();
  const range = sheet.getActiveRange();
  const rows = range.getNumRows();
  const rowPosition = range.getRow();
  const columnPosition = range.getColumn();

  const columns: GoogleAppsScript.Spreadsheet.Range[] = [];
  const nextColumns: GoogleAppsScript.Spreadsheet.Range[] = [];

  for (let index = 0; index < rows; index++) {
    const row = index + rowPosition;
    const column = columnPosition;

    columns.push(sheet.getRange(row, column));
    nextColumns.push(sheet.getRange(row, column + 1));
  }

  return {
    columns,
    nextColumns
  };
};

const getStargazersCount = (owner: string, repo: string): number => {
  const response = UrlFetchApp.fetch(
    `https://api.github.com/repos/${owner}/${repo}`
  );

  return JSON.parse(response.getContentText()).stargazers_count;
};

function main(): void {
  const { columns, nextColumns } = getSpreadsheetValues();

  for (let index = 0; index < columns.length; index++) {
    const column = columns[index];
    const nextColumn = nextColumns[index];

    const url = column.getValue() as string;

    try {
      const [, , owner, repo] = /(github.com)\/(.+)\/(.+)((\/|\?|#)(.*))?/.exec(
        url
      );

      const stargazersCount = getStargazersCount(owner, repo);
      nextColumn.setValue(stargazersCount);
    } catch {
      console.error(`Not GitHub repository: ${url}`);
    }
  }
}
