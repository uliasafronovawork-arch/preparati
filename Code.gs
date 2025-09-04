function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ICD Search')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function searchICDCodePublic(code) {
  try {
    const searchCode = normalizeCode(code || '');
    const spreadsheetId = '1zS0DoyXg6GE0rsOr1NLkDdljplc1GXU1cNVCZRgS1rQ';
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheetsToSearch = ['часть 1', 'часть 2'];
    const results = [];

    sheetsToSearch.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;

      const lastRow = sheet.getLastRow();
      if (!lastRow) return;

      const all = sheet.getRange(1, 1, lastRow, 2).getValues(); // только 2 колонки: код + название

      for (let r = 0; r < lastRow; r++) {
        const codeVal = all[r][0]; // первый столбец — код
        if (!codeVal) continue;

        const match = String(codeVal).trim().match(/^[A-Za-zА-Яа-я0-9.]+/);
        if (!match) continue;

        const cellCode = normalizeCode(match[0]);
        if (cellCode === searchCode) {
          let startRow = r + 1;
          let endRow = startRow;

          // идём вниз, пока не встретим пустую строку
          while (endRow <= lastRow) {
            const rowVals = all[endRow - 1];
            const isRowEmpty = rowVals.every(v => v === '' || v === null);
            if (isRowEmpty) break;
            endRow++;
          }

          const numRows = endRow - startRow;
          if (numRows <= 0) continue;

          const values = sheet.getRange(startRow, 1, numRows, 2).getValues();

          const block = values.map(row => ({
            code: row[0],
            name: row[1]
          }));

          const link = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheet.getSheetId()}&range=${startRow}:${endRow}`;

          results.push({
            link,
            sheet: sheetName,
            startRow,
            rows: numRows,
            cols: 2,
            data: block
          });

          break; // нашли — дальше можно не искать
        }
      }
    });

    return results;
  } catch (e) {
    return [{ error: e && e.message ? e.message : String(e) }];
  }
}

function normalizeCode(code) {
  return String(code).toUpperCase().replace(/\s+/g, '');
}

function colToLetter(col) {
  let temp, letter = '';
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}
