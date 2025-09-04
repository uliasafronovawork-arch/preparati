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
      const lastCol = sheet.getLastColumn();
      if (!lastRow || !lastCol) return;

      const all = sheet.getRange(1, 1, lastRow, lastCol).getValues();

      for (let r = 0; r < lastRow; r++) {
        for (let c = 0; c < lastCol; c++) {
          const val = all[r][c];
          if (!val) continue;

          const match = String(val).trim().match(/^[A-Za-zА-Яа-я0-9.]+/);
          if (!match) continue;

          const cellCode = normalizeCode(match[0]);
          if (cellCode === searchCode) {
            let startRow = r + 1; // начинаем ниже найденного кода
            let endRow = startRow;

            // идём вниз, пока строка не полностью пустая
            while (endRow <= lastRow) {
              const rowVals = all[endRow - 1];
              const isRowEmpty = rowVals.every(v => v === '' || v === null);
              if (isRowEmpty) break;
              endRow++;
            }

            const numRows = endRow - startRow;
            if (numRows <= 0) continue;

            const block = buildBlockWithMerges(sheet, startRow, numRows, lastCol);

            const anchorA1 = colToLetter(c + 1) + startRow;
            const link = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheet.getSheetId()}&range=${anchorA1}`;

            results.push({
              link,
              sheet: sheetName,
              startRow,
              rows: block.numRows,
              cols: block.numCols,
              cells: block.cells
            });

            r += numRows - 1;
            break;
          }
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

function buildBlockWithMerges(sheet, startRow, numRows, numCols) {
  const values = sheet.getRange(startRow, 1, numRows, numCols).getValues();
  const cells = values.map(row =>
    row.map(v => ({ v, rowspan: 1, colspan: 1, show: true }))
  );
  return { numRows, numCols, cells };
}
