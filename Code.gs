function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ICD Search')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function searchICDCodePublic(query) {
  try {
    const searchRaw = query ? String(query).trim() : '';
    if (!searchRaw) {
      return [{ error: 'Введите код или название заболевания' }];
    }

    const searchCode = normalizeCode(searchRaw); // вариант для кода
    const searchText = searchRaw.toLowerCase();  // вариант для текста

    const spreadsheetId = '1zS0DoyXg6GE0rsOr1NLkDdljplc1GXU1cNVCZRgS1rQ';
    const ss = SpreadsheetApp.openById(spreadsheetId);

    // два листа: в одном коды, в другом названия
    const sheetsToSearch = [
      { name: 'часть 1', mode: 'code' }, // ищем строго по коду
      { name: 'часть 0', mode: 'text' }  // ищем по словам
    ];

    const results = [];

    function cleanSheetText(str) {
      if (!str) return '';
      return String(str).replace(/[\u00A0]+/g, ' ').trim();
    }

    for (const sheetDef of sheetsToSearch) {
      const sheet = ss.getSheetByName(sheetDef.name);
      if (!sheet) continue;

      const lastRow = sheet.getLastRow();
      if (lastRow < 1) continue;

      const all = sheet.getRange(1, 1, lastRow, 2).getValues();

      let foundRow = null;
      for (let r = 0; r < lastRow; r++) {
        const cellA = cleanSheetText(all[r][0]);
        const cellB = cleanSheetText(all[r][1]);

        if (sheetDef.mode === 'code') {
          // проверка по коду (строгое совпадение в колонке A)
          const cellCode = normalizeCode(cellA);
          if (cellCode === searchCode) {
            foundRow = r + 1;
            break;
          }
        } else if (sheetDef.mode === 'text') {
          // поиск текста (вхождение в A или B)
          if ((cellA && cellA.toLowerCase().includes(searchText)) ||
              (cellB && cellB.toLowerCase().includes(searchText))) {
            foundRow = r + 1;
            break;
          }
        }
      }

      if (!foundRow) continue;

      // идём вниз до пустой строки
      let endRow = foundRow;
      while (endRow <= lastRow) {
        const rowVals = all[endRow - 1];
        const isEmptyRow = (!rowVals[0] && !rowVals[1]);
        if (isEmptyRow) break;
        endRow++;
      }

      const blockValues = all.slice(foundRow - 1, endRow).map(r => ({
        code: cleanSheetText(r[0]),
        name: cleanSheetText(r[1])
      }));

      const link = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheet.getSheetId()}&range=A${foundRow}:B${endRow-1}`;

      results.push({
        link,
        sheet: sheetDef.name,
        startRow: foundRow,
        rows: blockValues.length,
        cols: 2,
        data: blockValues
      });

      break; // выходим после первого найденного совпадения
    }

    if (results.length === 0) {
      return [{
        error: `По запросу "${query}" ничего не найдено в листах "часть 1" или "часть 0".`
      }];
    }

    return results;

  } catch (e) {
    return [{ error: e && e.message ? e.message : String(e) }];
  }
}

function normalizeCode(code) {
  return String(code).toUpperCase().replace(/\s+/g, '');
}
