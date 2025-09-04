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
    const MAX_BLOCK_ROWS = 500; // защита от "бесконечного" прохода

    const results = [];

    for (const sheetName of sheetsToSearch) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) continue;

      const lastRow = sheet.getLastRow();
      if (lastRow < 1) continue;

      // 1) Ищем код ТОЛЬКО в колонке A
      const colA = sheet.getRange(1, 1, lastRow, 1).getValues();
      let foundRow = null; // 1-based

      for (let i = 0; i < colA.length; i++) {
        const v = colA[i][0];
        if (v === '' || v === null) continue;

        // Берём первый токен вида "721" или "A12.3"
        const m = String(v).trim().match(/^[A-Za-zА-Яа-я0-9.]+/);
        if (!m) continue;

        if (normalizeCode(m[0]) === searchCode) {
          foundRow = i + 1; // в Apps Script индексация с 1
          break;
        }
      }

      if (!foundRow) continue;

      // 2) Собираем блок A+B от найденной строки до первой полностью пустой строки
      const tail = sheet.getRange(foundRow, 1, lastRow - foundRow + 1, 2).getValues();

      let length = 0;
      for (let i = 0; i < tail.length && i < MAX_BLOCK_ROWS; i++) {
        const a = tail[i][0];
        const b = tail[i][1];
        const isEmptyRow = (a === '' || a === null) && (b === '' || b === null);

        // Останавливаемся на первой пустой строке ПОСЛЕ того как захватили хотя бы заголовок
        if (isEmptyRow && i > 0) break;

        // Если прямо первая строка пустая (редкий случай) — всё равно возьмём её как 1 строку
        if (isEmptyRow && i === 0) { length = 1; break; }

        length++;
      }

      if (length === 0) length = 1; // на всякий случай

      const endRow = foundRow + length - 1;
      const values = tail.slice(0, length);
      const data = values.map(r => ({ code: r[0], name: r[1] }));

      const link = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheet.getSheetId()}&range=${foundRow}:${endRow}`;

      results.push({
        link,
        sheet: sheetName,
        startRow: foundRow,
        rows: length,
        cols: 2,
        data
      });

      // Если коды уникальные — можно сразу выходить. Если нужно искать во всех листах — убери break.
      break;
    }

    if (results.length === 0) {
      return [{
        error: `Код "${code}" не найден. Проверь, что:
- он находится в колонке A,
- листы называются "часть 1" и/или "часть 2",
- между группами есть пустая строка (как ты описала).`
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
