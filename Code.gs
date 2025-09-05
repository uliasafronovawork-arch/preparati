function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ICD Search')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function searchICDCodePublic(query) {
  try {
    const searchQuery = String(query || '').trim();
    if (!searchQuery) {
      return [{ error: "Введите код или название заболевания" }];
    }

    const spreadsheetId = '1zS0DoyXg6GE0rsOr1NLkDdljplc1GXU1cNVCZRgS1rQ';
    const ss = SpreadsheetApp.openById(spreadsheetId);

    const results = [];

    function cleanSheetText(str) {
      if (!str) return '';
      return String(str).replace(/[\u00A0]+/g, ' ').trim();
    }

    // === 1. Поиск по кодам (лист "часть 1") ===
    (function () {
      const sheet = ss.getSheetByName('часть 1');
      if (!sheet) return;

      const lastRow = sheet.getLastRow();
      if (lastRow < 1) return;

      const all = sheet.getRange(1, 1, lastRow, 2).getValues();

      let foundRow = null;
      const normCode = normalizeCode(searchQuery);

      for (let r = 0; r < lastRow; r++) {
        const val = all[r][0]; // колонка A
        if (!val) continue;

        const match = String(val).trim().match(/^[A-Za-zА-Яа-я0-9.]+/);
        if (!match) continue;

        const cellCode = normalizeCode(match[0]);
        if (cellCode === normCode) {
          foundRow = r + 1;
          break;
        }
      }

      if (foundRow) {
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

        const link = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheet.getSheetId()}&range=A${foundRow}:B${endRow - 1}`;

        results.push({
          link,
          sheet: 'часть 1',
          startRow: foundRow,
          rows: blockValues.length,
          cols: 2,
          data: blockValues
        });
      }
    })();

    // === 2. Поиск по словам (лист "часть 0") ===
    if (results.length === 0) {
      const sheet = ss.getSheetByName('часть 0');
      if (sheet) {
        const lastRow = sheet.getLastRow();
        if (lastRow > 0) {
          const all = sheet.getRange(1, 1, lastRow, 2).getValues();
          let foundRow = null;

          for (let r = 0; r < lastRow; r++) {
            const valA = cleanSheetText(all[r][0]);
            const valB = cleanSheetText(all[r][1]);
            if (isFuzzyMatch(valA, searchQuery) || isFuzzyMatch(valB, searchQuery)) {
              foundRow = r + 1;
              break;
            }
          }

          if (foundRow) {
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

            const link = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheet.getSheetId()}&range=A${foundRow}:B${endRow - 1}`;

            results.push({
              link,
              sheet: 'часть 0',
              startRow: foundRow,
              rows: blockValues.length,
              cols: 2,
              data: blockValues
            });
          }
        }
      }
    }

    if (results.length === 0) {
      return [{
        error: `По запросу "${query}" ничего не найдено. Проверьте листы "часть 1" или "часть 2".`
      }];
    }

    return results;

  } catch (e) {
    return [{ error: e && e.message ? e.message : String(e) }];
  }
}

// нормализация кода (для листа 1)
function normalizeCode(code) {
  return String(code).toUpperCase().replace(/\s+/g, '');
}

// ==== Умный поиск для листа 0 ====
function isFuzzyMatch(text, query) {
  if (!text || !query) return false;

  text = text.toLowerCase().replace(/\s+/g, ' ').trim();
  query = query.toLowerCase().replace(/\s+/g, ' ').trim();

  if (text.includes(query)) return true; // прямое вхождение

  // fuzzy по словам
  const words = text.split(' ');
  return words.some(word => levenshteinDistance(word, query) <= 2); // допускаем 2 ошибки
}

// Расстояние Левенштейна
function levenshteinDistance(a, b) {
  const m = [];
  for (let i = 0; i <= b.length; i++) m[i] = [i];
  for (let j = 0; j <= a.length; j++) m[0][j] = j;

  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      m[i][j] = b.charAt(i - 1) === a.charAt(j - 1)
        ? m[i - 1][j - 1]
        : Math.min(m[i - 1][j - 1] + 1, m[i][j - 1] + 1, m[i - 1][j] + 1);
    }
  }
  return m[b.length][a.length];
}
