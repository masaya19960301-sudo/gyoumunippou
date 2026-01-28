/**
 * 家具配送センター 業務日報アプリ - サーバーサイド
 * Excelから全列をそのまま貼り付けて使える
 */

var SHEET_NAME = '配送データ';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('家具配送センター 業務日報')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * スプレッドシートを取得（なければ作成）
 */
function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 貼り付けデータをパースして保存する
 */
function importCsvData(csvText) {
  try {
    var rows = parseText(csvText);
    if (rows.length === 0) {
      return { success: false, message: 'データが空です。' };
    }

    var sheet = getOrCreateSheet();
    var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');

    // ヘッダー行を判定
    var headers = null;
    var dataRows = rows;
    var firstRow = rows[0].map(function(c) { return c.trim(); });
    var headerKeywords = ['No', '号車', '伝票', '顧客', '住所', '金額', '配送', '支払', '方法', '店'];
    var matchCount = 0;
    headerKeywords.forEach(function(kw) {
      firstRow.forEach(function(cell) {
        if (cell.indexOf(kw) !== -1) matchCount++;
      });
    });
    if (matchCount >= 2) {
      headers = firstRow;
      dataRows = rows.slice(1);
    }

    if (dataRows.length === 0) {
      return { success: false, message: 'データ行がありません（ヘッダーのみ）。' };
    }

    // 列数を揃える（最大列数に合わせる）
    var maxCols = 0;
    dataRows.forEach(function(row) {
      if (row.length > maxCols) maxCols = row.length;
    });

    // シートにヘッダーがなければ設定する（日付列 + 元の列）
    var lastRow = sheet.getLastRow();
    var totalCols = maxCols + 1; // +1 for 日付
    if (lastRow === 0) {
      var sheetHeaders = ['日付'];
      if (headers) {
        sheetHeaders = sheetHeaders.concat(headers);
      } else {
        for (var i = 0; i < maxCols; i++) {
          sheetHeaders.push('列' + (i + 1));
        }
      }
      // 列数を揃える
      while (sheetHeaders.length < totalCols) sheetHeaders.push('');
      sheet.getRange(1, 1, 1, totalCols).setValues([sheetHeaders]);
      sheet.getRange(1, 1, 1, totalCols).setFontWeight('bold');
      lastRow = 1;
    }

    // 各行を保存（日付を先頭に付与）
    var saveData = dataRows.map(function(row) {
      var padded = row.concat(Array(maxCols).fill('')).slice(0, maxCols);
      return [today].concat(padded);
    });

    sheet.getRange(lastRow + 1, 1, saveData.length, totalCols).setValues(saveData);

    return {
      success: true,
      message: saveData.length + '件のデータを保存しました。'
    };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

/**
 * テキストをパース（CSV/TSV自動判定）
 */
function parseText(text) {
  var firstLine = text.split(/\r?\n/)[0] || '';
  var delimiter = (firstLine.indexOf('\t') !== -1) ? '\t' : ',';
  var rows = [];
  var lines = text.split(/\r?\n/);
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (line === '') continue;
    var cells = [];
    var current = '';
    var inQuotes = false;
    for (var j = 0; j < line.length; j++) {
      var ch = line[j];
      if (inQuotes) {
        if (ch === '"') {
          if (j + 1 < line.length && line[j + 1] === '"') {
            current += '"'; j++;
          } else {
            inQuotes = false;
          }
        } else {
          current += ch;
        }
      } else {
        if (ch === '"') { inQuotes = true; }
        else if (ch === delimiter) { cells.push(current); current = ''; }
        else { current += ch; }
      }
    }
    cells.push(current);
    rows.push(cells);
  }
  return rows;
}

/**
 * 保存済みデータをヘッダー付きで取得
 */
function getSavedData(date) {
  try {
    var sheet = getOrCreateSheet();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol === 0) {
      return { success: true, headers: [], data: [] };
    }

    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var allData = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    if (date) {
      allData = allData.filter(function(row) {
        var rowDate = '';
        if (row[0] instanceof Date) {
          rowDate = Utilities.formatDate(row[0], 'Asia/Tokyo', 'yyyy/MM/dd');
        } else {
          rowDate = String(row[0]);
        }
        return rowDate === date;
      });
    }

    // Date型を文字列に変換
    allData = allData.map(function(row) {
      return row.map(function(cell) {
        if (cell instanceof Date) {
          return Utilities.formatDate(cell, 'Asia/Tokyo', 'yyyy/MM/dd');
        }
        return cell;
      });
    });

    return { success: true, headers: headers, data: allData };
  } catch (e) {
    return { success: false, headers: [], data: [], message: 'エラー: ' + e.message };
  }
}

/**
 * データ行を削除
 */
function deleteRow(rowIndex) {
  try {
    var sheet = getOrCreateSheet();
    sheet.deleteRow(rowIndex + 2);
    return { success: true, message: '削除しました。' };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}
