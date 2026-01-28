/**
 * 家具配送センター 業務日報アプリ - サーバーサイド
 */

const SHEET_NAME = '配送データ';

/**
 * Webアプリとしてデプロイ時のエントリポイント
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('家具配送センター 業務日報')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * スプレッドシートを取得（なければ作成）
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = ['日付', '伝票No', '顧客名', '住所', '金額', '配送方法', '支払方法'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * CSVテキストをパースして保存する
 * @param {string} csvText - CSVテキスト
 * @return {Object} 結果オブジェクト
 */
function importCsvData(csvText) {
  try {
    const rows = parseCsv(csvText);
    if (rows.length === 0) {
      return { success: false, message: 'CSVデータが空です。' };
    }

    const sheet = getOrCreateSheet();
    const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');

    // ヘッダー行をスキップするか判定
    let dataRows = rows;
    const firstRow = rows[0].map(function(cell) { return cell.trim(); });
    const headerKeywords = ['伝票', '顧客', '住所', '金額', '配送', '支払'];
    const isHeader = headerKeywords.some(function(kw) {
      return firstRow.some(function(cell) { return cell.indexOf(kw) !== -1; });
    });
    if (isHeader) {
      dataRows = rows.slice(1);
    }

    if (dataRows.length === 0) {
      return { success: false, message: 'データ行がありません（ヘッダーのみ）。' };
    }

    // 各行を保存（日付を先頭に付与）
    var saveData = dataRows.map(function(row) {
      // CSVの列順: 伝票No, 顧客名, 住所, 金額, 配送方法, 支払方法
      var paddedRow = row.concat(Array(6).fill('')).slice(0, 6);
      return [today].concat(paddedRow);
    });

    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, saveData.length, saveData[0].length).setValues(saveData);

    return {
      success: true,
      message: saveData.length + '件のデータを保存しました。',
      data: saveData
    };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

/**
 * CSVテキストをパースする
 * @param {string} text - CSVテキスト
 * @return {Array} 2次元配列
 */
function parseCsv(text) {
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
            current += '"';
            j++;
          } else {
            inQuotes = false;
          }
        } else {
          current += ch;
        }
      } else {
        if (ch === '"') {
          inQuotes = true;
        } else if (ch === ',') {
          cells.push(current);
          current = '';
        } else {
          current += ch;
        }
      }
    }
    cells.push(current);
    rows.push(cells);
  }
  return rows;
}

/**
 * 保存済みデータを取得する
 * @param {string} date - 日付フィルタ（省略時は今日）
 * @return {Object} データ
 */
function getSavedData(date) {
  try {
    var sheet = getOrCreateSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, data: [], message: 'データがありません。' };
    }

    var allData = sheet.getRange(2, 1, lastRow - 1, 7).getValues();

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

    return { success: true, data: allData };
  } catch (e) {
    return { success: false, data: [], message: 'エラー: ' + e.message };
  }
}

/**
 * データを削除する
 * @param {number} rowIndex - 削除する行インデックス（シート上の行番号 - 2）
 * @return {Object} 結果
 */
function deleteRow(rowIndex) {
  try {
    var sheet = getOrCreateSheet();
    sheet.deleteRow(rowIndex + 2); // ヘッダー行 + 0-indexed offset
    return { success: true, message: '削除しました。' };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}
