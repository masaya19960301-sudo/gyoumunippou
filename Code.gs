/**
 * 家具配送センター 業務日報アプリ - サーバーサイド
 * Excelから全列をそのまま貼り付けて使える
 * 配送順データ（ルート）にも対応
 */

var SHEET_NAME = '配送データ';
var ROUTE_SHEET_NAME = '配送順データ';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('家具配送センター 業務日報')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- シート取得 ---

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateRouteSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ROUTE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ROUTE_SHEET_NAME);
    var headers = ['日付', '伝票No', '配送日', '電話番号', '店コード', '号車', '何件目', '配送センター', '契約日'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// --- 配送データ（メイン）取込 ---

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

    var maxCols = 0;
    dataRows.forEach(function(row) {
      if (row.length > maxCols) maxCols = row.length;
    });

    var lastRow = sheet.getLastRow();
    var totalCols = maxCols + 1;
    if (lastRow === 0) {
      var sheetHeaders = ['日付'];
      if (headers) {
        sheetHeaders = sheetHeaders.concat(headers);
      } else {
        for (var i = 0; i < maxCols; i++) {
          sheetHeaders.push('列' + (i + 1));
        }
      }
      while (sheetHeaders.length < totalCols) sheetHeaders.push('');
      sheet.getRange(1, 1, 1, totalCols).setValues([sheetHeaders]);
      sheet.getRange(1, 1, 1, totalCols).setFontWeight('bold');
      lastRow = 1;
    }

    var saveData = dataRows.map(function(row) {
      var padded = row.concat(Array(maxCols).fill('')).slice(0, maxCols);
      return [today].concat(padded);
    });

    sheet.getRange(lastRow + 1, 1, saveData.length, totalCols).setValues(saveData);

    return {
      success: true,
      message: saveData.length + '件の配送データを保存しました。'
    };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

// --- 配送順データ取込 ---

function importRouteData(csvText) {
  try {
    var rows = parseText(csvText);
    if (rows.length === 0) {
      return { success: false, message: 'データが空です。' };
    }

    var sheet = getOrCreateRouteSheet();
    var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');

    // 全列保存（日付を先頭に付与、最大8列分）
    var colCount = 8;
    var saveData = rows.map(function(row) {
      var padded = row.concat(Array(colCount).fill('')).slice(0, colCount);
      return [today].concat(padded);
    });

    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, saveData.length, colCount + 1).setValues(saveData);

    return {
      success: true,
      message: saveData.length + '件の配送順データを保存しました。'
    };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

// --- テキストパーサー（CSV/TSV自動判定）---

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

// --- 保存済みデータ取得（配送順データの配送日で絞り込み）---

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

    // Date型を文字列に変換
    allData = allData.map(function(row) {
      return row.map(function(cell) {
        if (cell instanceof Date) {
          return Utilities.formatDate(cell, 'Asia/Tokyo', 'yyyy/MM/dd');
        }
        return cell;
      });
    });

    // 配送順データを全件取得して結合
    var routeMap = getRouteMap();
    var slipIdx = -1;
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i]).indexOf('伝票') !== -1) { slipIdx = i; break; }
    }

    if (slipIdx !== -1 && Object.keys(routeMap).length > 0) {
      // ヘッダーに配送順情報を追加
      headers = headers.concat(['号車', '何件目', '配送日']);

      allData = allData.map(function(row) {
        var slipNo = String(row[slipIdx]).trim();
        var route = routeMap[slipNo];
        if (route) {
          return row.concat([route.truck, route.seq, route.deliveryDate]);
        } else {
          return row.concat(['', '', '']);
        }
      });

      // 配送日（配送順データの配送日）で絞り込む
      if (date) {
        var deliveryDateColIdx = headers.length - 1; // 配送日は最後の列
        allData = allData.filter(function(row) {
          var dd = String(row[deliveryDateColIdx]).trim();
          return dd === date;
        });
      }
    } else {
      // 配送順データがない場合は取込日で絞り込み
      if (date) {
        allData = allData.filter(function(row) {
          return String(row[0]).trim() === date;
        });
      }
    }

    return { success: true, headers: headers, data: allData };
  } catch (e) {
    return { success: false, headers: [], data: [], message: 'エラー: ' + e.message };
  }
}

/**
 * 配送順データをマップとして全件取得（伝票No → {号車, 何件目, 配送日}）
 * 配送日は配送日コードからyyyy/MM/dd形式に変換
 */
function getRouteMap() {
  var map = {};
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(ROUTE_SHEET_NAME);
    if (!sheet) return map;

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return map;

    // 列: 日付(0), 伝票No(1), 配送日(2), 電話番号(3), 店コード(4), 号車(5), 何件目(6), 配送センター(7), 契約日(8)
    var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

    data.forEach(function(row) {
      var slipNo = String(row[1]).trim();
      var rawDate = String(row[2]).trim();
      var deliveryDate = deliveryCodeToYMD(rawDate);

      map[slipNo] = {
        truck: String(row[5]).trim(),
        seq: String(row[6]).trim(),
        deliveryDate: deliveryDate
      };
    });
  } catch (e) {
    // エラー時は空のマップを返す
  }
  return map;
}

/**
 * 配送日コードをyyyy/MM/dd形式に変換
 * 例: 80120 → 令和8年1月20日 → 2026/01/20
 */
function deliveryCodeToYMD(code) {
  if (!code || code.length < 5) return code;
  var s = String(code);
  var yearStr, monthStr, dayStr;
  if (s.length === 5) {
    yearStr = s.substring(0, 1);
    monthStr = s.substring(1, 3);
    dayStr = s.substring(3, 5);
  } else if (s.length === 6) {
    yearStr = s.substring(0, 2);
    monthStr = s.substring(2, 4);
    dayStr = s.substring(4, 6);
  } else {
    return code;
  }
  // 令和 → 西暦 (令和1年 = 2019年)
  var reiwaYear = parseInt(yearStr, 10);
  var adYear = 2018 + reiwaYear;
  return adYear + '/' + monthStr + '/' + dayStr;
}

/**
 * 配送日コードを表示用に変換 (80120 → R8/01/20)
 */
function parseDeliveryDate(code) {
  if (!code || code.length < 5) return code;
  var s = String(code);
  var yearStr, monthStr, dayStr;
  if (s.length === 5) {
    yearStr = s.substring(0, 1);
    monthStr = s.substring(1, 3);
    dayStr = s.substring(3, 5);
  } else if (s.length === 6) {
    yearStr = s.substring(0, 2);
    monthStr = s.substring(2, 4);
    dayStr = s.substring(4, 6);
  } else {
    return code;
  }
  return 'R' + yearStr + '/' + monthStr + '/' + dayStr;
}

// --- データ削除 ---

function deleteRow(rowIndex) {
  try {
    var sheet = getOrCreateSheet();
    sheet.deleteRow(rowIndex + 2);
    return { success: true, message: '削除しました。' };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}
