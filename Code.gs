/**
 * 家具配送センター 業務日報アプリ - サーバーサイド
 * Excelから全列をそのまま貼り付けて使える
 * 配送順データ（ルート）にも対応
 */

var SHEET_NAME = '配送データ';
var ROUTE_SHEET_NAME = '配送順データ';
var REPORT_SHEET_NAME = '配送報告データ';

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

function getOrCreateReportSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(REPORT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(REPORT_SHEET_NAME);
    var headers = ['伝票No', '玄関先お渡し', '吊り作業', 'CS個数', '備考'];
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

// --- 利用可能な日付一覧を取得 ---

function getAvailableDates() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(ROUTE_SHEET_NAME);
    if (!sheet) return { success: true, dates: [] };

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, dates: [] };

    var data = sheet.getRange(2, 3, lastRow - 1, 1).getValues(); // 配送日列
    var dateSet = {};
    data.forEach(function(row) {
      var code = String(row[0]).trim();
      var ymd = deliveryCodeToYMD(code);
      if (ymd && ymd.match(/^\d{4}\/\d{2}\/\d{2}$/)) {
        dateSet[ymd] = true;
      }
    });
    var dates = Object.keys(dateSet).sort().reverse();
    return { success: true, dates: dates };
  } catch (e) {
    return { success: false, dates: [], message: e.message };
  }
}

// --- 指定日付の号車一覧を取得 ---

function getTrucksForDate(date) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(ROUTE_SHEET_NAME);
    if (!sheet) return { success: true, trucks: [] };

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, trucks: [] };

    var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    var truckSet = {};
    data.forEach(function(row) {
      var code = String(row[2]).trim();
      var ymd = deliveryCodeToYMD(code);
      if (ymd === date) {
        var truck = String(row[5]).trim();
        if (truck) truckSet[truck] = true;
      }
    });
    var trucks = Object.keys(truckSet).sort(function(a, b) {
      return parseInt(a, 10) - parseInt(b, 10);
    });
    return { success: true, trucks: trucks };
  } catch (e) {
    return { success: false, trucks: [], message: e.message };
  }
}

// --- 日付・号車でフィルタしたデータを取得 ---

function getFilteredData(date, truck) {
  try {
    var sheet = getOrCreateSheet();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol === 0) {
      return { success: true, data: [] };
    }

    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
      return String(h);
    });
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

    // 配送順データを取得
    var routeMap = getRouteMap();

    // 報告データを取得
    var reportMap = getReportMap();

    // 列インデックス取得（複数キーワードでフォールバック）
    var slipIdx = findColIdx(headers, '伝票');
    if (slipIdx === -1) {
      // ルートマップのキー（伝票No）と一致する値を持つ列を探す
      var routeKeys = Object.keys(routeMap);
      if (routeKeys.length > 0 && allData.length > 0) {
        var sampleKey = routeKeys[0];
        for (var ci = 0; ci < headers.length; ci++) {
          for (var ri = 0; ri < Math.min(5, allData.length); ri++) {
            if (String(allData[ri][ci]).trim() === sampleKey) {
              slipIdx = ci;
              break;
            }
          }
          if (slipIdx !== -1) break;
        }
      }
    }
    var addressIdx = findColIdx(headers, '住所');
    var amountIdx = findColIdx(headers, '金額');
    var paymentIdx = findColIdx(headers, '支払');

    // デバッグ用
    var debugInfo = {
      headers: headers.join(' | '),
      dataCount: allData.length,
      slipIdx: slipIdx,
      routeMapSize: Object.keys(routeMap).length,
      sampleSlips: [],
      sampleRouteKeys: Object.keys(routeMap).slice(0, 5)
    };

    if (slipIdx !== -1 && allData.length > 0) {
      for (var di = 0; di < Math.min(3, allData.length); di++) {
        debugInfo.sampleSlips.push(String(allData[di][slipIdx]));
      }
    }

    var result = [];

    allData.forEach(function(row) {
      var slipNo = slipIdx !== -1 ? String(row[slipIdx]).trim() : '';
      var route = routeMap[slipNo];

      if (!route) return;
      if (route.deliveryDate !== date) return;
      if (String(route.truck) !== String(truck)) return;

      var address = addressIdx !== -1 ? String(row[addressIdx]).trim() : '';
      var amount = amountIdx !== -1 ? String(row[amountIdx]) : '';
      var payment = paymentIdx !== -1 ? String(row[paymentIdx]).trim() : '';

      // 住所を分解
      var parsed = parseAddress(address);

      // 残金表示（支払方法が1なら着払）
      var zankin = (payment === '1') ? '着払' : '';

      // 報告データ
      var report = reportMap[slipNo] || {};

      result.push({
        seq: parseInt(route.seq, 10) || 999,
        slipNo: slipNo,
        city: parsed.city,
        area: parsed.area,
        amount: amount,
        zankin: zankin,
        genkan: report.genkan || '',
        tsuri: report.tsuri || '',
        cs: report.cs || '',
        biko: report.biko || ''
      });
    });

    // 何件目でソート
    result.sort(function(a, b) {
      return a.seq - b.seq;
    });

    return { success: true, data: result, debug: debugInfo };
  } catch (e) {
    return { success: false, data: [], message: 'エラー: ' + e.message };
  }
}

// --- 住所を分解 ---

function parseAddress(address) {
  if (!address) return { city: '', area: '' };

  // 市区町村を抽出（〇〇市、〇〇区、〇〇町、〇〇村）
  var cityMatch = address.match(/^(.+?[市区町村])/);
  var city = cityMatch ? cityMatch[1] : '';

  // 市区町村より後から番地（数字）までを抽出
  var afterCity = city ? address.substring(city.length) : address;
  var areaMatch = afterCity.match(/^([^\d０-９]+)/);
  var area = areaMatch ? areaMatch[1].trim() : '';

  return { city: city, area: area };
}

// --- 報告データのマップを取得 ---

function getReportMap() {
  var map = {};
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(REPORT_SHEET_NAME);
    if (!sheet) return map;

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return map;

    var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    data.forEach(function(row) {
      var slipNo = String(row[0]).trim();
      map[slipNo] = {
        genkan: String(row[1]),
        tsuri: String(row[2]),
        cs: String(row[3]),
        biko: String(row[4])
      };
    });
  } catch (e) {}
  return map;
}

// --- 報告データを保存 ---

function saveReportData(items) {
  try {
    var sheet = getOrCreateReportSheet();

    // 既存データを削除して新しいデータで置き換え
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      // 既存データを取得
      var existingData = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      var existingMap = {};
      existingData.forEach(function(row, idx) {
        existingMap[String(row[0]).trim()] = idx + 2; // 行番号
      });

      // 更新・追加
      items.forEach(function(item) {
        var slipNo = String(item.slipNo).trim();
        var rowData = [slipNo, item.genkan || '', item.tsuri || '', item.cs || '', item.biko || ''];

        if (existingMap[slipNo]) {
          // 既存行を更新
          sheet.getRange(existingMap[slipNo], 1, 1, 5).setValues([rowData]);
        } else {
          // 新規追加
          sheet.appendRow(rowData);
        }
      });
    } else {
      // データがない場合は全部追加
      items.forEach(function(item) {
        var slipNo = String(item.slipNo).trim();
        var rowData = [slipNo, item.genkan || '', item.tsuri || '', item.cs || '', item.biko || ''];
        sheet.appendRow(rowData);
      });
    }

    return { success: true, message: items.length + '件の報告データを保存しました。' };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

// --- 配送順データを更新（何件目の変更など）---

function updateRouteSeq(slipNo, newSeq) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(ROUTE_SHEET_NAME);
    if (!sheet) return { success: false, message: 'シートがありません' };

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: 'データがありません' };

    var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(slipNo).trim()) {
        sheet.getRange(i + 2, 7).setValue(newSeq); // 何件目は7列目
        return { success: true };
      }
    }
    return { success: false, message: '伝票Noが見つかりません' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// --- 保存済みデータ取得（旧版・互換用）---

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

    allData = allData.map(function(row) {
      return row.map(function(cell) {
        if (cell instanceof Date) {
          return Utilities.formatDate(cell, 'Asia/Tokyo', 'yyyy/MM/dd');
        }
        return cell;
      });
    });

    var routeMap = getRouteMap();
    var slipIdx = -1;
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i]).indexOf('伝票') !== -1) { slipIdx = i; break; }
    }

    if (slipIdx !== -1 && Object.keys(routeMap).length > 0) {
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

      if (date) {
        var deliveryDateColIdx = headers.length - 1;
        allData = allData.filter(function(row) {
          var dd = String(row[deliveryDateColIdx]).trim();
          return dd === date;
        });
      }
    } else {
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
  } catch (e) {}
  return map;
}

/**
 * 配送日コードをyyyy/MM/dd形式に変換
 * 例: 80120 → 令和8年1月20日 → 2026/01/20
 */
function deliveryCodeToYMD(code) {
  if (!code || String(code).length < 5) return String(code);
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
    return s;
  }
  var reiwaYear = parseInt(yearStr, 10);
  var adYear = 2018 + reiwaYear;
  return adYear + '/' + monthStr + '/' + dayStr;
}

/**
 * ヘッダーからキーワードを含む列のインデックスを返す
 */
function findColIdx(headers, keyword) {
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).indexOf(keyword) !== -1) return i;
  }
  return -1;
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
