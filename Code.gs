/**
 * 横浜高島屋 販売商品リスト — 在庫管理 Apps Script
 *
 * シート構成:
 *   商品一覧   … 商品マスタ（親データ）
 *   販売データ … 売上ログ
 *   売上ボタン … 操作UI（図形ボタン配置用）
 */

/* ===== 定数 ===== */
var SHEET_PRODUCTS = '商品一覧';
var SHEET_SALES    = '販売データ';
var TEMPLATE_ROW   = 2; // 販売データのテンプレ行

/* ===== メニュー / サイドバー ===== */

/**
 * スプレッドシートを開いたときにカスタムメニューを追加する。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('売上')
    .addItem('売上パネルを開く', 'showSalesSidebar')
    .addToUi();
}

/**
 * 売上登録用サイドバーを表示する。
 * 売上ボタンシートの図形にもこの関数を割り当てる。
 */
function showSalesSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('売上登録');
  SpreadsheetApp.getUi().showSidebar(html);
}

/* ===== ヘッダー解決ヘルパー ===== */

/**
 * ヘッダー配列から候補名に一致する列インデックス（0始まり）を返す。
 * 見つからなければ -1。
 * @param {string[]} headers
 * @param {string[]} candidates - 同義の候補名リスト
 * @return {number}
 */
function findColumnIndex_(headers, candidates) {
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim();
    for (var j = 0; j < candidates.length; j++) {
      if (h === candidates[j]) return i;
    }
  }
  return -1;
}

/**
 * ヘッダー行からマッピングオブジェクトを生成する。
 * @param {string[]} headers
 * @param {Object} spec - { key: [候補名, ...], ... }
 * @return {Object} { key: colIndex(0始まり), ... }  見つからないキーは -1
 */
function resolveHeaders_(headers, spec) {
  var map = {};
  for (var key in spec) {
    map[key] = findColumnIndex_(headers, spec[key]);
  }
  return map;
}

/* ===== 商品一覧 ===== */

/** 商品一覧ヘッダー候補 */
var PRODUCT_HEADER_SPEC = {
  code:    ['商品コード'],
  series:  ['シリーズ'],
  size:    ['サイズ'],
  name:    ['商品名'],
  price:   ['金額(本体)', '金額（本体）', '金額'],
  gender:  ["Men's/Women's", "Men's/Women's"],
  stock:   ['残数'],
  remarks: ['備考']
};

/**
 * 商品一覧シートから全商品データを取得する。
 * サイドバーから呼ばれる公開関数。
 * @return {Object[]}
 */
function getProductList() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_PRODUCTS);
  if (!sheet) throw new Error('シート「' + SHEET_PRODUCTS + '」が見つかりません。');

  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var col     = resolveHeaders_(headers, PRODUCT_HEADER_SPEC);

  if (col.code === -1) throw new Error('商品一覧に「商品コード」列がありません。');
  if (col.name === -1) throw new Error('商品一覧に「商品名」列がありません。');

  var products = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var code = String(row[col.code]).trim();
    if (!code) continue; // 空行スキップ

    products.push({
      code:    code,
      name:    col.name    !== -1 ? String(row[col.name])    : '',
      series:  col.series  !== -1 ? String(row[col.series])  : '',
      size:    col.size    !== -1 ? String(row[col.size])    : '',
      price:   col.price   !== -1 ? Number(row[col.price])   : 0,
      gender:  col.gender  !== -1 ? String(row[col.gender])  : '',
      stock:   col.stock   !== -1 ? Number(row[col.stock])   : 0,
      remarks: col.remarks !== -1 ? String(row[col.remarks]) : '',
      row:     i + 1 // シート上の行番号（1始まり）
    });
  }
  return products;
}

/* ===== 売上登録 ===== */

/** 販売データヘッダー候補 */
var SALES_HEADER_SPEC = {
  datetime: ['販売時刻', '売上時刻', '日時'],
  code:     ['商品コード'],
  series:   ['シリーズ'],
  size:     ['サイズ'],
  name:     ['商品名'],
  price:    ['金額(本体)', '金額（本体）', '金額'],
  gender:   ["Men's/Women's", "Men's/Women's"],
  stock:    ['残数'],
  remarks:  ['備考']
};

/**
 * 商品コードと数量を受け取り、販売データに追記＋在庫減算する。
 * @param {string} productCode
 * @param {number} [quantity=1]
 * @return {Object} { success, newStock, productName, message }
 */
function recordSaleByCode(productCode, quantity) {
  quantity = Math.max(1, Math.floor(Number(quantity) || 1));

  var lock = LockService.getDocumentLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: '他のユーザーが操作中です。少し待ってから再度お試しください。' };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    /* --- 商品一覧から対象商品を取得 --- */
    var prodSheet   = ss.getSheetByName(SHEET_PRODUCTS);
    if (!prodSheet) throw new Error('シート「' + SHEET_PRODUCTS + '」が見つかりません。');

    var prodData    = prodSheet.getDataRange().getValues();
    var prodHeaders = prodData[0];
    var prodCol     = resolveHeaders_(prodHeaders, PRODUCT_HEADER_SPEC);

    if (prodCol.code  === -1) throw new Error('商品一覧に「商品コード」列がありません。');
    if (prodCol.stock === -1) throw new Error('商品一覧に「残数」列がありません。');

    var targetRowIndex = -1; // data配列上のインデックス
    for (var i = 1; i < prodData.length; i++) {
      if (String(prodData[i][prodCol.code]).trim() === String(productCode).trim()) {
        targetRowIndex = i;
        break;
      }
    }
    if (targetRowIndex === -1) {
      return { success: false, message: '商品コード「' + productCode + '」が見つかりません。' };
    }

    var productRow = prodData[targetRowIndex];
    var currentStock = Number(productRow[prodCol.stock]) || 0;

    if (currentStock < quantity) {
      return {
        success: false,
        message: '在庫が不足しています。（残数: ' + currentStock + '、要求: ' + quantity + '）'
      };
    }

    /* --- 在庫減算 --- */
    var newStock = currentStock - quantity;
    var stockCell = prodSheet.getRange(targetRowIndex + 1, prodCol.stock + 1);
    stockCell.setValue(newStock);

    /* --- 販売データに追記 --- */
    var salesSheet = ss.getSheetByName(SHEET_SALES);
    if (!salesSheet) throw new Error('シート「' + SHEET_SALES + '」が見つかりません。');

    var salesHeaders = salesSheet.getRange(1, 1, 1, salesSheet.getLastColumn()).getValues()[0];
    var salesCol     = resolveHeaders_(salesHeaders, SALES_HEADER_SPEC);

    var newRowNum = Math.max(salesSheet.getLastRow() + 1, TEMPLATE_ROW + 1);

    // テンプレ行から書式・入力規則をコピー
    copyTemplateFormat_(salesSheet, TEMPLATE_ROW, newRowNum);

    // 書き込み値を配列で構築
    var newRow = new Array(salesHeaders.length);
    for (var c = 0; c < newRow.length; c++) newRow[c] = '';

    if (salesCol.datetime !== -1) newRow[salesCol.datetime] = new Date();
    if (salesCol.code     !== -1) newRow[salesCol.code]     = String(productRow[prodCol.code]);
    if (salesCol.series   !== -1 && prodCol.series  !== -1) newRow[salesCol.series]  = productRow[prodCol.series];
    if (salesCol.size     !== -1 && prodCol.size    !== -1) newRow[salesCol.size]    = productRow[prodCol.size];
    if (salesCol.name     !== -1 && prodCol.name    !== -1) newRow[salesCol.name]    = productRow[prodCol.name];
    if (salesCol.price    !== -1 && prodCol.price   !== -1) newRow[salesCol.price]   = productRow[prodCol.price];
    if (salesCol.gender   !== -1 && prodCol.gender  !== -1) newRow[salesCol.gender]  = productRow[prodCol.gender];
    if (salesCol.stock    !== -1) newRow[salesCol.stock] = newStock; // 減算後の残数
    if (salesCol.remarks  !== -1 && prodCol.remarks !== -1) newRow[salesCol.remarks] = productRow[prodCol.remarks];

    salesSheet.getRange(newRowNum, 1, 1, newRow.length).setValues([newRow]);

    var productName = prodCol.name !== -1 ? String(productRow[prodCol.name]) : productCode;

    return {
      success: true,
      newStock: newStock,
      productName: productName,
      message: '「' + productName + '」を' + quantity + '点登録しました。（残数: ' + newStock + '）'
    };

  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/* ===== テンプレ行コピー ===== */

/**
 * テンプレ行（通常2行目）の書式と入力規則を対象行にコピーする。
 * @param {Sheet} sheet
 * @param {number} templateRow - テンプレ行番号（1始まり）
 * @param {number} targetRow   - 対象行番号（1始まり）
 */
function copyTemplateFormat_(sheet, templateRow, targetRow) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;

  var srcRange = sheet.getRange(templateRow, 1, 1, lastCol);
  var dstRange = sheet.getRange(targetRow,   1, 1, lastCol);

  srcRange.copyTo(dstRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  srcRange.copyTo(dstRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
}
