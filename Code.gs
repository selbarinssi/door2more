/**
 * SALES VALIDATION — Apps Script backend
 *
 * Bind this script to the Google Sheet that has 4 tabs:
 *   Config     : SellerIDs | SellerName | PaymentMethods | ServiceNames  (header row 1, data from row 2)
 *   Catalog    : ItemID | ItemName | Price | Volume                     (header row 1, data from row 2)
 *   Sales      : SaleID | Timestamp | SellerID | CustomerName | CustomerPhone | CustomerAddress |
 *                OrderNumber | PaymentMethod | ServiceName | TotalAmount | TotalVolume
 *   SaleItems  : SaleID | ItemID | ItemName | Quantity | UnitPrice | LineTotal | UnitVolume | LineVolume
 *
 * Deploy: Extensions > Apps Script > paste this file > Deploy > New deployment
 *         Type: Web app | Execute as: Me | Who has access: Anyone
 *         Copy the /exec URL into SCRIPT_URL in index.html
 */

// Must match SHARED_TOKEN in index.html — change this to your own value.
var SHARED_TOKEN = 'RINCIGROUP';

var SHEET_CONFIG = 'Config';
var SHEET_CATALOG = 'Catalog';
var SHEET_SALES = 'Sales';
var SHEET_SALE_ITEMS = 'SaleItems';

function doGet(e) {
  var action = e.parameter.action;
  try {
    if (action === 'config') return jsonOutput(getConfig());
    if (action === 'search') return jsonOutput(searchCatalog(e.parameter.q || ''));
    return jsonOutput({ error: 'Unknown action' });
  } catch (err) {
    return jsonOutput({ error: err.message });
  }
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    if (payload.token !== SHARED_TOKEN) {
      return jsonOutput({ success: false, error: 'Invalid token' });
    }
    var result = recordSale(payload);
    return jsonOutput(result);
  } catch (err) {
    return jsonOutput({ success: false, error: err.message });
  }
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- CONFIG ----------

function getConfig() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG);
  var lastRow = sheet.getLastRow();

  var sellerIds = getColumn(sheet, 1, lastRow);
  var sellerNames = getColumn(sheet, 2, lastRow);
  var paymentMethods = getColumn(sheet, 3, lastRow);
  var serviceNames = getColumn(sheet, 4, lastRow);

  var sellers = [];
  for (var i = 0; i < sellerIds.length; i++) {
    if (sellerIds[i] === '') continue;
    sellers.push({ id: sellerIds[i], name: sellerNames[i] || '' });
  }

  return {
    sellers: sellers,
    paymentMethods: paymentMethods.filter(function (v) { return v !== ''; }),
    serviceNames: serviceNames.filter(function (v) { return v !== ''; })
  };
}

function getColumn(sheet, colIndex, lastRow) {
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, colIndex, lastRow - 1, 1).getValues();
  return values.map(function (r) {
    return (r[0] === null || r[0] === undefined) ? '' : String(r[0]).trim();
  });
}

// ---------- CATALOG SEARCH ----------
// Uses TextFinder (built-in sheet search) instead of loading the whole
// catalog into memory, so this stays fast even at 40,000+ rows.

function searchCatalog(query) {
  query = (query || '').trim();
  if (query.length < 2) return { items: [] };

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CATALOG);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { items: [] };

  var dataRange = sheet.getRange(2, 1, lastRow - 1, 4); // ItemID, ItemName, Price, Volume
  var finder = dataRange.createTextFinder(query).matchCase(false).useRegularExpression(false);
  var matches = finder.findAll();

  var rowsSeen = {};
  var items = [];
  for (var i = 0; i < matches.length && items.length < 25; i++) {
    var row = matches[i].getRow();
    if (rowsSeen[row]) continue;
    rowsSeen[row] = true;
    var rowValues = sheet.getRange(row, 1, 1, 4).getValues()[0];
    items.push({
      id: String(rowValues[0]),
      name: String(rowValues[1]),
      price: Number(rowValues[2]) || 0,
      volume: Number(rowValues[3]) || 0
    });
  }
  return { items: items };
}

// ---------- RECORD SALE ----------

function recordSale(payload) {
  var required = ['sellerId', 'customerName', 'customerPhone', 'customerAddress',
                   'orderNumber', 'paymentMethod', 'serviceName'];
  for (var i = 0; i < required.length; i++) {
    if (!payload[required[i]]) {
      return { success: false, error: 'Missing field: ' + required[i] };
    }
  }
  if (!payload.items || !payload.items.length) {
    return { success: false, error: 'No items in sale' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var salesSheet = ss.getSheetByName(SHEET_SALES);
  var itemsSheet = ss.getSheetByName(SHEET_SALE_ITEMS);

  var saleId = Utilities.getUuid();
  var timestamp = new Date();

  var totalAmount = 0;
  var totalVolume = 0;
  var itemRows = [];

  payload.items.forEach(function (item) {
    var qty = Number(item.qty) || 0;
    var unitPrice = Number(item.unitPrice) || 0;
    var unitVolume = Number(item.unitVolume) || 0;
    var lineTotal = qty * unitPrice;
    var lineVolume = qty * unitVolume;
    totalAmount += lineTotal;
    totalVolume += lineVolume;
    itemRows.push([saleId, item.itemId, item.itemName, qty, unitPrice, lineTotal, unitVolume, lineVolume]);
  });

  salesSheet.appendRow([
    saleId, timestamp, payload.sellerId, payload.customerName, payload.customerPhone,
    payload.customerAddress, payload.orderNumber, payload.paymentMethod, payload.serviceName,
    totalAmount, totalVolume
  ]);

  itemRows.forEach(function (row) {
    itemsSheet.appendRow(row);
  });

  return { success: true, saleId: saleId, totalAmount: totalAmount, totalVolume: totalVolume };
}
