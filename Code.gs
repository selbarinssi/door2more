/**
 * SALES VALIDATION — Apps Script backend
 *
 * Bind this script to the Google Sheet that has 4 tabs:
 *   Config     : A SellerIDs | B SellerName | C PaymentMethods | D ServiceNames | E ServiceMapping
 *                | F (unused) | G Cities | H..M pricing per city, one column per service
 *                (header row 1, data from row 2)
 *
 *                - Column E gives each ServiceName (col D) a short "mapping code".
 *                - Column G lists delivery cities, one per row.
 *                - Columns H..M hold that city's price for each service in the same row.
 *                  Row 1 (H1..M1) holds the header/code each column corresponds to.
 *                  H1 and I1 are always "Delivery1" and "Delivery2" — these two are NOT
 *                  matched by service name. Whichever ServiceName maps (via col E) to the
 *                  code "Delivery" (or "Livraison") uses Delivery1's price if the sale's
 *                  total basket volume is > 0.250 m³, otherwise it's treated as a "Parcel"
 *                  and uses Delivery2's price. Every other service column (J..M onward) is
 *                  matched by comparing its header text (row 1) to the service's mapping code.
 *
 *   Catalog    : ItemID | ItemName | Price | Volume                     (header row 1, data from row 2)
 *   Sales      : SaleID | Timestamp | SellerID | CustomerName | CustomerPhone | CustomerEmail |
 *                CustomerCity | CustomerAddress | OrderNumber | PaymentMethod | ServiceName |
 *                TotalAmount | ServiceFee | TotalVolume
 *   SaleItems  : SaleID | ItemID | ItemName | Quantity | UnitPrice | LineTotal | UnitVolume | LineVolume
 *
 * IMPORTANT: update the Sales sheet's header row to match the new column order above
 * (CustomerEmail was inserted after CustomerPhone, CustomerCity was inserted after that,
 * and ServiceFee was inserted after TotalAmount) before using this version — appendRow()
 * writes by position, not by header name. CustomerEmail is optional — an empty string is
 * written if the agent leaves it blank.
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
var SHEET_SALE_SELLERS = 'SaleSellers';

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
  var serviceCodes = getColumn(sheet, 5, lastRow);
  var cityNames = getColumn(sheet, 7, lastRow);

  var sellers = [];
  for (var i = 0; i < sellerIds.length; i++) {
    if (sellerIds[i] === '') continue;
    sellers.push({ id: sellerIds[i], name: sellerNames[i] || '' });
  }

  var services = [];
  for (var s = 0; s < serviceNames.length; s++) {
    if (serviceNames[s] === '') continue;
    services.push({ name: serviceNames[s], code: serviceCodes[s] || '' });
  }

  // Pricing grid: city names in col G, prices in H:M, header codes in row 1 (H1:M1).
  // H1/I1 are always "Delivery1"/"Delivery2" — see getServiceFeeForSale() for how
  // those two are applied.
  var pricingHeaders = sheet.getRange(1, 8, 1, 6).getValues()[0].map(function (v) {
    return String(v).trim();
  });

  var cities = [];
  if (lastRow >= 2) {
    var priceGrid = sheet.getRange(2, 8, lastRow - 1, 6).getValues();
    for (var c = 0; c < cityNames.length; c++) {
      if (cityNames[c] === '') continue;
      var prices = {};
      for (var h = 0; h < pricingHeaders.length; h++) {
        prices[pricingHeaders[h]] = Number(priceGrid[c][h]) || 0;
      }
      cities.push({ name: cityNames[c], prices: prices });
    }
  }

  return {
    sellers: sellers,
    paymentMethods: paymentMethods.filter(function (v) { return v !== ''; }),
    services: services,
    cities: cities,
    pricingHeaders: pricingHeaders
  };
}

// ---------- SERVICE FEE (city + service, Delivery1/Delivery2 split by basket volume) ----------

function getServiceFeeForSale(serviceName, cityName, totalVolume) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  // Look up the mapping code for the chosen service (col D -> col E).
  var serviceNames = getColumn(sheet, 4, lastRow);
  var serviceCodes = getColumn(sheet, 5, lastRow);
  var code = '';
  for (var i = 0; i < serviceNames.length; i++) {
    if (serviceNames[i] === serviceName) { code = serviceCodes[i]; break; }
  }
  code = (code || '').trim().toLowerCase();
  if (!code) return 0;

  // Find the pricing row for the chosen city (col G).
  var cityNames = getColumn(sheet, 7, lastRow);
  var cityRowOffset = -1;
  for (var j = 0; j < cityNames.length; j++) {
    if (cityNames[j] === cityName) { cityRowOffset = j; break; }
  }
  if (cityRowOffset === -1) return 0;

  var headerRow = sheet.getRange(1, 8, 1, 6).getValues()[0].map(function (v) {
    return String(v).trim();
  });
  var priceRow = sheet.getRange(2 + cityRowOffset, 8, 1, 6).getValues()[0];

  // "Delivery"/"Livraison" is special-cased by basket volume rather than matched
  // by name: index 0 = Delivery1 (H), index 1 = Delivery2 (I).
  if (code === 'delivery' || code === 'livraison') {
    var idx = (Number(totalVolume) > 0.250) ? 0 : 1; // >0.250 m3 -> Delivery1, else Parcel -> Delivery2
    return Number(priceRow[idx]) || 0;
  }

  // Every other service is matched by comparing its code to the header text,
  // skipping the first two columns reserved for Delivery1/Delivery2.
  for (var k = 2; k < headerRow.length; k++) {
    if (headerRow[k].toLowerCase() === code) {
      return Number(priceRow[k]) || 0;
    }
  }
  return 0;
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

  // Search only ItemID + ItemName (columns 1-2). Narrower range = fewer cells to
  // scan and no accidental matches against Price/Volume numbers, so results
  // come back faster and cleaner even with 40,000+ rows.
  var dataRange = sheet.getRange(2, 1, lastRow - 1, 2);
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
  var required = ['sellerIds', 'customerName', 'customerPhone', 'customerCity', 'customerAddress',
                   'orderNumber', 'paymentMethod', 'serviceName'];
  for (var i = 0; i < required.length; i++) {
    if (!payload[required[i]] || (Array.isArray(payload[required[i]]) && !payload[required[i]].length)) {
      return { success: false, error: 'Missing field: ' + required[i] };
    }
  }
  if (!payload.items || !payload.items.length) {
    return { success: false, error: 'No items in sale' };
  }
  if (!payload.sellerIds || payload.sellerIds.length < 2) {
    return { success: false, error: 'Minimum 2 agents requis' };
  }
  if (new Set(payload.sellerIds).size !== payload.sellerIds.length) {
    return { success: false, error: 'Le même agent est sélectionné plusieurs fois' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var salesSheet = ss.getSheetByName(SHEET_SALES);
  var itemsSheet = ss.getSheetByName(SHEET_SALE_ITEMS);
  var sellersSheet = ss.getSheetByName(SHEET_SALE_SELLERS);

  // Lock while we assign the next ID and write the row, so two agents
  // submitting at the same moment can never end up with the same SaleID.
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var saleId = getNextSaleId(salesSheet);
    var timestamp = new Date();
    var sellerIds = payload.sellerIds.join(', ');

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

    var serviceFee = getServiceFeeForSale(payload.serviceName, payload.customerCity, totalVolume);

    salesSheet.appendRow([
      saleId, timestamp, sellerIds, payload.customerName, payload.customerPhone,
      payload.customerEmail || '', payload.customerCity, payload.customerAddress, payload.orderNumber,
      payload.paymentMethod, payload.serviceName, totalAmount, serviceFee, totalVolume
    ]);

    itemRows.forEach(function (row) {
      itemsSheet.appendRow(row);
    });

    payload.sellerIds.forEach(function (id) {
      sellersSheet.appendRow([saleId, id]);
    });

    return {
      success: true,
      saleId: saleId,
      totalAmount: totalAmount,
      serviceFee: serviceFee,
      totalVolume: totalVolume
    };
  } finally {
    lock.releaseLock();
  }
}

// Simple incrementing SaleID (1, 2, 3…), based on the last row already in Sales.
// Runs inside the lock in recordSale(), so it's safe against concurrent submits.
function getNextSaleId(salesSheet) {
  var lastRow = salesSheet.getLastRow();
  if (lastRow < 2) return 1;
  var lastId = Number(salesSheet.getRange(lastRow, 1).getValue()) || 0;
  return lastId + 1;
}
