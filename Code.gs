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
//
// PERFORMANCE NOTE: this used to make 6 separate getColumn() calls, each of
// which issues its own getRange(...).getValues() round-trip covering the
// sheet's *entire* lastRow. If the Config tab's real data is small (a few
// dozen sellers/cities/services) but lastRow reports something like 38,000
// (very common — leftover formatting, a stray value, or a formula that once
// spilled down the sheet all push lastRow up even though the visible rows
// are empty), every one of those 6 calls pays the cost of a huge, mostly
// empty read. That was the actual bottleneck, not the small amount of real
// data.
//
// Fix #1 (this file): do ONE batched getRange(...).getValues() covering all
// the columns we need, instead of 6 separate calls. One large read is far
// cheaper than several.
// Fix #2 (this file): cache the parsed result for CACHE_TTL_SECONDS so
// repeat page loads/sale submissions don't hit the sheet at all.
// Fix #3 (you, in the Sheet): Config data changes rarely and should never
// need 38k rows. Select the row numbers below your real last row of data and
// "Delete rows" (not just clear content — clearing content does NOT lower
// lastRow, deleting rows does). That alone fixes the root cause.

var CACHE_KEY_CONFIG = 'config_v2';
var CACHE_TTL_SECONDS = 300; // 5 minutes — bump this up if Config changes rarely

function getConfig() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(CACHE_KEY_CONFIG);
  if (cached) return JSON.parse(cached);

  var result = buildConfig();
  try {
    cache.put(CACHE_KEY_CONFIG, JSON.stringify(result), CACHE_TTL_SECONDS);
  } catch (e) {
    // CacheService caps values at 100KB. If the Config sheet is still
    // carrying a lot of stray rows the JSON can exceed that — caching is a
    // nice-to-have, so just skip it rather than fail the whole request.
  }
  return result;
}

function buildConfig() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG);
  var lastRow = sheet.getLastRow();
  var headerRow = sheet.getRange(1, 8, 1, 6).getValues()[0].map(function (v) {
    return String(v).trim();
  });

  if (lastRow < 2) {
    return { sellers: [], paymentMethods: [], services: [], cities: [], pricingHeaders: headerRow };
  }

  // Single batched read of columns A..M instead of 8 separate range reads.
  var data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();

  var sellers = [];
  var paymentMethods = [];
  var seenPayment = {};
  var services = [];
  var cities = [];

  for (var r = 0; r < data.length; r++) {
    var row = data[r];

    var sellerId = cleanStr(row[0]);
    if (sellerId !== '') sellers.push({ id: sellerId, name: cleanStr(row[1]) });

    var pay = cleanStr(row[2]);
    if (pay !== '' && !seenPayment[pay]) {
      seenPayment[pay] = true;
      paymentMethods.push(pay);
    }

    var serviceName = cleanStr(row[3]);
    if (serviceName !== '') services.push({ name: serviceName, code: cleanStr(row[4]) });

    var cityName = cleanStr(row[6]);
    if (cityName !== '') {
      var prices = {};
      for (var h = 0; h < headerRow.length; h++) {
        prices[headerRow[h]] = Number(row[7 + h]) || 0;
      }
      cities.push({ name: cityName, prices: prices });
    }
  }

  return {
    sellers: sellers,
    paymentMethods: paymentMethods,
    services: services,
    cities: cities,
    pricingHeaders: headerRow
  };
}

function cleanStr(v) {
  return (v === null || v === undefined) ? '' : String(v).trim();
}

// ---------- SERVICE FEE (city + service, Delivery1/Delivery2 split by basket volume) ----------
//
// This used to re-scan the Config sheet from scratch (3 more getColumn() calls
// plus a row lookup) on every single sale. It now reuses getConfig(), which is
// cached, so a sale submission does zero extra Config-sheet reads in the
// common case.

function getServiceFeeForSale(serviceName, cityName, totalVolume) {
  var config = getConfig();

  var service = null;
  for (var i = 0; i < config.services.length; i++) {
    if (config.services[i].name === serviceName) { service = config.services[i]; break; }
  }
  if (!service) return 0;
  var code = (service.code || '').trim().toLowerCase();
  if (!code) return 0;

  var city = null;
  for (var j = 0; j < config.cities.length; j++) {
    if (config.cities[j].name === cityName) { city = config.cities[j]; break; }
  }
  if (!city) return 0;

  var headers = config.pricingHeaders;

  // "Delivery"/"Livraison" is special-cased by basket volume rather than matched
  // by name: index 0 = Delivery1 (H), index 1 = Delivery2 (I).
  if (code === 'delivery' || code === 'livraison') {
    var header = (Number(totalVolume) > 0.250) ? headers[0] : headers[1]; // >0.250 m3 -> Delivery1, else Parcel -> Delivery2
    return Number(city.prices[header]) || 0;
  }

  // Every other service is matched by comparing its code to the header text,
  // skipping the first two columns reserved for Delivery1/Delivery2.
  for (var k = 2; k < headers.length; k++) {
    if ((headers[k] || '').toLowerCase() === code) {
      return Number(city.prices[headers[k]]) || 0;
    }
  }
  return 0;
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
