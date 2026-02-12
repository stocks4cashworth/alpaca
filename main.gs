
// Main.gs


var PositionRowStart = 14;

/**
 * Makes a request to the Alpaca API.
 * Uses globally defined API keys and endpoint.
 * @param {string} path - The API endpoint path (e.g., "v2/account").
 * @param {Object} [params] - Optional parameters for the request (e.g., query string, method, payload).
 * @returns {Object|null} The parsed JSON response data, or null if an API error occurs.
 */function _request(path, params) {
  // 1. Force a reload of credentials from ScriptProperties
  initializeCredentials(); 

  // 2. Ensure keys are strings and not null/undefined
  var keyId = String(ALPAC_API_KEY_ID || "");
  var secretKey = String(ALPAC_API_SECRET_KEY || "");

  // 3. Safety check: If keys are empty, stop before the crash
  if (!keyId || !secretKey) {
    Logger.log("Critical Error: API Keys are missing. please use the Alpaca Tools menu to set them.");
    return null;
  }

  var headers = {
    "APCA-API-KEY-ID": keyId,
    "APCA-API-SECRET-KEY": secretKey,
  };

  var options = {
    "headers": headers,
    "muteHttpExceptions": true 
  };

  var url = (ALPAC_API_ENDPOINT || "https://paper-api.alpaca.markets/") + path;

  // ... rest of your existing _request logic (params handling and UrlFetchApp)
  if (params) {
    if (params.qs) {
      var kv = [];
      for (var k in params.qs) {
        kv.push(k + "=" + encodeURIComponent(params.qs[k]));
      }
      url += "?" + kv.join("&");
      delete params.qs;
    }
    for (var k in params) {
      options[k] = params[k];
    }
  }

  var response = UrlFetchApp.fetch(url, options);
  var responseCode = response.getResponseCode();
  var responseText = response.getContentText();

  if (responseCode >= 400) {
    Logger.log("API Request Error for " + url + ": " + responseCode + " - " + responseText);
    return null; 
  }

  var data = JSON.parse(responseText); 
  return data; 
}

/**
 * Sends a DELETE request to cancel an order by its ID.
 * Uses globally defined API keys and endpoint.
 * @param {string} orderId - The ID of the order to cancel.
 * @returns {Object} An object containing the response code and text.
 */
function _cancelRequest(orderId) {
  var headers = {
    "APCA-API-KEY-ID": ALPAC_API_KEY_ID, // Use global API key
    "APCA-API-SECRET-KEY": ALPAC_API_SECRET_KEY, // Use global API secret
  };

  var url = ALPAC_API_ENDPOINT + "v2/orders/" + orderId; // Use global endpoint
  var options = {
    "method": "DELETE",
    "headers": headers,
    "muteHttpExceptions": true 
  };

  var response = UrlFetchApp.fetch(url, options);
  var responseCode = response.getResponseCode();
  var responseText = response.getContentText();

  return {
    code: responseCode,
    text: responseText
  };
}

/**
 * Retrieves account information from Alpaca.
 * @returns {Object} Account data, or an empty object if the request fails.
 */
function getAccount() {
  var accountData = _request("v2/account");
  return accountData || {}; 
}

/**
 * Retrieves a list of orders from Alpaca.
 * Fetches all orders (open, filled, canceled etc.) from the last 30 days, up to 500.
 * @returns {Array} An array of order objects, or an empty array if the request fails.
 */
function listOrders() {
  var thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 300);
  var thirtyDaysAgoISO = thirtyDaysAgo.toISOString();

  var ordersData = _request("v2/orders", { qs: { status: "all", limit: 500, after: thirtyDaysAgoISO } }); 
  return ordersData || []; 
}

/**
 * Retrieves a list of current positions from Alpaca.
 * @returns {Array} An array of position objects, or an empty array if the request fails.
 */
function listPositions() {
  var positionsData = _request("v2/positions");
  return positionsData || []; 
}

/**
 * Submits an order to Alpaca.
 * This function is enhanced to support OCO orders by accepting additional parameters.
 *
 * @param {string} symbol - The symbol of the asset.
 * @param {number} qty - The quantity of the asset.
 * @param {string} side - "buy" or "sell".
 * @param {string} type - "market", "limit", "stop", "stop_limit", "trailing_stop".
 * @param {string} tif - "day", "gtc", "opg", "cls", "ioc", "fok".
 * @param {number} [limit_price] - Required for "limit" and "stop_limit" orders.
 * @param {number} [stop_price] - Required for "stop" and "stop_limit" orders.
 * @param {string} [order_class] - "simple", "bracket", "oco", "oto".
 * @param {number} [take_profit_limit_price] - For "bracket" or "oco" orders (take_profit leg).
 * @param {number} [stop_loss_stop_price] - For "bracket" or "oco" orders (stop_loss leg stop price).
 * @param {number} [stop_loss_limit_price] - For "bracket" or "oco" orders (stop_loss leg limit price, makes it stop-limit).
 * @returns {Object} The API response from the order submission, or an empty object if the request fails.
 */
function submitOrder(symbol, qty, side, type, tif, limit_price, stop_price, order_class, take_profit_limit_price, stop_loss_stop_price, stop_loss_limit_price) {
  var payload = {
    symbol: symbol,
    side: side,
    qty: qty,
    type: type,
    time_in_force: tif,
  };

  // Add limit_price and stop_price for simple orders if provided
  if (limit_price) {
    payload.limit_price = limit_price;
  }
  if (stop_price) {
    payload.stop_price = stop_price;
  }

  // Handle order_class specific parameters
  if (order_class) {
    payload.order_class = order_class;

    if (order_class === "oco" || order_class === "bracket") {
      if (take_profit_limit_price) {
        payload.take_profit = {
          limit_price: take_profit_limit_price
        };
      }
      if (stop_loss_stop_price) {
        payload.stop_loss = {
          stop_price: stop_loss_stop_price
        };
        if (stop_loss_limit_price) {
          payload.stop_loss.limit_price = stop_loss_limit_price;
        }
      }
    }
    // Add other order_class types if needed (e.g., OTO)
  }

  var response = _request("/v2/orders", {
    method: "POST",
    payload: JSON.stringify(payload),
  });
  return response || {}; 
}

function orderFromSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet.getRange("B1").setValue("submitting");

  // Read and clean values
  var side = sheet.getRange("G3").getValue().toString().toLowerCase().trim();
  var symbol = sheet.getRange("G4").getValue().toString().toUpperCase().trim();
  var qty = parseFloat(sheet.getRange("G5").getValue());
  var type = sheet.getRange("G6").getValue().toString().toLowerCase().trim();
  var tif = sheet.getRange("G7").getValue().toString().toLowerCase().trim();
  var limit = sheet.getRange("G8").getValue();
  var stop = sheet.getRange("G9").getValue();

  // Basic validation
  if (!symbol || isNaN(qty)) {
    sheet.getRange("B1").setValue("Error: Check Symbol (G4) and Quantity (G5).");
    return;
  }

  // Ensure limit and stop are numbers or null
  var limitPrice = limit ? parseFloat(limit) : null;
  var stopPrice = stop ? parseFloat(stop) : null;

  var resp = submitOrder(symbol, qty, side, type, tif, limitPrice, stopPrice);
  sheet.getRange("B1").setValue(JSON.stringify(resp, null, 2));
}

/**
 * Submits an OCO (One-Cancels-Other) order based on values from sheet K3:K10.
 * Reads symbol from G4 and forces it to uppercase.
 *
 * K3: Side (buy/sell)
 * G4: Symbol (shared with simple order)
 * K5: Quantity
 * K6: Primary Order Type (e.g., limit)
 * K7: Time in Force (e.g., gtc)
 * K8: Take Profit Limit Price
 * K9: Stop Loss Stop Price
 * K10: Stop Loss Limit Price (Optional, if provided, makes it a stop-limit)
 */
function OCOorderFromSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet.getRange("B1").setValue("submitting OCO"); // Update status cell

  var side = sheet.getRange("K3").getValue();
  var symbol = sheet.getRange("G4").getValue().toUpperCase(); // Read from G4 and force uppercase
  var qty = sheet.getRange("K5").getValue();
  var type = sheet.getRange("K6").getValue(); // e.g., "limit" for the primary order
  var tif = sheet.getRange("K7").getValue();
  var takeProfitLimit = sheet.getRange("K8").getValue();
  var stopLossStop = sheet.getRange("K9").getValue();
  var stopLossLimit = sheet.getRange("K10").getValue(); // Optional stop-limit price

  // Basic validation for symbol
  if (!symbol) {
    sheet.getRange("B1").setValue("Error: Symbol (G4) cannot be empty.");
    return;
  }

  // Validate required OCO parameters
  if (!takeProfitLimit || !stopLossStop) {
    sheet.getRange("B1").setValue("Error: Take Profit Limit (K8) and Stop Loss Stop (K9) are required for OCO orders.");
    return;
  }

  // Call the enhanced submitOrder function with OCO specific parameters
  var resp = submitOrder(
    symbol,
    qty,
    side,
    type, // Primary order type
    tif,
    null, // No simple limit_price for the primary order (it's handled by take_profit)
    null, // No simple stop_price for the primary order (it's handled by stop_loss)
    "oco", // Order Class
    takeProfitLimit, // Take Profit Limit Price
    stopLossStop,    // Stop Loss Stop Price
    stopLossLimit    // Stop Loss Limit Price (optional)
  );

  sheet.getRange("B1").setValue(JSON.stringify(resp, null, 2));
}

/**
 * Cancels an order by its ID, read from cell G11.
 * Displays status in cell B1.
 */
function cancelOrderFromSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var orderIdToCancel = sheet.getRange("G11").getValue(); // Read from G11
  var statusCell = sheet.getRange("B1"); // Status cell

  if (!orderIdToCancel) {
    statusCell.setValue("No Order ID provided in G11.");
    return;
  }

  statusCell.setValue("Attempting to cancel order: " + orderIdToCancel);

  var cancellationResult = _cancelRequest(orderIdToCancel);

  if (cancellationResult.code >= 200 && cancellationResult.code < 300) {
    statusCell.setValue("Successfully requested cancellation for order: " + orderIdToCancel);
  } else {
    statusCell.setValue("Error canceling order " + orderIdToCancel + ". Code: " + cancellationResult.code + ", Response: " + cancellationResult.text);
    Logger.log("Error canceling order " + orderIdToCancel + ". Code: " + cancellationResult.code + ", Response: " + cancellationResult.text);
  }
}

/**
 * Clears existing position data from the "Main" sheet.
 */
function clearPositions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Main"); // Explicitly get the "Main" sheet
  if (!sheet) { 
    sheet = ss.getActiveSheet();
    Logger.log("Warning: 'Main' sheet not found for clearPositions. Using active sheet.");
  }

  var rowIdx = PositionRowStart;
  while (true) {
    var symbol = sheet.getRange("A" + rowIdx).getValue(); // Assuming symbol is in Column A for clearing
    if (!symbol) {
      break;
    }
    rowIdx++;
  }
  var rows = rowIdx - PositionRowStart;
  if (rows > 0) {
    sheet.deleteRows(PositionRowStart, rows);
  }
}






/**
 * Updates the "Main" sheet with a unified view of Positions and Orders.
 * Extracts Stop Loss from OCO/Bracket order legs.
 */
function updateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Main");
  
  var headers = [
    "Ticker", "Owned Qty", "Cost Basis", "Unrealized Gain", "Profit %", 
    "Cost Basis/Share", "Current Price", "% of Portfolio", 
    "Buy Qty", "Buy Limit", 
    "Sell Qty", "Sell Limit", "Stop", "Order ID"
  ];
  
  sheet.getRange(13, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#f3f3f3");

  var accountInfo = getAccount();
  var positions = listPositions();
  var allOrders = listOrders();
  var openOrders = allOrders.filter(o => ['new', 'partially_filled', 'pending_cancel', 'accepted'].indexOf(o.status) !== -1);
  var portfolioValue = parseFloat(accountInfo.portfolio_value) || 1;

  var masterMap = {};
  positions.forEach(p => { masterMap[p.symbol] = { pos: p, buys: [], sells: [] }; });
  openOrders.forEach(o => {
    if (!masterMap[o.symbol]) masterMap[o.symbol] = { pos: null, buys: [], sells: [] };
    if (o.side === 'buy') masterMap[o.symbol].buys.push(o);
    else masterMap[o.symbol].sells.push(o);
  });

  var symbols = Object.keys(masterMap).sort();
  var outputRows = [];

  symbols.forEach(sym => {
    var data = masterMap[sym];
    var p = data.pos;
    var rowCount = Math.max(1, data.buys.length, data.sells.length);

    for (var i = 0; i < rowCount; i++) {
      var buy = data.buys[i] || {};
      var sell = data.sells[i] || {};
      
      // NEW: Logic to find the Stop Loss price within the order legs
      var stopLossPrice = sell.stop_price || ""; 
      if (!stopLossPrice && sell.legs) {
        // Search through legs for a stop or stop_limit order type
        sell.legs.forEach(leg => {
          if (leg.stop_price) stopLossPrice = leg.stop_price;
        });
      }

      var row = [
        sym,
        (i === 0 && p) ? p.qty : "",
        (i === 0 && p) ? p.market_value : "",
        (i === 0 && p) ? p.unrealized_pl : "",
        (i === 0 && p) ? p.unrealized_plpc : "",
        (i === 0 && p) ? (p.cost_basis / p.qty) : "",
        (i === 0 && p) ? p.current_price : "",
        (i === 0 && p) ? (p.market_value / portfolioValue) : "",
        buy.qty || "",
        buy.limit_price || "",
        sell.qty || "",
        sell.limit_price || "",
        stopLossPrice, // Corrected to pull from legs if necessary
        buy.id || sell.id || ""
      ];
      outputRows.push(row);
    }
  });

  sheet.getRange(14, 1, 500, headers.length).clearContent();
  if (outputRows.length > 0) {
    sheet.getRange(14, 1, outputRows.length, headers.length).setValues(outputRows);
    // Apply standard currency and percentage formatting...
    sheet.getRange(14, 12, outputRows.length, 2).setNumberFormat("$#,##0.00"); // Sell Limit & Stop
  }
}
