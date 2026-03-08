
// Main.gs


var PositionRowStart = 14;

/**
 * Makes a request to the Alpaca API.
 * Uses globally defined API keys and endpoint.
 * @param {string} path - The API endpoint path (e.g., "v2/account").
 * @param {Object} [params] - Optional parameters for the request (e.g., query string, method, payload).
 * @returns {Object|null} The parsed JSON response data, or null if an API error occurs.
 */
function _request(path, params) {
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

// Replace the error block in your _request function
if (responseCode >= 400) {
  Logger.log("API Request Error for " + url + ": " + responseCode + " - " + responseText);
  try {
    return JSON.parse(responseText); // Return the actual error object from Alpaca
  } catch (e) {
    return { message: "Unknown API Error: " + responseCode }; 
  }

}
  var data = JSON.parse(responseText); 
  return data; 
}

/**
 * UPDATED: Fixed cancellation request to prevent Header:null errors.
 */
function _cancelRequest(orderId) {
  initializeCredentials(); // Ensure variables are loaded

  var headers = {
    "APCA-API-KEY-ID": ALPAC_API_KEY_ID,
    "APCA-API-SECRET-KEY": ALPAC_API_SECRET_KEY,
  };

  var url = ALPAC_API_ENDPOINT + "v2/orders/" + orderId;
  var options = {
    "method": "DELETE",
    "headers": headers,
    "muteHttpExceptions": true 
  };

  var response = UrlFetchApp.fetch(url, options);
  return {
    code: response.getResponseCode(),
    text: response.getContentText()
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
 

/**
 * Cancels an order by its ID, read from cell c12.
 * Displays status in cell B1.
 */
function cancelOrderFromSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var orderIdToCancel = sheet.getRange("c12").getValue(); // Read from c12
  var statusCell = sheet.getRange("B1"); // Status cell

  if (!orderIdToCancel) {
    statusCell.setValue("No Order ID provided in c12.");
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




