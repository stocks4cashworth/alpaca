
// Main.gs

// --- GLOBAL CONFIGURATION ---
// IMPORTANT: Replace with your actual Alpaca API Key ID and Secret Key
// For paper trading, use "https://paper-api.alpaca.markets/"
//var ALPAC_API_KEY_ID = ""; // <<< Fill in your Alpaca API Key ID
//var ALPAC_API_SECRET_KEY = ""; // <<< Fill in your Alpaca API Secret Key
//var ALPAC_API_ENDPOINT = "https://paper-api.alpaca.markets/"; // Using paper API endpoint
// For live trading, use "https://api.alpaca.markets/"
var ALPAC_API_KEY_ID = ""; 
var ALPAC_API_SECRET_KEY = ""; 
var ALPAC_API_ENDPOINT = "https://api.alpaca.markets/"; 


var PositionRowStart = 14; 

/**
 * Makes a request to the Alpaca API.
 * Uses globally defined API keys and endpoint.
 * @param {string} path - The API endpoint path (e.g., "v2/account").
 * @param {Object} [params] - Optional parameters for the request (e.g., query string, method, payload).
 * @returns {Object|null} The parsed JSON response data, or null if an API error occurs.
 */
function _request(path, params) {
  var headers = {
    "APCA-API-KEY-ID": ALPAC_API_KEY_ID,
    "APCA-API-SECRET-KEY": ALPAC_API_SECRET_KEY,
  };

  var options = {
    "headers": headers,
    "muteHttpExceptions": true // Ensures UrlFetchApp doesn't throw on 4xx/5xx responses
  };
  var url = ALPAC_API_ENDPOINT + path; // Use global endpoint

  if (params) {
    // Handle query string parameters
    if (params.qs) {
      var kv = [];
      for (var k in params.qs) {
        kv.push(k + "=" + encodeURIComponent(params.qs[k]));
      }
      url += "?" + kv.join("&");
      delete params.qs; 
    }
    // Merge other options (method, payload etc.)
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
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
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
 * Submits a simple order based on values from sheet G3:G9.
 * Reads symbol from G4 and forces it to uppercase.
 */
function orderFromSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet.getRange("B1").setValue("submitting"); // Status cell for order submission

  var side = sheet.getRange("G3").getValue();
  var symbol = sheet.getRange("G4").getValue()//.toUpperCase(); // Read from G4 and force uppercase
  var qty = sheet.getRange("G5").getValue();
  var type = sheet.getRange("G6").getValue();
  var tif = sheet.getRange("G7").getValue();
  var limit = sheet.getRange("G8").getValue();
  var stop = sheet.getRange("G9").getValue();

  // Basic validation for symbol
  if (!symbol) {
    sheet.getRange("B1").setValue("Error: Symbol (G4) cannot be empty.");
    return;
  }

  // Calling submitOrder for a simple order
  var resp = submitOrder(symbol, qty, side, type, tif, limit, stop);
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
 * Updates the "Main" sheet with account information, positions,
 * and open orders.
 */
function updateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var accountSheet = ss.getSheetByName("Main");
  
  // Update Account Information
  var accountInfo = getAccount();
  accountSheet.getRange("B5").setValue(accountInfo.id || "");
  accountSheet.getRange("B6").setValue(accountInfo.buying_power || "");
  accountSheet.getRange("B7").setValue(accountInfo.cash || "");
  accountSheet.getRange("B8").setValue(accountInfo.portfolio_value || ""); // Portfolio Value
  accountSheet.getRange("B9").setValue(accountInfo.status || "");
  accountSheet.getRange("B6:B8").setNumberFormat("$#,##0.00");

  var portfolioValueCell = "B8"; // Cell containing portfolio value

  // Clear and Update Positions
  clearPositions(); 
  var positions = listPositions();
  if (positions.length > 0) {
    positions.sort(function(a, b) { return a.symbol < b.symbol ? -1 : 1 });
    for (var i = 0; i < positions.length; i++) {
      var rowIdx = PositionRowStart + i;
      accountSheet.getRange("A" + rowIdx).setValue(positions[i].symbol || "");
      accountSheet.getRange("B" + rowIdx).setValue(positions[i].qty || "");
      accountSheet.getRange("C" + rowIdx).setValue(positions[i].market_value || "");
      accountSheet.getRange("D" + rowIdx).setValue(positions[i].cost_basis || "");
      accountSheet.getRange("E" + rowIdx).setValue(positions[i].unrealized_pl || "");
      accountSheet.getRange("F" + rowIdx).setValue(positions[i].unrealized_plpc || "");
      accountSheet.getRange("G" + rowIdx).setValue(positions[i].current_price || "");
      
      // Calculate and set 'Percent of Portfolio' in Column H
      // Use IFERROR to prevent #DIV/0! if portfolio value is 0 or empty
      accountSheet.getRange("H" + rowIdx).setFormula("=IFERROR(C" + rowIdx + "/" + portfolioValueCell + ", \"\")"); 
    }
    var endIdx = PositionRowStart + positions.length - 1;
    accountSheet.getRange("B" + PositionRowStart + ":B" + endIdx).setNumberFormat("#,###");
    accountSheet.getRange("C" + PositionRowStart + ":C" + endIdx).setNumberFormat("$#,##0.00");
    accountSheet.getRange("D" + PositionRowStart + ":D" + endIdx).setNumberFormat("$#,##0.00");
    accountSheet.getRange("E" + PositionRowStart + ":E" + endIdx).setNumberFormat("$#,##0.00");
    accountSheet.getRange("F" + PositionRowStart + ":F" + endIdx).setNumberFormat("0.00%");
    accountSheet.getRange("G" + PositionRowStart + ":G" + endIdx).setNumberFormat("$#,##0.00");
    accountSheet.getRange("H" + PositionRowStart + ":H" + endIdx).setNumberFormat("0.00%"); // Format column H as percentage

    // Set "total", "average", "median" labels and make them bold
    accountSheet.getRange("C" + (endIdx + 1)).setValue("total").setFontWeight("bold");
    accountSheet.getRange("D" + (endIdx + 1)).setValue("total").setFontWeight("bold");
    accountSheet.getRange("E" + (endIdx + 1)).setValue("total").setFontWeight("bold");
    accountSheet.getRange("F" + (endIdx + 1)).setValue("average").setFontWeight("bold");
    accountSheet.getRange("G" + (endIdx + 1)).setValue("median").setFontWeight("bold");

    accountSheet.getRange("C" + (endIdx + 2)).setFormula("=sum(C" + PositionRowStart + ":C" + endIdx + ")");
    accountSheet.getRange("D" + (endIdx + 2)).setFormula("=sum(D" + PositionRowStart + ":D" + endIdx + ")");
    accountSheet.getRange("E" + (endIdx + 2)).setFormula("=sum(E" + PositionRowStart + ":E" + endIdx + ")");
    accountSheet.getRange("F" + (endIdx + 2)).setFormula("=average(F" + PositionRowStart + ":F" + endIdx + ")");
    accountSheet.getRange("G" + (endIdx + 2)).setFormula("=median(G" + PositionRowStart + ":G" + endIdx + ")");
  } else {
    // If no positions, clear any old totals/averages
    var clearStartRow = PositionRowStart;
    var clearEndRow = PositionRowStart + 5; // Clear a few rows below start
    accountSheet.getRange("C" + clearStartRow + ":G" + clearEndRow).clearContent();
    accountSheet.getRange("H" + clearStartRow + ":H" + clearEndRow).clearContent(); // Clear percentage column too
  }

  // List Open Orders on the same sheet, starting at J12
  var orders = listOrders(); 
  var openOrders = orders.filter(function(order) {
    return ['new', 'partially_filled', 'pending_cancel', 'accepted'].indexOf(order.status) !== -1;
  });

  var openOrdersStartRow = 12; 
  var openOrdersStartColumn = 10; // Column J is the 10th column

  // Define all headers for open orders (moved outside if/else for safety)
  var headers = [
    "ID", "Symbol", "Quantity", "Side", "Type", "Time in Force", "Limit Price", "Stop Price",
    "Status", "Submitted At", "Created At" , "Order Class",  "Legs" 
  ];

  if (openOrders.length > 0) {
    accountSheet.getRange(openOrdersStartRow, openOrdersStartColumn).setValue("Open Orders");
    accountSheet.getRange(openOrdersStartRow, openOrdersStartColumn).setFontWeight("bold");

    accountSheet.getRange(openOrdersStartRow + 1, openOrdersStartColumn, 1, headers.length).setValues([headers]).setFontWeight("bold");

    var openOrderData = openOrders.map(function(order) {
      return [
        order.id || "",
        order.symbol || "",
        order.qty || "",
        order.side || "",
        order.type || "",
        order.time_in_force || "",
        order.limit_price || "",
        order.stop_price || "",
        order.status || "",
        order.submitted_at || "",
        order.created_at || "",
        order.asset_class || "",
       (order.legs && order.legs.length > 0) ? JSON.stringify(order.legs) : "" 
      ];
    });
    accountSheet.getRange(openOrdersStartRow + 2, openOrdersStartColumn, openOrderData.length, openOrderData[0].length).setValues(openOrderData);
    accountSheet.autoResizeColumns(openOrdersStartColumn, openOrderData[0].length);
  } else {
    // Clear previous open orders if none exist now
    accountSheet.getRange(openOrdersStartRow, openOrdersStartColumn, 100, headers.length).clearContent(); 
  }
}
