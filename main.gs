// Main.gs

var PositionRowStart = 14; 

function _request(path, params) {
  var headers = {
    "APCA-API-KEY-ID": "PKY5JN04MODLJ577J0XG",
    "APCA-API-SECRET-KEY": "jhbCwzcDcMQ3XLDv3m1cdPInnZl4uwAfGmQf1oeW",
  };

  var endpoint = "https://paper-api.alpaca.markets/";
  var options = {
    "headers": headers,
  };
  var url = endpoint + path;
  if (params) {
    if (params.qs) {
      var kv = [];
      for (var k in params.qs) {
        kv.push(k + "=" + encodeURIComponent(params.qs[k]));
      }
      url += "?" + kv.join("&");
      delete params.qs
    }
    for (var k in params) {
      options[k] = params[k];
    }
  }

  var response = UrlFetchApp.fetch(url, options);
  var json = response.getContentText();
  var data = JSON.parse(json);
  return data;
}

function _cancelRequest(orderId) {
  var headers = {
    "APCA-API-KEY-ID": "PKY5JN04MODLJ577J0XG",
    "APCA-API-SECRET-KEY": "jhbCwzcDcMQ3XLDv3m1cdPInnZl4uwAfGmQf1oeW",
  };

  var endpoint = "https://paper-api.alpaca.markets/";
  var url = endpoint + "v2/orders/" + orderId;
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

function getAccount() {
  return _request("v2/account");
}

function listOrders() {
  // Request all orders (status: "all") submitted after 30 days ago, with a limit of 500
  var thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
  var thirtyDaysAgoISO = thirtyDaysAgo.toISOString();

  return _request("v2/orders", { qs: { status: "all", limit: 500, after: thirtyDaysAgoISO } }); 
}

function listPositions() {
  return _request("v2/positions");
}

function submitOrder(symbol, qty, side, type, tif, limit, stop) {
  var payload = {
    symbol: symbol,
    side: side,
    qty: qty,
    type: type,
    time_in_force: tif,
  };
  if (limit) {
    payload.limit_price = limit;
  }
  if (stop) {
    payload.stop_price = stop;
  }
  return _request("/v2/orders", {
    method: "POST",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
}

function orderFromSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet.getRange("E2").setValue("submitting")

  var side = sheet.getRange("J3").getValue()
  var symbol = sheet.getRange("J4").getValue()
  var qty = sheet.getRange("J5").getValue()
  var type = sheet.getRange("J6").getValue()
  var tif = sheet.getRange("J7").getValue()
  var limit = sheet.getRange("J8").getValue()
  var stop = sheet.getRange("J9").getValue()

  var resp = submitOrder(symbol, qty, side, type, tif, limit, stop);
  sheet.getRange("E2").setValue(JSON.stringify(resp, null, 2))
}

function cancelOrderFromSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var orderIdToCancel = sheet.getRange("H10").getValue();
  var statusCell = sheet.getRange("H11"); 

  if (!orderIdToCancel) {
    statusCell.setValue("No Order ID provided in H10.");
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

function clearPositions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Main"); // Explicitly get the "Main" sheet
  if (!sheet) { // Fallback if "Main" sheet doesn't exist, though it should for this setup
    sheet = ss.getActiveSheet();
    Logger.log("Warning: 'Main' sheet not found for clearPositions. Using active sheet.");
  }

  var rowIdx = PositionRowStart;
  while (true) {
    var symbol = sheet.getRange("E" + rowIdx).getValue();
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

function updateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var accountSheet = ss.getSheetByName("Main");
  
  // Update Account Information
  var accountInfo = getAccount();
  accountSheet.getRange("B5").setValue(accountInfo.id);
  accountSheet.getRange("B6").setValue(accountInfo.buying_power);
  accountSheet.getRange("B7").setValue(accountInfo.cash);
  accountSheet.getRange("B8").setValue(accountInfo.portfolio_value); // Portfolio Value
  accountSheet.getRange("B9").setValue(accountInfo.status);
  accountSheet.getRange("B6:B8").setNumberFormat("$#,##0.00");

  var portfolioValueCell = "B8"; // Cell containing portfolio value

  // Clear and Update Positions
  clearPositions(); // This will now explicitly target the "Main" sheet
  var positions = listPositions();
  if (positions.length > 0) {
    positions.sort(function(a, b) { return a.symbol < b.symbol ? -1 : 1 });
    for (var i = 0; i < positions.length; i++) {
      var rowIdx = PositionRowStart + i;
      accountSheet.getRange("A" + rowIdx).setValue(positions[i].symbol);
      accountSheet.getRange("B" + rowIdx).setValue(positions[i].qty);
      accountSheet.getRange("C" + rowIdx).setValue(positions[i].market_value);
      accountSheet.getRange("D" + rowIdx).setValue(positions[i].cost_basis);
      accountSheet.getRange("E" + rowIdx).setValue(positions[i].unrealized_pl);
      accountSheet.getRange("F" + rowIdx).setValue(positions[i].unrealized_plpc);
      accountSheet.getRange("G" + rowIdx).setValue(positions[i].current_price);
      
      // Calculate and set 'Percent of Portfolio' in Column H
      accountSheet.getRange("H" + rowIdx).setFormula("=C" + rowIdx + "/" + portfolioValueCell); // Formula for percentage
    }
    var endIdx = PositionRowStart + positions.length - 1;
    accountSheet.getRange("B" + PositionRowStart + ":B" + endIdx).setNumberFormat("#,###");
    accountSheet.getRange("C" + PositionRowStart + ":C" + endIdx).setNumberFormat("$#,##0.00");
    accountSheet.getRange("D" + PositionRowStart + ":D" + endIdx).setNumberFormat("$#,##0.00");
    accountSheet.getRange("E" + PositionRowStart + ":E" + endIdx).setNumberFormat("$#,##0.00");
    accountSheet.getRange("F" + PositionRowStart + ":F" + endIdx).setNumberFormat("0.00%");
    accountSheet.getRange("G" + PositionRowStart + ":G" + endIdx).setNumberFormat("$#,##0.00");
    accountSheet.getRange("H" + PositionRowStart + ":H" + endIdx).setNumberFormat("0.00%"); // Format column H as percentage

    accountSheet.getRange("C" + (endIdx + 1)).setValue("total");
    accountSheet.getRange("D" + (endIdx + 1)).setValue("total");
    accountSheet.getRange("E" + (endIdx + 1)).setValue("total");
    accountSheet.getRange("F" + (endIdx + 1)).setValue("average");
    accountSheet.getRange("G" + (endIdx + 1)).setValue("median");

    accountSheet.getRange("C" + (endIdx + 2)).setFormula("=sum(C" + PositionRowStart + ":C" + endIdx + ")");
    accountSheet.getRange("D" + (endIdx + 2)).setFormula("=sum(D" + PositionRowStart + ":D" + endIdx + ")");
    accountSheet.getRange("E" + (endIdx + 2)).setFormula("=sum(E" + PositionRowStart + ":E" + endIdx + ")");
    accountSheet.getRange("F" + (endIdx + 2)).setFormula("=average(F" + PositionRowStart + ":F" + endIdx + ")");
    accountSheet.getRange("G" + (endIdx + 2)).setFormula("=median(G" + PositionRowStart + ":G" + endIdx + ")");
  }

  // List Open Orders on the same sheet, starting at J12
  var orders = listOrders(); // This will now fetch ALL orders (last 30 days)
  var openOrders = orders.filter(function(order) {
    return ['new', 'partially_filled', 'pending_cancel', 'accepted'].indexOf(order.status) !== -1;
  });

  var openOrdersStartRow = 12; 
  var openOrdersStartColumn = 10; // Column J is the 10th column

  if (openOrders.length > 0) {
    accountSheet.getRange(openOrdersStartRow, openOrdersStartColumn).setValue("Open Orders");
    accountSheet.getRange(openOrdersStartRow, openOrdersStartColumn).setFontWeight("bold");

    var headers = ["ID", "Symbol", "Quantity", "Side", "Type", "Time in Force", "Limit Price", "Stop Price", "Status", "Submitted At"];
    accountSheet.getRange(openOrdersStartRow + 1, openOrdersStartColumn, 1, headers.length).setValues([headers]).setFontWeight("bold");

    var openOrderData = openOrders.map(function(order) {
      return [
        order.id,
        order.symbol,
        order.qty,
        order.side,
        order.type,
        order.time_in_force,
        order.limit_price || "",
        order.stop_price || "",
        order.status,
        order.submitted_at
      ];
    });
    accountSheet.getRange(openOrdersStartRow + 2, openOrdersStartColumn, openOrderData.length, openOrderData[0].length).setValues(openOrderData);
    accountSheet.autoResizeColumns(openOrdersStartColumn, openOrderData[0].length);
  } else {
    // Clear previous open orders if none exist now
    accountSheet.getRange(openOrdersStartRow, openOrdersStartColumn, 100, 10).clearContent(); // Clear a reasonable range (10 columns from J)
  }
}
