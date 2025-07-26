// FilledOrders.gs

function updatefilledSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var filledOrdersSheet = ss.getSheetByName("Filled Orders");

  Logger.log("--- Starting updatefilledSheet (for Filled Orders) ---");

  // Create sheet if it doesn't exist
  if (!filledOrdersSheet) {
    Logger.log("Creating new 'Filled Orders' sheet.");
    filledOrdersSheet = ss.insertSheet("Filled Orders");
  } else {
    Logger.log("Found 'Filled Orders' sheet.");
  }

  // Always write headers to ensure they are present and correct
  var headers = ["ID", "Symbol", "Quantity", "Side", "Type", "Time in Force", "Limit Price", "Stop Price", "Filled At", "Filled Avg Price"];
  filledOrdersSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  filledOrdersSheet.autoResizeColumns(1, headers.length); // Auto-resize headers
  Logger.log("Headers written to 'Filled Orders' sheet.");

  // --- Retrieve existing filled order IDs from the sheet ---
  // Get all values from column A, starting from row 2 (after headers)
  var existingDataRange = filledOrdersSheet.getRange("A2:A" + filledOrdersSheet.getLastRow());
  var existingOrderIds = new Set();
  if (existingDataRange.getValues().length > 0 && filledOrdersSheet.getLastRow() > 1) { // Check if there's actual data beyond headers
    existingDataRange.getValues().forEach(function(row) {
      if (row[0]) { // Ensure the cell is not empty
        existingOrderIds.add(row[0]);
      }
    });
  }
  Logger.log("Number of existing filled order IDs on sheet: " + existingOrderIds.size);

  // --- Step 1: Get all orders (limited to last 30 days by listOrders) ---
  Logger.log("Calling listOrders() to retrieve orders (last 30 days) from Alpaca.");
  // listOrders() is defined in Main.gs but is globally accessible
  var orders = listOrders(); 
  Logger.log("Total orders retrieved from Alpaca (last 30 days): " + orders.length);

  if (orders.length > 0) {
    Logger.log("Logging status for each retrieved order before filtering:");
    orders.forEach(function(order) {
      Logger.log("  Order ID: " + order.id + ", Symbol: " + order.symbol + ", Status: " + order.status + ", Type: " + order.type);
    });
  } else {
    Logger.log("No orders retrieved at all from Alpaca API for the last 30 days.");
  }

  // --- Step 2: Filter for ONLY new filled orders ---
  Logger.log("Filtering for new 'filled' orders not already on the sheet.");
  var newFilledOrders = orders.filter(function(order) {
    var isFilled = (order.status === "filled");
    var isNew = !existingOrderIds.has(order.id); // Check if ID is NOT already on the sheet

    if (isFilled && isNew) {
      Logger.log("  NEW FILLED ORDER: Order ID: " + order.id + ", Status: " + order.status);
    } else if (isFilled && !isNew) {
      Logger.log("  EXISTING FILLED ORDER (skipped): Order ID: " + order.id + ", Status: " + order.status);
    }
    return isFilled && isNew;
  });

  Logger.log("Number of NEW 'filled' orders found to write: " + newFilledOrders.length);

  // --- Step 3: Write new filled orders data to sheet ---
  if (newFilledOrders.length > 0) {
    Logger.log("Preparing data for " + newFilledOrders.length + " new filled orders.");
    var filledOrderData = newFilledOrders.map(function(order) {
      return [
        order.id,
        order.symbol,
        order.qty,
        order.side,
        order.type,
        order.time_in_force,
        order.limit_price || "",
        order.stop_price || "",
        order.filled_at || "", 
        order.filled_avg_price || "" 
      ];
    });
    // Append data starting from the last row + 1
    filledOrdersSheet.getRange(filledOrdersSheet.getLastRow() + 1, 1, filledOrderData.length, filledOrderData[0].length).setValues(filledOrderData);
    filledOrdersSheet.autoResizeColumns(1, filledOrderData[0].length);
    Logger.log("Successfully wrote new filled order data to 'Filled Orders' sheet.");
  } else {
    Logger.log("No NEW 'filled' orders to write to 'Filled Orders' sheet.");
  }
  Logger.log("--- Finished updatefilledSheet ---");
}


/**
 * Copies the "Filled Orders" tab to a new sheet,
 * transforms the data by adjusting quantity for sell orders,
 * calculates position net cost,
 * orders the data by symbol, and then groups it by symbol
 * showing the total adjusted quantity and total position net cost.
 */
function processFilledOrdersSnapshot() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var filledOrdersSheet = ss.getSheetByName("Filled Orders");

  if (!filledOrdersSheet) {
    Logger.log("Error: 'Filled Orders' sheet not found. Please run updatefilledSheet() first.");
    SpreadsheetApp.getUi().alert("Error", "'Filled Orders' sheet not found. Please run updatefilledSheet() first.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Create a new sheet with a timestamped name
  var timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyyMMdd_HHmmss");
  var newSheetName = "Filled_Summary_" + timestamp;
  var summarySheet = ss.insertSheet(newSheetName);
  Logger.log("Created new summary sheet: " + newSheetName);

  // Get all data from "Filled Orders" sheet (including headers)
  var allFilledData = filledOrdersSheet.getDataRange().getValues();

  if (allFilledData.length <= 1) { // Only headers or no data
    Logger.log("No filled order data to process beyond headers.");
    summarySheet.appendRow(["No filled orders to summarize."]);
    return;
  }

  // Extract headers and data rows
  var headers = allFilledData[0];
  var dataRows = allFilledData.slice(1); // All rows except the header

  // Find column indices dynamically
  var symbolCol = headers.indexOf("Symbol");
  var quantityCol = headers.indexOf("Quantity");
  var sideCol = headers.indexOf("Side");
  var filledAvgPriceCol = headers.indexOf("Filled Avg Price"); // New: Index for Filled Avg Price

  if (symbolCol === -1 || quantityCol === -1 || sideCol === -1 || filledAvgPriceCol === -1) {
    Logger.log("Error: Required columns (Symbol, Quantity, Side, Filled Avg Price) not found in 'Filled Orders' sheet.");
    SpreadsheetApp.getUi().alert("Error", "Required columns (Symbol, Quantity, Side, Filled Avg Price) not found in 'Filled Orders' sheet.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // --- Step 1: Add Adjusted Quantity column and prepare for sorting ---
  var processedData = [];
  dataRows.forEach(function(row) {
    var symbol = row[symbolCol];
    var quantity = parseFloat(row[quantityCol]); // Ensure quantity is a number
    var side = row[sideCol];
    var filledAvgPrice = parseFloat(row[filledAvgPriceCol]); // Ensure filledAvgPrice is a number

    var adjustedQuantity = quantity;
    if (side && typeof side === 'string' && side.toLowerCase() === 'sell') {
      adjustedQuantity = -quantity;
    }

    // Step 1b: Calculate Position Net Cost
    var positionNetCost = adjustedQuantity * filledAvgPrice;
    // Handle potential NaN if filledAvgPrice is not a valid number
    if (isNaN(positionNetCost)) {
      positionNetCost = 0; // Default to 0 or handle as appropriate
      Logger.log("Warning: Position Net Cost calculated as NaN for order: " + symbol + ". Check 'Filled Avg Price' data.");
    }


    // Create a new row with original data + adjusted quantity + position net cost
    var newRow = row.slice(); // Create a copy of the original row
    newRow.push(adjustedQuantity); // Add adjusted quantity as a new column
    newRow.push(positionNetCost); // Add position net cost as another new column
    processedData.push(newRow);
  });

  // Add "Adjusted Quantity" and "Position Net Cost" to headers
  var newHeaders = headers.slice();
  newHeaders.push("Adjusted Quantity");
  newHeaders.push("Position Net Cost"); // New header

  // --- Step 2: Order the data by Symbol ---
  // Sort by Symbol (assuming symbolCol is 1 for 'Symbol' column in the original headers)
  processedData.sort(function(a, b) {
    var symbolA = a[symbolCol].toString().toUpperCase(); // Ensure string comparison
    var symbolB = b[symbolCol].toString().toUpperCase();
    if (symbolA < symbolB) return -1;
    if (symbolA > symbolB) return 1;
    return 0;
  });
  Logger.log("Data sorted by Symbol.");

  // --- Step 3: Group the data by Symbol and sum Adjusted Quantity and Position Net Cost ---
  var groupedData = {}; // Object to store sums by symbol

  // The adjusted quantity is now the second to last element, and net cost is the last
  var adjustedQuantityIndex = newHeaders.length - 2; 
  var positionNetCostIndex = newHeaders.length - 1;

  processedData.forEach(function(row) {
    var symbol = row[symbolCol];
    var adjQty = row[adjustedQuantityIndex];
    var netCost = row[positionNetCostIndex];

    if (!groupedData[symbol]) {
      groupedData[symbol] = { totalAdjQty: 0, totalNetCost: 0 };
    }
    groupedData[symbol].totalAdjQty += adjQty;
    groupedData[symbol].totalNetCost += netCost;
  });
  Logger.log("Data grouped by Symbol.");

  // Prepare data for the summary sheet
  var summaryData = [];
  summaryData.push(["Symbol", "Total Adjusted Quantity", "Total Position Net Cost"]); // Updated Summary headers

  for (var symbol in groupedData) {
    summaryData.push([symbol, groupedData[symbol].totalAdjQty, groupedData[symbol].totalNetCost]);
  }

  // Write summary data to the new sheet
  summarySheet.getRange(1, 1, summaryData.length, summaryData[0].length).setValues(summaryData);
  summarySheet.autoResizeColumns(1, summaryData[0].length);
  // Format the quantity and cost columns
  summarySheet.getRange(2, 2, summaryData.length -1, 1).setNumberFormat("#,###"); // Total Adjusted Quantity
  summarySheet.getRange(2, 3, summaryData.length -1, 1).setNumberFormat("$#,##0.00"); // Total Position Net Cost
  Logger.log("Summary data written to " + newSheetName);
}
