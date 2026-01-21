// FilledOrders.gs
/**
 * Fetches filled orders from Alpaca, filters for new ones, and writes them
 * to the "Filled Orders" sheet with all available data points.
 */
function updatefilledSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var filledOrdersSheet = ss.getSheetByName("Filled Orders");

  Logger.log("--- Starting updatefilledSheet (for Filled Orders) ---");

  // Create sheet if it doesn't exist
  if (!filledOrdersSheet) {
    Logger.log("Creating new 'Filled Orders' sheet.");
    filledOrdersSheet = ss.insertSheet("Filled Orders");
  } else {
    Logger.log("Found 'Filled Orders' sheet. Clearing existing data (A2:Z).");
    filledOrdersSheet.getRange("A2:Z").clearContent(); // Clear all data below header
  }

  // Define comprehensive headers for filled orders
  var headers = [
    "ID", "Symbol", "Quantity", "Side", "Type", "Time in Force", "Limit Price", "Stop Price",
    "Status", "Submitted At", "Created At", "Updated At", "Expired At", "Filled At",
    "Client Order ID", "Extended Hours", "Asset ID", "Asset Class", "Filled Quantity",
    "Filled Avg Price", "Order Class", "Trail Price", "Trail Percent", "HWM", "Commission",
    "Legs" 
  ];
  filledOrdersSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  filledOrdersSheet.autoResizeColumns(1, headers.length); // Auto-resize headers
  Logger.log("Headers written to 'Filled Orders' sheet.");

  // --- Retrieve existing filled order IDs from the sheet to prevent duplicates ---
  // This logic is commented out for now to ensure all filled orders from API are written,
  // as the sheet is cleared initially. If you want true appending without clearing,
  // uncomment this and adjust the clearContent() above.
  /*
  var existingOrderIds = new Set();
  if (filledOrdersSheet.getLastRow() > 1) { 
    var existingDataRange = filledOrdersSheet.getRange("A2:A" + filledOrdersSheet.getLastRow());
    existingDataRange.getValues().forEach(function(row) {
      if (row[0]) { 
        existingOrderIds.add(row[0]);
      }
    });
  }
  Logger.log("Number of existing filled order IDs on sheet: " + existingOrderIds.size);
  */

  // --- Step 1: Get all orders (limited to last 30 days by listOrders) ---
  Logger.log("Calling listOrders() to retrieve ALL orders from Alpaca (last 30 days).");
  // listOrders() is defined in Main.gs and is globally accessible
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

  // --- Step 2: Filter for ONLY filled orders ---
  Logger.log("Filtering for orders with status 'filled'.");
  var filledOrders = orders.filter(function(order) {
    var isFilled = (order.status === "filled");
    // If using incremental update, also check isNew: var isNew = !existingOrderIds.has(order.id);
    if (isFilled) { // && isNew
      Logger.log("  MATCH: Order ID: " + order.id + ", Status: " + order.status);
    }
    return isFilled; // && isNew;
  });

  Logger.log("Number of 'filled' orders found: " + filledOrders.length);

  // --- Step 3: Write filled orders data to sheet ---
  if (filledOrders.length > 0) {
    Logger.log("Preparing data for " + filledOrders.length + " filled orders.");
    var filledOrderData = filledOrders.map(function(order) {
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
        order.updated_at || "",
        order.expired_at || "",
        order.filled_at || "",
        order.client_order_id || "",
        order.extended_hours || "",
        order.asset_id || "",
        order.asset_class || "",
        order.filled_qty || "",
        order.filled_avg_price || "",
        order.order_class || "",
        order.trail_price || "",
        order.trail_percent || "",
        order.hwm || "",
        order.commission || "",
        (order.legs && order.legs.length > 0) ? JSON.stringify(order.legs) : "" 
      ];
    });
    filledOrdersSheet.getRange(2, 1, filledOrderData.length, filledOrderData[0].length).setValues(filledOrderData);
    filledOrdersSheet.autoResizeColumns(1, filledOrderData[0].length);
    Logger.log("Successfully wrote filled order data to 'Filled Orders' sheet.");
  } else {
    Logger.log("No 'filled' orders to write to 'Filled Orders' sheet.");
  }
  Logger.log("--- Finished updatefilledSheet ---");
}
