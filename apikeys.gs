// --- PRIVATE CREDENTIALS ---
// Keep this file local; do not sync this file to your public repository.

var ALPAC_API_KEY_ID = "YOUR_LIVE_KEY_HERE"; 
var ALPAC_API_SECRET_KEY = "YOUR_LIVE_SECRET_HERE"; 
var ALPAC_API_ENDPOINT = "https://api.alpaca.markets/"; // Use 'https://paper-api.alpaca.markets/' for paper

/**
 * Simple trigger to create the menu on open.
 * We removed the safety checks since the keys are now hardcoded here.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Alpaca Tools')
      .addItem('📊 Update Portfolio Now', 'updateSheet')
      .addSeparator()
      .addItem('🧪 Switch to Paper Trading', 'usePaperTrading')
      .addItem('💰 Switch to Live Trading', 'useLiveTrading')
      .addToUi();
}

// These functions now update the variable directly for the current session
function usePaperTrading() {
  ALPAC_API_ENDPOINT = "https://paper-api.alpaca.markets/";
  SpreadsheetApp.getUi().alert("Endpoint set to PAPER. Note: Hardcoded variables reset on refresh.");
}

function useLiveTrading() {
  ALPAC_API_ENDPOINT = "https://api.alpaca.markets/";
  SpreadsheetApp.getUi().alert("Endpoint set to LIVE.");
}
