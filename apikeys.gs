// --- GLOBAL CONFIGURATION ---
// These will still run every time a script is triggered
var scriptProperties = PropertiesService.getScriptProperties();
var ALPAC_API_KEY_ID = scriptProperties.getProperty('ALPACA_KEY'); 
var ALPAC_API_SECRET_KEY = scriptProperties.getProperty('ALPACA_SECRET'); 
var ALPAC_API_ENDPOINT = scriptProperties.getProperty('ALPACA_ENDPOINT'); 

/**
 * Simple trigger that runs when the spreadsheet is opened.
 */
function onOpen() {
  // 1. Run the Safety Check immediately
  checkCredentials();

  // 2. Create the Menu
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Alpaca Tools')
      .addItem('🔐 Set/Update API Credentials', 'uiSetApiKeys')
      .addSeparator()
      .addItem('📊 Update Portfolio Now', 'updateSheet')
      .addToUi();
}

/**
 * Validation Logic
 */
function checkCredentials() {
  if (!ALPAC_API_KEY_ID || !ALPAC_API_SECRET_KEY || !ALPAC_API_ENDPOINT) {
    // We use an alert instead of 'throw' here so it pops up nicely for the user
    SpreadsheetApp.getUi().alert("⚠️ Missing API Credentials. Please go to 'Alpaca Tools' > 'Set API Credentials' to begin.");
    return false;
  }
  return true;
}
