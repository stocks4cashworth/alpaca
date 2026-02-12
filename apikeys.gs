// --- GLOBAL VARIABLES ---
var ALPAC_API_KEY_ID;
var ALPAC_API_SECRET_KEY;
var ALPAC_API_ENDPOINT;

/**
 * Safely loads credentials from Script Properties into global variables.
 */
function initializeCredentials() {
  var props = PropertiesService.getScriptProperties();
  // Use String() to ensure nulls become empty strings, preventing "Header:null" errors
  ALPAC_API_KEY_ID = String(props.getProperty('ALPACA_KEY') || "");
  ALPAC_API_SECRET_KEY = String(props.getProperty('ALPACA_SECRET') || "");
  ALPAC_API_ENDPOINT = props.getProperty('ALPACA_ENDPOINT') || "https://paper-api.alpaca.markets/";
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  initializeCredentials();

  ui.createMenu('🚀 Alpaca Tools')
      .addItem('📊 Update Portfolio Now', 'updateSheet')
      .addSeparator()
      .addItem('🔐 Setup API Keys', 'promptForApiKeys')
      .addItem('🧪 Switch to Paper Trading', 'usePaperTrading')
      .addItem('💰 Switch to Live Trading', 'useLiveTrading')
      .addToUi();

  if (!ALPAC_API_KEY_ID || !ALPAC_API_SECRET_KEY) {
    promptForApiKeys();
  }
}

function promptForApiKeys() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  
  var keyResp = ui.prompt('Security Setup', 'Enter Alpaca API Key ID:', ui.ButtonSet.OK_CANCEL);
  if (keyResp.getSelectedButton() !== ui.Button.OK) return;
  
  var secretResp = ui.prompt('Security Setup', 'Enter Alpaca Secret Key:', ui.ButtonSet.OK_CANCEL);
  if (secretResp.getSelectedButton() !== ui.Button.OK) return;

  props.setProperty('ALPACA_KEY', keyResp.getResponseText().trim());
  props.setProperty('ALPACA_SECRET', secretResp.getResponseText().trim());
  
  initializeCredentials();
  ui.alert('Keys saved successfully.');
}

function usePaperTrading() {
  PropertiesService.getScriptProperties().setProperty('ALPACA_ENDPOINT', "https://paper-api.alpaca.markets/");
  initializeCredentials();
  SpreadsheetApp.getUi().alert("Endpoint set to PAPER.");
}

function useLiveTrading() {
  PropertiesService.getScriptProperties().setProperty('ALPACA_ENDPOINT', "https://api.alpaca.markets/");
  initializeCredentials();
  SpreadsheetApp.getUi().alert("Endpoint set to LIVE.");
}
