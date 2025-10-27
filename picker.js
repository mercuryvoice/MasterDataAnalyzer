// @ts-check

/**
 * 獲取 API 金鑰、App ID 和 OAuth 權杖。
 * (V14：移除了 SpreadsheetApp.getUi().alert() 以修復 TransportError)
 * @returns {{apiKey: string, appId: string, oauthToken: string}}
 */
function getPickerKeys() {
  try {
    const userProperties = PropertiesService.getScriptProperties();
    const apiKey = userProperties.getProperty('GOOGLE_API_KEY');
    const appId = userProperties.getProperty('GOOGLE_APP_ID'); // <-- 必須是 12 位數字的 Project Number
    const oauthToken = ScriptApp.getOAuthToken();

    if (!apiKey || !appId) {
      throw new Error("API Key or App ID not found in Script Properties. Please set 'GOOGLE_API_KEY' and 'GOOGLE_APP_ID'.");
    }

    if (!oauthToken) {
      throw new Error("Could not retrieve OAuth token. Please ensure the add-on is authorized.");
    }

    return {
      apiKey: apiKey,
      appId: appId,
      oauthToken: oauthToken
    };
  } catch (e) {
    Logger.log(`Error getting Picker keys: ${e.message}`);
    // **重要**：在此處呼叫 .alert() 會導致 google.script.run 失敗
    // SpreadsheetApp.getUi().alert('獲取 Picker 憑證失敗：\n' + e.message);
    throw e; // 將錯誤傳回給 HTML 的 .withFailureHandler
  }
}

/**
 * 開啟一個對話方塊來測試 Google Picker。
 * (V14：返回 IFRAME 模式)
 */
function showPickerTestDialog_V14() {
  // @ts-ignore
  const T = MasterData.getTranslations();
  
  const htmlTemplate = HtmlService.createTemplateFromFile('picker_ui.html');
  htmlTemplate.T = T;
  
  // V14：使用 IFRAME 模式
  // 這是唯一的解決方案，因為 NATIVE 模式會載入損壞的 API
  const htmlOutput = htmlTemplate.evaluate()
      .setWidth(600)
      .setHeight(450)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME); // <-- 必須使用 IFRAME
  
  // 顯示對話方塊
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Google Picker Test (V14)');
}

/**
 * [NEW] 處理從 Picker 選擇的檔案，並驗證 drive.file 權限。
 * @param {string} fileId 來自 Google Picker 的檔案 ID。
 * @returns {string} 成功或失敗的訊息。
 */
function processPickedFile(fileId) {
  try {
    if (!fileId) {
      throw new Error("File ID is missing.");
    }
    
    // 使用 DriveApp 存取檔案，這一步會驗證 drive.file 權限是否生效
    const file = DriveApp.getFileById(fileId);
    const fileName = file.getName();
    
    Logger.log(`Successfully accessed file: "${fileName}" (ID: ${fileId})`);
    
    // 將成功訊息傳回給前端
    return `成功讀取檔案： "${fileName}"`;
    
  } catch (e) {
    Logger.log(`Error processing picked file (ID: ${fileId}): ${e.message}`);
    // 將錯誤訊息傳回給前端
    throw new Error(`無法讀取檔案。請確認您已授權 drive.file 權限。\n錯誤詳情： ${e.message}`);
  }
}