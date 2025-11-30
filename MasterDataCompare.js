// @ts-check

/**
 * Gets the API Key, App ID, and OAuth Token for the Compare feature.
 * @returns {{apiKey: string, appId: string, oauthToken: string}}
 */
function compare_getPickerKeys() {
  const T = MasterData.getTranslations();
  try {
    const userProperties = PropertiesService.getScriptProperties();
    const apiKey = userProperties.getProperty('GOOGLE_API_KEY');
    const appId = userProperties.getProperty('GOOGLE_APP_ID');
    const oauthToken = ScriptApp.getOAuthToken();

    if (!apiKey || !appId) {
      throw new Error(T.errorApiKeyOrAppIdNotFound || "API Key or App ID not found in Script Properties. Please set 'GOOGLE_API_KEY' and 'GOOGLE_APP_ID'.");
    }

    if (!oauthToken) {
      throw new Error(T.errorOAuthToken || "Could not retrieve OAuth token. Please ensure the add-on is authorized.");
    }

    return {
      apiKey: apiKey,
      appId: appId,
      oauthToken: oauthToken
    };
  } catch (e) {
    Logger.log(`Error getting Picker keys: ${e.message}`);
    throw e;
  }
}

/**
 * Gets all sheet names from a given spreadsheet ID for the Compare feature.
 * @param {string} fileId The ID of the spreadsheet.
 * @returns {string[]} An array of sheet names.
 */
function compare_getSheetNames(fileId) {
    const T = MasterData.getTranslations();
    try {
        // Handle the case where we want sheets from the currently active spreadsheet
        if (!fileId) {
            const activeSs = SpreadsheetApp.getActiveSpreadsheet();
            if (activeSs) {
                return activeSs.getSheets().map(sheet => sheet.getName());
            }
            throw new Error(T.errorNoActiveSpreadsheet || "No active spreadsheet found.");
        }

        // Use the Sheets API for files selected via Picker
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${fileId}?fields=sheets.properties.title`;
        const options = {
            method: 'get',
            headers: {
                Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
            },
            muteHttpExceptions: true,
        };

        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        const responseBody = response.getContentText();

        if (responseCode === 200) {
            const data = JSON.parse(responseBody);
            if (data.sheets && data.sheets.length > 0) {
                return data.sheets.map(sheet => sheet.properties.title);
            }
            throw new Error(T.noSheetsFound);
        } else {
            Logger.log(`Sheets API Error for fileId ${fileId}: ${responseCode} - ${responseBody}`);
            throw new Error(T.errorInvalidUrl + ` (API Error: ${responseCode})`);
        }

    } catch (e) {
        Logger.log(`Critical Error in getSheetNamesFromFileId for fileId '${fileId}': ${e.message}`);
        // Provide a more specific error message based on the error type if possible
        if (e.message.includes(T.noSheetsFound)) {
            throw e;
        }
        throw new Error(T.errorInvalidUrl);
    }
}
