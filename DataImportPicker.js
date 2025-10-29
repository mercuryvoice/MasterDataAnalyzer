// @ts-check

// This file is a new test version combining MasterDataImport.js and Picker.js functionalities.

/**
 * Gets the API Key, App ID, and OAuth Token.
 * @returns {{apiKey: string, appId: string, oauthToken: string}}
 */
function getPickerKeys() {
  try {
    const userProperties = PropertiesService.getScriptProperties();
    const apiKey = userProperties.getProperty('GOOGLE_API_KEY');
    const appId = userProperties.getProperty('GOOGLE_APP_ID');
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
    throw e;
  }
}

/**
 * Saves Import Settings for a specific sheet.
 * @param {object} settings The settings object from the UI.
 * @param {string} sheetName The target sheet name.
 * @returns {{success: boolean, message: string}} Result object.
 */
function saveImportPickerSettings(settings, sheetName) {
    const T = MasterData.getTranslations();
    try {
        if (!sheetName) {
            throw new Error("Sheet name is required to save settings.");
        }
        if (!settings.sourceFileId || !settings.sourceSheetName || !settings.targetSheetName) {
            throw new Error(T.errorUrlRequired); // You might want to change this error message to be more generic
        }

        // Validation logic can be adapted here if needed

        const properties = PropertiesService.getDocumentProperties();
        const key = `importSettings_${sheetName}`;
        properties.setProperty(key, JSON.stringify(settings));

        return {
            success: true,
            message: T.saveSuccess
        };
    } catch (e) {
        Logger.log(`Error saving import settings for sheet ${sheetName}: ${e.message}`);
        return {
            success: false,
            message: `${T.saveFailure}: ${e.message}`
        };
    }
}

/**
 * Gets Import Settings for a specific sheet.
 * @param {string} sheetName The target sheet name.
 * @returns {object} The saved settings object.
 */
function getImportSettings(sheetName) {
    const currentSheetName = sheetName || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
    const properties = PropertiesService.getDocumentProperties();
    const key = `importSettings_${currentSheetName}`;
    const settingsString = properties.getProperty(key);

    let settings;
    if (!settingsString) {
        settings = {};
    } else {
        try {
            settings = JSON.parse(settingsString);
        } catch (e) {
            Logger.log(`Error parsing import settings for sheet ${currentSheetName}: ${e.message}`);
            settings = {};
        }
    }

    return {
        sourceFileId: settings.sourceFileId || '',
        sourceFileName: settings.sourceFileName || '',
        sourceSheetName: settings.sourceSheetName || '',
        targetSheetName: settings.targetSheetName || currentSheetName,
        targetHeaderRow: settings.targetHeaderRow ? parseInt(settings.targetHeaderRow, 10) : '',
        targetStartRow: settings.targetStartRow ? parseInt(settings.targetStartRow, 10) : '',
        sourceIdentifierRange: settings.sourceIdentifierRange || '',
        sourceHeaderRange: settings.sourceHeaderRange || '',
        sourceValueMatrixRange: settings.sourceValueMatrixRange || '',
        rawImportFilterHeaders: settings.rawImportFilterHeaders || '',
        keywordFilters: settings.keywordFilters || []
    };
}

/**
 * Gets all sheet names from a given spreadsheet ID.
 * @param {string} fileId The ID of the spreadsheet.
 * @returns {string[]} An array of sheet names.
 */
function getSheetNamesFromFileId(fileId) {
    const T = MasterData.getTranslations();
    try {
        // Handle the case where we want sheets from the currently active spreadsheet
        if (!fileId) {
            const activeSs = SpreadsheetApp.getActiveSpreadsheet();
            if (activeSs) {
                return activeSs.getSheets().map(sheet => sheet.getName());
            }
            throw new Error("No active spreadsheet found.");
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

/**
 * Fetches header options for the filter dropdown.
 */
function getFilterHeaderOptions(fileId, sourceSheetName, sourceIdentifierRange, importFilterHeadersString) {
    if (!fileId || !sourceSheetName) return [];

    try {
        const sourceSs = SpreadsheetApp.openById(fileId);
        const sourceSh = sourceSs.getSheetByName(sourceSheetName);
        if (!sourceSh) return [];

        if (importFilterHeadersString) {
            if (importFilterHeadersString.includes(':') || importFilterHeadersString.includes('：')) {
                const rangeString = importFilterHeadersString.replace(/：/g, ':');
                const headers = sourceSh.getRange(rangeString).getValues()[0];
                return headers.map(h => h.toString().trim()).filter(h => h);
            } else {
                return importFilterHeadersString.split(/,|，/g).map(h => h.trim()).filter(h => h);
            }
        }

        if (sourceIdentifierRange) {
            const idRange = sourceSh.getRange(sourceIdentifierRange);
            const headerRow = idRange.getRow() - 1;
            if (headerRow < 1) return [];
            const idHeaderRange = sourceSh.getRange(headerRow, idRange.getColumn(), 1, idRange.getNumColumns());
            const headers = idHeaderRange.getValues()[0];
            return headers.map(h => h.toString().trim()).filter(h => h);
        }

        return [];
    } catch (e) {
        Logger.log(`Error in getFilterHeaderOptions: ${e.message}`);
        return [];
    }
}


/**
 * Fetches unique data values from a specified column for dependent dropdowns.
 */
function getUniqueValuesForHeader(fileId, sourceSheetName, sourceIdentifierRange, headerName) {
    if (!fileId || !sourceSheetName || !sourceIdentifierRange || !headerName) {
        return [];
    }
    try {
        const sourceSs = SpreadsheetApp.openById(fileId);
        const sourceSh = sourceSs.getSheetByName(sourceSheetName);
        if (!sourceSh) return [];

        const idRange = sourceSh.getRange(sourceIdentifierRange);
        const headerRow = idRange.getRow() - 1;
        if (headerRow < 1) return [];

        const idHeaderRange = sourceSh.getRange(headerRow, idRange.getColumn(), 1, idRange.getNumColumns());
        const headers = idHeaderRange.getValues()[0].map(h => h.toString().trim());

        const headerIndex = headers.indexOf(headerName);
        if (headerIndex === -1) {
            return [];
        }

        const columnIndex = idRange.getColumn() + headerIndex;
        const values = sourceSh.getRange(idRange.getRow(), columnIndex, idRange.getNumRows(), 1)
            .getValues()
            .flat()
            .map(v => v.toString().trim())
            .filter(v => v);

        return [...new Set(values)];
    } catch (e) {
        Logger.log(`Error in getUniqueValuesForHeader: ${e.message}`);
        return [];
    }
}

/**
 * Validates the core source and target fields.
 */
function validateSourceAndTarget(fileId, sourceSheetName, targetSheetName) {
    const T = MasterData.getTranslations();
    const errors = {
        sourceFileIdError: '',
        sourceSheetError: '',
        targetSheetError: ''
    };

    if (targetSheetName) {
        const activeSs = SpreadsheetApp.getActiveSpreadsheet();
        if (!activeSs.getSheetByName(targetSheetName)) {
            errors.targetSheetError = T.errorTargetSheetNotFound.replace('{SHEET_NAME}', targetSheetName);
        }
    }

    if (fileId && sourceSheetName) {
        try {
            const sourceSs = SpreadsheetApp.openById(fileId);
            if (!sourceSs.getSheetByName(sourceSheetName)) {
                errors.sourceSheetError = T.errorSheetNotFoundInUrl.replace('{SHEET_NAME}', sourceSheetName);
            }
        } catch (e) {
            errors.sourceFileIdError = T.errorInvalidUrl; // Or a better error message
        }
    }

    return errors;
}

/**
 * Validates all user inputs for the Import UI.
 */
function validateAllInputs(inputs) {
    const T = MasterData.getTranslations();
    const {
        sourceFileId,
        sourceSheetName,
        sourceIdentifierRange,
        importFilterHeaders,
        keywordFilters
    } = inputs;
    const results = {
        duplicateHeaderWarning: '',
        importFilterErrors: [],
        keywordFilterErrors: []
    };

    if (!sourceFileId || !sourceSheetName) return results;

    try {
        const sourceSs = SpreadsheetApp.openById(sourceFileId);
        const sourceSh = sourceSs.getSheetByName(sourceSheetName);
        if (!sourceSh) return results;

        const sourceHeaders = getFilterHeaderOptions(sourceFileId, sourceSheetName, sourceIdentifierRange, importFilterHeaders);

        if (!importFilterHeaders && sourceIdentifierRange) {
            const idRange = sourceSh.getRange(sourceIdentifierRange);
            const headerRow = idRange.getRow() - 1;
            if (headerRow > 0) {
                const rangeHeaders = sourceSh.getRange(headerRow, idRange.getColumn(), 1, idRange.getNumColumns()).getValues()[0];
                const seen = new Set();
                const duplicates = new Set();
                rangeHeaders.forEach(h => {
                    const trimmedHeader = h.toString().trim();
                    if (trimmedHeader) {
                        if (seen.has(trimmedHeader)) {
                            duplicates.add(trimmedHeader);
                        }
                        seen.add(trimmedHeader);
                    }
                });
                if (duplicates.size > 0) {
                    results.duplicateHeaderWarning = T.duplicateHeaderWarning.replace('{HEADERS}', [...duplicates].join(', '));
                }
            }
        }

        if (importFilterHeaders) {
            const allSourceHeaders = sourceSh.getDataRange().getValues().flat().map(h => h.toString().trim());
            const inputHeaders = importFilterHeaders.includes(':') ? getFilterHeaderOptions(sourceFileId, sourceSheetName, null, importFilterHeaders) : importFilterHeaders.split(/,|，/g).map(h => h.trim());
            const invalidHeaders = inputHeaders.filter(h => !allSourceHeaders.includes(h));
            if (invalidHeaders.length > 0) {
                results.importFilterErrors = invalidHeaders;
            }
        }

        if (keywordFilters && keywordFilters.length > 0) {
            keywordFilters.forEach((filter, index) => {
                if (!filter.header) return;

                if (!sourceHeaders.includes(filter.header)) {
                    results.keywordFilterErrors.push({
                        index,
                        invalidHeader: filter.header
                    });
                    return;
                }

                if (filter.keywords) {
                    const validKeywords = getUniqueValuesForHeader(sourceFileId, sourceSheetName, sourceIdentifierRange, filter.header);
                    const inputKeywords = filter.keywords.split(/,|，/g).map(k => k.trim()).filter(k => k);
                    const invalidKeywords = inputKeywords.filter(k => !validKeywords.includes(k));

                    if (invalidKeywords.length > 0) {
                        results.keywordFilterErrors.push({
                            index,
                            invalidKeywords,
                            header: filter.header
                        });
                    }
                }
            });
        }

        return results;
    } catch (e) {
        Logger.log(`Error during validation: ${e.message}`);
        return results;
    }
}

/**
 * Shows the new Data Import Picker UI.
 */
function showDataImportPicker() {
    const T = MasterData.getTranslations();
    const htmlTemplate = HtmlService.createTemplateFromFile('DataImportPicker.html');
    htmlTemplate.T = T;
    const htmlOutput = htmlTemplate.evaluate().setWidth(450).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Data Import Settings (Picker)");
}
