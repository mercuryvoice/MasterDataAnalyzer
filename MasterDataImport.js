// @ts-check

// =================================================================================================
// =================================== VERSION & FEATURE SUMMARY ===================================
// =================================================================================================
//
// V1.0 (Pre-release version):
// - Noted.
//
// =================================================================================================

/**
 * MasterDataAnalyzer - A Google Sheets Add-on for intelligent data operations.
 *
 * Copyright (c) 2025 Tata Sum (mda.design)
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by

 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program. If not, see <https://www.gnu.org/licenses/>.
 */


// =================================================================================================
// ===================================== SECTION 1: USER INTERFACE =================================
// =================================================================================================

/**
 * [REFACTORED] Adds all custom menus under a single "MasterDataAnalyzer" menu.
 */
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    const T = MasterData.getTranslations();

    const mainMenu = ui.createMenu(T.mainMenuTitle);

    // --- Sub-Menu: Data Import Tool ---
    const importSubMenu = ui.createMenu(T.importMenuTitle)
        .addItem(T.runImportItem, 'runImportProcess')
        .addItem(T.settingsItem, 'showImportSettingsSidebar')
        .addSeparator()
        .addItem(T.runCompareItem, 'runCompareProcess')
        .addItem(T.compareSettingsItem, 'showCompareSettingsSidebar')
        .addSeparator()
        .addItem(T.stopImportItem, 'requestStopImport')
        .addItem(T.resetImportItem, 'resetAllState');
    
    // --- Sub-Menu: Data Validation Tool ---
    const validationSubMenu = ui.createMenu(T.validationMenuTitle)
        .addItem(T.runMsModeItem, 'runValidationMsOnly')
        .addItem(T.verifySettingsItem, 'showVerifySettingsSidebar')
        .addSeparator()
        .addItem(T.stopValidationItem, 'requestStopValidation')
        .addSeparator()
        .addItem(T.verifySumsItem, 'verifySumAndCumulativeValues')
        .addSeparator()
        .addItem(T.cleanupItem, 'runReset'); // UPDATED to call the new reset function

    // --- Sub-Menu: Data Management Tool ---
    const monitorSubMenu = ui.createMenu(T.monitorMenuName)
        .addItem(T.enableNotifyItem, 'createOnChangeTrigger')
        .addItem(T.disableNotifyItem, 'deleteOnChangeTrigger');

    const managementSubMenu = ui.createMenu(T.manageMenuTitle)
        .addItem(T.manageSettingsItem, 'showManageSettingsSidebar')
        .addSeparator()
        .addSubMenu(monitorSubMenu)
        .addSeparator()
        .addItem(T.reportSettingsItem, 'showReportSettingsDialog')
        .addSeparator()
        .addItem(T.quickDeleteItem, 'showQuickDeleteSheetUI');

    // --- [NEW] Sub-Menu: Guides & Examples ---
    const guideSubMenu = ui.createMenu(T.guideMenuTitle)
        .addItem(T.businessGuide, 'generateBusinessExample')
        .addItem(T.manufacturingGuide, 'generateManufacturingExample')
        // .addItem(T.hrGuide, 'showHrGuide')
        .addSeparator()
        .addItem(T.startBusinessGuide, 'startBusinessTutorial')
        .addItem(T.startManufacturingGuide, 'startManufacturingTutorial')
        .addSeparator()
        .addItem(T.deleteExamplesItem, 'deleteExampleSheets');

    // --- Add all sub-menus and items to the main menu ---
    mainMenu.addSubMenu(importSubMenu);
    mainMenu.addSubMenu(validationSubMenu);
    mainMenu.addSubMenu(managementSubMenu);
    mainMenu.addSeparator();
    mainMenu.addSubMenu(guideSubMenu);
    mainMenu.addSeparator();
    mainMenu.addItem(T.privacyPolicyItem, 'showPrivacyPolicy');

    mainMenu.addToUi();
}

// function showHrGuide() { SpreadsheetApp.getUi().alert('人資管理範例即將推出！'); }
/**
 * Shows the Privacy Policy UI.
 */
function showPrivacyPolicy() {
    const T = MasterData.getTranslations();
    const htmlTemplate = HtmlService.createTemplateFromFile('PrivacyPolicy.html');
    htmlTemplate.T = T;
    const htmlOutput = htmlTemplate.evaluate().setWidth(600).setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, T.privacyPolicyTitle);
}


/**
 * Shows the HTML settings for Data Import.
 */
function showImportSettingsSidebar() {
    const T = MasterData.getTranslations();
    const htmlTemplate = HtmlService.createTemplateFromFile('SettingsPageImport');
    htmlTemplate.T = T;
    const htmlOutput = htmlTemplate.evaluate().setWidth(450).setHeight(600);
    SpreadsheetApp.getUi().showModelessDialog(htmlOutput, T.importSettingsTitle);
}

/**
 * Shows the HTML settings for Data Comparison.
 */
function showCompareSettingsSidebar() {
    const T = MasterData.getTranslations();
    const htmlTemplate = HtmlService.createTemplateFromFile('SettingsPageCompare.html'); // Ensure this file exists
    htmlTemplate.T = T;
    const htmlOutput = htmlTemplate.evaluate().setWidth(450).setHeight(600);
    SpreadsheetApp.getUi().showModelessDialog(htmlOutput, T.compareSettingsTitle);
}

/**
 * Shows the HTML settings for Report Generation.
 */
function showReportSettingsDialog() {
    const T = MasterData.getTranslations();
    const htmlTemplate = HtmlService.createTemplateFromFile('SettingPageReport.html');
    htmlTemplate.T = T; // Pass translations to the HTML file
    const htmlOutput = htmlTemplate.evaluate().setWidth(500).setHeight(650);
    SpreadsheetApp.getUi().showModelessDialog(htmlOutput, T.reportSettingsTitle);
}

// =================================================================================================
// ===================================== PropertiesService Functions ===============================
// =================================================================================================

/**
 * Saves Import Settings for a specific sheet.
 * @param {object} settings The settings object from the UI.
 * @param {string} sheetName The target sheet name.
 * @returns {{success: boolean, message: string}} Result object.
 */
function saveImportSettings(settings, sheetName) {
    const T = MasterData.getTranslations();
    try {
        if (!sheetName) {
            throw new Error("Sheet name is required to save settings.");
        }
        if (!settings.sourceUrl || !settings.sourceSheetName || !settings.targetSheetName) {
            throw new Error(T.errorUrlRequired);
        }

        const validationResults = validateAllInputs({
            sourceUrl: settings.sourceUrl,
            sourceSheetName: settings.sourceSheetName,
            sourceIdentifierRange: settings.sourceIdentifierRange,
            importFilterHeaders: settings.rawImportFilterHeaders,
            keywordFilters: settings.keywordFilters
        });

        let errorMessages = [];
        if (validationResults.importFilterErrors && validationResults.importFilterErrors.length > 0) {
            errorMessages.push(T.invalidHeaders.replace('{HEADERS}', validationResults.importFilterErrors.join(', ')));
        }
        if (validationResults.keywordFilterErrors && validationResults.keywordFilterErrors.length > 0) {
            validationResults.keywordFilterErrors.forEach(err => {
                if (err.invalidHeader) {
                    errorMessages.push(T.invalidHeaders.replace('{HEADERS}', err.invalidHeader));
                } else if (err.invalidKeywords) {
                    errorMessages.push(T.invalidKeywords.replace('{HEADER}', err.header).replace('{KEYWORDS}', err.invalidKeywords.join(', ')));
                }
            });
        }

        if (errorMessages.length > 0) {
            return {
                success: false,
                message: errorMessages.join('\n')
            };
        }

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
        sourceUrl: settings.sourceUrl || '',
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
 * Saves Compare Settings for a specific sheet.
 * @param {object} settings The settings object from the UI.
 * @param {string} sheetName The name of the sheet to save settings for.
 * @returns {{success: boolean, error?: string}} Result object.
 */
function saveCompareSettings(settings, sheetName) {
    try {
        if (!sheetName) {
            throw new Error("Sheet name is required to save settings.");
        }
        if (!settings.sourceUrl || !settings.sourceSheetName || !settings.targetSheetName) {
            throw new Error("Source/Target URL and Sheet Name are required.");
        }
        if (!settings.sourceCompareRange || !settings.targetLookupCol || !settings.sourceLookupCol || !settings.sourceReturnCol || !settings.targetWriteCol) {
            throw new Error("All comparison and mapping fields are required.");
        }

        const properties = PropertiesService.getDocumentProperties();
        const key = `compareSettings_${sheetName}`;
        properties.setProperty(key, JSON.stringify(settings));
        return {
            success: true
        };
    } catch (e) {
        Logger.log(`Error saving compare settings for sheet ${sheetName}: ${e.message}`);
        return {
            success: false,
            error: e.message
        };
    }
}

/**
 * Gets Compare Settings for a specific sheet.
 * @param {string} sheetName The name of the sheet to get settings for.
 * @returns {object} The comparison settings object.
 */
function getCompareSettings(sheetName) {
    const currentSheetName = sheetName || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
    const properties = PropertiesService.getDocumentProperties();
    const key = `compareSettings_${currentSheetName}`;
    const settingsString = properties.getProperty(key);

    let settings;
    if (!settingsString) {
        settings = {};
    } else {
        try {
            settings = JSON.parse(settingsString);
        } catch (e) {
            Logger.log(`Error parsing compare settings for sheet ${currentSheetName}: ${e.message}`);
            settings = {};
        }
    }

    return {
        targetSheetName: settings.targetSheetName || currentSheetName,
        sourceUrl: settings.sourceUrl || '',
        sourceSheetName: settings.sourceSheetName || '',
        targetHeaderRow: settings.targetHeaderRow ? parseInt(settings.targetHeaderRow, 10) : '',
        targetStartRow: settings.targetStartRow ? parseInt(settings.targetStartRow, 10) : '',
        sourceCompareRange: settings.sourceCompareRange || '',
        targetLookupCol: settings.targetLookupCol || '',
        sourceLookupCol: settings.sourceLookupCol || '',
        sourceReturnCol: settings.sourceReturnCol || '',
        targetWriteCol: settings.targetWriteCol || ''
    };
}

/**
 * Gets the active sheet name for the UI.
 * @returns {string} The active sheet name.
 */
function getActiveSheetName() {
    return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}


/**
 * Shows the HTML settings for Data Validation.
 */
function showVerifySettingsSidebar() {
    const T = MasterData.getTranslations();
    const htmlTemplate = HtmlService.createTemplateFromFile('SettingsPageVerify.html');
    htmlTemplate.T = T;
    const htmlOutput = htmlTemplate.evaluate().setWidth(500).setHeight(650);
    SpreadsheetApp.getUi().showModelessDialog(htmlOutput, T.verifySettingsTitle);
}

/**
 * Gets all sheet names from a given spreadsheet URL.
 * @param {string} url The URL of the spreadsheet.
 * @returns {string[]} An array of sheet names.
 */
function getSheetNames(url) {
    const T = MasterData.getTranslations();
    try {
        const ss = url ? SpreadsheetApp.openByUrl(url) : SpreadsheetApp.getActiveSpreadsheet();
        const sheets = ss.getSheets();
        if (sheets.length === 0) {
            throw new Error(T.noSheetsFound);
        }
        return sheets.map(sheet => sheet.getName());
    } catch (e) {
        Logger.log(`Error in getSheetNames: ${e.message}`);
        throw new Error(T.errorInvalidUrl);
    }
}

/**
 * Gets settings for the Data Verify HTML interface.
 * This is a wrapper to call the function in the Data Verify script.
 * @param {string} sheetName The name of the sheet to get settings for.
 */
function getVerifySettingsForHtml(sheetName) {
    return getVerifySettings(sheetName);
}

/**
 * Gets default templates for notification emails.
 */
function getNotificationDefaultTemplates() {
    const T = MasterData.getTranslations();
    return {
        subject: T.defaultSubjectTemplate,
        body: T.defaultBodyTemplate
    };
}

/**
 * Fetches header options for the filter dropdown.
 */
function getFilterHeaderOptions(sourceUrl, sourceSheetName, sourceIdentifierRange, importFilterHeadersString) {
    if (!sourceUrl || !sourceSheetName) return [];

    try {
        const sourceSs = SpreadsheetApp.openByUrl(sourceUrl);
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
function getUniqueValuesForHeader(sourceUrl, sourceSheetName, sourceIdentifierRange, headerName) {
    if (!sourceUrl || !sourceSheetName || !sourceIdentifierRange || !headerName) {
        return [];
    }
    try {
        const sourceSs = SpreadsheetApp.openByUrl(sourceUrl);
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
function validateSourceAndTarget(sourceUrl, sourceSheetName, targetSheetName) {
    const T = MasterData.getTranslations();
    const errors = {
        sourceUrlError: '',
        sourceSheetError: '',
        targetSheetError: ''
    };

    if (targetSheetName) {
        const activeSs = SpreadsheetApp.getActiveSpreadsheet();
        if (!activeSs.getSheetByName(targetSheetName)) {
            errors.targetSheetError = T.errorTargetSheetNotFound.replace('{SHEET_NAME}', targetSheetName);
        }
    }

    if (sourceUrl) {
        const validUrlPattern = /^https:\/\/docs\.google\.com\/spreadsheets\/d\//;
        if (!validUrlPattern.test(sourceUrl)) {
            errors.sourceUrlError = T.errorInvalidUrl;
            return errors;
        }
    }

    if (sourceUrl && sourceSheetName && !errors.sourceUrlError) {
        try {
            const sourceSs = SpreadsheetApp.openByUrl(sourceUrl);
            if (!sourceSs.getSheetByName(sourceSheetName)) {
                errors.sourceSheetError = T.errorSheetNotFoundInUrl.replace('{SHEET_NAME}', sourceSheetName);
            }
        } catch (e) {
            errors.sourceUrlError = T.errorInvalidUrl;
        }
    }

    return errors;
}

/**
 * Validates the core source and target fields for the Compare UI.
 */
function validateCompareCoreInputs(sourceUrl, sourceSheetName, targetSheetName) {
    return validateSourceAndTarget(sourceUrl, sourceSheetName, targetSheetName);
}


/**
 * Validates the core source and target fields for the Verify UI.
 */
function validateVerifyInputs(sourceUrl, sourceSheetName, targetSheetName) {
    return validateSourceAndTarget(sourceUrl, sourceSheetName, targetSheetName);
}


/**
 * Validates all user inputs for the Import UI.
 */
function validateAllInputs(inputs) {
    const T = MasterData.getTranslations();
    const {
        sourceUrl,
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

    if (!sourceUrl || !sourceSheetName) return results;

    try {
        const sourceSs = SpreadsheetApp.openByUrl(sourceUrl);
        const sourceSh = sourceSs.getSheetByName(sourceSheetName);
        if (!sourceSh) return results;

        const sourceHeaders = getFilterHeaderOptions(sourceUrl, sourceSheetName, sourceIdentifierRange, importFilterHeaders);

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
            const inputHeaders = importFilterHeaders.includes(':') ? getFilterHeaderOptions(sourceUrl, sourceSheetName, null, importFilterHeaders) : importFilterHeaders.split(/,|，/g).map(h => h.trim());
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
                    const validKeywords = getUniqueValuesForHeader(sourceUrl, sourceSheetName, sourceIdentifierRange, filter.header);
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
 * Saves settings from the Data Verify HTML interface.
 * This is a wrapper to call the function in the Data Verify script.
 * @param {object} settings The settings object from the UI.
 * @param {string} sheetName The name of the sheet to save settings for.
 */
function saveVerifySettingsFromHtml(settings, sheetName) {
    return saveVerifySettings(settings, sheetName);
}

/**
 * Runs the auto-mapping process based on header names for a specific type.
 */
function runAutoMapping(settings) {
    const T = MasterData.getTranslations();
    const sourceUrl = settings.sourceUrl || settings.sourceDataUrl;
    const sourceSheetName = settings.sourceSheetName || settings.sourceDataSheetName;
    const {
        sourceHeaderRow,
        targetSheetName,
        targetHeaderRow
    } = settings;

    if (!sourceUrl || !sourceSheetName || !sourceHeaderRow || !targetSheetName || !targetHeaderRow) {
        throw new Error("Cannot perform auto-mapping. Required settings are missing.");
    }

    const targetSs = SpreadsheetApp.getActiveSpreadsheet();
    const targetSh = targetSs.getSheetByName(targetSheetName);
    if (!targetSh) throw new Error(T.errorTargetSheetNotFound.replace('{SHEET_NAME}', targetSheetName));
    let targetHeaders;
    try {
        targetHeaders = targetSh.getRange(targetHeaderRow, 1, 1, targetSh.getMaxColumns()).getValues()[0];
    } catch (e) {
        throw new Error(T.errorInvalidHeaderRow.replace('{ROW_NUM}', targetHeaderRow));
    }

    let sourceSh;
    try {
        const sourceSs = SpreadsheetApp.openByUrl(sourceUrl);
        sourceSh = sourceSs.getSheetByName(sourceSheetName);
        if (!sourceSh) throw new Error(T.errorSheetNotFoundInUrl.replace('{SHEET_NAME}', sourceSheetName));
    } catch (e) {
        throw new Error(T.errorInvalidUrl);
    }
    let sourceHeaders;
    try {
        sourceHeaders = sourceSh.getRange(sourceHeaderRow, 1, 1, sourceSh.getMaxColumns()).getValues()[0];
    } catch (e) {
        throw new Error(T.errorInvalidHeaderRow.replace('{ROW_NUM}', sourceHeaderRow));
    }

    const targetHeaderMap = new Map();
    targetHeaders.forEach((h, i) => {
        const header = h.toString().trim().normalize('NFC');
        if (header) targetHeaderMap.set(header, i + 1);
    });

    const sourceHeaderMap = new Map();
    sourceHeaders.forEach((h, i) => {
        const header = h.toString().trim().normalize('NFC');
        if (header) sourceHeaderMap.set(header, i + 1);
    });

    const mappings = [];
    targetHeaderMap.forEach((targetColIndex, header) => {
        if (sourceHeaderMap.has(header)) {
            const sourceColIndex = sourceHeaderMap.get(header);
            mappings.push({
                targetCol: columnToLetter(targetColIndex),
                sourceCol: columnToLetter(sourceColIndex)
            });
        }
    });

    return mappings;
}


// =================================================================================================
// ===================================== SECTION 2: CORE LOGIC =====================================
// =================================================================================================

/**
 * Validates the source comparison column for empty or duplicate values.
 * @param {object} settings The comparison settings.
 * @returns {{isValid: boolean, message: string}} An object containing the validation result.
 */
function checkSourceCompareField(settings) {
    const {
        sourceUrl,
        sourceSheetName,
        sourceCompareRange,
        sourceLookupCol
    } = settings;
    const T = MasterData.getTranslations();

    if (!sourceUrl || !sourceSheetName || !sourceCompareRange || !sourceLookupCol) {
        return {
            isValid: false,
            message: "請先填寫來源 URL、分頁、比對範圍與比對欄位。"
        };
    }

    try {
        const sourceSs = SpreadsheetApp.openByUrl(sourceUrl);
        const sourceSheet = sourceSs.getSheetByName(sourceSheetName);
        if (!sourceSheet) {
            return {
                isValid: false,
                message: T.errorSheetNotFoundInUrl.replace('{SHEET_NAME}', sourceSheetName)
            };
        }

        const sourceRange = sourceSheet.getRange(sourceCompareRange);
        const sourceData = sourceRange.getValues();
        const lookupColNum = letterToColumn(sourceLookupCol);

        if (lookupColNum < sourceRange.getColumn() || lookupColNum >= (sourceRange.getColumn() + sourceRange.getNumColumns())) {
            const errorMessage = T.sourceCompareFieldCheckError
                .replace('{COLUMN}', sourceLookupCol)
                .replace('{RANGE}', sourceCompareRange);
            return {
                isValid: false,
                message: errorMessage
            };
        }

        const lookupColIndex = lookupColNum - sourceRange.getColumn();

        const valuesSeen = new Map();
        const duplicates = new Map();
        const emptyRows = [];

        sourceData.forEach((row, index) => {
            const key = row[lookupColIndex];
            const a1NotationRow = sourceRange.getRow() + index;

            if (key === null || key === "") {
                emptyRows.push(a1NotationRow);
            } else {
                const keyStr = key.toString().trim();
                if (valuesSeen.has(keyStr)) {
                    if (!duplicates.has(keyStr)) {
                        duplicates.set(keyStr, [valuesSeen.get(keyStr)]);
                    }
                    duplicates.get(keyStr).push(a1NotationRow);
                } else {
                    valuesSeen.set(keyStr, a1NotationRow);
                }
            }
        });

        let messages = [];
        if (emptyRows.length > 0) {
            messages.push(T.emptyRowsFound.replace('{ROWS}', emptyRows.join(', ')));
        }
        if (duplicates.size > 0) {
            let duplicateMessages = [];
            duplicates.forEach((rows, value) => {
                duplicateMessages.push(T.duplicateValuesFound.replace('{VALUE}', value).replace('{ROWS}', rows.sort((a, b) => a - b).join(', ')));
            });
            messages.push(T.multipleDuplicateValuesFound.replace('{DETAILS}', duplicateMessages.join('; ')));
        }

        if (messages.length > 0) {
            return {
                isValid: false,
                message: messages.join('\n')
            };
        }

        return {
            isValid: true,
            message: T.checkSuccess
        };

    } catch (e) {
        Logger.log(`Error in checkSourceCompareField: ${e.message}`);
        return {
            isValid: false,
            message: `檢查時發生錯誤: ${e.message}`
        };
    }
}


/**
 * Validates the target lookup column for empty or duplicate values.
 * @param {object} settings The comparison settings.
 * @returns {{isValid: boolean, message: string}} An object containing the validation result.
 */
function checkTargetLookupField(settings) {
    const {
        targetSheetName,
        targetStartRow,
        targetLookupCol
    } = settings;
    const T = MasterData.getTranslations();

    if (!targetSheetName || !targetStartRow || !targetLookupCol) {
        return {
            isValid: false,
            message: "請先填寫目標分頁、起始列與查找欄位。"
        };
    }

    try {
        const targetSs = SpreadsheetApp.getActiveSpreadsheet();
        const targetSheet = targetSs.getSheetByName(targetSheetName);
        if (!targetSheet) {
            return {
                isValid: false,
                message: T.errorTargetSheetNotFound.replace('{SHEET_NAME}', targetSheetName)
            };
        }

        const startRow = parseInt(targetStartRow, 10);
        const lastRow = targetSheet.getLastRow();

        if (lastRow < startRow) {
            return {
                isValid: true,
                message: "目標欄位無資料可檢查。"
            }; // Not an error, just empty.
        }

        const lookupColNum = letterToColumn(targetLookupCol);
        const range = targetSheet.getRange(startRow, lookupColNum, lastRow - startRow + 1, 1);
        const data = range.getValues();

        const valuesSeen = new Map();
        const duplicates = new Map();
        const emptyRows = [];

        data.forEach((row, index) => {
            const key = row[0];
            const a1NotationRow = startRow + index;

            if (key === null || key === "") {
                emptyRows.push(a1NotationRow);
            } else {
                const keyStr = key.toString().trim();
                if (valuesSeen.has(keyStr)) {
                    if (!duplicates.has(keyStr)) {
                        duplicates.set(keyStr, [valuesSeen.get(keyStr)]);
                    }
                    duplicates.get(keyStr).push(a1NotationRow);
                } else {
                    valuesSeen.set(keyStr, a1NotationRow);
                }
            }
        });

        let messages = [];
        if (emptyRows.length > 0) {
            messages.push(T.emptyRowsFound.replace('{ROWS}', emptyRows.join(', ')));
        }
        if (duplicates.size > 0) {
            let duplicateMessages = [];
            duplicates.forEach((rows, value) => {
                duplicateMessages.push(T.duplicateValuesFound.replace('{VALUE}', value).replace('{ROWS}', rows.sort((a, b) => a - b).join(', ')));
            });
            messages.push(T.multipleDuplicateValuesFound.replace('{DETAILS}', duplicateMessages.join('; ')));
        }

        if (messages.length > 0) {
            return {
                isValid: false,
                message: messages.join('\n')
            };
        }

        return {
            isValid: true,
            message: T.checkSuccess
        };

    } catch (e) {
        Logger.log(`Error in checkTargetLookupField: ${e.message}`);
        const errorMessage = T.targetLookupFieldCheckError.replace('{COLUMN}', targetLookupCol);
        return {
            isValid: false,
            message: `${errorMessage}\n詳細錯誤: ${e.message}`
        };
    }
}


function requestStopImport() {
    PropertiesService.getScriptProperties().setProperty('stopImportRequested', 'true');
    SpreadsheetApp.getActiveSpreadsheet().toast('Stop request sent. The script will stop after finishing the current item.', 'Info', 5);
}

function requestStopValidation() {
    PropertiesService.getScriptProperties().setProperty('stopValidationRequested', 'true');
    SpreadsheetApp.getActiveSpreadsheet().toast('Stop request sent. The script will stop after finishing the current item.', 'Info', 5);
}

function resetAllState() {
    const ui = SpreadsheetApp.getUi();
    const scriptProperties = PropertiesService.getScriptProperties();
    try {
        scriptProperties.deleteProperty('lastCompletedAllocationKey');
        const activeSheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
        const settings = getSettings(activeSheetName).importSettings;
        clearTargetSheetData(settings);
        SpreadsheetApp.flush();
        SpreadsheetApp.getActiveSpreadsheet().toast('Reset complete! All progress and sheet data have been cleared.', 'Success', 5);
    } catch (e) {
        Logger.log(`Reset failed: ${e.message}`);
        ui.alert('Reset Failed', e.message, ui.ButtonSet.OK);
    }
}

/**
 * Main function to run the data comparison process.
 */
function runCompareProcess() {
    const ui = SpreadsheetApp.getUi();
    const activeSheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
    const T = MasterData.getTranslations();
    try {
        const settings = getCompareSettings(activeSheetName);
        if (!settings || Object.keys(settings).length <= 1 || !settings.sourceUrl) {
            throw new Error(T.errorNoCompareSettingsFound);
        }

        // --- Pre-flight check for SOURCE ---
        const sourceValidation = checkSourceCompareField(settings);
        if (!sourceValidation.isValid) {
            const confirmationMessage = T.preCheckWarningBody
                .replace('{COLUMN}', settings.sourceLookupCol)
                .replace('{MESSAGE}', sourceValidation.message);
            const response = ui.alert(T.preCheckWarningTitle, confirmationMessage, ui.ButtonSet.YES_NO);
            if (response !== ui.Button.YES) {
                SpreadsheetApp.getActiveSpreadsheet().toast(T.preCheckCancelled, 'Cancelled', 5);
                return;
            }
        }

        // --- Pre-flight check for TARGET ---
        const targetValidation = checkTargetLookupField(settings);
        if (!targetValidation.isValid) {
            const confirmationMessage = T.preCheckWarningBodyTarget
                .replace('{COLUMN}', settings.targetLookupCol)
                .replace('{MESSAGE}', targetValidation.message);
            const response = ui.alert(T.preCheckWarningTitle, confirmationMessage, ui.ButtonSet.YES_NO);
            if (response !== ui.Button.YES) {
                SpreadsheetApp.getActiveSpreadsheet().toast(T.preCheckCancelled, 'Cancelled', 5);
                return;
            }
        }

        SpreadsheetApp.getActiveSpreadsheet().toast('開始執行資料比對...', '處理中', 10);

        const targetSs = SpreadsheetApp.getActiveSpreadsheet();
        const targetSsId = targetSs.getId(); // Get target ID
        const targetSheet = targetSs.getSheetByName(settings.targetSheetName);
        if (!targetSheet) throw new Error(`找不到目標分頁: ${settings.targetSheetName}`);

        const sourceSs = SpreadsheetApp.openByUrl(settings.sourceUrl);
        const sourceSsId = sourceSs.getId(); // Get source ID
        // *** FIX: Corrected a syntax error in the following line ***
        const sourceSheet = sourceSs.getSheetByName(settings.sourceSheetName);
        if (!sourceSheet) throw new Error(`在來源檔案中找不到分頁: ${settings.sourceSheetName}`);

        const sourceGid = sourceSheet.getSheetId();
        const sourceUrlForLink = sourceSs.getUrl().replace('/edit', '');

        const sourceRange = sourceSheet.getRange(settings.sourceCompareRange);
        const sourceData = sourceRange.getValues();
        const lookupColIndex = letterToColumn(settings.sourceLookupCol) - sourceRange.getColumn();
        const returnColIndex = letterToColumn(settings.sourceReturnCol) - sourceRange.getColumn();

        if (lookupColIndex < 0 || returnColIndex < 0 || lookupColIndex >= sourceRange.getNumColumns() || returnColIndex >= sourceRange.getNumColumns()) {
            throw new Error("來源比對欄位或返回欄位不在指定的比對範圍內。");
        }

        const lookupMap = new Map();
        sourceData.forEach((row, index) => {
            const key = row[lookupColIndex];
            if (key !== null && key !== "") {
                lookupMap.set(key.toString().trim(), {
                    value: row[returnColIndex],
                    sourceRowIndex: index
                });
            }
        });

        const targetStartRow = settings.targetStartRow;
        const lastRow = targetSheet.getLastRow();
        if (lastRow < targetStartRow) {
            SpreadsheetApp.getActiveSpreadsheet().toast('目標分頁沒有需要比對的資料。', '提示', 5);
            return;
        }

        const targetLookupColNum = letterToColumn(settings.targetLookupCol);
        const targetWriteColNum = letterToColumn(settings.targetWriteCol);

        const targetLookupRange = targetSheet.getRange(targetStartRow, targetLookupColNum, lastRow - targetStartRow + 1, 1);
        const targetLookupValues = targetLookupRange.getValues();

        const resultsToWrite = [];
        const richTextTasks = [];

        targetLookupValues.forEach((row, index) => {
            const lookupValue = row[0];
            if (lookupValue !== null && lookupValue !== "") {
                const lookupKey = lookupValue.toString().trim();
                if (lookupMap.has(lookupKey)) {
                    const result = lookupMap.get(lookupKey);
                    const foundValue = result.value;

                    if (foundValue === null || foundValue === "") {
                        const sourceRowForDisplay = sourceRange.getRow() + result.sourceRowIndex;
                        const sourceColLetterForDisplay = settings.sourceReturnCol;
                        const displayText = `${sourceColLetterForDisplay}${sourceRowForDisplay}${T.noSourceDataSuffix}`;

                        resultsToWrite.push([displayText]);
                        richTextTasks.push({
                            targetRow: targetStartRow + index,
                            targetCol: targetWriteColNum,
                            sourceRow: sourceRowForDisplay,
                            sourceColLetter: sourceColLetterForDisplay
                        });
                    } else {
                        resultsToWrite.push([foundValue]);
                    }
                } else {
                    resultsToWrite.push([""]);
                }
            } else {
                resultsToWrite.push([""]);
            }
        });

        if (resultsToWrite.length > 0) {
            targetSheet.getRange(targetStartRow, targetWriteColNum, resultsToWrite.length, 1).setValues(resultsToWrite);

            // Logic to create internal or external links
           const isInternalLink = (targetSsId === sourceSsId);
           richTextTasks.forEach(task => {
            const cell = targetSheet.getRange(task.targetRow, task.targetCol);
            const linkFragment = `#gid=${sourceGid}&range=${task.sourceColLetter}${task.sourceRow}`;
            const linkUrl = isInternalLink ? linkFragment : `${sourceUrlForLink}${linkFragment}`;
            const newText = cell.getValue();
            if (typeof newText === 'string' && newText.includes(T.noSourceDataSuffix)) {
            const richText = SpreadsheetApp.newRichTextValue()
                .setText(newText)
                .setLinkUrl(0, newText.length, linkUrl)
                .build();
            cell.setRichTextValue(richText);
                }
           });

            SpreadsheetApp.getActiveSpreadsheet().toast(`資料比對完成！已更新 ${resultsToWrite.length} 筆資料。`, '成功', 5);
        } else {
            SpreadsheetApp.getActiveSpreadsheet().toast('沒有找到任何可比對的資料。', '提示', 5);
        }

    } catch (e) {
        Logger.log(`Data comparison failed: ${e.stack}`);
        ui.alert(T.compareFailedTitle, e.message, ui.ButtonSet.OK);
    }
}


function runImportProcess() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetSsId = ss.getId(); // Get target ID
    const activeSheetName = ss.getActiveSheet().getName();
    const ui = SpreadsheetApp.getUi();
    const scriptProperties = PropertiesService.getScriptProperties();
    const T = MasterData.getTranslations();

    scriptProperties.deleteProperty('stopImportRequested');

    try {
        const allSettings = getSettings(activeSheetName);
        const settings = allSettings.importSettings;

        if (!settings || !settings.sourceUrl) {
            throw new Error(T.errorNoImportSettingsFound.replace('{SHEET_NAME}', activeSheetName));
        }

        if (settings.importFilterHeaders && settings.importFilterHeaders.length > 0) {
            const sourceSheet = SpreadsheetApp.openByUrl(settings.sourceUrl).getSheetByName(settings.sourceSheetName);
            const range = sourceSheet.getRange(settings.sourceIdentifierRange);
            const rangeHeaders = sourceSheet.getRange(range.getRow() - 1, range.getColumn(), 1, range.getNumColumns()).getValues()[0].map(h => h.toString().trim());
            const filterHeaders = settings.importFilterHeaders;
            const rangeHeadersSorted = JSON.stringify([...rangeHeaders].sort());
            const filterHeadersSorted = JSON.stringify([...filterHeaders].sort());

            if (rangeHeadersSorted !== filterHeadersSorted) {
                let body = T.asymmetryWarningBody
                    .replace('{RANGE_HEADERS}', rangeHeaders.join(', '))
                    .replace('{FILTER_HEADERS}', filterHeaders.join(', '));
                const response = ui.alert(T.asymmetryWarningTitle, body, ui.ButtonSet.YES_NO);
                if (response !== ui.Button.YES) {
                    ss.toast(T.importCancelled, 'Cancelled', 5);
                    return;
                }
            }
        }

        const warningMessages = [];
        const filters = settings.keywordFilters || [];
        const incompleteFilter = filters.some(f => f.header && !f.keywords);
        if (incompleteFilter) {
            warningMessages.push(T.preflightFilterWarning);
        }

        if (warningMessages.length > 0) {
            const prompt = `${T.preflightWarning}\n\n${warningMessages.join('\n')}\n\n${T.preflightSuggestion}\n\n${T.preflightConfirmation}`;
            const response = ui.alert(T.preflightTitle, prompt, ui.ButtonSet.YES_NO);
            if (response !== ui.Button.YES) {
                ss.toast(T.importCancelled, 'Cancelled', 5);
                return;
            }
        }

        if (settings.targetHeaderRow >= settings.targetStartRow) {
            throw new Error(T.headerLessThanStartError);
        }

        if (!settings.sourceUrl || !settings.sourceUrl.startsWith('http')) {
            throw new Error("The 'Source Sheet URL' is empty or has an invalid format in the Settings sheet.");
        }

        const targetSheet = ss.getSheetByName(settings.targetSheetName);
        if (!targetSheet) {
            throw new Error(`Target sheet named "${settings.targetSheetName}" not found.`);
        }

        const isFlatteningMode = !!(settings.sourceHeaderRange && settings.sourceValueMatrixRange);
        const modeMessage = isFlatteningMode ? "Data Flattening Mode" : "Direct Import Mode";
        ss.toast(`Reading source data in ${modeMessage}...`, 'Processing', 5);

        const {
            allTasks,
            sourceInfo,
            hasSourceData
        } = buildFullTaskList(settings, targetSheet, isFlatteningMode);

        if (!hasSourceData) {
            ss.toast('No data to import in the source. The target sheet will be cleared.', 'Info', 8);
            clearTargetSheetData(settings);
            ss.toast('Target sheet has been cleared.', 'Complete', 5);
            return;
        }

        if (allTasks.length === 0) {
            const alertBody = T.filterMismatchBody.replace('{FILTER_HEADER}', 'your conditions');
            ui.alert(T.filterMismatchTitle, alertBody, ui.ButtonSet.OK);
            clearTargetSheetData(settings);
            ss.toast('Target sheet has been cleared.', 'Complete', 5);
            return;
        }

        ss.toast('Clearing target sheet for synchronization...', 'Processing', 5);
        clearTargetSheetData(settings);

        const valuesForBulkWrite = allTasks.map(task => task.valuesToWrite);

        if (valuesForBulkWrite.length > 0) {
            ss.toast(`Writing ${allTasks.length} records...`, 'Processing', 10);
            const rangeToWrite = targetSheet.getRange(settings.targetStartRow, 1, valuesForBulkWrite.length, valuesForBulkWrite[0].length);
            rangeToWrite.setValues(valuesForBulkWrite);

            for (let i = 0; i < allTasks.length; i++) {
                if (scriptProperties.getProperty('stopImportRequested') === 'true') {
                    scriptProperties.deleteProperty('stopImportRequested');
                    Logger.log('Import script was manually stopped by the user.');
                    ss.toast('Import process was manually stopped. Data may be incomplete.', 'Stopped', 8);
                    return;
                }

                const task = allTasks[i];
                const currentRow = settings.targetStartRow + i;

                //Logic to create internal or external links
               const isInternalLink = (targetSsId === sourceInfo.id);

                task.richTextChecks.forEach(check => {
                 const cell = targetSheet.getRange(currentRow, check.targetCol);
                 const sourceRow = sourceInfo.startRow + task.sourceRowIndex;
                 const sourceColLetter = columnToLetter(check.originalCol);
                 const linkFragment = `#gid=${sourceInfo.gid}&range=${sourceColLetter}${sourceRow}`;
                 const linkUrl = isInternalLink ? linkFragment : `${sourceInfo.url.replace('/edit', '')}${linkFragment}`;
                 const newText = `${sourceColLetter}${sourceRow}${T.noSourceDataSuffix}`;
                 const richText = SpreadsheetApp.newRichTextValue()
                     .setText(newText)
                     .setLinkUrl(0, newText.length, linkUrl)
                     .build();
                 cell.setRichTextValue(richText);
                });
            }

            const commitColumnValues = Array(allTasks.length).fill(['OK']);
            const commitColIndex = findColumnIndexByHeader(targetSheet.getRange(settings.targetHeaderRow, 1, 1, targetSheet.getMaxColumns()).getValues()[0], 'Commit');
            if (commitColIndex !== -1) {
                targetSheet.getRange(settings.targetStartRow, commitColIndex, commitColumnValues.length, 1).setValues(commitColumnValues);
                targetSheet.hideColumns(commitColIndex);
            } else {
                Logger.log("Warning: 'Commit' header not found in target sheet. Cannot write commit status.");
            }
        }

        ss.toast('Data synchronization complete!', 'Success', 5);

    } catch (e) {
        Logger.log(`Data synchronization failed: ${e.stack}`);
        ui.alert('Data Synchronization Failed', e.message, ui.ButtonSet.OK);
    } finally {
        scriptProperties.deleteProperty('stopImportRequested');
    }
}

// =================================================================================================
// ===================================== SECTION 3: HELPER FUNCTIONS ===============================
// =================================================================================================

function clearTargetSheetData(settings) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = ss.getSheetByName(settings.targetSheetName);
    if (!targetSheet) throw new Error(`Target sheet named "${settings.targetSheetName}" not found.`);
    const commitColIndex = findColumnIndexByHeader(targetSheet.getRange(settings.targetHeaderRow, 1, 1, targetSheet.getMaxColumns()).getValues()[0], 'Commit');
    if (commitColIndex !== -1) {
        targetSheet.showColumns(commitColIndex);
    }
    const lastRow = targetSheet.getLastRow();
    if (lastRow >= settings.targetStartRow) {
        targetSheet.getRange(settings.targetStartRow, 1, lastRow - settings.targetStartRow + 1, targetSheet.getMaxColumns()).clear();
    }
}

/**
 * Applies all configured keyword filters to a single row of data.
 */
function applyAllFilters(rowData, sourceHeaders, filters) {
    if (!filters || filters.length === 0) return true;
    return filters.every(filter => {
        if (!filter.header) return true;
        const colIndex = sourceHeaders.findIndex(h => h === filter.header);
        if (colIndex === -1) {
            Logger.log(`Warning: Filter header "${filter.header}" not found in imported data. Ignoring this filter condition.`);
            return true;
        }
        const keywords = (filter.keywords || '').split(/,|，/g).map(kw => kw.trim()).filter(kw => kw);
        if (keywords.length === 0) return true;
        const cellValue = rowData[colIndex] ? rowData[colIndex].toString().trim() : '';
        return keywords.includes(cellValue);
    });
}


/**
 * Builds the full list of tasks based on the determined import mode.
 */
function buildFullTaskList(settings, targetSheet, isFlatteningMode) {
    const targetHeaderRow = settings.targetHeaderRow;
    if (targetHeaderRow <= 0) throw new Error(`'Target Header Row' in Settings must be a positive number.`);
    const targetHeaders = targetSheet.getRange(targetHeaderRow, 1, 1, targetSheet.getMaxColumns()).getValues()[0];

    if (isFlatteningMode) {
        const {
            identifierValues,
            sourceIdentifierHeaders,
            originalColumnIndices,
            headerValues,
            matrixValues,
            sourceInfo,
            masterHeaderName
        } = fetchImportData(settings);

        const qtyHeaderName = "Q'ty";
        const allRequiredHeaders = [...sourceIdentifierHeaders, qtyHeaderName, masterHeaderName];
        const headerMap = {};
        allRequiredHeaders.forEach(header => {
            if (!header) return;
            const colIndex = findColumnIndexByHeader(targetHeaders, header);
            if (colIndex === -1) {
                const message = `Required header "${header}" (from source sheet) was not found in the target sheet "${settings.targetSheetName}" on row ${targetHeaderRow}.\n\nPlease ensure all required headers exist: ${allRequiredHeaders.join(', ')}.`;
                throw new Error(message);
            }
            headerMap[header] = colIndex;
        });

        const allTasks = [];
        const hasSourceData = identifierValues.some(row => row.some(cell => cell.toString().trim() !== ''));
        const loopLength = Math.min(identifierValues.length, matrixValues.length);
        if (identifierValues.length !== matrixValues.length) {
            Logger.log(`Warning: Identifier range has ${identifierValues.length} rows, but matrix range has ${matrixValues.length} rows. Processing the first ${loopLength} common rows.`);
        }

        for (let i = 0; i < loopLength; i++) {
            if (identifierValues[i].every(cell => cell.toString().trim() === '')) {
                continue;
            }

            const nTimes = matrixValues[i];
            if (applyAllFilters(identifierValues[i], sourceIdentifierHeaders, settings.keywordFilters)) {
                const nKeys = getNonEmptyIndices(nTimes);
                if (nKeys.length === 0) {
                    let valuesToWrite = new Array(targetHeaders.length).fill("");
                    sourceIdentifierHeaders.forEach((header, colIndex) => {
                        if (headerMap[header]) {
                            valuesToWrite[headerMap[header] - 1] = identifierValues[i][colIndex];
                        }
                    });
                    valuesToWrite[headerMap[qtyHeaderName] - 1] = 'No source data';
                    valuesToWrite[headerMap[masterHeaderName] - 1] = '';

                    const richTextChecks = [];
                    sourceIdentifierHeaders.forEach((headerToValidate, sourceColIndex) => {
                        const dataIsMissing = !identifierValues[i][sourceColIndex] || identifierValues[i][sourceColIndex].toString().trim() === '';
                        if (dataIsMissing && headerMap[headerToValidate]) {
                            const originalColIndex = originalColumnIndices[sourceColIndex];
                            richTextChecks.push({
                                originalCol: originalColIndex + 1,
                                targetCol: headerMap[headerToValidate]
                            });
                        }
                    });

                    allTasks.push({
                        valuesToWrite: valuesToWrite,
                        richTextChecks: richTextChecks,
                        sourceRowIndex: i
                    });
                } else {
                    for (let z = 0; z < nKeys.length; z++) {
                        const keyIndex = nKeys[z];
                        let qtyValue = nTimes[keyIndex];
                        let configValue = headerValues[0][keyIndex];
                        let valuesToWrite = new Array(targetHeaders.length).fill("");

                        sourceIdentifierHeaders.forEach((header, colIndex) => {
                            if (headerMap[header]) {
                                valuesToWrite[headerMap[header] - 1] = identifierValues[i][colIndex];
                            }
                        });

                        valuesToWrite[headerMap[qtyHeaderName] - 1] = qtyValue;
                        valuesToWrite[headerMap[masterHeaderName] - 1] = configValue;

                        const richTextChecks = [];
                        sourceIdentifierHeaders.forEach((headerToValidate, sourceColIndex) => {
                            const dataIsMissing = !identifierValues[i][sourceColIndex] || identifierValues[i][sourceColIndex].toString().trim() === '';
                            if (dataIsMissing && headerMap[headerToValidate]) {
                                const originalColIndex = originalColumnIndices[sourceColIndex];
                                richTextChecks.push({
                                    originalCol: originalColIndex + 1,
                                    targetCol: headerMap[headerToValidate]
                                });
                            }
                        });

                        allTasks.push({
                            valuesToWrite: valuesToWrite,
                            richTextChecks: richTextChecks,
                            sourceRowIndex: i
                        });
                    }
                }
            }
        }
        return {
            allTasks,
            sourceInfo,
            hasSourceData
        };
    } else {
        // --- MODE 2: DIRECT IMPORT ---
        const {
            identifierValues,
            sourceIdentifierHeaders,
            originalColumnIndices,
            sourceInfo
        } = fetchImportData(settings);
        const hasSourceData = identifierValues.some(row => row.some(cell => cell.toString().trim() !== ''));

        const headersToImport = sourceIdentifierHeaders;
        const headerMap = {};
        headersToImport.forEach(header => {
            const colIndex = findColumnIndexByHeader(targetHeaders, header);
            if (colIndex === -1) {
                throw new Error(`Required header "${header}" (from source sheet) was not found in the target sheet "${settings.targetSheetName}" on row ${settings.targetHeaderRow}.`);
            }
            headerMap[header] = colIndex;
        });

        const allTasks = [];
        for (let i = 0; i < identifierValues.length; i++) {
            if (identifierValues[i].every(cell => cell.toString().trim() === '')) {
                continue;
            }

            const sourceRowData = identifierValues[i];

            if (applyAllFilters(sourceRowData, sourceIdentifierHeaders, settings.keywordFilters)) {
                let valuesToWrite = new Array(targetHeaders.length).fill("");

                headersToImport.forEach((header, idx) => {
                    const targetCol = headerMap[header];
                    if (targetCol) {
                        valuesToWrite[targetCol - 1] = sourceRowData[idx];
                    }
                });

                const richTextChecks = [];
                headersToImport.forEach((headerToValidate, sourceColIndex) => {
                    const dataIsMissing = !sourceRowData[sourceColIndex] || sourceRowData[sourceColIndex].toString().trim() === '';
                    if (dataIsMissing && headerMap[headerToValidate]) {
                        const originalColIndex = originalColumnIndices[sourceColIndex];
                        richTextChecks.push({
                            originalCol: originalColIndex + 1,
                            targetCol: headerMap[headerToValidate]
                        });
                    }
                });

                allTasks.push({
                    valuesToWrite: valuesToWrite,
                    richTextChecks: richTextChecks,
                    sourceRowIndex: i
                });
            }
        }
        return {
            allTasks,
            sourceInfo,
            hasSourceData
        };
    }
}

/**
 * Fetches data based on the new hybrid logic and tracks original column indices.
 */
function fetchImportData(settings) {
    const sourceSpreadsheet = SpreadsheetApp.openByUrl(settings.sourceUrl);
    const sourceSheet = sourceSpreadsheet.getSheetByName(settings.sourceSheetName);
    if (!sourceSheet) throw new Error(`Sheet named "${settings.sourceSheetName}" not found in the source file.`);

    const isFlatteningMode = !!(settings.sourceHeaderRange && settings.sourceValueMatrixRange);

    if (!settings.sourceIdentifierRange) {
        throw new Error("Required setting 'Source Data Import Range' is missing. Please check your 'Data Import Settings' in the Settings sheet.");
    }
    if (isFlatteningMode && (!settings.sourceHeaderRange || !settings.sourceValueMatrixRange)) {
        throw new Error("For Data Flattening Mode, 'Header Start Row for Other Blocks' and 'Data Range within Header of Other Blocks' are required.");
    }

    const sourceGid = sourceSheet.getSheetId();
    const sourceUrl = sourceSpreadsheet.getUrl();
    const identifierRange = sourceSheet.getRange(settings.sourceIdentifierRange);
    const identifierStartRow = identifierRange.getRow();
    const identifierStartCol = identifierRange.getColumn();
    const identifierNumRows = identifierRange.getNumRows();

    let headerValues, matrixValues, masterHeaderName;

    if (isFlatteningMode) {
        const configHeaderDefinitionRange = sourceSheet.getRange(settings.sourceHeaderRange);
        const matrixRange = sourceSheet.getRange(settings.sourceValueMatrixRange);
        matrixValues = matrixRange.getValues();
        const configHeaderRow = configHeaderDefinitionRange.getRow();
        const alignedHeaderRange = sourceSheet.getRange(configHeaderRow, matrixRange.getColumn(), 1, matrixRange.getNumColumns());
        headerValues = alignedHeaderRange.getValues();
        const originalHeaderValues = configHeaderDefinitionRange.getValues();
        const headerRowForDetection = originalHeaderValues[0];
        const firstHeaderIndex = headerRowForDetection.findIndex(h => h.toString().trim() !== '');
        if (firstHeaderIndex === -1) throw new Error(`Could not find headers. Please check your "Source Header Range" (${settings.sourceHeaderRange}) setting in "${settings.sourceSheetName}".`);
        const masterHeaderColumn = configHeaderDefinitionRange.getColumn() + firstHeaderIndex;
        const masterHeaderCell = sourceSheet.getRange(configHeaderRow - 1, masterHeaderColumn);
        masterHeaderName = masterHeaderCell.getValue().toString().trim();
    }

    let sourceIdentifierHeaders;
    let identifierValues;
    let originalColumnIndices = [];
    const useFilter = settings.importFilterHeaders && settings.importFilterHeaders.length > 0;

    if (useFilter) {
        sourceIdentifierHeaders = settings.importFilterHeaders;
        const fullHeaderRange = sourceSheet.getRange(identifierStartRow - 1, 1, 1, sourceSheet.getMaxColumns());
        const fullHeaders = fullHeaderRange.getValues()[0].map(h => h.toString().trim());
        const colIndicesToKeep = sourceIdentifierHeaders.map(header => {
            const index = fullHeaders.indexOf(header);
            if (index === -1) throw new Error(`The header "${header}" specified in the 'Source Header Import Filter' was not found in the source sheet.`);
            return index;
        });
        originalColumnIndices = colIndicesToKeep;
        const fullDataRange = sourceSheet.getRange(identifierStartRow, 1, identifierNumRows, sourceSheet.getMaxColumns());
        const fullDataValues = fullDataRange.getValues();
        identifierValues = fullDataValues.map(row => {
            return colIndicesToKeep.map(colIndex => row[colIndex]);
        });
    } else {
        const idRange = sourceSheet.getRange(settings.sourceIdentifierRange);
        identifierValues = idRange.getValues();
        const idHeaderRange = sourceSheet.getRange(idRange.getRow() - 1, idRange.getColumn(), 1, idRange.getNumColumns());
        sourceIdentifierHeaders = idHeaderRange.getValues()[0].map(h => h.toString().trim());
        const numCols = idRange.getNumColumns();
        const startCol = idRange.getColumn();
        for (let i = 0; i < numCols; i++) {
            originalColumnIndices.push(startCol + i - 1);
        }
    }
    return {
        identifierValues,
        sourceIdentifierHeaders,
        originalColumnIndices,
        headerValues,
        matrixValues,
        masterHeaderName,
        sourceInfo: {
            id: sourceSpreadsheet.getId(),
            gid: sourceGid,
            url: sourceUrl,
            startRow: identifierStartRow,
            startCol: identifierStartCol
        }
    };
}

function getNonEmptyIndices(rowArray) {
    if (!rowArray) return [];
    const nonEmptyIndices = [];
    for (let i = 0; i < rowArray.length; i++) {
        if (rowArray[i] !== "" && rowArray[i] !== null && rowArray[i] !== undefined) {
            nonEmptyIndices.push(i);
        }
    }
    return nonEmptyIndices;
}

/**
 * Converts a 1-based column index to a letter.
 * @param {number} column The 1-based column index.
 * @return {string} The column letter.
 */
function columnToLetter(column) {
    let temp, letter = '';
    while (column > 0) {
        temp = (column - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = (column - temp - 1) / 26;
    }
    return letter;
}

/**
 * Converts a column letter to a 1-based column index.
 * @param {string} letter The column letter.
 * @return {number} The 1-based column index.
 */
function letterToColumn(letter) {
    let column = 0,
        length = letter.length;
    for (let i = 0; i < length; i++) {
        column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
    }
    return column;
}


function findColumnIndexByHeader(headerRow, headerName) {
    if (!headerName) return -1;
    const index = headerRow.findIndex(header => header.toString().trim() === headerName.trim());
    return index !== -1 ? index + 1 : -1;
}

/**
 * Reads all settings, combining PropertiesService and the "Settings" sheet for a transition period.
 * @param {string} [sheetName] - The name of the sheet to get settings for. Defaults to the active sheet.
 * @returns {{importSettings: object, verifySettings: object, monitorSettings: object}}
 */
function getSettings(sheetName) {
    const currentSheetName = sheetName || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();

    const importSettings = getImportSettings(currentSheetName);
    const verifySettings = getVerifySettings(currentSheetName);
    const monitorSettings = getMonitorSettings();

    // Post-process header filter range
    if (importSettings.rawImportFilterHeaders && importSettings.sourceUrl && importSettings.sourceSheetName) {
        const T = MasterData.getTranslations();
        if (importSettings.rawImportFilterHeaders.includes(':') || importSettings.rawImportFilterHeaders.includes('：')) {
            try {
                const rangeString = importSettings.rawImportFilterHeaders.replace(/：/g, ':');
                const sourceSs = SpreadsheetApp.openByUrl(importSettings.sourceUrl);
                const sourceSh = sourceSs.getSheetByName(importSettings.sourceSheetName);
                importSettings.importFilterHeaders = sourceSh.getRange(rangeString).getValues()[0].map(h => h.toString().trim()).filter(h => h);
            } catch (e) {
                throw new Error(T.errorInvalidHeaderRange.replace('{RANGE}', importSettings.rawImportFilterHeaders));
            }
        } else {
            importSettings.importFilterHeaders = importSettings.rawImportFilterHeaders.split(/,|，/g).map(h => h.trim()).filter(h => h);
        }
    } else {
        importSettings.importFilterHeaders = [];
    }

    return {
        importSettings,
        verifySettings,
        monitorSettings
    };
}
