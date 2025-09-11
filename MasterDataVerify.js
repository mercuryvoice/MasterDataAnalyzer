/**
 * @OnlyCurrentDoc
 */

// =================================================================================================
// =================================== VERSION & FEATURE SUMMARY ===================================
// =================================================================================================
//
// V1.0 (Pre-release version):
// - Noted.
//
// =================================================================================================


// ================================================================
// SECTION 0: TRANSLATIONS & SETTINGS
// ================================================================

const _MDV_TRANSLATIONS = {
  en: {
    validationFailedTitle: 'Data Validation Failed',
    errorNoValidationSettingsFound: 'No validation settings found for the current sheet "{SHEET_NAME}". Please configure them first.',
    errorNoValidationMappings: 'No validation mappings found. Please configure "Field Validation Conditions".',
    errorDialogTitle: 'Error',
    majorPrefix: 'Major',
    minorPrefix: 'Minor',
    mismatchSuffix: 'Mismatch',
    sourceValueMismatch: 'Source Data Mismatch',
    sourceValueBlank: 'Source Data is Blank',
    noSourceDataSuffix: '_No source data',
    outputErrorRequiredMismatch: 'Required validation conditions do not match, cannot output result',
    headerMismatchErrorTitle: 'Header Mismatch',
    headerMismatchErrorBody: 'The following headers do not match between the target and source sheets for the configured mappings:\n\n{MISMATCH_DETAILS}\n\nPlease correct the mappings or the sheet headers and try again.',
    headerMismatchDetail: 'Target column {TARGET_COL} ("{TARGET_HEADER}") AND Source column {SOURCE_COL} ("{SOURCE_HEADER}") NOT MATCH',
    checkResultTitle: 'Field Check Notice',
    checkResultMismatches: 'Found Mismatched Headers:',
    checkResultSuggestions: 'Found Potentially Mismatched Headers (Suggestions):',
    suggestionText: '- Target header "{TARGET_HEADER}" (Col {TARGET_COL}) might correspond to Source header "{SOURCE_HEADER}" (Col {SOURCE_COL})?',
    noIssuesFound: 'No issues found. All mapped headers match perfectly.',
    checkSuccessHeaders: 'Header check passed. All mapped columns correspond correctly.',
    checkFailureHeaders: 'Header Mismatch Found:',
    checkSuccessColumn: 'Column check passed. No empty values found.',
    checkFailureColumn: 'Empty Values Found:',
    emptyValuesInRows: 'Empty values found in the following rows of source column {COLUMN}: {ROWS}',
    columnIsEmpty: 'The entire source column {COLUMN} is empty or could not be read.',
    noMappingsToCheck: 'No field mappings have been set up, no check required.',
    emptyCheckTitle: 'Source Column Empty Value Check',
    emptyCheckSuccess: 'Check complete. No empty values found in any mapped source columns.',
    emptyCheckFailure: 'Empty values were found in the following source columns:',
    checkEmptyValuesButton: 'Check Source for Empty Values',
    firstColumnMismatch: 'First key column mismatch',
  },
  'zh_TW': {
    validationFailedTitle: '資料驗證失敗',
    errorNoValidationSettingsFound: '找不到工作表 "{SHEET_NAME}" 的資料驗證設定，請先設定。',
    errorNoValidationMappings: '找不到任何驗證對應。請設定「欄位驗證條件」。',
    errorDialogTitle: '錯誤',
    majorPrefix: '主要',
    minorPrefix: '非主要',
    mismatchSuffix: '不匹配',
    sourceValueMismatch: '來源數值不匹配',
    sourceValueBlank: '來源數值為空白',
    noSourceDataSuffix: '_無來源資料',
    outputErrorRequiredMismatch: '必要驗證條件不吻合無法輸出驗證結果',
    headerMismatchErrorTitle: '標頭不符',
    headerMismatchErrorBody: '根據您的欄位對應設定，目標與來源工作表的以下標頭不相符：\n\n{MISMATCH_DETAILS}\n\n請修正您的設定或工作表標頭後再試一次。',
    headerMismatchDetail: '目標欄位 {TARGET_COL} ("{TARGET_HEADER}") 與 來源欄位 {SOURCE_COL} ("{SOURCE_HEADER}") 不一致',
    checkResultTitle: '欄位檢查提示',
    checkResultMismatches: '發現不匹配的標頭：',
    checkResultSuggestions: '發現疑似不匹配的標頭 (建議)：',
    suggestionText: '- 目標標頭 "{TARGET_HEADER}" (欄 {TARGET_COL}) 是否應對應到來源標頭 "{SOURCE_HEADER}" (欄 {SOURCE_COL})？',
    noIssuesFound: '檢查完畢，未發現問題。所有已設定的對應欄位標頭均完全匹配。',
    checkSuccessHeaders: '標頭檢查通過。所有對應欄位的標頭均正確對應。',
    checkFailureHeaders: '發現標頭不符:',
    checkSuccessColumn: '欄位檢查通過，沒有發現空值。',
    checkFailureColumn: '發現空值:',
    emptyValuesInRows: '在來源欄位 {COLUMN} 的以下列中發現空值: {ROWS}',
    columnIsEmpty: '整個來源欄位 {COLUMN} 為空或無法讀取。',
    noMappingsToCheck: '未設定任何欄位對應，無需檢查。',
    emptyCheckTitle: '來源欄位空值檢查',
    emptyCheckSuccess: '檢查完畢，所有已對應的來源欄位中均未發現空值。',
    emptyCheckFailure: '在以下來源欄位中發現空值：',
    checkEmptyValuesButton: '檢查來源空值',
    firstColumnMismatch: '第一起始欄位不匹配',
  }
};

/**
 * Gets the appropriate translation object based on the user's locale.
 */
function _MDV_getTranslations() {
  const locale = Session.getActiveUserLocale();
  return _MDV_TRANSLATIONS[locale] || _MDV_TRANSLATIONS.en;
}

/**
 * Saves Data Validation Settings for a specific sheet using PropertiesService.
 * @param {object} settings The settings object from the UI.
 * @param {string} sheetName The name of the sheet to save settings for.
 * @returns {{success: boolean, message: string}} Result object.
 */
function saveVerifySettings(settings, sheetName) {
  const T = getTranslations();
  try {
    if (!sheetName) {
      throw new Error("工作表名稱為必填項，無法儲存設定。");
    }
    if (!settings.sourceDataUrl || !settings.sourceDataSheetName || !settings.targetSheetName) {
      throw new Error(T.errorVerifyUrlRequired);
    }
    const properties = PropertiesService.getDocumentProperties();
    const key = `verifySettings_${sheetName}`;
    properties.setProperty(key, JSON.stringify(settings));
    return { success: true, message: T.saveSuccess };
  } catch (e) {
    Logger.log(`Error saving verify settings for sheet ${sheetName}: ${e.message}`);
    return { success: false, message: `${T.saveFailure}: ${e.message}` };
  }
}

/**
 * Gets Data Validation Settings for a specific sheet from PropertiesService.
 * @param {string} sheetName The name of the sheet to get settings for. If null, uses the active sheet.
 * @returns {object} The saved settings object.
 */
function getVerifySettings(sheetName) {
  const currentSheetName = sheetName || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  const properties = PropertiesService.getDocumentProperties();
  const key = `verifySettings_${currentSheetName}`;
  const settingsString = properties.getProperty(key);
  let settings;
  if (!settingsString) {
    settings = {};
  } else {
    try {
      settings = JSON.parse(settingsString);
    } catch (e) {
      Logger.log(`Error parsing verify settings for sheet ${currentSheetName}: ${e.message}`);
      settings = {};
    }
  }
  return {
    targetSheetName: settings.targetSheetName || currentSheetName,
    startRow: settings.startRow ? parseInt(settings.startRow, 10) : '',
    sourceDataUrl: settings.sourceDataUrl || '',
    sourceDataSheetName: settings.sourceDataSheetName || '',
    mismatchColumn: settings.mismatchColumn || '',
    targetHeaderRow: settings.targetHeaderRow ? parseInt(settings.targetHeaderRow, 10) : '',
    sourceHeaderRow: settings.sourceHeaderRow ? parseInt(settings.sourceHeaderRow, 10) : '',
    validationMappings: settings.validationMappings || [],
    outputMappings: settings.outputMappings || []
  };
}


// ================================================================
// SECTION 1: Main Validation Functions
// ================================================================

function runValidationMsOnly() { runDataValidation('MS_ONLY'); }

/**
 * NEW: Triggers the new, safer reset process.
 */
function runReset() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheetName = ss.getActiveSheet().getName();
    try {
        ss.toast('Resetting validation data...', 'Processing', 5);
        const settings = getVerifySettings(activeSheetName);
        if (!settings.targetSheetName) {
          throw new Error(`No settings found for sheet "${activeSheetName}".`);
        }
        resetValidationData(settings);
        ss.toast('Validation data has been reset.', 'Success', 5);
    } catch (e) {
        ss.toast(`Error during reset: ${e.message}`, 'Error', 10);
        Logger.log('Error during reset: ' + e.stack);
    }
}

/**
 * Main validation function with header check and quota-based post-processing.
 */
function runDataValidation(mode) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const activeSheetName = ss.getActiveSheet().getName();
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteProperty('stopValidationRequested');

     try {
        const T = _MDV_getTranslations(); // Get translations
        const modeText = mode === 'MS_ONLY' ? '(MS Mode)' : '(EX Expanded Mode)';
        ss.toast(`Starting "Data Validation" ${modeText}...`, 'Processing', 3);
        const settings = getVerifySettings(activeSheetName);

        // UPDATED: Use ui.alert for pre-flight checks
        if (!settings.sourceDataUrl) {
            const errorMessage = T.errorNoValidationSettingsFound.replace('{SHEET_NAME}', activeSheetName);
            ui.alert(T.validationFailedTitle, errorMessage, ui.ButtonSet.OK);
            return;
        }
        if (!settings.validationMappings || settings.validationMappings.length === 0) {
            ui.alert(T.validationFailedTitle, T.errorNoValidationMappings, ui.ButtonSet.OK);
            return;
        }

        const sheet = ss.getSheetByName(settings.targetSheetName);
        if (!sheet) throw new Error(`Target sheet "${settings.targetSheetName}" not found.`);
        const sheetUrl = ss.getUrl();
        const sheetGid = sheet.getSheetId();

        const headerCheckResult = checkSourceHeaderMapping(settings);
        if (!headerCheckResult.isPerfect) {
            const mismatchMessage = headerCheckResult.mismatches.map(m => m.message).join('\n');
            const suggestionMessage = headerCheckResult.suggestions.map(s => s.message).join('\n');
            let alertMsg = '';
            if (mismatchMessage) {
                alertMsg += T.checkResultMismatches + '\n' + mismatchMessage + '\n\n';
            }
            if (suggestionMessage) {
                alertMsg += T.checkResultSuggestions + '\n' + suggestionMessage;
            }
            ui.alert(T.checkResultTitle, alertMsg, ui.ButtonSet.OK);
            return;
        }

        const lastRow = sheet.getLastRow();
        if (lastRow < settings.startRow) {
            ui.alert('No Data to Process', 'No data rows found below the configured "Data Start Row".', ui.ButtonSet.OK);
            return;
        }
        
        const initialRange = sheet.getRange(settings.startRow, 1, lastRow - settings.startRow + 1, sheet.getMaxColumns());
        const initialData = initialRange.getValues();
        const initialRichText = initialRange.getRichTextValues();

        const tasksToProcess = buildTaskListFromData(initialData, settings);
        const taskMap = new Map(tasksToProcess.map(t => [t.originalRowIndex, t]));
        
        if (tasksToProcess.length === 0) {
            ui.alert('No Processable Rows Found', 'No rows with data in the validation columns were found.', ui.ButtonSet.OK);
            return;
        }

        const externalDataMap = fetchExternalData(settings);
        let processedResults = new Map();
        for (const task of tasksToProcess) {
            if (scriptProperties.getProperty('stopValidationRequested') === 'true') {
                throw new Error('Validation process was manually stopped by the user.');
            }
            const childResult = processSingleTask(task, externalDataMap, settings, mode, sheetUrl, sheetGid, T);
            processedResults.set(task.originalRowIndex, childResult);
        }

        processedResults = cleanupMsRowGroups(processedResults, taskMap);

        let finalSheetLayout = [];
        let finalFormattingTasks = [];

        for (let i = 0; i < initialData.length; i++) {
            const originalRowData = initialData[i];
            const originalRichText = initialRichText[i];
            const currentRowIndexInLayout = finalSheetLayout.length;

            finalSheetLayout.push(originalRowData);
            
            for (let j = 0; j < originalRichText.length; j++) {
                const rtValue = originalRichText[j];
                if (rtValue.getLinkUrl() || rtValue.getRuns().length > 1) {
                    finalFormattingTasks.push({
                        type: 'richText',
                        row: settings.startRow + currentRowIndexInLayout,
                        col: j + 1,
                        value: rtValue
                    });
                }
            }

            if (processedResults.has(i)) {
                const result = processedResults.get(i);
                result.childRows.forEach(child => {
                    const childRowIndexInLayout = finalSheetLayout.length;
                    finalSheetLayout.push(child.rowData);
                    child.formattingTasks.forEach(fmt => {
                        fmt.row = settings.startRow + childRowIndexInLayout;
                        finalFormattingTasks.push(fmt);
                    });
                });
            }
        }

        initialRange.clear();
        SpreadsheetApp.flush();

        if (finalSheetLayout.length > 0) {
            const outputRange = sheet.getRange(settings.startRow, 1, finalSheetLayout.length, sheet.getMaxColumns());
            outputRange.setValues(finalSheetLayout);
        }

        finalFormattingTasks.forEach(fmtTask => {
            if (fmtTask.type === 'richText') {
                const cell = sheet.getRange(fmtTask.row, fmtTask.col);
                cell.setRichTextValue(fmtTask.value);
            } else if (fmtTask.type === 'background') {
                const range = sheet.getRange(fmtTask.row, 1, 1, sheet.getMaxColumns());
                range.setBackground(fmtTask.value);
            }
        });

        ss.toast('Data validation complete!', 'Success', 5);

    } catch (e) {
        ss.toast(`Error: ${e.message}`, 'Error', 10);
        Logger.log('Error during data validation: ' + e.stack);
    } finally {
        scriptProperties.deleteProperty('stopValidationRequested');
    }
}

// ================================================================
// SECTION 2: Core Logic & Pre-Check Functions
// ================================================================

/**
 * Calculates the Levenshtein distance between two strings.
 * @param {string} a The first string.
 * @param {string} b The second string.
 * @returns {number} The Levenshtein distance.
 */
function levenshteinDistance(a, b) {
    if (a.length === 0) return b.length;
    if (b.length === 0) return a.length;
    const matrix = [];
    for (let i = 0; i <= b.length; i++) { matrix[i] = [i]; }
    for (let j = 0; j <= a.length; j++) { matrix[0][j] = j; }
    for (let i = 1; i <= b.length; i++) {
        for (let j = 1; j <= a.length; j++) {
            if (b.charAt(i - 1) === a.charAt(j - 1)) {
                matrix[i][j] = matrix[i - 1][j - 1];
            } else {
                matrix[i][j] = Math.min(
                    matrix[i - 1][j - 1] + 1,
                    matrix[i][j - 1] + 1,
                    matrix[i - 1][j] + 1
                );
            }
        }
    }
    return matrix[b.length][a.length];
}

/**
 * Scans all headers and returns a comprehensive report of perfect,
 * suggested (fuzzy), and unmatched headers, with an option to exclude existing mappings.
 * @param {object} settings The validation settings for the current sheet.
 * @param {Array<object>} [mappingsToExclude=[]] An array of mappings to exclude from the results.
 * @returns {object} An object containing arrays for perfectMatches, suggestedMatches, unmatchedTarget, and unmatchedSource.
 */
function getSmartMappingResults(settings, mappingsToExclude = []) {
    const { targetHeaderRow, sourceHeaderRow, targetSheetName, sourceDataUrl, sourceDataSheetName } = settings;

    if (!targetHeaderRow || !sourceHeaderRow || !sourceDataUrl || !sourceDataSheetName) {
        throw new Error("請先完整填寫來源與目標的 URL、分頁名稱與標頭起始列。");
    }

    try {
        const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName);
        if (!targetSheet) throw new Error(`找不到目標分頁 "${targetSheetName}"。`);
        
        const sourceSs = SpreadsheetApp.openByUrl(sourceDataUrl);
        const sourceSheet = sourceSs.getSheetByName(sourceDataSheetName);
        if (!sourceSheet) throw new Error(`在來源檔案中找不到分頁 "${sourceDataSheetName}"。`);

        const targetHeadersRaw = targetSheet.getRange(targetHeaderRow, 1, 1, targetSheet.getMaxColumns()).getValues()[0];
        const sourceHeadersRaw = sourceSheet.getRange(sourceHeaderRow, 1, 1, sourceSheet.getMaxColumns()).getValues()[0];

        const excludedTargetCols = new Set(mappingsToExclude.map(m => m.targetCol));
        const excludedSourceCols = new Set(mappingsToExclude.map(m => m.sourceCol));

        let targetHeaders = targetHeadersRaw
            .map((h, i) => ({ header: (h || '').toString().trim(), col: columnToLetter(i + 1), used: false }))
            .filter(h => h.header && !excludedTargetCols.has(h.col));
            
        let sourceHeaders = sourceHeadersRaw
            .map((h, i) => ({ header: (h || '').toString().trim(), col: columnToLetter(i + 1), used: false }))
            .filter(h => h.header && !excludedSourceCols.has(h.col));

        const perfectMatches = [];
        const suggestedMatches = [];

        // Pass 1: Find perfect matches
        targetHeaders.forEach(target => {
            if (target.used) return;
            const perfectMatch = sourceHeaders.find(source => !source.used && source.header === target.header);
            if (perfectMatch) {
                perfectMatches.push({
                    targetCol: target.col,
                    targetHeader: target.header,
                    sourceCol: perfectMatch.col,
                    sourceHeader: perfectMatch.header
                });
                target.used = true;
                perfectMatch.used = true;
            }
        });

        // Pass 2: Find fuzzy matches
        targetHeaders.forEach(target => {
            if (target.used) return;
            let bestMatch = { distance: Infinity, source: null };

            sourceHeaders.forEach(source => {
                if (source.used) return;
                const distance = levenshteinDistance(target.header, source.header);
                if (distance > 0 && distance <= 2 && distance < bestMatch.distance) {
                    bestMatch = { distance, source };
                }
            });

            if (bestMatch.source) {
                suggestedMatches.push({
                    targetCol: target.col,
                    targetHeader: target.header,
                    sourceCol: bestMatch.source.col,
                    sourceHeader: bestMatch.source.header
                });
                target.used = true;
                bestMatch.source.used = true;
            }
        });

        const unmatchedTarget = targetHeaders.filter(h => !h.used);
        const unmatchedSource = sourceHeaders.filter(h => !h.used);

        return { perfectMatches, suggestedMatches, unmatchedTarget, unmatchedSource };

    } catch (e) {
        Logger.log(`Smart mapping failed: ${e.stack}`);
        throw new Error(`智慧對應時發生錯誤: ${e.message}`);
    }
}

/**
 * Checks headers for currently mapped fields and returns structured mismatch data.
 * @param {object} settings The validation settings for the current sheet.
 * @returns {{isPerfect: boolean, mismatches: Array<object>}} An object with detailed results.
 */
function checkSourceHeaderMapping(settings) {
    const T = _MDV_getTranslations();
    const { validationMappings, outputMappings, targetHeaderRow, sourceHeaderRow, targetSheetName, sourceDataUrl, sourceDataSheetName } = settings;
    
    if (!targetHeaderRow || !sourceHeaderRow || !sourceDataUrl || !sourceDataSheetName) {
        return { isPerfect: false, mismatches: [{ message: "請先完整填寫來源與目標的 URL、分頁名稱與標頭起始列。" }] };
    }

    const allMappings = [...(validationMappings || []), ...(outputMappings || [])];
    if (allMappings.length === 0) {
        return { isPerfect: true, mismatches: [] };
    }

    try {
        const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName);
        if (!targetSheet) return { isPerfect: false, mismatches: [{ message: `找不到目標分頁 "${targetSheetName}"。` }] };
        
        const sourceSs = SpreadsheetApp.openByUrl(sourceDataUrl);
        const sourceSheet = sourceSs.getSheetByName(sourceDataSheetName);
        if (!sourceSheet) return { isPerfect: false, mismatches: [{ message: `在來源檔案中找不到分頁 "${sourceDataSheetName}"。` }] };

        const targetHeadersRaw = targetSheet.getRange(targetHeaderRow, 1, 1, targetSheet.getMaxColumns()).getValues()[0];
        const sourceHeadersRaw = sourceSheet.getRange(sourceHeaderRow, 1, 1, sourceSheet.getMaxColumns()).getValues()[0];

        const mismatches = [];
        allMappings.forEach(mapping => {
            if (!mapping.targetCol || !mapping.sourceCol) return;
            const targetColNum = columnToNumber(mapping.targetCol);
            const sourceColNum = columnToNumber(mapping.sourceCol);
            if (targetColNum === 0 || sourceColNum === 0) return;
            const targetHeader = (targetHeadersRaw[targetColNum - 1] || '').toString().trim();
            const sourceHeader = (sourceHeadersRaw[sourceColNum - 1] || '').toString().trim();

            if (targetHeader.normalize('NFC') !== sourceHeader.normalize('NFC')) {
                mismatches.push({
                  targetCol: mapping.targetCol,
                  sourceCol: mapping.sourceCol,
                  targetHeader: targetHeader || '空白',
                  sourceHeader: sourceHeader || '空白',
                  message: T.headerMismatchDetail
                    .replace('{TARGET_COL}', mapping.targetCol)
                    .replace('{TARGET_HEADER}', targetHeader || '空白')
                    .replace('{SOURCE_COL}', mapping.sourceCol)
                    .replace('{SOURCE_HEADER}', sourceHeader || '空白')
                });
            }
        });

        return { isPerfect: mismatches.length === 0, mismatches: mismatches };

    } catch (e) {
        Logger.log(`Header check failed: ${e.stack}`);
        return { isPerfect: false, mismatches: [{ message: `檢查時發生錯誤: ${e.message}` }] };
    }
}


/**
 * Checks all currently mapped source columns for empty values.
 * @param {object} settings The validation settings object.
 * @returns {{isValid: boolean, message: string}} An object containing the validation result.
 */
function checkAllSourceColumnsForEmptyValues(settings) {
    const T = _MDV_getTranslations();
    const { sourceDataUrl, sourceDataSheetName, sourceHeaderRow, validationMappings, outputMappings } = settings;

    if (!sourceDataUrl || !sourceDataSheetName || !sourceHeaderRow) {
        return { isValid: false, message: "請先完整填寫來源 URL、分頁與標頭列。" };
    }

    const allMappings = [...(validationMappings || []), ...(outputMappings || [])];
    const sourceColumns = [...new Set(allMappings.map(m => m.sourceCol).filter(Boolean))];

    if (sourceColumns.length === 0) {
        return { isValid: true, message: T.noMappingsToCheck };
    }

    try {
        const sourceSs = SpreadsheetApp.openByUrl(sourceDataUrl);
        const sourceSheet = sourceSs.getSheetByName(sourceDataSheetName);
        if (!sourceSheet) return { isValid: false, message: `在來源檔案中找不到分頁 "${sourceDataSheetName}"。` };

        const lastRow = sourceSheet.getLastRow();
        const dataStartRow = parseInt(sourceHeaderRow, 10) + 1;

        if (lastRow < dataStartRow) {
            return { isValid: false, message: T.columnIsEmpty.replace('{COLUMN}', sourceColumns.join(', ')) };
        }

        const errorMessages = [];
        sourceColumns.forEach(col => {
            const colNum = columnToNumber(col);
            const values = sourceSheet.getRange(dataStartRow, colNum, lastRow - dataStartRow + 1, 1).getValues();
            const emptyRows = [];
            values.forEach((row, index) => {
                if (row[0] === null || row[0] === '') {
                    emptyRows.push(dataStartRow + index);
                }
            });
            if (emptyRows.length > 0) {
                errorMessages.push(T.emptyValuesInRows
                    .replace('{COLUMN}', col)
                    .replace('{ROWS}', emptyRows.join(', ')));
            }
        });

        if (errorMessages.length > 0) {
            return { isValid: false, message: `${T.emptyCheckFailure}\n- ${errorMessages.join('\n- ')}` };
        }

        return { isValid: true, message: T.emptyCheckSuccess };
    } catch (e) {
        Logger.log(`Bulk column check for empty values failed: ${e.message}`);
        return { isValid: false, message: `檢查時發生錯誤: ${e.message}` };
    }
}


/**
 * Cleans up MS row groups using "Fill the Quota" logic.
 */
function cleanupMsRowGroups(processedResults, taskMap) {
    processedResults.forEach((result, originalRowIndex) => {
        const childRows = result.childRows;
        if (!childRows || childRows.length === 0) {
            return;
        }

        const parentTask = taskMap.get(originalRowIndex);
        if (!parentTask) return;

        const parentQty = parseInt(parentTask.originalRowData[5], 10) || 0;
        if (parentQty <= 0) {
            return;
        }

        const perfectMatches = childRows.filter(child => child.isPerfect);
        const imperfectMatches = childRows.filter(child => !child.isPerfect);

        let finalChildRows = [];
        finalChildRows.push(...perfectMatches);

        const quotaNeeded = parentQty - finalChildRows.length;
        if (quotaNeeded > 0) {
            const itemsToFill = imperfectMatches.slice(0, quotaNeeded);
            finalChildRows.push(...itemsToFill);
        }
        
        result.childRows = finalChildRows;
        processedResults.set(originalRowIndex, result);
    });
    return processedResults;
}

/**
 * Generates structured mismatch information.
 */
function generateInfoTask(sourceMatch, targetValues, settings) {
    const T = _MDV_getTranslations();
    const primaryMismatches = [];
    const secondaryMismatches = [];
    const generalMismatchLabels = new Set();
    const { validationMappings } = settings;

    for (const item of validationMappings) {
        const sourceValue = (sourceMatch.values[item.sourceCol] || '').toString().trim();
        const targetValue = (targetValues[item.targetCol] || '').toString().trim();

        if (sourceValue !== targetValue) {
            const prefix = item.isRequired ? T.majorPrefix : T.minorPrefix;
            const suffix = T.mismatchSuffix;
            const sourceLocation = `${item.sourceCol}${sourceMatch.originalRowNum}`;
            const valueSegment = sourceValue ? `_${sourceValue}` : '';
            const specificLabel = `${prefix}_${sourceLocation}${valueSegment}_${suffix}`;
            
            const sourceUrl = sourceMatch.sourceUrl.split('#')[0];
            const linkUrl = `${sourceUrl}#gid=${sourceMatch.sourceGid}&range=${item.sourceCol}${sourceMatch.originalRowNum}`;

            const mismatchDetail = {
                targetCol: item.targetCol,
                label: specificLabel,
                linkUrl: linkUrl
            };

            if (item.isRequired) {
                primaryMismatches.push(mismatchDetail);
            } else {
                secondaryMismatches.push(mismatchDetail);
            }

            if (sourceValue === '') {
                generalMismatchLabels.add(T.sourceValueBlank);
            } else {
                generalMismatchLabels.add(T.sourceValueMismatch);
            }
        }
    }

    return {
        primaryMismatches: primaryMismatches,
        secondaryMismatches: secondaryMismatches,
        mismatchColumnText: [...generalMismatchLabels].join('\n')
    };
}


/**
 * Processes a single task, with plain text mismatch info.
 */
function processSingleTask(task, externalDataMap, settings, mode, sheetUrl, sheetGid, T) {
    // const T = _MDV_getTranslations();
    const { validationMappings, outputMappings } = settings;
    const majorValidationMappings = validationMappings.filter(m => m.isRequired);

    const keyMappings = majorValidationMappings.length > 0 ? majorValidationMappings : [validationMappings[0]];

    const lookupKey = keyMappings
        .map(m => task.values[m.targetCol] || '')
        .join('||');

    const candidateMatches = lookupKey ? (externalDataMap.get(lookupKey) || []) : [];

    if (candidateMatches.length === 0) {
        let newChildRow = [...task.originalRowData];
        newChildRow[0] = "MS";

        if (outputMappings) {
            outputMappings.forEach(mapping => {
                const targetColNum = columnToNumber(mapping.targetCol);
                if (targetColNum > 0) newChildRow[targetColNum - 1] = '';
            });
        }
        
        const childFormattingTasks = [{ type: 'background', value: '#fff2cc' }];

        const mismatchColNum = columnToNumber(settings.mismatchColumn);
        if (mismatchColNum > 0) {
            const firstKeyMapping = keyMappings[0];
            const originalParentRow = settings.startRow + task.originalRowIndex;
            const targetCellA1 = `${firstKeyMapping.targetCol}${originalParentRow}`;
            const val = task.values[firstKeyMapping.targetCol] || '空白';
            const mismatchText = `${T.firstColumnMismatch}_${targetCellA1}_${val}`;
            newChildRow[mismatchColNum - 1] = mismatchText;
            newChildRow[mismatchColNum - 1] = mismatchText;
        }
        
        newChildRow[5] = 1;

        return {
            childRows: [{
                rowData: newChildRow,
                formattingTasks: childFormattingTasks,
                isPerfect: false
            }]
        };
    }

    let childRows = [];
    candidateMatches.forEach(match => {
        let newChildRow = [...task.originalRowData];
        newChildRow[0] = "MS";
        let childFormattingTasks = [{ type: 'background', value: '#d9ead3' }];

        const infoResult = generateInfoTask(match, task.values, settings);
        const isPerfectMatch = infoResult.primaryMismatches.length === 0 && infoResult.secondaryMismatches.length === 0;
        
        if (majorValidationMappings.length > 0 && infoResult.primaryMismatches.length > 0) {
            return; 
        }

        outputMappings.forEach(mapping => {
            const targetColNum = columnToNumber(mapping.targetCol);
            if (targetColNum <= 0) return;

            const sourceValue = match.values[mapping.sourceCol];
            if (sourceValue === undefined || sourceValue === null || sourceValue.toString().trim() === '') {
                const label = `${mapping.sourceCol}${match.originalRowNum}${T.noSourceDataSuffix}`;
                newChildRow[targetColNum - 1] = label;
                const sourceUrl = match.sourceUrl.split('#')[0];
                const linkUrl = `${sourceUrl}#gid=${match.sourceGid}&range=${mapping.sourceCol}${match.originalRowNum}`;
                const richText = SpreadsheetApp.newRichTextValue().setText(label).setLinkUrl(0, label.length, linkUrl).build();
                childFormattingTasks.push({ type: 'richText', col: targetColNum, value: richText });
            } else {
                newChildRow[targetColNum - 1] = sourceValue;
            }
        });
        
        const mismatchColNum = columnToNumber(settings.mismatchColumn);
        if (mismatchColNum > 0) {
            const secondaryMismatchText = infoResult.secondaryMismatches.map(d => d.label).join('\n');
            const fullMismatchText = [infoResult.mismatchColumnText, secondaryMismatchText].filter(Boolean).join('\n');
            newChildRow[mismatchColNum - 1] = fullMismatchText;
            
            if (infoResult.secondaryMismatches.length > 0) {
                 const richTextBuilder = SpreadsheetApp.newRichTextValue().setText(fullMismatchText);
                 let currentIndex = (infoResult.mismatchColumnText ? infoResult.mismatchColumnText.length + 1 : 0);
                 infoResult.secondaryMismatches.forEach(detail => {
                     richTextBuilder.setLinkUrl(currentIndex, currentIndex + detail.label.length, detail.linkUrl);
                     currentIndex += detail.label.length + 1;
                 });
                 childFormattingTasks.push({ type: 'richText', col: mismatchColNum, value: richTextBuilder.build() });
            }
        }

        newChildRow[5] = 1;

        childRows.push({
            rowData: newChildRow,
            formattingTasks: childFormattingTasks,
            isPerfect: isPerfectMatch
        });
    });
    
    return {
        childRows: childRows
    };
}


/**
 * NEW: Resets validation data by deleting generated rows and clearing specific columns in parent rows.
 * This approach is more direct and robust than the previous cleanup/memory mode.
 * @param {object} settings The validation settings for the current sheet.
 */
function resetValidationData(settings) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(settings.targetSheetName);
    if (!sheet) throw new Error(`Sheet named "${settings.targetSheetName}" not found.`);

    const startRow = settings.startRow;
    let lastRow = sheet.getLastRow();
    if (lastRow < startRow) return; // Nothing to process.

    // Get the full data range to process in memory
    const range = sheet.getRange(startRow, 1, lastRow - startRow + 1, sheet.getMaxColumns());
    const values = range.getValues();

    // Identify which columns to clear in parent rows
    const columnsToClear = settings.outputMappings.map(m => columnToNumber(m.targetCol));
    if (settings.mismatchColumn) {
        columnsToClear.push(columnToNumber(settings.mismatchColumn));
    }
    const uniqueColumnsToClear = [...new Set(columnsToClear)].filter(c => c > 0);

    const rowsToDelete = [];
    const rangesToClear = [];

    // Process the data in memory to decide what to delete and what to clear
    values.forEach((row, index) => {
        const currentRowInSheet = startRow + index;
        const flag = row[0].toString().trim();

        if (flag === 'MS' || flag.startsWith('EX-')) {
            // Mark generated rows for deletion
            rowsToDelete.push(currentRowInSheet);
        } else {
            // For parent rows, identify specific cells to clear
            uniqueColumnsToClear.forEach(colNum => {
                rangesToClear.push(`${columnToLetter(colNum)}${currentRowInSheet}`);
            });
        }
    });

    // Perform batch clearing of content for parent rows
    // This is done first as it doesn't affect row indices. It preserves formatting and hyperlinks.
    if (rangesToClear.length > 0) {
        // Use getRangeList for efficiency. Chunking to avoid potential limits.
        const chunkSize = 250; 
        for (let i = 0; i < rangesToClear.length; i += chunkSize) {
            const chunk = rangesToClear.slice(i, i + chunkSize);
            sheet.getRangeList(chunk).clearContent();
        }
    }

    // Perform batch deletion of generated rows from the bottom up
    if (rowsToDelete.length > 0) {
        // Group contiguous rows together to delete in a single call
        const reversedRows = rowsToDelete.sort((a, b) => b - a); // Sort descending
        let i = 0;
        while (i < reversedRows.length) {
            const rowNum = reversedRows[i];
            let count = 1;
            // Find how many contiguous rows are next in the list
            while (i + count < reversedRows.length && reversedRows[i + count] === rowNum - count) {
                count++;
            }
            // Delete the contiguous block of rows
            sheet.deleteRows(rowNum - count + 1, count);
            i += count;
        }
    }

    SpreadsheetApp.flush();
}



// ================================================================
// SECTION 3: Utility Functions
// ================================================================

function buildTaskListFromData(initialData, settings) {
    const tasks = [];
    const { validationMappings } = settings;

    if (!validationMappings || validationMappings.length === 0) {
        throw new Error("Configuration error: No validation mappings are defined.");
    }

    initialData.forEach((row, index) => {
        const flag = row[0].toString().trim();
        if (flag === 'MS' || flag.startsWith('EX-')) return;

        const hasDataToValidate = validationMappings.some(map => {
            const colNum = columnToNumber(map.targetCol);
            return row[colNum - 1] !== '';
        });

        if (hasDataToValidate) {
            let taskData = { originalRowIndex: index, values: {}, originalRowData: row };
            const allMappings = [...settings.validationMappings, ...settings.outputMappings];
            allMappings.forEach(map => {
                const colNum = columnToNumber(map.targetCol);
                if (colNum > 0) {
                  taskData.values[map.targetCol] = (row[colNum - 1] || '').toString().trim();
                }
            });
            tasks.push(taskData);
        }
    });
    return tasks;
}

function fetchExternalData(settings) {
    const sourceSpreadsheet = SpreadsheetApp.openByUrl(settings.sourceDataUrl);
    const sourceSheet = sourceSpreadsheet.getSheetByName(settings.sourceDataSheetName);
    if (!sourceSheet) throw new Error(`Sheet named "${settings.sourceDataSheetName}" not found in the source file.`);

    const sourceGid = sourceSheet.getSheetId();
    const sourceUrl = sourceSpreadsheet.getUrl();

    const { validationMappings, outputMappings } = settings;
    const allMappings = [...validationMappings, ...outputMappings];
    const allSourceColumns = [...new Set(allMappings.map(m => m.sourceCol))];
    
    if (allSourceColumns.length === 0) {
        throw new Error("No source columns have been defined in the settings mappings.");
    }
    if (validationMappings.length === 0) {
        throw new Error("At least one validation field must be defined in Settings for the script to run.");
    }
    
    const majorValidationMappings = validationMappings.filter(m => m.isRequired);
    const keyMappings = majorValidationMappings.length > 0 ? majorValidationMappings : [validationMappings[0]];

    const maxColNum = Math.max(...allSourceColumns.map(c => columnToNumber(c)));
    const lastRow = sourceSheet.getLastRow();
    if (lastRow < 2) return new Map();

    const externalRange = sourceSheet.getRange(2, 1, lastRow - 1, maxColNum);
    const externalValues = externalRange.getDisplayValues();

    const externalDataMap = new Map();

    externalValues.forEach((row, index) => {
        let rowData = { values: {}, originalRowNum: index + 2, sourceUrl: sourceUrl, sourceGid: sourceGid };
        allSourceColumns.forEach(colLetter => {
            const colNum = columnToNumber(colLetter);
            rowData.values[colLetter] = (row[colNum - 1] || '').toString().trim();
        });

        const key = keyMappings
            .map(m => rowData.values[m.sourceCol] || '')
            .join('||');

        if (key) {
            if (!externalDataMap.has(key)) {
                externalDataMap.set(key, []);
            }
            externalDataMap.get(key).push(rowData);
        }
    });

    return externalDataMap;
}

function columnToNumber(col) {
    if (!col || typeof col !== 'string') return 0;
    let column = 0, length = col.length;
    for (let i = 0; i < length; i++) {
        column += (col.toUpperCase().charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
    }
    return column;
}

function columnToLetter(column) {
    let temp, letter = '';
    while (column > 0) {
        temp = (column - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = (column - temp - 1) / 26;
    }
    return letter;
}

function verifySumAndCumulativeValues() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheetName = ss.getActiveSheet().getName();
    const settings = getVerifySettings(activeSheetName);

    if (!settings.targetSheetName) {
        ss.toast(`No settings found for sheet "${activeSheetName}".`, 'Error', 5);
        return;
    }

    const targetSheet = ss.getSheetByName(settings.targetSheetName);
    if (!targetSheet) {
        ss.toast(`Target sheet named "${settings.targetSheetName}" not found.`, 'Error', 5);
        return;
    }

    ss.toast('Starting verification of sum and cumulative values...', 'Processing', 5);

    const startRow = settings.startRow;
    const lastRow = targetSheet.getLastRow();
    if (lastRow < startRow) {
        ss.toast('No data in the sheet to verify.', 'Info', 5);
        return;
    }

    const range = targetSheet.getRange(startRow, 1, lastRow - startRow + 1, 6);
    const values = range.getValues();
    const backgrounds = range.getBackgrounds();

    for (let i = 0; i < values.length; i++) {
        const currentRow = values[i];
        const flag = currentRow[0].toString().trim();

        if (!flag.startsWith('MS') && !flag.startsWith('EX-')) {
            let isParent = false;
            let sumQty = 0;
            let childRowCount = 0;
            let hasExRows = false;

            if (i + 1 < values.length && (values[i + 1][0].toString().trim() === 'MS' || values[i + 1][0].toString().trim().startsWith('EX-'))) {
                isParent = true;
                for (let j = i + 1; j < values.length; j++) {
                    const childFlag = values[j][0].toString().trim();
                    if (childFlag.startsWith('EX-')) {
                        hasExRows = true;
                        break;
                    }
                    if (childFlag !== 'MS') {
                        break;
                    }
                }

                for (let j = i + 1; j < values.length; j++) {
                    const childFlag = values[j][0].toString().trim();
                    if (childFlag === 'MS' || childFlag.startsWith('EX-')) {
                        let shouldSum = false;
                        if (hasExRows && childFlag.startsWith('EX-')) {
                            shouldSum = true;
                        } else if (!hasExRows && childFlag === 'MS') {
                            shouldSum = true;
                        }

                        if (shouldSum) {
                            const childQty = parseFloat(values[j][5]);
                            if (!isNaN(childQty)) {
                                sumQty += childQty;
                            }
                        }
                        childRowCount++;
                    } else {
                        break;
                    }
                }
            }

            if (isParent) {
                const parentRowIndexInSheet = i + startRow;
                const parentQty = parseFloat(currentRow[5]);
                if (isNaN(parentQty)) continue;

                const qtyCell = targetSheet.getRange(parentRowIndexInSheet, 6);
                const currentBackground = backgrounds[i][5];
                const redColor = '#ff0000';

                if (parentQty !== sumQty) {
                    if (currentBackground.toLowerCase() !== redColor) {
                        qtyCell.setBackground(redColor);
                    }
                } else {
                    if (currentBackground.toLowerCase() === redColor) {
                        qtyCell.clearFormat();
                    }
                }
                i += childRowCount;
            }
        }
    }

    SpreadsheetApp.flush();
    ss.toast('Verification of sum and cumulative values complete!', 'Success', 5);
}
