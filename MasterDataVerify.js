/**
 * @OnlyCurrentDoc
 */

// =================================================================================================
// =================================== VERSION & FEATURE SUMMARY ===================================
// =================================================================================================
//
// V1.9 (Mismatch UI Fix):
// - [FIXED] Hyperlinks now intelligently switch between internal (#gid=...) and external (https://...) formats.
// - [FIXED] Mismatch Info text format simplified to "Cell_Value" (e.g., "C7_Motherboard") matching Import module style.
// - Removed verbose "!= TargetValue" comparison string from the output cell.
// - Maintained all previous REST API and Permission fixes.
//
// =================================================================================================

/**
 * MasterDataAnalyzer - A Google Sheets Add-on for intelligent data operations.
 * Copyright (c) 2025 Tata Sum (mda.design)
 */

// ================================================================
// SECTION 0: 基礎設置與輔助函式 (API & Auth)
// ================================================================

/**
 * 獲取 Google Picker 所需的 API Key 與 Token。
 */
function verify_getPickerKeys() {
  try {
    const userProperties = PropertiesService.getScriptProperties();
    const apiKey = userProperties.getProperty('GOOGLE_API_KEY');
    const appId = userProperties.getProperty('GOOGLE_APP_ID');
    const oauthToken = ScriptApp.getOAuthToken();

    if (!apiKey || !appId) {
      throw new Error("未在指令碼屬性中找到 'GOOGLE_API_KEY' 或 'GOOGLE_APP_ID'。請檢查專案設定。");
    }

    return {
      apiKey: apiKey,
      appId: appId,
      oauthToken: oauthToken
    };
  } catch (e) {
    Logger.log(`Error in verify_getPickerKeys: ${e.message}`);
    throw e;
  }
}

/**
 * 獲取當前活動的工作表名稱。
 */
function verify_getActiveSheetName() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

/**
 * [核心函式] 透過 Google Sheets API v4 讀取資料。
 */
function verify_fetchDataFromApi_(fileId, sheetName, rangeA1) {
    if (!fileId || !sheetName || !rangeA1) {
        throw new Error("API 呼叫失敗：缺少必要的檔案 ID、分頁名稱或範圍。");
    }

    const token = ScriptApp.getOAuthToken();
    const encodedRange = encodeURIComponent(`'${sheetName}'!${rangeA1}`);
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${fileId}/values/${encodedRange}`;
    
    const options = {
        method: 'get',
        headers: {
            Authorization: 'Bearer ' + token
        },
        muteHttpExceptions: true
    };
    
    try {
        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        const responseBody = response.getContentText();
        
        if (responseCode !== 200) {
            Logger.log(`API Error fetching ${url}: ${responseCode} - ${responseBody}`);
            throw new Error(`無法讀取外部檔案 (API 錯誤代碼: ${responseCode})。請確認您已選取該檔案。`);
        }
        
        const result = JSON.parse(responseBody);
        return result.values || []; 
    } catch (e) {
        Logger.log(`verify_fetchDataFromApi_ Exception: ${e.message}`);
        throw e;
    }
}

/**
 * 透過 API 獲取指定分頁名稱的 Sheet ID (GID)。
 */
function verify_getSheetGidByName_(fileId, sheetName) {
    const T = MasterData.getTranslations();
    if (!fileId || !sheetName) return null;

    try {
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${fileId}?fields=sheets(properties(sheetId,title))`;
        const options = {
            method: 'get',
            headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
            muteHttpExceptions: true,
        };

        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        
        if (responseCode !== 200) {
            Logger.log(`API Error getting GID: ${responseCode}`);
            throw new Error(T.errorInvalidUrl + ` (API Error: ${responseCode})`);
        }

        const data = JSON.parse(response.getContentText());
        const sheet = data.sheets.find(s => s.properties.title === sheetName);
        
        return sheet ? sheet.properties.sheetId.toString() : null;

    } catch (e) {
        Logger.log(`Error in verify_getSheetGidByName_: ${e.message}`);
        throw new Error(`無法獲取分頁 ID: ${e.message}`);
    }
}

/**
 * 獲取分頁名稱。
 */
function verify_getSheetNames(fileId) {
    const T = MasterData.getTranslations();
    
    // 情況 1: 目標試算表 (當前檔案)
    if (!fileId) {
        try {
            const activeSs = SpreadsheetApp.getActiveSpreadsheet();
            return activeSs.getSheets().map(sheet => sheet.getName());
        } catch (e) {
            Logger.log(`Error getting active sheets: ${e.message}`);
            throw new Error("無法讀取當前試算表的分頁。");
        }
    }
    
    // 情況 2: 來源試算表 (外部檔案，透過 API)
    try {
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
        
        if (responseCode === 200) {
            const data = JSON.parse(response.getContentText());
            if (data.sheets && data.sheets.length > 0) {
                return data.sheets.map(sheet => sheet.properties.title);
            }
            return [];
        } else {
            Logger.log(`Sheets API Error (GetNames) for fileId ${fileId}: ${responseCode}`);
            throw new Error((T.errorInvalidUrl || "無效的 URL") + ` (API Error: ${responseCode})`);
        }
    } catch (e) {
        Logger.log(`Error in verify_getSheetNames: ${e.message}`);
        throw new Error(`無法讀取檔案分頁: ${e.message}`);
    }
}

/**
 * 驗證使用者輸入的來源與目標資訊是否有效。
 */
function verify_validateInputs(fileId, sourceSheetName, targetSheetName) {
    const T = MasterData.getTranslations();
    const errors = {
        sourceFileIdError: '',
        sourceSheetError: '',
        targetSheetError: ''
    };

    if (targetSheetName) {
        const activeSs = SpreadsheetApp.getActiveSpreadsheet();
        if (!activeSs.getSheetByName(targetSheetName)) {
            const msgTemplate = T.errorTargetSheetNotFound || "找不到目標分頁 '{SHEET_NAME}'";
            errors.targetSheetError = msgTemplate.replace('{SHEET_NAME}', targetSheetName);
        }
    }

    if (fileId) {
        try {
            const sheetNames = verify_getSheetNames(fileId);
            if (sourceSheetName && !sheetNames.includes(sourceSheetName)) {
                const msgTemplate = T.errorSheetNotFoundInUrl || "在來源檔案中找不到分頁 '{SHEET_NAME}'";
                errors.sourceSheetError = msgTemplate.replace('{SHEET_NAME}', sourceSheetName);
            }
        } catch (e) {
            errors.sourceFileIdError = "無法存取此檔案或 ID 無效 (請確認您已透過 Picker 選擇此檔案)。";
        }
    } else {
        if (sourceSheetName) {
             errors.sourceFileIdError = "請先選擇來源檔案。";
        }
    }

    return errors;
}

// ================================================================
// SECTION 1: 設定檔存取
// ================================================================

function saveVerifySettings(settings, sheetName) {
  const T = MasterData.getTranslations();
  try {
    if (!sheetName) {
      throw new Error("工作表名稱為必填項，無法儲存設定。");
    }
    
    if ((!settings.sourceFileId && !settings.sourceDataUrl) || !settings.sourceDataSheetName || !settings.targetSheetName) {
      throw new Error(T.errorVerifyUrlRequired || "請填寫完整的來源與目標資訊。");
    }
    
    const properties = PropertiesService.getDocumentProperties();
    const key = `verifySettings_${sheetName}`;
    properties.setProperty(key, JSON.stringify(settings));
    return { success: true, message: T.saveSuccess || "設定已儲存" };
  } catch (e) {
    Logger.log(`Error saving verify settings for sheet ${sheetName}: ${e.message}`);
    const failMsg = T.saveFailure || "儲存失敗";
    return { success: false, message: `${failMsg}: ${e.message}` };
  }
}

function getVerifySettings(sheetName) {
  const currentSheetName = sheetName || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  const properties = PropertiesService.getDocumentProperties();
  const key = `verifySettings_${currentSheetName}`;
  const settingsString = properties.getProperty(key);
  let settings = {};
  
  if (settingsString) {
    try {
      settings = JSON.parse(settingsString);
    } catch (e) {
      Logger.log(`Error parsing verify settings for sheet ${currentSheetName}: ${e.message}`);
      settings = {};
    }
  }

  if (!settings.sourceFileId && settings.sourceDataUrl) {
      const match = settings.sourceDataUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
      if (match && match[1]) {
          settings.sourceFileId = match[1];
      }
  }

  return {
    targetSheetName: settings.targetSheetName || currentSheetName,
    startRow: settings.startRow ? parseInt(settings.startRow, 10) : '',
    sourceDataUrl: settings.sourceDataUrl || '',
    sourceFileId: settings.sourceFileId || '', 
    sourceFileName: settings.sourceFileName || '', 
    sourceDataSheetName: settings.sourceDataSheetName || '',
    mismatchColumn: settings.mismatchColumn || '',
    targetHeaderRow: settings.targetHeaderRow ? parseInt(settings.targetHeaderRow, 10) : '',
    sourceHeaderRow: settings.sourceHeaderRow ? parseInt(settings.sourceHeaderRow, 10) : '',
    validationMappings: settings.validationMappings || [],
    outputMappings: settings.outputMappings || []
  };
}

// ================================================================
// SECTION 2: 執行入口
// ================================================================

function runValidationMsOnly() { 
    runDataValidation('MS_ONLY'); 
}

function runReset() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheetName = ss.getActiveSheet().getName();
    const T = MasterData.getTranslations();
    try {
        ss.toast(T.resetValidationStart || 'Starting verification data reset...', 'Processing', 5);
        const settings = getVerifySettings(activeSheetName);
        if (!settings.targetSheetName) {
            const errorMsg = (T.errorSettingsNotFound || 'Settings not found for sheet "{SHEET_NAME}".').replace('{SHEET_NAME}', activeSheetName);
            throw new Error(errorMsg);
        }
        resetValidationData(settings);
        ss.toast(T.resetValidationComplete || 'Verification data reset complete.', 'Success', 5);
    } catch (e) {
        const errorTemplate = T.resetValidationError || 'Error during reset: {MESSAGE}';
        ss.toast(errorTemplate.replace('{MESSAGE}', e.message), 'Error', 10);
        Logger.log('Error during reset: ' + e.stack);
    }
}

function runDataValidation(mode) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const activeSheetName = ss.getActiveSheet().getName();
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteProperty('stopValidationRequested');

     try {
        const T = MasterData.getTranslations();
        const modeText = mode === 'MS_ONLY' ? (T.validationModeMs || '(MS Mode)') : (T.validationModeEx || '(EX Expanded Mode)');
        const startMsg = (T.validationStartToast || 'Starting "Data Validation" {MODE}...').replace('{MODE}', modeText);
        ss.toast(startMsg, T.toastTitleProcessing || 'Processing', 3);
        
        const settings = getVerifySettings(activeSheetName);

        if (!settings.sourceFileId && !settings.sourceDataUrl) {
            const msgTemplate = T.errorNoValidationSettingsFound || "找不到工作表 '{SHEET_NAME}' 的驗證設定。";
            const errorMessage = msgTemplate.replace('{SHEET_NAME}', activeSheetName);
            ui.alert(T.validationFailedTitle || "驗證失敗", errorMessage, ui.ButtonSet.OK);
            return;
        }
        if (!settings.validationMappings || settings.validationMappings.length === 0) {
            ui.alert(T.validationFailedTitle || "驗證失敗", T.errorNoValidationMappings || "未設定驗證對應欄位。", ui.ButtonSet.OK);
            return;
        }

        // [NEW] Integrity Check
        const integrityResult = verify_checkTargetColumnsForIntegrity(settings);
        if (!integrityResult.isValid) {
            const response = ui.alert(
                integrityResult.title || "資料完整性警告",
                integrityResult.warning,
                ui.ButtonSet.YES_NO
            );
            
            // 如果使用者選擇「否」 (NO)，則中止
            if (response === ui.Button.NO) {
                ss.toast('已取消驗證程序。', 'Info', 3);
                return;
            }
            // 如果選擇「是」 (YES)，則繼續往下執行
        }

        const sheet = ss.getSheetByName(settings.targetSheetName);
        if (!sheet) throw new Error(`找不到目標分頁 "${settings.targetSheetName}"。`);
        
        const headerCheckResult = checkSourceHeaderMapping(settings);
        if (!headerCheckResult.isPerfect) {
            const mismatchMessage = headerCheckResult.mismatches.map(m => m.message).join('\n');
            let alertMsg = (T.checkResultMismatches || "發現標頭不符：") + '\n' + mismatchMessage + '\n\n';
            ui.alert(T.checkResultTitle || "檢查結果", alertMsg, ui.ButtonSet.OK);
            return;
        }

        const lastRow = sheet.getLastRow();
        if (lastRow < settings.startRow) {
            ui.alert('無資料', '在設定的起始列之後找不到任何資料。', ui.ButtonSet.OK);
            return;
        }
        
        const initialRange = sheet.getRange(settings.startRow, 1, lastRow - settings.startRow + 1, sheet.getMaxColumns());
        const initialData = initialRange.getValues();
        const initialRichText = initialRange.getRichTextValues();

        const tasksToProcess = buildTaskListFromData(initialData, settings);
        const taskMap = new Map(tasksToProcess.map(t => [t.originalRowIndex, t]));
        
        if (tasksToProcess.length === 0) {
            ui.alert('無可處理資料', '找不到包含驗證欄位資料的列。', ui.ButtonSet.OK);
            return;
        }

        const externalDataMap = fetchExternalData(settings);
        
        let processedResults = new Map();
        // [UPDATED] Pass current Spreadsheet ID for internal link checking
        const targetSsId = ss.getId();

        for (const task of tasksToProcess) {
            if (scriptProperties.getProperty('stopValidationRequested') === 'true') {
                throw new Error('使用者手動停止了驗證程序。');
            }
            // [UPDATED] Pass targetSsId to processSingleTask
            const childResult = processSingleTask(task, externalDataMap, settings, mode, targetSsId, T);
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
            outputRange.clear({contentsOnly: true, formatOnly: true}); // [FIX] Clear residual formats (like hyperlinks)
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

        ss.toast(T.validationCompleteToast || 'Data validation complete!', T.toastTitleSuccess || 'Success', 5);

    } catch (e) {
        ss.toast(`錯誤: ${e.message}`, 'Error', 10);
        Logger.log('Error during data validation: ' + e.stack);
    } finally {
        scriptProperties.deleteProperty('stopValidationRequested');
    }
}

// ================================================================
// SECTION 3: 核心邏輯 (API 與 演算法)
// ================================================================

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

function getSmartMappingResults(settings, mappingsToExclude = []) {
    const { targetHeaderRow, sourceHeaderRow, targetSheetName, sourceFileId, sourceDataUrl, sourceDataSheetName } = settings;
    
    const fileId = sourceFileId || (sourceDataUrl ? sourceDataUrl.match(/\/d\/([a-zA-Z0-9-_]+)/)[1] : null);

    if (!targetHeaderRow || !sourceHeaderRow || !fileId || !sourceDataSheetName) {
        throw new Error("請先完整填寫來源檔案、分頁名稱與標頭起始列。");
    }

    try {
        const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName);
        if (!targetSheet) throw new Error(`找不到目標分頁 "${targetSheetName}"。`);
        
        const targetHeadersRaw = targetSheet.getRange(targetHeaderRow, 1, 1, targetSheet.getMaxColumns()).getValues()[0];
        
        const sourceRange = `${sourceHeaderRow}:${sourceHeaderRow}`;
        const sourceValues = verify_fetchDataFromApi_(fileId, sourceDataSheetName, sourceRange);
        const sourceHeadersRaw = sourceValues && sourceValues.length > 0 ? sourceValues[0] : [];

        if (sourceHeadersRaw.length === 0) {
             throw new Error("無法從來源檔案讀取到標頭列資料。");
        }

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

function checkSourceHeaderMapping(settings) {
    const T = MasterData.getTranslations();
    const { targetHeaderRow, sourceHeaderRow, targetSheetName, sourceFileId, sourceDataUrl, sourceDataSheetName, validationMappings, outputMappings } = settings;
    const fileId = sourceFileId || (sourceDataUrl ? sourceDataUrl.match(/\/d\/([a-zA-Z0-9-_]+)/)[1] : null);

    if (!targetHeaderRow || !sourceHeaderRow || !fileId || !sourceDataSheetName) {
        return { isPerfect: false, mismatches: [{ message: "請先完整填寫設定。" }] };
    }

    const allMappings = [...(validationMappings || []), ...(outputMappings || [])];
    if (allMappings.length === 0) {
        return { isPerfect: true, mismatches: [] };
    }

    try {
        const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName);
        if (!targetSheet) return { isPerfect: false, mismatches: [{ message: `找不到目標分頁 "${targetSheetName}"。` }] };
        
        const targetHeadersRaw = targetSheet.getRange(targetHeaderRow, 1, 1, targetSheet.getMaxColumns()).getValues()[0];
        
        const sourceRange = `${sourceHeaderRow}:${sourceHeaderRow}`;
        const sourceValues = verify_fetchDataFromApi_(fileId, sourceDataSheetName, sourceRange);
        const sourceHeadersRaw = sourceValues && sourceValues.length > 0 ? sourceValues[0] : [];

        const mismatches = [];
        allMappings.forEach(mapping => {
            if (!mapping.targetCol || !mapping.sourceCol) return;
            const targetColNum = columnToNumber(mapping.targetCol);
            const sourceColNum = columnToNumber(mapping.sourceCol);
            if (targetColNum === 0 || sourceColNum === 0) return;
            
            const targetHeader = (targetHeadersRaw[targetColNum - 1] || '').toString().trim();
            const sourceHeader = (sourceHeadersRaw[sourceColNum - 1] || '').toString().trim();

            if (targetHeader.normalize('NFC') !== sourceHeader.normalize('NFC')) {
                const msgTemplate = T.headerMismatchDetail || "目標 '{TARGET_COL}' ({TARGET_HEADER}) 與來源 '{SOURCE_COL}' ({SOURCE_HEADER}) 不符";
                mismatches.push({
                  targetCol: mapping.targetCol,
                  sourceCol: mapping.sourceCol,
                  targetHeader: targetHeader || '空白',
                  sourceHeader: sourceHeader || '空白',
                  message: msgTemplate
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


function checkAllSourceColumnsForEmptyValues(settings) {
    const T = MasterData.getTranslations();
    const { sourceFileId, sourceDataUrl, sourceDataSheetName, sourceHeaderRow, validationMappings, outputMappings } = settings;
    const fileId = sourceFileId || (sourceDataUrl ? sourceDataUrl.match(/\/d\/([a-zA-Z0-9-_]+)/)[1] : null);

    if (!fileId || !sourceDataSheetName || !sourceHeaderRow) {
        return { isValid: false, message: "請先完整填寫來源檔案、分頁與標頭列。" };
    }

    const allMappings = [...(validationMappings || []), ...(outputMappings || [])];
    const sourceColumns = [...new Set(allMappings.map(m => m.sourceCol).filter(Boolean))];

    if (sourceColumns.length === 0) {
        return { isValid: true, message: T.noMappingsToCheck || "沒有需要檢查的對應欄位。" };
    }

    try {
        const startRow = parseInt(sourceHeaderRow, 10) + 1;
        const errorMessages = [];
        
        for (const colLetter of sourceColumns) {
            const range = `${colLetter}${startRow}:${colLetter}`;
            const values = verify_fetchDataFromApi_(fileId, sourceDataSheetName, range);
            
            const emptyRows = [];
            values.forEach((row, index) => {
                if (!row || row.length === 0 || row[0] === null || row[0] === '') {
                    emptyRows.push(startRow + index);
                }
            });

            if (emptyRows.length > 0) {
                const displayRows = emptyRows.slice(0, 15).join(', ') + (emptyRows.length > 15 ? '...' : '');
                const msgTemplate = T.emptyValuesInRows || "{COLUMN} 欄: 第 {ROWS} 列發現空值";
                errorMessages.push(msgTemplate
                    .replace('{COLUMN}', colLetter)
                    .replace('{ROWS}', displayRows));
            }
        }

        if (errorMessages.length > 0) {
            const failTitle = T.emptyCheckFailure || "檢查發現空值";
            return { isValid: false, message: `${failTitle}\n- ${errorMessages.join('\n- ')}` };
        }

        return { isValid: true, message: T.emptyCheckSuccess || "檢查通過，來源欄位無空值。" };
    } catch (e) {
        Logger.log(`Empty check failed: ${e.message}`);
        return { isValid: false, message: `檢查時發生錯誤: ${e.message}` };
    }
}

/**
 * 檢查目標工作表中的資料驗證欄位是否包含「無來源資料」的標記 (如 '_No source data' 或 '_無來源資料')。
 * 這些標記通常來自之前的資料匯入失敗。
 */
function verify_checkTargetColumnsForIntegrity(settings) {
    const T = MasterData.getTranslations();
    const { targetSheetName, startRow, validationMappings } = settings;
    
    // 檢查兩種語言的後綴
    const suffixes = ['_No source data', '_無來源資料'];

    if (!targetSheetName || !startRow || !validationMappings || validationMappings.length === 0) {
        return { isValid: true };
    }

    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName);
        if (!sheet) return { isValid: true };

        const lastRow = sheet.getLastRow();
        if (lastRow < startRow) return { isValid: true };

        // 取得所有需要檢查的目標欄位 (去除重複)
        const targetCols = [...new Set(validationMappings.map(m => m.targetCol).filter(Boolean))];
        
        if (targetCols.length === 0) return { isValid: true };

        const errorCells = [];
        const numRowsToCheck = lastRow - startRow + 1;
        const CELL_LIMIT = 20;

        for (const colLetter of targetCols) {
            // Check if we've already hit the limit to avoid unnecessary processing
            if (errorCells.length >= CELL_LIMIT) break;

            const colIdx = columnToNumber(colLetter);
            const range = sheet.getRange(startRow, colIdx, numRowsToCheck, 1);
            const values = range.getDisplayValues(); 

            for (let i = 0; i < values.length; i++) {
                const cellVal = values[i][0];
                if (cellVal) {
                    // 只要包含任一種後綴即視為異常
                    if (suffixes.some(s => cellVal.toString().includes(s))) {
                        errorCells.push(`${colLetter}${startRow + i}`);
                        if (errorCells.length >= CELL_LIMIT) break;
                    }
                }
            }
        }

        if (errorCells.length > 0) {
            const warningTitle = T.dataIntegrityWarningTitle || "資料完整性警告";
            const msgTemplate = T.dataIntegrityWarningBody || "在以下欄位中發現資料不完整標記 ('_No source data' 或 '_無來源資料')：\n\n[ {COLS} ]\n\n這表示之前的資料匯入可能不完整，若繼續執行，將導致驗證結果錯誤。\n\n請問是否忽略此警告並繼續執行？";
            
            const displayCells = errorCells.join(', ') + (errorCells.length >= CELL_LIMIT ? '...' : '');
            // Using the same placeholder {COLS} to minimize translation changes, but it now contains cells
            const warningMsg = msgTemplate.replace('{COLS}', displayCells);
            
            return { isValid: false, warning: warningMsg, title: warningTitle };
        }

        return { isValid: true };

    } catch (e) {
        Logger.log(`Integrity check failed: ${e.message}`);
        return { isValid: true }; 
    }
}

function fetchExternalData(settings) {
    const fileId = settings.sourceFileId || (settings.sourceDataUrl ? settings.sourceDataUrl.match(/\/d\/([a-zA-Z0-9-_]+)/)[1] : null);
    const sheetName = settings.sourceDataSheetName;
    const startRow = parseInt(settings.sourceHeaderRow, 10) + 1;

    const cols = [...settings.validationMappings, ...settings.outputMappings].map(m => m.sourceCol);
    if (cols.length === 0) throw new Error("未設定任何來源欄位。");

    const maxColNum = Math.max(...cols.map(c => columnToNumber(c)));
    const maxColLetter = columnToLetter(maxColNum);
    const range = `A${startRow}:${maxColLetter}`; 
    
    const values = verify_fetchDataFromApi_(fileId, sheetName, range);
    const map = new Map();
    
    // [FIXED] Fetch correct GID using helper
    const sourceGid = verify_getSheetGidByName_(fileId, sheetName) || '0';

    const keys = settings.validationMappings.filter(m => m.isRequired);
    const keyMaps = keys.length > 0 ? keys : [settings.validationMappings[0]];

    values.forEach((row, i) => {
        let rowData = { 
            values: {}, 
            originalRowNum: startRow + i, 
            sourceFileId: fileId, // Pass fileId here
            sourceGid: sourceGid 
        };
        
        cols.forEach(colLetter => {
            const colIdx = columnToNumber(colLetter) - 1;
            rowData.values[colLetter] = (row[colIdx] || '').toString().trim();
        });

        const k = keyMaps.map(m => rowData.values[m.sourceCol] || '').join('||');
        if (k) {
            if (!map.has(k)) map.set(k, []);
            map.get(k).push(rowData);
        }
    });

    return map;
}

// ================================================================
// SECTION 4: 通用工具函式
// ================================================================

function resetValidationData(settings) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(settings.targetSheetName);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < settings.startRow) return;

    const range = sheet.getRange(settings.startRow, 1, lastRow - settings.startRow + 1, 1);
    const values = range.getValues();

    const rowsToDelete = [];
    const colsToClear = [...settings.outputMappings.map(m=>columnToNumber(m.targetCol)), columnToNumber(settings.mismatchColumn)].filter(c=>c>0);
    
    values.forEach((r, i) => {
        const flag = r[0].toString().trim();
        if (flag.match(/^(MS|EX)/)) {
            rowsToDelete.push(settings.startRow + i);
        }
    });

    if (rowsToDelete.length > 0) {
        const reversedRows = rowsToDelete.sort((a, b) => b - a);
        let i = 0;
        while (i < reversedRows.length) {
            const rowNum = reversedRows[i];
            let count = 1;
            while (i + count < reversedRows.length && reversedRows[i + count] === rowNum - count) {
                count++;
            }
            sheet.deleteRows(rowNum - count + 1, count);
            i += count;
        }
    }

    SpreadsheetApp.flush();
}

function buildTaskListFromData(initialData, settings) {
    const tasks = [];
    initialData.forEach((row, index) => {
        const flag = row[0].toString().trim();
        if (flag.match(/^(MS|EX)/)) return; 

        const hasData = settings.validationMappings.some(m => row[columnToNumber(m.targetCol) - 1]);
        if (hasData) {
            let values = {};
            [...settings.validationMappings, ...settings.outputMappings].forEach(m => {
                values[m.targetCol] = (row[columnToNumber(m.targetCol) - 1] || '').toString().trim();
            });
            tasks.push({ originalRowIndex: index, values, originalRowData: row });
        }
    });
    return tasks;
}

/**
 * [FIXED] Process tasks with intelligent hyperlinks and simplified text.
 */
function processSingleTask(task, extMap, settings, mode, targetSsId, T) {
    const keys = settings.validationMappings.filter(m => m.isRequired);
    const keyMaps = keys.length > 0 ? keys : [settings.validationMappings[0]];
    const k = keyMaps.map(m => task.values[m.targetCol]).join('||');
    const matches = extMap.get(k) || [];
    const childRows = [];

    if (matches.length === 0) {
        let newRow = [...task.originalRowData];
        newRow[0] = "MS";
        
        settings.outputMappings.forEach(map => {
            const idx = columnToNumber(map.targetCol) - 1;
            if (idx >= 0) newRow[idx] = '';
        });

        if (settings.mismatchColumn) {
            const idx = columnToNumber(settings.mismatchColumn) - 1;
            if (idx >= 0) {
                const msgTemplate = T.firstColumnMismatch || "首欄不符";
                newRow[idx] = `${msgTemplate}_${keyMaps[0].targetCol}${settings.startRow + task.originalRowIndex}_${task.values[keyMaps[0].targetCol]}`;
            }
        }
        newRow[5] = 1; 
        childRows.push({ rowData: newRow, formattingTasks: [{ type: 'background', value: '#fff2cc' }], isPerfect: false });
    } else {
        matches.forEach(m => {
            let newRow = [...task.originalRowData];
            newRow[0] = "MS";
            let fmt = [{ type: 'background', value: '#d9ead3' }];

            settings.outputMappings.forEach(map => {
                const idx = columnToNumber(map.targetCol) - 1;
                if (idx >= 0) {
                    const val = m.values[map.sourceCol];
                    if (val === '' || val === null || val === undefined) {
                        const sourceCol = map.sourceCol;
                        const sourceRow = m.originalRowNum;
                        const suffix = T.noSourceDataSuffix || '_No source data';
                        const text = `${sourceCol}${sourceRow}${suffix}`;
                        newRow[idx] = text; 
                        
                        const isInternal = (m.sourceFileId === targetSsId);
                        const linkFragment = `#gid=${m.sourceGid}&range=${sourceCol}${sourceRow}`;
                        const linkUrl = isInternal ? linkFragment : `https://docs.google.com/spreadsheets/d/${m.sourceFileId}/${linkFragment}`;
                        
                        const builder = SpreadsheetApp.newRichTextValue();
                        builder.setText(text);
                        builder.setLinkUrl(0, text.length, linkUrl);
                        
                        fmt.push({
                            type: 'richText',
                            col: idx + 1,
                            value: builder.build()
                        });

                    } else {
                        newRow[idx] = val;
                    }
                }
            });

            let mismatches = [];
            let mismatchDetails = []; 

            settings.validationMappings.forEach(vm => {
                if (m.values[vm.sourceCol] !== task.values[vm.targetCol]) {
                    // [FIXED] Simplified text format: "Cell_Value" (or Cell_NoDataSuffix)
                    const sourceVal = m.values[vm.sourceCol];
                    const displayVal = (sourceVal === '' || sourceVal === null) ? (T.noSourceDataSuffix || 'No source data') : sourceVal;
                    const text = `${vm.sourceCol}${m.originalRowNum}_${displayVal}`;
                    
                    mismatches.push(text);
                    mismatchDetails.push({
                        text: text,
                        sourceCol: vm.sourceCol,
                        row: m.originalRowNum,
                        fileId: m.sourceFileId,
                        gid: m.sourceGid
                    });
                }
            });

            if (settings.mismatchColumn && mismatches.length > 0) {
                const idx = columnToNumber(settings.mismatchColumn) - 1;
                if (idx >= 0) {
                    const fullText = mismatches.join('\n');
                    newRow[idx] = fullText;
                    
                    const builder = SpreadsheetApp.newRichTextValue();
                    builder.setText(fullText);
                    
                    let currentStart = 0;
                    mismatchDetails.forEach(detail => {
                        // [FIXED] Logic for Internal vs External Link
                        const isInternal = (detail.fileId === targetSsId);
                        const linkFragment = `#gid=${detail.gid}&range=${detail.sourceCol}${detail.row}`;
                        const linkUrl = isInternal ? linkFragment : `https://docs.google.com/spreadsheets/d/${detail.fileId}/${linkFragment}`;
                        
                        // Apply link to the full text segment "Cell_Value"
                        // Ensure we don't go out of bounds
                        const linkLength = detail.text.length;
                        
                        if (linkLength > 0 && (currentStart + linkLength) <= fullText.length) {
                             builder.setLinkUrl(currentStart, currentStart + linkLength, linkUrl);
                        }
                        currentStart += linkLength + 1; // +1 for newline
                    });
                    
                    fmt.push({
                        type: 'richText',
                        col: idx + 1,
                        value: builder.build()
                    });
                }
            }

            if (keys.length > 0 && mismatches.some(msg => keys.some(k => msg.startsWith(k.sourceCol)))) return;

            newRow[5] = 1;
            childRows.push({ rowData: newRow, formattingTasks: fmt, isPerfect: mismatches.length === 0 });
        });
    }
    return { childRows };
}

function cleanupMsRowGroups(processedResults, taskMap) {
    processedResults.forEach((result, rowIndex) => {
        const parent = taskMap.get(rowIndex);
        if (!parent) return;
        const parentQty = parseInt(parent.originalRowData[5], 10) || 0;
        if (parentQty <= 0) return;

        const perfect = result.childRows.filter(c => c.isPerfect);
        const imperfect = result.childRows.filter(c => !c.isPerfect);
        
        let final = [...perfect];
        const needed = parentQty - final.length;
        if (needed > 0) {
            final.push(...imperfect.slice(0, needed));
        }
        
        result.childRows = final;
    });
    return processedResults;
}

function columnToNumber(col) {
    if (!col) return 0;
    let c = 0;
    for (let i = 0; i < col.length; i++) c += (col.charCodeAt(i) - 64) * Math.pow(26, col.length - i - 1);
    return c;
}

function columnToLetter(col) {
    let l = '';
    while (col > 0) {
        let t = (col - 1) % 26;
        l = String.fromCharCode(t + 65) + l;
        col = (col - t - 1) / 26;
    }
    return l;
}

function verifySumAndCumulativeValues() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheetName = ss.getActiveSheet().getName();
    const settings = getVerifySettings(activeSheetName);

    if (!settings.targetSheetName) {
        ss.toast(`找不到工作表設定`, 'Error', 5);
        return;
    }

    const targetSheet = ss.getSheetByName(settings.targetSheetName);
    const startRow = settings.startRow;
    const lastRow = targetSheet.getLastRow();

    if (lastRow < startRow) return;

    const range = targetSheet.getRange(startRow, 1, lastRow - startRow + 1, 6);
    const values = range.getValues();
    const backgrounds = range.getBackgrounds();

    for (let i = 0; i < values.length; i++) {
        const row = values[i];
        const flag = row[0].toString().trim();
        
        if (!flag.match(/^(MS|EX)/)) { // 是父列
            let sumQty = 0;
            let childCount = 0;
            
            for (let j = i + 1; j < values.length; j++) {
                const childFlag = values[j][0].toString().trim();
                if (!childFlag.match(/^(MS|EX)/)) break; 
                
                const qty = parseFloat(values[j][5]);
                if (!isNaN(qty)) sumQty += qty;
                childCount++;
            }
            
            const parentQty = parseFloat(row[5]);
            const cell = targetSheet.getRange(startRow + i, 6);
            
            if (!isNaN(parentQty) && parentQty !== sumQty) {
                if (backgrounds[i][5] !== '#ff0000') cell.setBackground('#ff0000');
            } else {
                if (backgrounds[i][5] === '#ff0000') cell.clearFormat();
            }
            
            i += childCount; 
        }
    }
    ss.toast('數量驗證完成', 'Success');
}