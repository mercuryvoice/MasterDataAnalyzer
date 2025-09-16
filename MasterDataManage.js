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


// ================================================================
// SECTION 1: UI & Settings Functions
// ================================================================

/**
 * Shows the HTML settings for Data Management.
 */
function showManageSettingsSidebar() {
  const T = getTranslations();
  const htmlTemplate = HtmlService.createTemplateFromFile('SettingsPageManage');
  htmlTemplate.T = T;
  const htmlOutput = htmlTemplate.evaluate().setWidth(450).setHeight(500);
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, T.manageSettingsTitle);
}

/**
 * Gets settings for the Data Management HTML interface.
 */
function getManageSettingsForHtml() {
    const monitorSettings = getSettings().monitorSettings;
    return {
        monitorRange: monitorSettings.monitorRange || '',
        monitorRecipientEmail: monitorSettings.monitorRecipientEmail || '',
        monitorSubject: monitorSettings.monitorSubject || '',
        monitorBody: monitorSettings.monitorBody || ''
    };
}

/**
 * Saves settings from the Data Management HTML interface.
 */
function saveManageSettingsFromHtml(manageSettings) {
    const T = getTranslations();
    const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
    if (!settingsSheet) throw new Error('Sheet named "Settings" not found.');

    const values = settingsSheet.getDataRange().getValues();
    const sectionHeaderPair = ["Data Management Settings", "資料管理設定"];
    const settingLabels = {
        monitorRange: ["文件內容變更自動通知", "Automatic notification for document content changes"],
        monitorRecipientEmail: ["通知接收者 Email", "Notification Recipient Email"],
        monitorSubject: ["通知標題", "Notification Subject"],
        monitorBody: ["通知內文", "Notification Body"]
    };

    for (const key in manageSettings) {
        if (settingLabels[key]) {
            const row = findRowInSection(values, sectionHeaderPair, ...settingLabels[key]);
            if (row !== -1) {
                settingsSheet.getRange(row, 2).setValue(manageSettings[key]);
            } else {
                Logger.log(`Setting label for "${key}" not found in the Settings sheet.`);
            }
        }
    }
    SpreadsheetApp.flush();
    return T.saveSuccess;
}


// ================================================================
// SECTION 2: Data Monitoring & Notification
// ================================================================

const NOTIFY_TRIGGER_FUNCTION = 'checkAndNotify';
const NOTIFY_PROPERTY_KEY = 'monitorRangePreviousState';

function checkAndNotifyWrapper() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast('正在手動檢查變更...', '處理中', 5);
    const changed = checkAndNotify();
    if (changed) {
        ss.toast('偵測到變更並已發送通知！', '完成', 5);
    } else {
        ss.toast('檢查完成，未偵測到任何變更。', '完成', 5);
    }
}

function createOnChangeTrigger() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    deleteOnChangeTrigger(true);

    ScriptApp.newTrigger(NOTIFY_TRIGGER_FUNCTION)
        .forSpreadsheet(ss)
        .onChange()
        .create();

    storeCurrentState();

    ss.toast('已啟用自動通知功能。', '成功', 5);
}

function deleteOnChangeTrigger(silent = false) {
    const triggers = ScriptApp.getProjectTriggers();
    let deleted = false;
    triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === NOTIFY_TRIGGER_FUNCTION) {
            ScriptApp.deleteTrigger(trigger);
            deleted = true;
        }
    });

    if (!silent) {
        if (deleted) {
            SpreadsheetApp.getActiveSpreadsheet().toast('已停用自動通知功能。', '成功', 5);
        } else {
            SpreadsheetApp.getActiveSpreadsheet().toast('目前沒有已啟用的自動通知。', '資訊', 5);
        }
    }
}

function storeCurrentState() {
    try {
        const allSettings = getSettings();
        const monitorSheetName = allSettings.verifySettings.targetSheetName;
        const monitorRangeA1 = allSettings.monitorSettings.monitorRange;

        if (!monitorSheetName || !monitorRangeA1) {
            return;
        }

        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(monitorSheetName);
        if (!sheet) return;

        const range = sheet.getRange(monitorRangeA1);
        const values = range.getDisplayValues();
        const currentState = JSON.stringify(values);

        PropertiesService.getScriptProperties().setProperty(NOTIFY_PROPERTY_KEY, currentState);
        Logger.log("Stored current state for monitored range.");
    } catch (e) {
        Logger.log(`Error storing current state: ${e.message}`);
    }
}

function checkAndNotify() {
    try {
        const allSettings = getSettings();
        const monitorSettings = allSettings.monitorSettings;
        const verifySettings = allSettings.verifySettings;
        
        const monitorSheetName = verifySettings.targetSheetName;
        const monitorRangeA1 = monitorSettings.monitorRange;
        const recipientEmail = monitorSettings.monitorRecipientEmail;

        if (!monitorSheetName || !monitorRangeA1 || !recipientEmail) {
            Logger.log("Monitoring settings are incomplete, skipping notification.");
            return false;
        }

        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(monitorSheetName);
        if (!sheet) {
            Logger.log(`Sheet not found: ${monitorSheetName}`);
            return false;
        }

        const range = sheet.getRange(monitorRangeA1);
        const newValues = range.getDisplayValues();
        const storedState = PropertiesService.getScriptProperties().getProperty(NOTIFY_PROPERTY_KEY);

        if (!storedState) {
            PropertiesService.getScriptProperties().setProperty(NOTIFY_PROPERTY_KEY, JSON.stringify(newValues));
            Logger.log("No previous state found. Storing current state.");
            return false;
        }

        const oldValues = JSON.parse(storedState);
        const changes = [];
        const startRow = range.getRow();
        const startCol = range.getColumn();

        for (let r = 0; r < newValues.length; r++) {
            for (let c = 0; c < newValues[r].length; c++) {
                const oldValue = (oldValues[r] && oldValues[r][c]) ? oldValues[r][c] : "";
                const newValue = newValues[r][c] ? newValues[r][c] : "";
                if (oldValue !== newValue) {
                    const cellNotation = sheet.getRange(startRow + r, startCol + c).getA1Notation();
                    changes.push({
                        cell: cellNotation,
                        from: oldValue || "(空白)",
                        to: newValue || "(空白)"
                    });
                }
            }
        }

        if (changes.length > 0) {
            const T = getTranslations();
            const changeDetailsText = changes.slice(0, 20).map(c => `- 儲存格 ${c.cell}: 從 "${c.from}" 變為 "${c.to}"`).join('\n');
            const hasMoreChanges = changes.length > 20 ? `...還有 ${changes.length - 20} 項其他變更。\n` : '';
            const fullChangeDetails = `${changeDetailsText}\n${hasMoreChanges}`;

            let subject = monitorSettings.monitorSubject || T.defaultSubjectTemplate;
            let body = monitorSettings.monitorBody || T.defaultBodyTemplate;

            subject = subject.replace(/{SHEET_NAME}/g, monitorSheetName);
            
            body = body.replace(/{SHEET_NAME}/g, monitorSheetName)
                       .replace(/{RANGE}/g, monitorRangeA1)
                       .replace(/{TIMESTAMP}/g, new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' }))
                       .replace(/{CHANGES_COUNT}/g, changes.length)
                       .replace(/{CHANGE_DETAILS}/g, fullChangeDetails)
                       .replace(/{SHEET_URL}/g, SpreadsheetApp.getActiveSpreadsheet().getUrl());

            GmailApp.sendEmail(recipientEmail, subject, body);

            PropertiesService.getScriptProperties().setProperty(NOTIFY_PROPERTY_KEY, JSON.stringify(newValues));
            Logger.log(`Detected ${changes.length} changes and sent email to ${recipientEmail}`);
            return true;
        } else {
            Logger.log("No changes detected.");
            return false;
        }
    } catch (e) {
        Logger.log(`Error during notification: ${e.stack}`);
        return false;
    }
}


// ================================================================
// SECTION 3: Quick Sheet Deletion
// ================================================================

/**
 * Shows the UI for the Quick Delete Sheets tool.
 * This function should be called from the add-on menu.
 */
function showQuickDeleteSheetUI() {
  const T = getTranslations(); 
  const htmlTemplate = HtmlService.createTemplateFromFile('QuickDeleteSheet');
  htmlTemplate.T = T;
  const htmlOutput = htmlTemplate.evaluate().setWidth(400).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, T.quickDeleteTitle || '快速刪除分頁');
}

/**
 * Gets a list of all sheet names that can be deleted.
 * Excludes protected sheets like "Settings".
 * @returns {string[]} An array of sheet names.
 */
function getDeletableSheetNames() {
  const protectedSheets = ['Settings']; // Add other sheet names to protect here
  const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const allSheetNames = allSheets.map(sheet => sheet.getName());
  return allSheetNames.filter(name => !protectedSheets.includes(name));
}

/**
 * Deletes multiple sheets based on their names.
 * @param {string[]} sheetNames An array of names of the sheets to delete.
 * @returns {{success: boolean, message: string}} A result object.
 */
function deleteSheets(sheetNames) {
  if (!sheetNames || sheetNames.length === 0) {
    return { success: false, message: '未選擇任何分頁。' };
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let deletedCount = 0;
  let failedCount = 0;
  let failedNames = [];

  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      try {
        ss.deleteSheet(sheet);
        deletedCount++;
      } catch (e) {
        failedCount++;
        failedNames.push(name);
        Logger.log(`Failed to delete sheet "${name}": ${e.message}`);
      }
    } else {
      failedCount++;
      failedNames.push(name);
       Logger.log(`Sheet not found for deletion: "${name}"`);
    }
  });

  let message = `成功刪除 ${deletedCount} 個分頁。`;
  if (failedCount > 0) {
    message += `\n${failedCount} 個分頁刪除失敗: ${failedNames.join(', ')}。`;
  }
  
  return { success: deletedCount > 0, message: message };
}

