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
// SECTION 1: MANUFACTURING EXAMPLE DATA & FUNCTIONS
// ================================================================

/**
 * [REFACTORED] Entry point: Shows confirmation dialog (non-blocking).
 */
function generateManufacturingExample() {
  const T = MasterData.getTranslations();

  const htmlTemplate = HtmlService.createTemplateFromFile('MasterDataDialog');
  htmlTemplate.title = T.manufacturingGuide;
  htmlTemplate.message = T.exampleGenerationConfirmBody.replace('三個', '四個').replace('three', 'four');
  htmlTemplate.type = 'confirm';
  htmlTemplate.callback = 'generateManufacturingExample_Step2';
  htmlTemplate.args = [];
  htmlTemplate.T = T;

  SpreadsheetApp.getUi().showModalDialog(htmlTemplate.evaluate().setWidth(400).setHeight(300), T.manufacturingGuide);
}

/**
 * [REFACTORED] Step 2: Generates the complete set of sheets and data for the Manufacturing use case.
 */
function generateManufacturingExample_Step2() {
  const ui = SpreadsheetApp.getUi();
  const T = MasterData.getTranslations();

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(T.generatingExampleBody, T.generatingExampleProcess, 10);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Define configurations using translation keys
    const TARGET_SHEET_CONFIG_MFG = {
      name: T.exampleTargetSheet,
      isComplexTarget: true,
      title: "Master Data Example\n" + T.manufacturingProductionTitle,
      headers: [
        "Number", "Items", "Category", "Owner", "WareHouse", "Q'ty",
        "Item Skus", "Shipment Tracking Number", "Packing List Number", "Asset Number", "Mismatch Info",
        "Shipment Status"
      ],
      values: [],
      headerColors: [
        ["#f9a825", "#f9a825", "#f9a825", "#f9a825", "#f9a825", "#f9a825", "#1976d2", "#bbdefb", "#bbdefb", "#bbdefb", "#bbdefb", "#c8e6c9"]
      ]
    };

    const SOURCE_IMPORT_SHEET_CONFIG_MFG = {
      name: T.exampleImportSourceSheet,
      isComplexImportSource: true
    };

    const SOURCE_VERIFY_SHEET_CONFIG_MFG = {
      name: T.exampleVerifySourceSheet,
      headers: ["Item Skus", "Items", "Category", "Owner", "WareHouse", "Shipment Tracking Number", "Packing List Number", "Asset Number"],
      values: [
        ["Model 2", "Samsung v1", "Screen", "Mark", "TW", "STN-101", "PLN-101", "AN-101"],
        ["Model 2", "Samsung v1", "Screen", "Mark", "", "STN-102", "PLN-102", ""], // E3 cleared
        ["Model 3", "BOE V2", "Screen", "Mark", "TW", "STN-120", "PLN-120", "AN-120"],
        ["Model 3", "BOE V2", "Screen", "Mark", "TW", "STN-121", "", "AN-121"],
        ["Model 3", "BOE V2", "Screen", "Mark", "TW", "STN-122", "PLN-122", "AN-122"],
        ["Model 3", "NV TK1", "Motherboard", "Tom", "US", "STN-103", "PLN-103", "AN-103"],
        ["Model 3", "NV TK1", "Motherboard", "Tom", "US", "STN-104", "PLN-104", ""],
        ["Model 3", "NV TK1", "Motherboard", "Tom", "US", "STN-105", "PLN-105", "AN-105"],
        ["Model 4", "Liteon v1", "Battery", "Linda", "Japan", "STN-106", "PLN-106", "AN-106"], // E10 changed
        ["Model 4", "Liteon v1", "Battery", "Linda", "China", "STN-107", "PLN-107", "AN-107"], // E11 changed
        ["Model 4", "Liteon v1", "Battery", "Linda", "China", "STN-108", "", ""],
        ["Model 4", "Liteon v1", "Battery", "Linda3", "China", "STN-109", "PLN-109", "AN-109"], // D13, E13 changed
        ["Model 2", "Carbon v1", "Back Cover", "Amanda", "Japan", "STN-110", "PLN-110", "AN-110"],
        ["Model 2", "Carbon v1", "Back Cover", "Amanda", "Japan", "STN-111", "PLN-111", "AN-111"],
        ["Model 3", "Carbon v1", "Back Cover", "Amanda", "Japan", "STN-112", "", "AN-112"],
        ["", "Carbon v1", "Back Cover", "Amanda", "Japan", "STN-113", "PLN-113", ""], // A17 cleared
        ["Model 1", "Qualcomm v1", "RF Module", "Sam", "Korea", "STN-114", "PLN-114", "AN-114"],
        ["Model 1", "Qualcomm v1", "RF Module", "Sam", "Korea", "STN-115", "", "AN-115"],
        ["Model 4", "Qualcomm v1", "RF Module", "Sam", "Korea", "STN-116", "PLN-116", "AN-116"],
        ["Model 4", "Qualcomm v1", "RF Module", "Sam", "Korea", "STN-117", "PLN-117", "AN-117"],
        ["Model 4", "Qualcomm v1", "RF Module", "Sam", "Korea", "STN-118", "PLN-118", "AN-118"]
      ]
    };
    
    const SHIPPING_STATUS_CONFIG = {
      name: T.exampleCompareSourceSheet,
      headers: ["Shipment Tracking Number", "Shipment Status"],
      values: (function() {
        const trackingNumbers = SOURCE_VERIFY_SHEET_CONFIG_MFG.values
          .map(row => row[5])
          .filter(stn => stn && stn.trim() !== "");

        for (let i = trackingNumbers.length - 1; i > 0; i--) {
          const j = Math.floor(Math.random() * (i + 1));
          [trackingNumbers[i], trackingNumbers[j]] = [trackingNumbers[j], trackingNumbers[i]];
        }
        
        const statuses = ["Shipped", "Packing", "Pending", "Delivered", "In Transit"];
        
        return trackingNumbers.map((stn, index) => {
          const status = statuses[index % statuses.length]; 
          return [stn, status];
        });
      })()
    };

    const sheetsToCreate = [
      SHIPPING_STATUS_CONFIG,
      SOURCE_VERIFY_SHEET_CONFIG_MFG,
      SOURCE_IMPORT_SHEET_CONFIG_MFG,
      TARGET_SHEET_CONFIG_MFG
    ];
    
    const createdSheets = sheetsToCreate.map(config => createAndFormatSheet_(ss, config, false));
    
    SpreadsheetApp.flush();

    createdSheets.forEach(sheet => {
      if (sheet && sheet.getLastColumn() > 0) {
        sheet.autoResizeColumns(1, sheet.getLastColumn());
      }
    });

    SpreadsheetApp.getActiveSpreadsheet().toast(T.generationSuccessBody, T.generationSuccessTitle, 5);

  } catch (e) {
    Logger.log(`Error generating manufacturing example: ${e.stack}`);
    ui.alert('錯誤', `生成範例時發生錯誤: ${e.message}`);
  }
}

// ================================================================
// SECTION 2: BUSINESS & SALES EXAMPLE DATA & FUNCTIONS
// ================================================================

/**
 * [REFACTORED] Entry point: Shows confirmation dialog (non-blocking).
 */
function generateBusinessExample() {
  const T = MasterData.getTranslations();

  if (Session.getActiveUserLocale().startsWith('zh')) {
    T.exampleCustomerMasterSheet_Sales = '[來源] 客戶清單';
  } else {
    T.exampleCustomerMasterSheet_Sales = 'Source | Customer List';
  }

  const htmlTemplate = HtmlService.createTemplateFromFile('MasterDataDialog');
  htmlTemplate.title = T.businessGuide;
  htmlTemplate.message = T.exampleGenerationConfirmBody.replace('三個', '四個').replace('three', 'four');
  htmlTemplate.type = 'confirm';
  htmlTemplate.callback = 'generateBusinessExample_Step2';
  htmlTemplate.args = [];
  htmlTemplate.T = T;

  SpreadsheetApp.getUi().showModalDialog(htmlTemplate.evaluate().setWidth(400).setHeight(300), T.businessGuide);
}

/**
 * [REFACTORED] Step 2: Generates the complete set of sheets and data for the Business use case.
 */
function generateBusinessExample_Step2() {
  const ui = SpreadsheetApp.getUi();
  const T = MasterData.getTranslations();

  if (Session.getActiveUserLocale().startsWith('zh')) {
    T.exampleCustomerMasterSheet_Sales = '[來源] 客戶清單';
  } else {
    T.exampleCustomerMasterSheet_Sales = 'Source | Customer List';
  }

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(T.generatingExampleBody, T.generatingExampleProcess, 10);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const DASHBOARD_CONFIG = {
      name: T.exampleDashboardSheet_Sales,
      isDashboard: true,
      headers: [
        T.headerOrderDate, T.headerCustomerID, T.headerCustomerName, T.headerRegion, T.headerSalesperson,
        T.headerProductName, T.headerQuantity, T.headerUnitPrice, T.headerTotalSales, T.headerMonthlyTarget, T.headerAchievementRate
      ]
    };

    const SALES_LOG_CONFIG = {
      name: T.exampleSalesLogSheet_Sales,
      headers: [T.headerOrderDate, T.headerCustomerID, T.headerProductName, T.headerQuantity, T.headerUnitPrice],
      values: [
        ["2025/07/01", "C001", T.productLaptop, 2, 45000],
        ["2025/07/02", "C003", T.productKeyboard, 5, 3200],
        ["2025/07/03", "C002", T.productMonitor, 3, 12500],
        ["2025/07/05", "C004", T.productLaptop, 1, 48000],
        ["2025/07/06", "C001", T.productMouse, 10, 1500],
        ["2025/07/08", "C005", T.productMonitor, 2, 13000],
        ["2025/07/11", "C002", T.productLaptop, 1, 46000],
        ["2025/07/15", "C003", T.productMouse, 8, 1550]
      ]
    };

    const CUSTOMER_MASTER_CONFIG = {
      name: T.exampleCustomerMasterSheet_Sales,
      headers: [T.headerCustomerID, T.headerCustomerName, T.headerRegion, T.headerSalesperson],
      values: [
        ["C001", T.customerA, T.regionNorth, "Tom"],
        ["C002", T.customerB, T.regionCentral, "Amanda"],
        ["C003", T.customerC, T.regionSouth, "Tom"],
        ["C004", T.customerD, T.regionNorth, "David"],
        ["C005", T.customerE, T.regionCentral, "Amanda"]
      ]
    };

    const PRODUCT_TARGETS_CONFIG = {
      name: T.exampleProductTargetsSheet_Sales,
      headers: [T.headerProductName, T.headerMonthlyTarget],
      values: [
        [T.productLaptop, 200000],
        [T.productMonitor, 80000],
        [T.productKeyboard, 20000],
        [T.productMouse, 25000]
      ]
    };
    
    const sheetsToCreate = [
      PRODUCT_TARGETS_CONFIG,
      CUSTOMER_MASTER_CONFIG,
      SALES_LOG_CONFIG,
      DASHBOARD_CONFIG
    ];
    
    const createdSheets = sheetsToCreate.map(config => createAndFormatSheet_(ss, config, false));

    SpreadsheetApp.flush();

    createdSheets.forEach(sheet => {
      if (sheet && sheet.getLastColumn() > 0) {
        sheet.autoResizeColumns(1, sheet.getLastColumn());
      }
    });

    SpreadsheetApp.getActiveSpreadsheet().toast(T.generationSuccessBodySales, T.generationSuccessTitle, 5);

  } catch (e) {
    Logger.log(`Error generating business example: ${e.stack}`);
    ui.alert('錯誤', `生成範例時發生錯誤: ${e.message}`);
  }
}

// ================================================================
// SECTION 3: DELETE EXAMPLE FUNCTION
// ================================================================

/**
 * [REFACTORED] Entry point: Check sheets and show confirmation (non-blocking).
 */
function deleteExampleSheets() {
  const ui = SpreadsheetApp.getUi();
  const T = MasterData.getTranslations();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();

  const en_T = MasterData.TRANSLATIONS.en;
  const zh_TW_T = MasterData.TRANSLATIONS.zh_TW;
  
  // Ensure we check for all potential localized names
  en_T.exampleCustomerMasterSheet_Sales = 'Source | Customer List';
  zh_TW_T.exampleCustomerMasterSheet_Sales = '[來源] 客戶清單';

  const allPossibleExampleSheetNames = new Set([
    en_T.exampleTargetSheet, zh_TW_T.exampleTargetSheet,
    en_T.exampleImportSourceSheet, zh_TW_T.exampleImportSourceSheet,
    en_T.exampleVerifySourceSheet, zh_TW_T.exampleVerifySourceSheet,
    en_T.exampleCompareSourceSheet, zh_TW_T.exampleCompareSourceSheet,
    en_T.exampleDashboardSheet_Sales, zh_TW_T.exampleDashboardSheet_Sales,
    en_T.exampleSalesLogSheet_Sales, zh_TW_T.exampleSalesLogSheet_Sales,
    en_T.exampleCustomerMasterSheet_Sales, zh_TW_T.exampleCustomerMasterSheet_Sales,
    en_T.exampleProductTargetsSheet_Sales, zh_TW_T.exampleProductTargetsSheet_Sales,
  ]);

  const sheetsToDelete = allSheets.filter(sheet => allPossibleExampleSheetNames.has(sheet.getName()));

  if (sheetsToDelete.length === 0) {
    ui.alert(T.noExampleSheetsFound);
    return;
  }

  const sheetNames = sheetsToDelete.map(sheet => sheet.getName());
  const sheetList = sheetNames.map(name => `- ${name}`).join('\n');
  const confirmMessage = T.deleteExampleConfirmBody.replace('{SHEET_LIST}', sheetList);
  
  // Use MasterDataDialog for confirmation
  const htmlTemplate = HtmlService.createTemplateFromFile('MasterDataDialog');
  htmlTemplate.title = T.deleteExampleConfirmTitle;
  htmlTemplate.message = confirmMessage;
  htmlTemplate.type = 'confirm';
  htmlTemplate.callback = 'deleteExampleSheets_Step2';
  htmlTemplate.args = [sheetNames];
  htmlTemplate.T = T;

  SpreadsheetApp.getUi().showModalDialog(htmlTemplate.evaluate().setWidth(400).setHeight(400), T.deleteExampleConfirmTitle);
}

/**
 * [REFACTORED] Step 2: Execute deletion.
 */
function deleteExampleSheets_Step2(sheetNames) {
  const T = MasterData.getTranslations();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  SpreadsheetApp.getActiveSpreadsheet().toast('Deleting sheets...', T.generatingExampleProcess, 10);

  if (sheetNames && Array.isArray(sheetNames)) {
    sheetNames.forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (sheet) {
        ss.deleteSheet(sheet);
      }
    });
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(T.deleteExampleSuccess, T.generationSuccessTitle, 5);
}


// ================================================================
// SECTION 4: SHARED HELPER FUNCTIONS
// ================================================================

function createAndFormatSheet_(ss, sheetConfig, autoResize = true) {
  let sheet = ss.getSheetByName(sheetConfig.name);
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  sheet = ss.insertSheet(sheetConfig.name);

  if (sheetConfig.isComplexTarget) {
    const titleRange = sheet.getRange("A1");
    titleRange.setValue(sheetConfig.title)
      .setFontWeight('bold')
      .setFontSize(12)
      .setVerticalAlignment('middle');
    sheet.getRange(1, 1, 2, sheetConfig.headers.length).merge();
    
    const headerRange = sheet.getRange(3, 1, 1, sheetConfig.headers.length);
    headerRange.setValues([sheetConfig.headers]);
    headerRange.setBackgrounds(sheetConfig.headerColors).setFontWeight('bold');
    
    sheet.getRange("G3").setFontColor("#ffffff");

    if (sheetConfig.values && sheetConfig.values.length > 0) {
      const dataRange = sheet.getRange(4, 1, sheetConfig.values.length, sheetConfig.values[0].length);
      dataRange.setValues(sheetConfig.values);
    }
    
  } else if (sheetConfig.isComplexImportSource) {
    sheet.getRange("F1:I1").merge().setValue("Item Skus").setBackground("#1976d2").setFontColor("#ffffff").setHorizontalAlignment("center").setVerticalAlignment("middle").setFontWeight('bold');
    sheet.getRange("F2:I2").setValues([["Model 1", "Model 2", "Model 3", "Model 4"]]).setFontWeight('bold');
    
    const verticalText = "CPU:\nMEM:\nSSD:\nDisplay:\nBattey:";
    sheet.getRange("F3:F7").merge().setValue(verticalText).setVerticalAlignment("top").setWrap(true);
    sheet.getRange("G3:G7").merge().setValue(verticalText).setVerticalAlignment("top").setWrap(true);
    sheet.getRange("H3:H7").merge().setValue(verticalText).setVerticalAlignment("top").setWrap(true);
    sheet.getRange("I3:I7").merge().setValue(verticalText).setVerticalAlignment("top").setWrap(true);

    sheet.getRange("F8:I8").merge().setValue("Q'ty").setHorizontalAlignment("center").setFontWeight('bold');

    const mainHeaderRange = sheet.getRange("A8:E8");
    mainHeaderRange.setValues([["Number", "Owner", "Items", "Category", "WareHouse"]]);
    mainHeaderRange.setBackground("#f9a825").setFontWeight('bold');

    const mainData = [
        ["1", "Mark", "Samsung v1", "Screen", "TW"],
        ["2", "Tom", "NV TK1", "", "US"],
        ["3", "Linda", "Liteon v1", "Battery", ""],
        ["4", "Amanda", "Carbon v1", "", "Japan"],
        ["5", "Sam", "Qualcomm v1", "RF Module", "Korea"],
        ["6", "Mark", "BOE V2", "Screen", ""]
    ];
    sheet.getRange("A9:E14").setValues(mainData);

    const matrixData = [
        ["", 2, "", ""],
        ["", "", 3, ""],
        ["", "", "", 4],
        ["", 2, 1, 1],
        [2, "", "", 3],
        ["", "", 3, ""]
    ];
    sheet.getRange("F9:I14").setValues(matrixData);

  } else if (sheetConfig.isDashboard) {
    const headerRange = sheet.getRange(1, 1, 1, sheetConfig.headers.length);
    headerRange.setValues([sheetConfig.headers]);
    headerRange.setFontWeight('bold').setBackground("#e0e7ff");
  } else {
    const data = [sheetConfig.headers, ...sheetConfig.values];
    if (data.length > 1 && data[0].length > 0) {
        const range = sheet.getRange(1, 1, data.length, data[0].length);
        range.setValues(data);
        sheet.getRange(1, 1, 1, data[0].length).setFontWeight('bold');
    }
  }
  
  if (autoResize && sheet.getLastColumn() > 0) {
    sheet.autoResizeColumns(1, sheet.getLastColumn());
  }
  
  return sheet;
}
