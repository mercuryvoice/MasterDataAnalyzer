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

function showTutorialSidebar(tutorialType) {
  const T = MasterData.getTranslations();
  let title = '';
  let htmlFile = 'GuidePageTraining.html';

  if (tutorialType === 'business') {
   // generateBusinessExample(); 
    title = T.businessGuideTutorialTitle;
  } else if (tutorialType === 'manufacturing') {
   // generateManufacturingExample();
    title = T.manufacturingGuideTutorialTitle;
  } else {
    throw new Error('Unknown tutorial type specified.');
  }

  const htmlTemplate = HtmlService.createTemplateFromFile(htmlFile);
  htmlTemplate.tutorialType = tutorialType;
  htmlTemplate.translations = JSON.stringify(T); 

  const htmlOutput = htmlTemplate.evaluate().setTitle(title);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function startBusinessTutorial() {
  showTutorialSidebar('business');
}

function startManufacturingTutorial() {
  showTutorialSidebar('manufacturing');
}

function tutorial_highlightSheet(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.activate();
    }
  } catch (e) {
    Logger.log(`Could not highlight sheet "${sheetName}": ${e.message}`);
  }
}

function tutorial_highlightRange(sheetName, rangeA1) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.activate();
      const range = sheet.getRange(rangeA1);
      range.activate();
      // Removed sleep and background color change to prevent "Running script" toast
      // and rely on client-side delay for UI locking.
    }
  } catch (e) {
    Logger.log(`Could not highlight range "${rangeA1}" on sheet "${sheetName}": ${e.message}`);
  }
}

function findColumnIndexByHeader(headers, headerName) {
  if (!headers || !Array.isArray(headers)) return -1;
  const index = headers.findIndex(h => h.toString().trim() === (headerName || '').toString().trim());
  return index !== -1 ? index + 1 : -1;
}

function insertSalesMetricsFormulas() {
  const T = MasterData.getTranslations();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(T.exampleDashboardSheet_Sales);

  if (!targetSheet) {
    throw new Error(`Could not find the example sheet: "${T.exampleDashboardSheet_Sales}"`);
  }

  const lastRow = targetSheet.getLastRow();
  const headerRow = 1; 
  const startRow = 2;

  if (lastRow < startRow) {
    return; 
  }
  const numRows = lastRow - startRow + 1;
  const headers = targetSheet.getRange(headerRow, 1, 1, targetSheet.getMaxColumns()).getValues()[0];

  const qtyCol = findColumnIndexByHeader(headers, T.headerQuantity);
  const unitPriceCol = findColumnIndexByHeader(headers, T.headerUnitPrice);
  const totalSalesCol = findColumnIndexByHeader(headers, T.headerTotalSales);
  const monthlyTargetCol = findColumnIndexByHeader(headers, T.headerMonthlyTarget);
  const achievementRateCol = findColumnIndexByHeader(headers, T.headerAchievementRate);

  if (totalSalesCol > 0 && qtyCol > 0 && unitPriceCol > 0) {
    const totalSalesRange = targetSheet.getRange(startRow, totalSalesCol, numRows, 1);
    const formulaR1C1 = `=RC[${qtyCol - totalSalesCol}]*RC[${unitPriceCol - totalSalesCol}]`;
    totalSalesRange.setFormulaR1C1(formulaR1C1);
  } else {
    Logger.log("Could not find all required columns for Total Sales calculation.");
  }

  if (achievementRateCol > 0 && totalSalesCol > 0 && monthlyTargetCol > 0) {
    const achievementRateRange = targetSheet.getRange(startRow, achievementRateCol, numRows, 1);
    const formulaR1C1 = `=IF(RC[${monthlyTargetCol - achievementRateCol}]=0, 0, RC[${totalSalesCol - achievementRateCol}]/RC[${monthlyTargetCol - achievementRateCol}])`;
    achievementRateRange.setFormulaR1C1(formulaR1C1);
    achievementRateRange.setNumberFormat("0.00%");
  } else {
     Logger.log("Could not find all required columns for Achievement Rate calculation.");
  }
}
