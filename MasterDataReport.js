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

/**
 * Saves the report settings for a specific sheet.
 * @param {object} settings The settings object from the UI.
 * @param {string} sheetName The name of the sheet to save settings for.
 * @returns {{success: boolean, message: string}} Result object.
 */
function saveReportSettings(settings, sheetName) {
    const T = getTranslations();
    try {
        if (!sheetName) throw new Error(T.saveFailureMissingSheetName); //mod
        const properties = PropertiesService.getDocumentProperties();
        const key = `reportSettings_${sheetName}`;
        properties.setProperty(key, JSON.stringify(settings));
        return { success: true, message: T.saveSuccess }; //mod
    } catch (e) {
        Logger.log(`Error saving report settings for ${sheetName}: ${e.message}`);
        return { success: false, message: `${T.saveFailure}: ${e.message}` }; //mod
    }
}

/**
 * Gets the report settings for a specific sheet.
 * @param {string} sheetName The name of the sheet to get settings for.
 * @returns {object} The saved settings object.
 */
function getReportSettings(sheetName) {
    try {
        const properties = PropertiesService.getDocumentProperties();
        const key = `reportSettings_${sheetName}`;
        const settingsString = properties.getProperty(key);
        return settingsString ? JSON.parse(settingsString) : null;
    } catch (e) {
        Logger.log(`Error getting report settings for ${sheetName}: ${e.message}`);
        return null;
    }
}


/**
 * Validates the report inputs for source URL, sheet name, and range.
 * @param {string} url The source spreadsheet URL.
 * @param {string} sheetName The source sheet name.
 * @param {string} rangeA1 The source data range.
 * @returns {object} An object containing any validation error messages.
 */
function validateReportInputs(url, sheetName, rangeA1) {
    const T = getTranslations();
    const errors = { sourceUrlError: '', sheetError: '', rangeError: '' };
    let ss;

    try {
        if (url) {
            const validUrlPattern = /^https:\/\/docs\.google\.com\/spreadsheets\/d\//;
            if (!validUrlPattern.test(url)) {
                errors.sourceUrlError = T.errorInvalidUrl; //mod
                return errors;
            }
            ss = SpreadsheetApp.openByUrl(url);
        } else {
            ss = SpreadsheetApp.getActiveSpreadsheet();
        }

        if (!sheetName) {
             errors.sheetError = T.reportSheetNameRequired; //mod
             return errors;
        }

        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
            errors.sheetError = T.sheetNotFound.replace('{SHEET_NAME}', sheetName); //mod
            return errors;
        }

        if (rangeA1) {
            try {
                sheet.getRange(rangeA1);
            } catch (e) {
                errors.rangeError = T.errorInvalidHeaderRange.replace('{RANGE}', rangeA1); //mod
            }
        } else {
            errors.rangeError = T.reportRangeRequired; //mod
        }

    } catch (e) {
        errors.sourceUrlError = T.errorInvalidUrl; //mod
    }

    return errors;
}

/**
 * Fetches headers from a specified range.
 * @param {string} url The source spreadsheet URL.
 * @param {string} sheetName The name of the source sheet.
 * @param {string} rangeA1 The A1 notation of the data range.
 * @returns {string[]} An array of header strings.
 */
function getHeadersFromRange(url, sheetName, rangeA1) {
    try {
        if (!sheetName || !rangeA1) return [];
        const ss = url ? SpreadsheetApp.openByUrl(url) : SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) throw new Error(`找不到工作表: ${sheetName}`);

        const range = sheet.getRange(rangeA1);
        const headers = sheet.getRange(range.getRow(), range.getColumn(), 1, range.getNumColumns()).getValues()[0];
        return headers.filter(h => h.toString().trim() !== '');
    } catch (e) {
        Logger.log(`Error getting headers: ${e.stack}`);
        throw e;
    }
}

/**
 * Checks specified metric fields for non-numeric values.
 * @param {string} url The source spreadsheet URL.
 * @param {string} sheetName The name of the source sheet.
 * @param {string} rangeA1 The A1 notation of the data range.
 * @param {string[]} metricHeaders An array of header names to check.
 * @returns {string[]} An array of header names that contain non-numeric data.
 */
function checkMetricFields(url, sheetName, rangeA1, metricHeaders) {
    if (!sheetName || !rangeA1 || !metricHeaders || metricHeaders.length === 0) return [];
    try {
        const ss = url ? SpreadsheetApp.openByUrl(url) : SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) return metricHeaders;

        const range = sheet.getRange(rangeA1);
        const allValues = range.getValues();
        const allHeaders = allValues[0];
        const dataRows = allValues.slice(1);
        const nonNumericFields = [];

        metricHeaders.forEach(metricHeader => {
            const colIndex = allHeaders.indexOf(metricHeader);
            if (colIndex === -1) return;

            for (let i = 0; i < Math.min(dataRows.length, 5); i++) { // Check first 5 rows
                const cellValue = dataRows[i][colIndex];
                if (cellValue !== null && cellValue !== '' && isNaN(Number(cellValue))) {
                    nonNumericFields.push(metricHeader);
                    return;
                }
            }
        });
        return nonNumericFields;
    } catch (e) {
        Logger.log(`Error checking metric fields: ${e.stack}`);
        return metricHeaders;
    }
}

/**
 * Main dynamic analysis function with smart metric handling.
 * @param {object} settings The settings object from the UI, including analysisFields.
 * @returns {object} A structured object with all analysis results or an error.
 */
function runDynamicAnalysis(settings) {
    const T = getTranslations(); //mod
    try {
        const { sheetName, rangeA1, analysisFields } = settings;
        if (!sheetName || !rangeA1) throw new Error(T.reportSheetOrRangeMissing); //mod
        if (!analysisFields || analysisFields.length === 0) throw new Error(T.reportAnalysisFieldsMissing); //mod

        const dimensions = analysisFields.filter(f => f.type === 'dimension').map(f => f.header);
        const metrics = analysisFields.filter(f => f.type === 'metric').map(f => f.header);
        
        if (dimensions.length === 0) throw new Error(T.reportDimensionMissing); //mod
        if (metrics.length === 0) throw new Error(T.reportMetricMissing); //mod

        const ss = settings.sourceUrl ? SpreadsheetApp.openByUrl(settings.sourceUrl) : SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) throw new Error(T.sheetNotFound.replace('{SHEET_NAME}', sheetName)); //mod

        const dataValues = sheet.getRange(rangeA1).getValues();
        const headers = dataValues[0];
        const data = dataValues.slice(1).map(row => {
            let obj = {};
            headers.forEach((header, i) => {
                 obj[header] = (metrics.includes(header)) ? (parseFloat(row[i]) || 0) : row[i];
            });
            return obj;
        });

        const overallMetrics = {};
        metrics.forEach(metric => {
            const totalSum = data.reduce((sum, row) => sum + (row[metric] || 0), 0);
            const totalCount = data.length;
            overallMetrics[metric] = {
                sum: totalSum,
                average: totalCount > 0 ? totalSum / totalCount : 0
            };
        });

        const groupedData = {};
        data.forEach(row => {
            const key = dimensions.map(d => row[d]).join(' | ');
            if (!groupedData[key]) {
                groupedData[key] = {};
                dimensions.forEach(d => { groupedData[key][d] = row[d]; });
                metrics.forEach(m => { 
                    groupedData[key][m] = 0;
                    groupedData[key][`__count_${m}`] = 0;
                });
            }
            metrics.forEach(m => {
                groupedData[key][m] += row[m];
                groupedData[key][`__count_${m}`]++;
            });
        });
        
        let finalData = Object.values(groupedData);
        
        finalData.forEach(row => {
          metrics.forEach(metric => {
            const isRateField = /率|%|百分比|Rate|Percentage/i.test(metric);
            if (isRateField && row[`__count_${metric}`] > 0) {
              row[metric] = row[metric] / row[`__count_${metric}`];
            }
            delete row[`__count_${metric}`];
          });
        });
        
        return {
            success: true,
            headers: [...dimensions, ...metrics],
            dimensions: dimensions,
            metrics: metrics,
            data: finalData,
            overallMetrics: overallMetrics
        };

    } catch (e) {
        Logger.log(`Dynamic analysis failed: ${e.stack}`);
        return { success: false, error: e.message };
    }
}

const REPORT_COLORS = ['#4285F4', '#DB4437', '#F4B400', '#0F9D58', '#AB47BC', '#00ACC1', '#FF7043', '#7E57C2', '#5C6BC0', '#26A69A'];

/**
 * Returns the standard color palette for reports.
 * @returns {string[]} An array of hex color codes.
 */
function getReportColors() {
  return REPORT_COLORS;
}

/**
 * Exports the selected report data based on the chosen format.
 * @param {object} exportOptions The export settings from the UI.
 * @param {object} reportData The full analysis result data.
 * @returns {{success: boolean, message?: string, sheetName?: string, url?: string, error?: string}} Result object.
 */
function exportReport(exportOptions, reportData) {
  const T = getTranslations(); //mod
  try {
    const { format } = exportOptions;
    switch (format) {
      case 'sheet':
        return exportToSheet(exportOptions, reportData);
      case 'doc':
        return exportToDoc(exportOptions, reportData);
      case 'pdf':
        return exportToPdf(exportOptions, reportData);
      default:
        throw new Error(T.reportUnsupportedFormat); //mod
    }
  } catch(e) {
    Logger.log(`Report export failed: ${e.stack}`);
    return { success: false, error: e.message };
  }
}

/**
 * Helper function to export data to a new Google Sheet.
 */
function exportToSheet(exportOptions, reportData) {
    const T = getTranslations(); //mod
    const { selections } = exportOptions;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const prefix = T.reportExportSheetNamePrefix || 'MasterDataAnalyzer';
    const sheetName = prefix + new Date().toLocaleString('sv-se').replace(/ /g, '_').replace(/:/g, '-');
    const newSheet = ss.insertSheet(sheetName);
    let currentRow = 1;

    const formatValue = (value, metricName) => {
        const isRateField = /率|%|百分比|Rate|Percentage/i.test(metricName);
        if (isRateField && typeof value === 'number') return value;
        return value;
    };
    
    const tabOrder = [T.overviewTab, ...reportData.dimensions.map(d => T.analysisTab.replace('{DIMENSION}', d)), T.rawDataTab];

    for (const tabName of tabOrder) {
        if (selections[tabName] && selections[tabName].length > 0) {
            newSheet.getRange(currentRow, 1).setValue(tabName).setFontWeight('bold').setFontSize(14);
            currentRow += 2;

            let sectionStartRow = currentRow;
            let dataWriteRow = currentRow;
            
            const kpis = selections[tabName].filter(id => id.startsWith('kpi-'));
            const charts = selections[tabName].filter(id => id.startsWith('chart-'));
            const tables = selections[tabName].filter(id => id.startsWith('table-'));

            if (kpis.length > 0) {
              kpis.forEach(cardId => {
                  const metric = cardId.replace('kpi-', '');
                  const isRateField = /率|%|百分比|Rate|Percentage/i.test(metric);
                  const displayValue = isRateField ? reportData.overallMetrics[metric].average : reportData.overallMetrics[metric].sum;
                  const cardTitle = (T.locale === 'zh_TW' && !metric.startsWith('總')) ? `總 ${metric}` : metric; //mod
                  const range = newSheet.getRange(dataWriteRow, 1, 1, 2);
                  range.setValues([[cardTitle, displayValue]]);
                  range.getCell(1, 1).setFontWeight('bold');
                  if (isRateField) range.getCell(1, 2).setNumberFormat('0.0%');
                  else range.getCell(1, 2).setNumberFormat('#,##0');
                  dataWriteRow++;
              });
            }

            let chartDataRanges = {};

            if (charts.length > 0) {
              const uniqueChartData = {};
              charts.forEach(cardId => {
                  let dimension, metric;
                   if (cardId.startsWith('chart-overview')) {
                        dimension = reportData.dimensions[0];
                        metric = reportData.metrics[0];
                   } else {
                        const idString = cardId.replace(/^chart-/, '');
                        const lastHyphenIndex = idString.lastIndexOf('-');
                        if (lastHyphenIndex !== -1) {
                            dimension = idString.substring(0, lastHyphenIndex).replace(/_/g, ' ');
                            metric = idString.substring(lastHyphenIndex + 1).replace(/_/g, ' ');
                        } else {
                            return;
                        }
                   }
                    const dataKey = `${dimension}|${metric}`;
                    if (!uniqueChartData[dataKey]) {
                        uniqueChartData[dataKey] = {dimension, metric};
                    }
              });

              for (const key in uniqueChartData) {
                  const {dimension, metric} = uniqueChartData[key];
                  const chartHeaders = [dimension, metric];
                  const isRateField = /率|%|百分比|Rate|Percentage/i.test(metric);
                  const aggregatedData = {}; 

                  reportData.data.forEach(row => {
                      const groupKey = row[dimension];
                      if(groupKey === undefined || row[metric] === undefined) return;
                      if (!aggregatedData[groupKey]) {
                          aggregatedData[groupKey] = { sum: 0, count: 0 };
                      }
                      aggregatedData[groupKey].sum += row[metric];
                      aggregatedData[groupKey].count++;
                  });

                  const chartData = Object.keys(aggregatedData).map(k => {
                      const group = aggregatedData[k];
                      const finalValue = isRateField && group.count > 0
                          ? group.sum / group.count
                          : group.sum;
                      return [k, finalValue];
                  });
                  
                  chartData.sort((a, b) => b[1] - a[1]);

                  const dataToInsert = [chartHeaders, ...chartData];
                  const dataRange = newSheet.getRange(dataWriteRow, 1, dataToInsert.length, chartHeaders.length);
                  if(isRateField){
                    dataRange.offset(1, 1, dataRange.getNumRows() - 1, 1).setNumberFormat('0.0%');
                  }
                  dataRange.setValues(dataToInsert);
                  chartDataRanges[key] = dataRange;
                  dataWriteRow += dataToInsert.length + 1;
              }
            }
            
            let chartWriteCol = 3;
            let maxChartHeight = 0;

            charts.forEach(cardId => {
              let chartTitle, chartType, chartOptions = {}, dataRange, dimension, metric;
              
              if (cardId.startsWith('chart-overview')) {
                  dimension = reportData.dimensions[0];
                  metric = reportData.metrics[0];
                  dataRange = chartDataRanges[`${dimension}|${metric}`];
                  
                  if (cardId.includes('pie')) {
                      chartType = Charts.ChartType.PIE;
                      chartTitle = `${dimension} ${T.pieChart}`;
                      chartOptions = { pieHole: 0.4, colors: REPORT_COLORS };
                      if(cardId.includes('labeled')) {
                        chartTitle += T.valuesSuffix;
                        chartOptions.pieSliceText = 'value';
                        chartOptions.pieSliceTextStyle = { color: 'white' };
                      }
                  } else {
                      chartType = Charts.ChartType.BAR;
                      chartTitle = `${dimension} ${T.barChart}`;
                      chartOptions.colors = [REPORT_COLORS[0]];
                  }
              } else {
                  const idString = cardId.replace(/^chart-/, '');
                  const lastHyphenIndex = idString.lastIndexOf('-');
                   if (lastHyphenIndex !== -1) {
                      dimension = idString.substring(0, lastHyphenIndex).replace(/_/g, ' ');
                      metric = idString.substring(lastHyphenIndex + 1).replace(/_/g, ' ');
                   } else {
                      return;
                   }
                  dataRange = chartDataRanges[`${dimension}|${metric}`];
                  chartType = Charts.ChartType.BAR;
                  chartTitle = `${metric} by ${dimension}`;
                  chartOptions.colors = [REPORT_COLORS[0]];
              }

              if (dataRange) {
                let chartBuilder = newSheet.newChart()
                    .addRange(dataRange)
                    .setChartType(chartType)
                    .setOption('title', chartTitle)
                    .setPosition(sectionStartRow, chartWriteCol, 0, 0);

                if (Object.keys(chartOptions).length > 0) {
                    for (const option in chartOptions) {
                        chartBuilder.setOption(option, chartOptions[option]);
                    }
                }

                // Specific options for bar charts that are not part of the generic options
                if (chartType === Charts.ChartType.BAR) {
                    const isRateField = /率|%|百分比|Rate|Percentage/i.test(metric);
                    const seriesOptions = { 
                        0: { 
                            annotations: { 
                                textStyle: { color: 'white', fontSize: 10 }
                            },
                            dataLabel: 'value'
                        }
                    };
                     if(isRateField) {
                        seriesOptions[0].format = '0.0%';
                     }
                    chartBuilder.setOption('series', seriesOptions);
                    chartBuilder.setOption('legend', { position: 'none' });
                }

                newSheet.insertChart(chartBuilder.build());
                chartWriteCol += 6;
                maxChartHeight = Math.max(maxChartHeight, 15);
              }
            });
            
            if (tables.length > 0) {
              tables.forEach(cardId => {
                if (cardId === 'table-raw' && tabName === T.rawDataTab) {
                  const headers = reportData.headers;
                  const data = reportData.data.map(row => headers.map(header => {
                      const val = row[header];
                      return val === null || val === undefined ? '' : val;
                  }));
                  const tableData = [headers, ...data];
                  if (tableData.length > 0) {
                    const tableRange = newSheet.getRange(dataWriteRow, 1, tableData.length, headers.length);
                    tableRange.setValues(tableData);
                    newSheet.getRange(dataWriteRow, 1, 1, headers.length).setFontWeight('bold');
                    dataWriteRow += tableData.length + 1;
                  }
                }
              });
            }
            
            currentRow = Math.max(dataWriteRow, sectionStartRow + maxChartHeight) + 6;
        }
    }
    
    newSheet.autoResizeColumns(1, 2);
    return { success: true, message: T.reportExportSuccess, sheetName: sheetName }; //mod
}



/**
 * Helper function to export data to a new Google Doc.
 */
function exportToDoc(exportOptions, reportData) {
  const T = getTranslations(); //mod
  const { selections, chartImages } = exportOptions;
  const prefix = T.reportExportDocNamePrefix || 'MasterDataAnalyzer';
  const doc = DocumentApp.create(prefix + new Date().toLocaleString('sv-se').replace(/ /g, '_').replace(/:/g, '-'));
  const body = doc.getBody();

  // Define styles
  const titleStyle = {};
  titleStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  titleStyle[DocumentApp.Attribute.FONT_SIZE] = 18;
  titleStyle[DocumentApp.Attribute.BOLD] = true;

  const heading1Style = {};
  heading1Style[DocumentApp.Attribute.FONT_SIZE] = 14;
  heading1Style[DocumentApp.Attribute.BOLD] = true;
  
  const tableHeaderStyle = {};
  tableHeaderStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#F2F2F2';
  tableHeaderStyle[DocumentApp.Attribute.BOLD] = true;

  // --- Start building document ---
  body.appendParagraph(T.reportAnalysisReportTitle).setAttributes(titleStyle); //mod
  body.appendParagraph(''); // Spacer
  
  const tabOrder = [T.overviewTab, ...reportData.dimensions.map(d => T.analysisTab.replace('{DIMENSION}', d)), T.rawDataTab];

  for (const tabName of tabOrder) {
    if (selections[tabName] && selections[tabName].length > 0) {
      body.appendParagraph(tabName).setAttributes(heading1Style);
      body.appendHorizontalRule();
      
      const kpis = [];
      const charts = [];
      const tables = [];

      selections[tabName].forEach(cardId => {
        if(cardId.startsWith('kpi-')) kpis.push(cardId);
        else if (cardId.startsWith('chart-')) charts.push(cardId);
        else if (cardId.startsWith('table-')) tables.push(cardId);
      });

      // Render KPIs in a table
      if (kpis.length > 0) {
        const kpiTableCells = [];
        kpis.forEach(cardId => {
          const metric = cardId.replace('kpi-', '');
          const isRateField = /率|%|百分比|Rate|Percentage/i.test(metric);
          const displayValue = isRateField ? (reportData.overallMetrics[metric].average * 100).toFixed(1) + '%' : reportData.overallMetrics[metric].sum.toLocaleString();
          const cardTitle = (T.locale === 'zh_TW' && !metric.startsWith('總')) ? `總 ${metric}` : metric; //mod
          kpiTableCells.push([cardTitle, displayValue]);
        });
        const kpiTable = body.appendTable(kpiTableCells);
        for(let i = 0; i < kpiTable.getNumRows(); i++){
            kpiTable.getRow(i).getCell(0).editAsText().setBold(true);
        }
      }

      // Render Charts
      charts.forEach((cardId, index) => {
        if (chartImages[cardId]) {
            let chartTitle = '';
            const dimensionName = reportData.dimensions[0];
            if (cardId === 'chart-overview-pie') {
                chartTitle = T.pieChartLegendLabel.replace('{DIMENSION}', dimensionName);
            } else if (cardId === 'chart-overview-pie-labeled') {
                chartTitle = T.pieChartValuesLabel.replace('{DIMENSION}', dimensionName);
            } else if (cardId === 'chart-overview-bar') {
                chartTitle = T.barChartLabel.replace('{DIMENSION}', dimensionName);
            } else {
              const parts = cardId.replace('chart-', '').split('_');
              const metric = parts.pop().replace(/_/g, ' ');
              const dimension = parts.join('_').replace(/_/g, ' ');
              chartTitle = T.chartByTitle.replace('{METRIC}', metric).replace('{DIMENSION}', dimension);
            }

            body.appendParagraph(chartTitle).setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
            const imageDataString = chartImages[cardId].split(',')[1];
            const imageBlob = Utilities.newBlob(Utilities.base64Decode(imageDataString), 'image/png', `${cardId}.png`);
            
            const p = body.appendParagraph('');
            const image = p.appendInlineImage(imageBlob);
            
            const newWidth = 500;
            const imageWidth = image.getWidth();
            if (imageWidth > newWidth) {
                const aspect = image.getHeight() / imageWidth;
                image.setWidth(newWidth).setHeight(newWidth * aspect);
            }

            p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

            // Add separator if there are more charts to follow in the same section
            if (index < charts.length - 1) {
                body.appendHorizontalRule();
                body.appendParagraph('');
            }
        }
      });
      
      // Render Tables
      tables.forEach(cardId => {
          if (cardId === 'table-raw') {
              body.appendParagraph('原始數據').setBold(true);
              const headers = reportData.headers;
              const data = reportData.data.map(row => headers.map(header => row[header] === null || row[header] === undefined ? '' : String(row[header])));
              const table = body.appendTable([headers, ...data]);
              const headerRow = table.getRow(0);
              for(let i = 0; i < headerRow.getNumCells(); i++){
                   headerRow.getCell(i).setAttributes(tableHeaderStyle);
              }
          }
      });

      body.appendParagraph(''); // Add space
    }
  }

  doc.saveAndClose();
  return { success: true, url: doc.getUrl() };
}

/**
 * Helper function to export data to a PDF file in Google Drive.
 */
function exportToPdf(exportOptions, reportData) {
  const docResult = exportToDoc(exportOptions, reportData);
  if (docResult.success) {
    const doc = DocumentApp.openByUrl(docResult.url);
    const pdfBlob = doc.getAs('application/pdf').setName(doc.getName() + '.pdf');
    const pdfFile = DriveApp.createFile(pdfBlob);
    
    // Clean up the temporary Google Doc
    DriveApp.getFileById(doc.getId()).setTrashed(true);
    
    return { success: true, url: pdfFile.getUrl() };
  } else {
    return docResult; // Return the error from the Doc creation
  }
}
