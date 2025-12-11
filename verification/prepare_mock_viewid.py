
import json
import re

# Read files
with open("MasterDataTranslation.js", "r") as f:
    translation_js = f.read()

with open("SettingPageReport.html", "r") as f:
    html_content = f.read()

# Extract Translations (Simplified Mock)
T = {
    "settingsForSheetHint": "Current settings for sheet: {SHEET_NAME}",
    "loadingMessage": "Loading...",
    "loadSettingsFailure": "Failed to load settings",
    "fileSelected": "File selected: ",
    "operationCancelled": "Operation cancelled.",
    "selectButton": "Select",
    "gapiLoading": "Status: Loading Google API...",
    "gapiLoaded": "Status: GAPI loaded. Loading Picker API...",
    "pickerApiLoaded": "Status: Picker API loaded. Fetching keys from server...",
    "keyFetchError": "Error: Failed to fetch keys:",
    "initFailed": "Initialization failed",
    "statusReady": "Status: Ready.",
    "pickerCreationError": "Error: ",
    "dragAndActionHelp": "Click the empty space...",
    "dataSourceTitle": "Data Source",
    "sourceSpreadsheetUrlLabel": "Select Source Spreadsheet",
    "noFileSelected": "No file selected",
    "sourceDataSheetNameLabel": "Source Data Sheet Name",
    "sourceDataRangeLabel": "Data Range",
    "sourceDataRangePlaceholder": "e.g. A1:B2",
    "fieldMappingTitle": "Field Mapping",
    "addAnalysisFieldButton": "+ Add Analysis Field",
    "analysisResultsTitle": "Analysis Results",
    "closeButton": "Close",
    "saveButton": "Save Settings",
    "exportReportButton": "Export Report",
    "generateReportButton": "Generate Report",
    "okButton": "OK",
    "sheetSelectionTitle": "Select a Sheet",
    "exportSettingsTitle": "Export Settings",
    "exportFormatTitle": "Select Format",
    "exportContentTitle": "Select Content",
    "exportToSheetLabel": "Export to Sheet",
    "cancelButton": "Cancel",
    "confirmExportButton": "Confirm",
    "dimensionOption": "Dimension",
    "metricOption": "Metric",
    "metricWarningHint": "Warning",
    "expandHint": "Expand",
    "collapseHint": "Collapse",
    "minimizeHint": "Minimize",
    "unminimizeHint": "Unminimize",
    "pinHint": "Pin",
    "unpinHint": "Unpin",
    "reportExportSuccess": "Success",
    "pieChartTitle": "Pie Chart",
    "barChartTitle": "Bar Chart",
    "toggleValuesHint": "Toggle Values",
    "noDataForChart": "No Data",
    "distributionChartTitle": "Distribution",
    "chartByTitle": "Chart By",
    "kpiCardTitle": "KPI",
    "analysisTab": "Analysis",
    "overviewTab": "Overview",
    "rawDataTab": "Raw Data",
    "pieChartLegendLabel": "Legend",
    "pieChartValuesLabel": "Values",
    "barChartLabel": "Bar Chart"
}

# Replace scriptlets
def replace_t(match):
    key = match.group(1).strip()
    return T.get(key, f"[{key}]")

html_content = re.sub(r'<\?!= T\.(\w+) \?>', replace_t, html_content)
html_content = html_content.replace('<?!= JSON.stringify(T) ?>', json.dumps(T))

# Mock google.script.run and gapi
mock_script = """
<script>
    // Mock google object
    window.google = {
        script: {
            host: {
                close: () => console.log('Host close'),
                setWidth: (w) => console.log('Set width', w),
                setHeight: (h) => console.log('Set height', h),
                origin: 'https://mock-origin'
            },
            run: {
                withSuccessHandler: function(callback) {
                    const clone = Object.create(this);
                    clone._success = callback;
                    return clone;
                },
                withFailureHandler: function(callback) {
                    const clone = Object.create(this);
                    clone._failure = callback;
                    return clone;
                },
                report_getReportColors: function() {
                     const successCb = this._success;
                     setTimeout(() => { if(successCb) successCb(['#000']); }, 100);
                },
                report_getPickerKeys: function() {
                    const successCb = this._success;
                    setTimeout(() => { if(successCb) successCb({apiKey: 'mock', appId: 'mock', oauthToken: 'mock'}); }, 100);
                },
                report_getActiveSheetName: function() {
                    const successCb = this._success;
                    setTimeout(() => { if(successCb) successCb('Sheet1'); }, 100);
                },
                report_getReportSettings: function(sheetName) {
                    const successCb = this._success;
                    const settings = window.mockSettings || null;
                    setTimeout(() => { if(successCb) successCb(settings); }, 100);
                },
                report_validateReportInputs: function(id, name, range) {
                     const successCb = this._success;
                     setTimeout(() => { if(successCb) successCb({sourceUrlError:'', sheetError:'', rangeError:''}); }, 100);
                },
                report_getHeadersFromRange: function() {
                     const successCb = this._success;
                     setTimeout(() => { if(successCb) successCb(['Header1', 'Header2']); }, 100);
                },
                report_getSheetNames: function() {
                    const successCb = this._success;
                    setTimeout(() => { if(successCb) successCb(['Sheet1', 'Sheet2']); }, 100);
                }
            }
        },
        picker: {
            ViewId: { SPREADSHEETS: 'SPREADSHEETS', DOCS: 'DOCS' },
            Action: { PICKED: 'picked', CANCEL: 'cancel' },
            Feature: { NAV_HIDDEN: 'nav_hidden', MULTISELECT_ENABLED: 'multiselect_enabled' },
            DocsView: class {
                constructor(viewId) { this.viewId = viewId; }
                setIncludeFolders() { return this; }
            },
            PickerBuilder: class {
                setDeveloperKey() { return this; }
                setAppId() { return this; }
                setOAuthToken() { return this; }
                addView(view) {
                    this.view = view;
                    return this;
                }
                setCallback(cb) { this.callback = cb; return this; }
                setOrigin() { return this; }
                enableFeature() { return this; }
                disableFeature() { return this; }
                build() {
                    // Capture the viewId used to build this picker
                    window.lastPickerViewId = this.view && this.view.viewId;
                    return {
                        setVisible: (v) => {
                            console.log('Picker visible, ViewID:', window.lastPickerViewId);
                            window.mockPickerCallback = this.callback;
                        }
                    };
                }
            }
        },
        visualization: {
            arrayToDataTable: () => {},
            PieChart: class { draw() {} },
            BarChart: class { draw() {} },
            events: { addListener: () => {} }
        },
        charts: {
            load: (pkg, opts) => {},
            setOnLoadCallback: (cb) => cb()
        }
    };

    // Mock gapi
    window.gapi = {
        load: function(api, opts) {
            if (opts && opts.callback) opts.callback();
        }
    };

    // Helper to simulate picking a file
    window.simulateFilePick = function(id, name) {
        if (window.mockPickerCallback) {
            window.mockPickerCallback({
                action: 'picked',
                docs: [{id: id, name: name}]
            });
        }
    };
</script>
"""

# Inject mock script
html_content = html_content.replace('<script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>',
                                    '<script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>' + mock_script)

# Remove actual gapi load script
html_content = re.sub(r'<script[^>]*src="https://apis.google.com/js/api.js"[^>]*>.*?</script>', '', html_content, flags=re.DOTALL)

# Replace document.head.appendChild(script) with script.onload()
html_content = html_content.replace("document.head.appendChild(script);", "script.onload();")

with open("verification/mock_report_settings_viewid.html", "w") as f:
    f.write(html_content)
