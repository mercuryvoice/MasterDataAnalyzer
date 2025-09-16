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
// ===================================== SECTION 1: USER INTERFACE & TRANSLATIONS ==================
// =================================================================================================

const TRANSLATIONS = {
    en: {
        // --- Main Menu ---
        mainMenuTitle: 'MasterDataAnalyzer',
        // --- Sub Menus ---
        importMenuTitle: 'Data Import Tool',
        validationMenuTitle: 'Data Validation Tool',
        manageMenuTitle: 'Data Management Tool',
        guideMenuTitle: 'Guides & Examples',
        // --- Items ---
        settingsItem: '⚙️ Open Settings',
        manageSettingsItem: '⚙️ Monitoring Management Settings',
        quickDeleteItem: '🗑️ Quick Delete Sheets',
        reportSettingsItem: '📊 Report Generation Settings',
        runImportItem: '▶️ Run Import (Sync)',
        stopImportItem: '⏹️ Stop Import',
        resetImportItem: '🔄 Clear Import Data & Progress',
        runCompareItem: '▶️ Run Data Comparison',
        compareSettingsItem: '⚙️ Data Comparison Settings',
        verifySettingsItem: '⚙️ Data Validation Settings',
        runMsModeItem: '▶️ Run Validation (MS Mode)',
        stopValidationItem: '⏹️ Stop Validation',
        verifySumsItem: '🔍 Verify Sums & Cumulative Values',
        cleanupItem: '🔄 Clear Validation Data & Progress',
        monitorMenuName: 'Data Change Monitoring',
        enableNotifyItem: '🟢 Enable Automatic Notifications',
        disableNotifyItem: '🔴 Disable Automatic Notifications',
        checkNowItem: '✉️ Check and Notify Now',
        privacyPolicyItem: 'Privacy Policy',
        // --- Guide Sub-menu Items (Placeholders) ---
        manufacturingGuide: 'Manufacturing Example',
        manufacturingProductionTitle: 'Manufacturing Production',
        businessGuide: 'Business & Sales Example',
        // hrGuide: 'Human Resources Example',
        startBusinessGuide: '▶️ Start Interactive Guide (Sales)',
        startManufacturingGuide: '▶️ Start Interactive Guide (Manufacturing)',
        // --- HTML UI Titles ---
        importSettingsTitle: 'Data Import Settings',
        compareSettingsTitle: 'Data Comparison Settings',
        verifySettingsTitle: 'Data Validation Settings',
        manageSettingsTitle: 'Data Monitoring Management Settings',
        quickDeleteTitle: 'Quick Delete Sheets',
        reportSettingsTitle: 'Report Generation Settings',
        sheetSelectionTitle: 'Select a Sheet',
        privacyPolicyTitle: 'Privacy Policy',
         // Privacy Policy Content
        privacyLastUpdated: 'Last Updated: September 2, 2025',
        privacyIntro: "Thank you for using MasterDataAnalyzer (hereinafter referred to as 'this add-on'). We are committed to protecting your privacy and ensuring you understand how your data is handled. All operations of this add-on are completed within your Google account. We do not collect, store, or share your personal information or document content with any third parties.",
        privacyDataCollectionTitle: 'Data Collection and Use',
        privacyDataCollectionP1: "This add-on is a tool built on Google Apps Script, designed to help you process data within Google Sheets.",
        privacyDataCollectionL1: '**No Personal Information Collected**: We do not ask for, collect, or store any personally identifiable information, such as your name, email address, or contact details.',
        privacyDataCollectionL2: '**Data Processing**: All data reading, processing, and writing operations occur directly within the Google Sheets, Google Docs, or Google Drive files you authorize. Data is not transmitted to our servers or any external services.',
        privacyDataCollectionL3: "**Settings Storage**: The settings you configure for this add-on (e.g., data sources, field mappings) are stored using Google Apps Script's built-in `PropertiesService` as properties linked to your Google document. These settings are only accessible to the add-on for its operation within your account; we cannot access them externally.",
        privacySecurityTitle: 'Security Statement and Permissions Explanation',
        privacySecurityP1: 'To provide its full functionality, this add-on will request your authorization for the following Google services upon installation. These permissions are used exclusively for the specific functions described and will never be used for any other purpose.',
        privacySecurityP2: 'The following is a detailed explanation based on the `oauthScopes` in your `appsscript.json` file:',
        privacyScopeSpreadsheetsTitle: '`https://www.googleapis.com/auth/spreadsheets`',
        privacyScopeSpreadsheetsPermission: 'Permission: View, edit, create, and delete your Google Sheets spreadsheets.',
        privacyScopeSpreadsheetsPurpose: "Purpose: This is the core permission for the add-on. We need this to:",
        privacyScopeSpreadsheetsUse1: 'Read data from your specified source spreadsheets.',
        privacyScopeSpreadsheetsUse2: 'Write processed data into your target spreadsheets.',
        privacyScopeSpreadsheetsUse3: 'Create new sheets for reports, sample data, etc.',
        privacyScopeSpreadsheetsUse4: 'Execute the "Quick Delete Sheets" feature.',
        privacyScopeDocumentsTitle: '`https://www.googleapis.com/auth/documents`',
        privacyScopeDocumentsPermission: 'Permission: View, edit, create, and delete your Google Docs documents.',
        privacyScopeDocumentsPurpose: 'Purpose: This permission is used solely for the "Export Report" feature. When you choose to export analysis results to a Google Doc, the add-on creates a new document and writes the report content into it.',
        privacyScopeDriveTitle: '`https://www.googleapis.com/auth/drive`',
        privacyScopeDrivePermission: 'Permission: View, edit, create, and delete specific files in your Google Drive.',
        privacyScopeDrivePurpose: 'Purpose: This permission primarily supports the "Export Report" feature:',
        privacyScopeDriveUse1: 'When exporting to Google Docs, this permission is needed to create the document in your Drive.',
        privacyScopeDriveUse2: 'When exporting as a PDF, the add-on first creates a Google Doc, converts it to a PDF file saved in your Drive, and may delete the temporary document.',
        privacyScopeGmailTitle: '`https://www.googleapis.com/auth/gmail.send`',
        privacyScopeGmailPermission: 'Permission: Allow this add-on to send email on your behalf. **(Note: This add-on cannot read any of your emails)**',
        privacyScopeGmailPurpose: 'Purpose: This permission is used only for the "Data Change Monitoring" feature. When a change occurs in a cell range you have configured for monitoring, the add-on will automatically send a notification email to the address you have specified.',
        privacyScopeUITitle: '`https://www.googleapis.com/auth/script.container.ui`',
        privacyScopeUIPermission: 'Permission: Display user interfaces in Google Sheets.',
        privacyScopeUIPurpose: 'Purpose: This permission is required for the add-on to display all user interfaces, such as settings windows, sidebars, dialog boxes, and custom menus within your spreadsheet.',
        privacyScopeScriptAppTitle: '`https://www.googleapis.com/auth/script.scriptapp`',
        privacyScopeScriptAppPermission: 'Permission: Allow Apps Script to create and manage script triggers.',
        privacyScopeScriptAppPurpose: 'Purpose: This permission is used to create an `onOpen` trigger to automatically generate the "MasterDataAnalyzer" menu when you open a spreadsheet. It is also used by the "Data Change Monitoring" feature to create triggers that detect worksheet changes.',
        privacyScopeExternalRequestTitle: '`https://www.googleapis.com/auth/script.external_request`',
        privacyScopeExternalRequestPermission: 'Permission: Allow Apps Script to connect to external services.',
        privacyScopeExternalRequestPurpose: 'Purpose: In the current version, this add-on **does not** actively send requests or transmit your data to any non-Google external servers. This permission is included to allow for potential future feature enhancements (e.g., connecting to public API services) but is not currently in use.',
        privacyScopeStorageTitle: '`https://www.googleapis.com/auth/script.storage`',
        privacyScopeStoragePermission: 'Permission: Allow Apps Script to store a small amount of data.',
        privacyScopeStoragePurpose: 'Purpose: As described in "Settings Storage" above, we use this permission to save your configurations for various features, so you do not need to re-enter them each time you use the add-on.',
        privacyChangesTitle: 'Policy Changes',
        privacyChangesP1: 'We may update this Privacy Policy from time to time. Any changes will be posted on this page, and we encourage you to review it periodically.',
        privacyContactTitle: 'Contact Us',
        privacyContactP1: 'If you have any questions about this Privacy Policy, please contact us at [tsengmercury@gmail.com].',
        // --- NEW: Dashboard Generator UI ---
        sectionFieldMapping: 'Field Mapping',
        regionColumnLabel: 'Region Field',
        productColumnLabel: 'Product/Item Field',
        salesColumnLabel: 'Sales Value Field',
        generateReportButton: 'Generate Report',
        step1Title: 'Step 1: Prepare Raw Data',
        step1Description: 'This is a typical sales ledger, containing information such as date, region, product, and sales amount.',
        // sourceSpreadsheetUrlLabel: 'Source Spreadsheet URL',
        sourceDataSheetNameLabel: 'Source Data Sheet Name',
        sourceDataRangeLabel: 'Data Range (including header)',
        sourceDataRangePlaceholder: 'e.g., A1:G100',
        generateDashboardButton: 'Generate Dashboard with One Click',
        generatingDashboard: 'Generating Dashboard...',
        step2Title: 'Step 2: View Generated Results',
        overviewTab: 'Overview',
        productAnalysisTab: 'Product Analysis',
        regionAnalysisTab: 'Region Analysis',
        rawDataTab: 'Raw Data',
        salesOverviewTitle: 'Sales Overview',
        totalSalesLabel: 'Total Sales',
        totalOrdersLabel: 'Total Orders',
        regionSalesDistributionTitle: 'Sales Distribution by Region',
        productSalesAnalysisTitle: 'Sales Analysis by Product',
        totalSalesAxisTitle: 'Total Sales',
        productAxisTitle: 'Product',
        regionSalesAnalysisTitle: 'Sales Analysis by Region',
        regionAxisTitle: 'Region',
        errorTitle: 'Error',
        requiredFieldsError: 'The selected range must contain the following headers: {HEADERS}. Please check your data range.',
        // Tutorial Steps
        businessGuideTutorialTitle: 'Business & Sales Example - Tutorial',
        manufacturingGuideTutorialTitle: 'Manufacturing Example - Tutorial',
        // Business Tutorial
        businessWelcomeTitle: "Welcome! (Business & Sales)",
        businessWelcomeInstruction: "Welcome to the interactive tutorial for the 'Business & Sales Example'.\n\nOur goal is to transform an incomplete sales record into a complete analysis report using the features of MasterDataAnalyzer.\n\nClick 'Next' to start our first task!",
        businessTask1Step1Title: "Task 1: Data Import (1/6)",
        businessTask1Step1Instruction: "First, please ensure you have activated the <code>{SHEET_NAME}</code> sheet.\n\nOur goal is to filter and import sales records from <code>{SOURCE_SHEET_NAME}</code> into the current dashboard.",
        businessTask1Step2Title: "Task 1: Data Import (2/6)",
        businessTask1Step2Instruction: "Please click on <code>MasterDataAnalyzer > Data Import Tool > ⚙️ Data Import Settings</code> in the top menu.",
        businessTask1Step3Title: "Task 1: Data Import (3/6)",
        businessTask1Step3Instruction: "In the settings window, please configure the following:\n1. **Source Spreadsheet URL**: (Paste the URL of the current file)\n2. **Source Data Sheet Name**: Select <code>{SOURCE_SHEET_NAME}</code>\n3. **Target Data Sheet Name**: Should be auto-filled with <code>{TARGET_SHEET_NAME}</code>",
        businessTask1Step4Title: "Task 1: Data Import (4/6)",
        businessTask1Step4Instruction: "Next, set the data ranges:\n1. **Header Start Row for Data Import**: <code>1</code>\n2. **Data Import Start Row**: <code>2</code>\n3. **Source Data Import Range**: <code>A2:E9</code>",
        businessTask1Step5Title: "Task 1: Data Import (5/6)",
        businessTask1Step5Instruction: "Suppose we only want to analyze sales for the \"{PRODUCT_NAME}\". Please add a filter condition:\n1. **Header Name**: <code>{HEADER_NAME}</code>\n2. **Keywords**: <code>{PRODUCT_NAME}</code>",
        businessTask1Step6Title: "Task 1: Data Import (6/6)",
        businessTask1Step6Instruction: "Great! All settings are complete.\n\nPlease click 'Save Settings', close the window, and then run <code>MasterDataAnalyzer > Data Import Tool > ▶️ Run Import</code> from the menu.",
        businessCheckpointTitle: "Task 1 Complete!",
        businessCheckpointInstruction: "Congratulations! You have successfully imported the raw data into the dashboard.\n\nWould you like to continue learning about the next core feature, 'Data Comparison'?",
        businessTask2Step1Title: "Task 2: Enrich Customer Data (1/3)",
        businessTask2Step1Instruction: "Excellent! The dashboard now has raw data but lacks detailed customer information.\n\nNext, we'll use the 'Data Comparison' feature to look up and fill in customer data from the \"[Source] Customer Master\".\n\nOpen <code>Data Import Tool > ⚙️ Data Comparison Settings</code> and begin configuring Task 2:",
        businessTask2Step2Title: "Task 2: Enrich Customer Data (2/3)",
        businessTask2Step2Instruction: "Please apply the following settings:\n1. **Source Sheets**: Choose \"Source Data Sheet Name\" as <code>{SOURCE_SHEET_NAME}</code> and \"Target Data Sheet Name\" as <code>{TARGET_SHEET_NAME}</code>.\n2. **Ranges**: Set \"Target Data Start Row\" to <code>2</code> and \"Source Data Compare Range\" to <code>A2:D6</code>.\n3. **Field Mapping**: Set up the fields in the following order:\n  - **Target Lookup Column**: <code>B</code> (Customer ID)\n  - **Source Compare Column**: <code>A</code> (Customer ID)\n  - **Source Return Column**: <code>B</code> (Customer Name)\n  - **Target Write Column**: <code>C</code> (Customer Name)",
        businessTask2Step3Title: "Task 2: Enrich Customer Data (3/3)",
        businessTask2Step3Instruction: "After saving the settings, run <code>MasterDataAnalyzer > Data Import Tool > ▶️ Run Data Comparison</code>.\n\nYou will see the customer names have been successfully imported. Repeat this process to also fill in the \"Region\" and \"Salesperson\" to complete the dashboard.",
        businessTask3Step1Title: "Task 3: Compare Sales Targets (1/2)",
        businessTask3Step1Instruction: "The dashboard data is becoming more complete! For the final step, let's compare the sales targets for each product.\n\nPlease open the 'Data Comparison Settings' again.",
        businessTask3Step2Title: "Task 3: Compare Sales Targets (2/2)",
        businessTask3Step2Instruction: "This time, use \"Product Name\" as the lookup key and select \"[Source] Product Targets\" as the source sheet.<br>Set the \"Source Data Compare Range\" to <code>A2:B5</code>, and configure the field mapping as follows:<br>  - **Target Lookup Column**: <code>F</code> (Product Name)<br>  - **Source Compare Column**: <code>A</code> (Product Name)<br>  - **Source Return Column**: <code>B</code> (Monthly Target)<br>  - **Target Write Column**: <code>J</code> (Monthly Target)<br><br>Then, run the 'Data Comparison' again.",
        businessFinalStepTitle: "Congratulations, Analysis Complete!",
        businessFinalStepInstruction: "All data is now in place! You now have a clean, complete dataset ready for analysis.\n\nNext, you can manually enter or click the button below to insert formulas in the corresponding columns to complete the final calculations:\n- **Total Sales**: <code>=H2*I2</code> (Quantity * Unit Price)\n- **Achievement Rate**: <code>=G2/J2</code> (Total Sales / Monthly Target)",
        // Manufacturing Tutorial
        mfgWelcomeTitle: "Welcome! (Manufacturing)",
        mfgWelcomeInstruction: "Welcome to the interactive tutorial for the 'Manufacturing Example'.\n\nThis guide will walk you through using MasterDataAnalyzer's 'Data Import (Array Mode)' and 'Data Validation' features to convert a complex bill of materials into a standardized format and compare it against master data.",
        mfgTask1Step1Title: "Task 1: Data Import (1/6)",
        mfgTask1Step1Instruction: "First, please ensure you have activated the <code>{SHEET_NAME}</code> sheet.\n\nOur goal is to convert and import unstructured data from <code>{SOURCE_SHEET_NAME}</code> into the current target sheet.",
        mfgTask1Step2Title: "Task 1: Import Settings (2/6)",
        mfgTask1Step2Instruction: "Please click on <code>MasterDataAnalyzer > Data Import Tool > ⚙️ Data Import Settings</code> in the top menu.",
        mfgTask1Step3Title: "Task 1: Import Settings (3/6)",
        mfgTask1Step3Instruction: "In the settings window, please configure the following:\n1. **Source Spreadsheet URL**: (Paste the URL of the current file)\n2. **Source Data Sheet Name**: Select <code>{SOURCE_SHEET_NAME}</code>\n3. **Target Data Sheet Name**: Should be auto-filled with <code>{TARGET_SHEET_NAME}</code>",
        mfgTask1Step4Title: "Task 1: Import Settings (4/6)",
        mfgTask1Step4Instruction: "Next is the key step, enabling 'Data Array Comparison' mode:\n1. **Header Start Row for Data Import**: <code>3</code>\n2. **Data Import Start Row**: <code>4</code>\n3. **Source Data Import Range**: <code>A9:E14</code>\n4. **Header Start Row for Other Blocks**: <code>F2:I2</code>\n5. **Data Range within Header of Other Blocks**: <code>F9:I14</code>",
        mfgTask1Step5Title: "Task 1: Filter Settings (5/6)",
        mfgTask1Step5Instruction: "In the Filter & Validation section, you can select which headers and keywords to import.\n1. In \"Source Header Import Range Settings\", enter <code>A8:E8</code>\n2. For \"Keyword Filter Conditions (AND)\", select <code>Owner</code>\n3. Click the \"Select\" box to the right of the keyword and choose <code>Mark</code>, <code>Linda</code>, and <code>Mary</code>.",
        mfgTask1Step6Title: "Task 1: Execute Import (6/6)",
        mfgTask1Step6Instruction: "Once the import settings are configured, don't forget to click \"Save Settings\". After saving, proceed to run <code>▶️ Run Import</code>.\n\nYou should now see imported data in columns A - G, indicating that the data has been successfully filtered and imported.",
        mfgCheckpointTitle: "Task 1 Complete!",
        mfgCheckpointInstruction: "Congratulations! You have successfully transformed and imported the array data.\n\nNext, would you like to learn how to use the 'Data Validation' feature to compare the imported data with master data?",
        mfgTask2Step1Title: "Task 2: Validation Settings (1/4)",
        mfgTask2Step1Instruction: "Great! Now let's verify the accuracy of the imported data.\n\nPlease open <code>MasterDataAnalyzer > Data Validation Tool > ⚙️ Data Validation Settings</code>.",
        mfgTask2Step2Title: "Task 2: Data Range Settings (2/4)",
        mfgTask2Step2Instruction: "In the [Data Validation] settings window, please configure the following:<br>1. **Source Spreadsheet URL**: (Paste the URL of the current file)<br>2. **Source Data Sheet Name**: Select <code>{SOURCE_SHEET_NAME}</code><br>3. Please fill in the data range settings in the following order:<br>   - Data Import Start Row after Validation: <code>4</code><br>   - Target Data Header Start Row: <code>3</code><br>   - Source Data Header Start Row: <code>1</code><br><br>Proceed to the next step to begin setting up the 'Field Validation Conditions' and 'Validation Result Outputs'.",
        mfgTask2Step3Title: "Task 2: Field Validation Settings (3/4)",
        mfgTask2Step3Instruction: "In the Field and Validation Conditions section, the first step is to map the 'Target Column' and 'Source Column'.<br>We recommend using the <b>Auto-map Validation Fields</b> feature first, which will <b>automatically match and recommend</b> suitable headers for you.<br>For this task, the validation conditions are as follows:<br>Target Column - Source Column<br><code>B</code> - <code>B</code><br><code>C</code> - <code>C</code><br><code>D</code> - <code>D</code><br><code>E</code> - <code>E</code><br><code>G</code> - <code>A</code>",
        mfgTask2Step4Title: "Task 2: Validation Output Settings (4/4)",
        mfgTask2Step4Instruction: "Next, configure the data columns to be returned from the source upon successful validation.<br>Similarly, you can use the <b>Auto-map Output Fields</b> feature to speed up the setup.<br>The output field conditions are as follows:<br>Target Column - Source Column<br><code>H</code> - <code>F</code><br><code>I</code> - <code>G</code><br><code>J</code> - <code>H</code><br><br>Finally, we will set the \"Mismatch Info Output Column\" to <code>K</code>, so the script can write error messages there.",
        mfgFinalStepTitle: "All Settings Are Ready!",
        mfgFinalStepInstruction: "Please remember to click [Save Settings], then start by running <code>Data Validation Tool > ▶️ Run Validation (MS Mode)</code>!<br><br>After execution, please check column K in the <code>{SHEET_NAME}</code> sheet. You will see that the script has automatically flagged all mismatched items and their reasons.",
        sectionSourceAndTarget: 'Source & Target',
        sectionDataRanges: 'Data Ranges',
        sectionFilterAndValidate: 'Filter & Validation',
        sectionValidationConditions: 'Field Validation Conditions',
        sectionValidationOutputs: 'Validation Result Outputs',
        sectionMismatchOutput: 'Mismatch Information Output',
        sectionDataManagement: 'Data Management Settings',
        sectionRangesAndConditions: 'Data Ranges & Conditions',
        sectionFieldMapping: 'Field Mapping',
        closeButton: 'Close',
        saveButton: 'Save Settings',
        selectButton: 'Select',
        okButton: 'OK',
        defaultTemplateButton: 'Default Template',
        removeAllButton: 'Remove All Fields',
        checkButton: 'Check',
        pinWindowTooltip: 'Pin Window',
        unpinWindowTooltip: 'Unpin Window',
        expandWindowTooltip: 'Expand Window',
        collapseWindowTooltip: 'Collapse Window',
        dragAndActionHelp: 'Click the empty space in the title bar above to drag.\nUse the icons on the right to expand or pin.',
        // sourceSpreadsheetUrlLabel: 'Source Spreadsheet URL',
        sourceDataSheetNameLabel: 'Source Data Sheet Name',
        // Import Settings UI
        currentTargetSheetNameLabel: "Target Data Sheet Name",
        settingsForSheetHint: 'Current settings for sheet: {SHEET_NAME}',
        importHeaderStartRowLabel: 'Header Start Row for Data Import',
        importDataStartRowLabel: 'Data Import Start Row',
        sourceDataRangeLabel: 'Source Data Import Range',
        validationHeaderStartRowLabel: 'Header Start Row for Other Blocks',
        validationMatrixRangeLabel: 'Data Range within Header of Other Blocks',
        headerImportFilterLabel: 'Source Header Import Range Settings',
        keywordFiltersLabel: 'Keyword Filter Conditions (AND)',
        addFilterConditionLabel: '+ Add Filter Condition',
        headerPlaceholder: 'Header Name',
        keywordsPlaceholder: 'Keywords (comma-separated)',
        dataStartRowLabel: 'Start Row for Data Validation Import',
        verifySourceUrlLabel: 'Source Spreadsheet URL',
        verifySourceSheetNameLabel: 'Source Data Sheet Name',
        mismatchColumnLabel: 'Mismatch Info Output Column',
        targetHeaderRowLabel: 'Target Data Header Start Row',
        sourceHeaderRowLabel: 'Source Data Header Start Row',
        targetColumnLabel: 'Target Column',
        sourceColumnLabel: 'Source Column',
        primaryValidationLabel: 'Major Validation',
        selectAllLabel: 'Select/Deselect All',
        fieldMappingJSONLabel: 'Field Mapping & Validation Settings',
        checkEmptyValuesButton: 'Check Source for Empty Values',
        monitorRangeLabel: 'Automatic notification for document content changes',
        monitorEmailLabel: 'Notification Recipient Email',
        monitorSubjectLabel: 'Notification Subject',
        monitorBodyLabel: 'Notification Body',
        addValidationMappingButton: '+ Add Validation Field',
        autoMapValidationButton: 'Auto-map Validation Fields',
        addOutputMappingButton: '+ Add Output Field',
        autoMapOutputButton: 'Auto-map Output Fields',
        targetStartRowLabel: 'Target Data Start Row',
        targetStartRowHelp: 'Start reading target data and writing comparison results from this row number.',
        sourceCompareRangeLabel: 'Source Data Compare Range',
        sourceCompareRangeHelp: 'The range in the source sheet that includes the "Compare Column" and "Return Column".',
        targetLookupColLabel: 'Target Lookup Column',
        targetLookupColHelp: 'The column in the target sheet used to look up values.',
        targetWriteColLabel: 'Target Write Column',
        targetWriteColHelp: 'Write the comparison results to this column.',
        sourceLookupColLabel: 'Source Compare Column',
        sourceLookupColHelp: 'The column within the "Compare Range" to be compared against.',
        sourceReturnColLabel: 'Source Return Column',
        sourceReturnColHelp: 'The column within the "Compare Range" to return upon a match.',
        sourceUrlHelp: 'Paste the full URL of the source Google Sheet.',
        importHeaderStartRowHelp: "The row containing the classification names for your data columns is the header row.",
        importDataStartRowHelp: 'The row number where data import should begin in the target sheet.',
        sourceDataRangeHelp: "The row and column range of the data after the header.",
        validationHeaderStartRowHelp: "If filled, enables 'Data Array Comparison' mode. Compares import headers against this row to list statistical item data.",
        headerImportFilterHelp: 'Optional. Specify headers as a list (Header1,Header2) or a range (A1:D1). If blank, all columns from the "Source Data Import Range" will be used.',
        verifyStartRowHelp: "The data validation will start updating based on the set row number.",
        verifySourceUrlHelp: 'URL of the spreadsheet to use for data verification.',
        mismatchColumnHelp: 'Column letter (e.g., K) to output mismatch information.',
        targetHeaderRowHelp: 'The row number in the target sheet where the headers are located.',
        sourceHeaderRowHelp: 'The row number in the source sheet where the headers are located.',
        monitorRangeHelp: 'Specify the cell range (e.g., A2:E20) to monitor for changes.',
        monitorEmailHelp: 'Enter the email address to receive notifications.',
        monitorSubjectHelp: 'Set a custom subject for the notification email. Use {SHEET_NAME} as a placeholder.',
        monitorBodyHelp: 'Set a custom body. Use placeholders: {SHEET_NAME}, {RANGE}, {TIMESTAMP}, {CHANGES_COUNT}, {CHANGE_DETAILS}, {SHEET_URL}.',
        defaultSubjectTemplate: '[Change Notification] Sheet "{SHEET_NAME}" has been updated',
        defaultBodyTemplate: 'Hello,\n\nThe system has detected a change in the monitored sheet.\n\nSheet Name: {SHEET_NAME}\nMonitored Range: {RANGE}\nChange Time: {TIMESTAMP}\n\nChange Details ({CHANGES_COUNT} items):\n{CHANGE_DETAILS}\n\nPlease click the link below to view the latest content:\n{SHEET_URL}',
        savingMessage: 'Saving...',
        validatingMessage: 'Validating...',
        saveSuccess: 'Settings saved successfully!',
        saveFailure: 'Failed to save settings',
        autoMappingMessage: 'Auto-mapping...',
        autoMapSuccessBody: 'Found and created {COUNT} matching field(s). Please review and save.',
        autoMapNoMatchTitle: 'No Matches Found',
        autoMapNoMatchBody: 'Could not find any headers with matching names between the source and target sheets. Please check your settings.',
        checkingMessage: 'Checking...',
        checkSuccess: 'Validation passed. No empty or duplicate values found.',
        checkFailure: 'Field Check Notice:',
        importCancelled: 'Import cancelled by user.',
        settingsError: 'Settings Error',
        headerLessThanStartError: 'Header Start Row for Data Import must be less than Data Import Start Row.',
        preflightTitle: 'Settings Confirmation',
        preflightWarning: 'Warning:',
        preflightFilterWarning: '- A filter condition has a header but no keywords; it will be ignored.',
        preflightSuggestion: 'This might be unintended. It is recommended to review your settings.',
        preflightConfirmation: 'Are you sure you want to continue?',
        asymmetryWarningTitle: "Asymmetric Settings Warning",
        asymmetryWarningBody: "The columns in your 'Source Data Import Range' do not match the columns in your 'Source Header Import Filter'.\n\nThe script will prioritize the 'Source Header Import Filter' and only import the columns you specified there.\n\nRange Columns: {RANGE_HEADERS}\nFilter Columns: {FILTER_HEADERS}\n\nDo you want to continue?",
        filterMismatchTitle: "No Matching Data Found",
        filterMismatchBody: "Source data was found, but no rows matched your filter criteria.\n\nPlease check your 'Keyword Filter Conditions' and ensure they correctly correspond to the source data.\n\nThe target sheet will now be cleared.",
        preCheckWarningTitle: 'Comparison Field Warning',
        preCheckWarningBody: 'Before running the comparison, the following issues were found in the "Source Compare Field" ({COLUMN}):\n\n{MESSAGE}\n\nContinuing may lead to unexpected results. Are you sure you want to proceed?',
        preCheckWarningBodyTarget: 'Before running the comparison, the following issues were found in the "Target Lookup Field" ({COLUMN}):\n\n{MESSAGE}\n\nContinuing may lead to unexpected results. Are you sure you want to proceed?',
        preCheckCancelled: 'Comparison cancelled by user.',
        errorUrlRequired: 'Source Spreadsheet URL cannot be empty.',
        errorSheetNameRequired: 'Source Data Sheet Name cannot be empty.',
        errorTargetSheetRequired: "Current Spreadsheet's Sheet Name cannot be empty.",
        errorVerifyUrlRequired: 'Source Spreadsheet URL cannot be empty.',
        errorVerifySheetNameRequired: 'Source Data Sheet Name cannot be empty.',
        errorInvalidUrl: "Could not access the provided URL. Please check if it is correct and that you have access permissions.",
        errorInvalidHeaderRange: "Could not read headers from the specified range '{RANGE}'. Please check the range and source sheet.",
        errorInvalidHeaderRow: "Could not read headers from row {ROW_NUM}. Please check the row number and sheet name.",
        errorInvalidColumnFormat: "Invalid format. Please enter a valid column letter (e.g., A, B, AA).",
        errorInvalidColumnSave: "Cannot save. Please correct the invalid column letters (marked in red).",
        noSheetsFound: "No sheets were found in the spreadsheet. It might be empty or inaccessible.",
        emptyRowsFound: 'Empty values found in rows: {ROWS}.',
        duplicateValuesFound: 'Duplicate value found: \'{VALUE}\' is repeated in rows: {ROWS}',
        multipleDuplicateValuesFound: 'Duplicate values found: {DETAILS}',
        sourceCompareFieldCheckError: 'The specified Source Compare Field ({COLUMN}) is not within the Source Compare Range ({RANGE}).',
        targetLookupFieldCheckError: 'Could not find data in the specified Target Lookup Field ({COLUMN}). Please check the column letter and the Target Start Row.',
        unsavedWarningTitle: 'Warning',
        unsavedWarningBody: 'Unsaved settings will be lost when switching sheets. It is recommended to save your current settings first.',
        saveAndContinueButton: 'Save Current Settings',
        continueAnywayButton: 'Continue Anyway',
        duplicateHeaderWarning: "Duplicate headers found in source data: {HEADERS}. To ensure accuracy, please specify a unique header range in the 'Source Header Import Filter' field.",
        invalidHeaders: "The following headers were not found in the source sheet: {HEADERS}.",
        invalidKeywords: "The following keywords were not found under header '{HEADER}': {KEYWORDS}.",
        errorDialogTitle: 'Error',
        cleanupError: 'Error during cleanup: {MESSAGE}',
        sumVerificationError: 'Error during sum verification: {MESSAGE}',
        sheetNotFound: 'Sheet named "{SHEET_NAME}" not found.',
        requiredMismatch: "Required {FIELDS}_Mismatch",
        noDataTitle: "No Data to Process",
        noDataBody: "No data rows were found in the target sheet to validate.\n\nPlease ensure there is data below the configured \"Data Start Row\".",
        noProcessableRowsTitle: "No Processable Rows Found",
        noProcessableRowsBody: "No rows containing data in the primary key column(s) {KEY_NAME_DISPLAY} could be found for processing.\n\nPlease check:\n1. That your required (Y) column(s) have data.\n2. That at least one field is marked as required (Y) in your Validation Conditions.",
        noSourceDataSuffix: '_No source data',
        perfectMatch: 'Perfect Matches',
        matchFailedLabel: 'Failed Matches',
        unmatchedTarget: 'Unmatched Target Headers',
        unmatchedSource: 'Unmatched Source Headers',
        mappingSuccessFormat: 'Match',
        mappingFailureFormat: 'Not Match',
        // Example Generation
        exampleTargetSheet: 'Target Sheet (Manufacturing)',
        exampleImportSourceSheet: 'Source_Import (Manufacturing)',
        exampleVerifySourceSheet: 'Source_Verify (Manufacturing)',
        exampleCompareSourceSheet: 'Source_Compare (Manufacturing)',
        exampleDashboardSheet_Sales: 'Dashboard | Sales Analysis',
        exampleSalesLogSheet_Sales: 'Source | Sales Log',
        exampleCustomerMasterSheet_Sales: 'Source | Customer Master',
        exampleProductTargetsSheet_Sales: 'Source | Product Targets',
        exampleGenerationConfirmBody: 'This action will create three new sheets in the current file and apply sample data and formatting.\n\nIf sheets with the same names already exist, their content will be [OVERWRITTEN].\n\nAre you sure you want to continue?\n\n If you do not want to overwrite current example sheets, please click NO then continously running example guide processing of right side.',
        generatingExampleProcess: 'Processing',
        generatingExampleBody: 'Generating example, please wait...',
        generationSuccessTitle: 'Complete',
        generationSuccessBody: 'Manufacturing example has been generated successfully!',
        generationSuccessBodySales: 'Business & Sales example has been generated successfully!',
        operationCancelled: 'Operation cancelled.',
        deleteExamplesItem: '🗑️ Delete Example Sheets',
        deleteExampleConfirmTitle: 'Confirm Deletion',
        deleteExampleConfirmBody: 'This will permanently delete the following example sheets:\n\n{SHEET_LIST}\n\nAre you sure you want to continue?',
        noExampleSheetsFound: 'No example sheets found to delete.',
        deleteExampleSuccess: 'Example sheets have been deleted successfully.',
        // Sales Example Data
        productLaptop: 'High-Performance Laptop',
        productKeyboard: 'Wireless Mechanical Keyboard',
        productMonitor: '27-inch 4K Monitor',
        productMouse: 'Wireless Mouse',
        customerA: 'Tech Giant Inc.',
        customerB: 'Creative Design Studio',
        customerC: 'Global Trade Ltd.',
        customerD: 'Digital Trends International',
        customerE: 'Apex Manufacturing Industries',
        regionNorth: 'North',
        regionCentral: 'Central',
        regionSouth: 'South',
        // Sales Example Headers
        headerOrderDate: 'Order Date',
        headerCustomerID: 'Customer ID',
        headerCustomerName: 'Customer Name',
        headerRegion: 'Region',
        headerSalesperson: 'Salesperson',
        headerProductName: 'Product Name',
        headerTotalSales: 'Total Sales',
        headerMonthlyTarget: 'Monthly Target',
        headerAchievementRate: 'Achievement Rate',
        headerQuantity: 'Quantity',
        headerUnitPrice: 'Unit Price',
        headerQuantity: 'Quantity',
        headerUnitPrice: 'Unit Price',
        // Error Messages
        errorNoImportSettingsFound: 'Could not find data import settings for sheet "{SHEET_NAME}". Please configure them first.',
        compareFailedTitle: 'Data Comparison Failed',
        errorNoCompareSettingsFound: 'Could not find data comparison settings for this sheet. Please configure them first via "Data Import Tool > Data Comparison Settings".',
        // --- Report Generator ---
        reportSettingsTitle: 'Report Generation Settings',
        sourceUrlPlaceholder: "Leave blank for current file, or paste URL for external file",
        exportReportButton: "Export Report",
        cancelButton: "Cancel",
        confirmExportButton: "Confirm Export",
        loadingMessage: "Loading...",
        generatingMessage: "Generating...",
        exportingMessage: "Exporting report, please wait...",
        exportSuccess: "Export successful!",
        exportFailure: "Export failed",
        exportSuccessSheet: "Report exported successfully! You can find it in the new sheet named \"{SHEET_NAME}\".",
        exportSuccessLink: "Report exported successfully! Click here to open the document:",
        backendValidationError: "Backend validation error",
        analysisError: "Analysis Error",
        executionError: "Execution Error",
        errorMissingSheetName: "Missing sheet name, cannot save settings.",
        errorSheetNotFound: "Sheet named \"{SHEET_NAME}\" not found in the specified spreadsheet.",
        errorInvalidRange: "Invalid range \"{RANGE}\". Please check your input.",
        errorAccessUrl: "Cannot access the provided URL. Please check the URL and your permissions.",
        errorNoHeaders: "Could not find any header data in the specified range {RANGE}. Please check your data range and ensure the first row contains field names.",
        errorNoAnalysisFields: "Please configure at least one analysis field.",
        errorNoDimensions: "Please configure at least one 'Dimension'.",
        errorNoMetrics: "Please configure at least one 'Metric'.",
        errorUnsupportedFormat: "Unsupported export format.",
        errorExportContent: "Please select content to export.",
        dataSourceTitle: "Data Source",
        fieldMappingTitle: "Field Mapping",
        analysisResultsTitle: "Analysis Results",
        exportSettingsTitle: "Export Settings",
        exportFormatTitle: "1. Select Export Format",
        exportContentTitle: "2. Select Content to Export",
        exportToSheetLabel: "Add to Google Sheet tab",
        exportToDocLabel: "Export to Google Docs",
        exportToPdfLabel: "Export as PDF",
        addAnalysisFieldButton: "+ Add Analysis Field",
        dimensionOption: "Dimension (Group By)",
        metricOption: "Metric (Calculate Value)",
        distributionChartTitle: "{DIMENSION} Distribution",
        chartByTitle: "{METRIC} by {DIMENSION}",
        kpiCardTitle: "Total {METRIC}",
        overviewTab: "Overview",
        analysisTab: "{DIMENSION} Analysis",
        minimizeHint: "Minimize",
        expandHint: "Expand Window",
        pinHint: "Pin Window",
        unminimizeHint: "Expand",
        collapseHint: "Collapse Window",
        unpinHint: "Unpin Window",
        metricWarningHint: "This field may contain non-numeric data and is not suitable as a metric.",
        pieChartLegendLabel: "Pie Chart: {DIMENSION} (Legend)",
        pieChartValuesLabel: "Pie Chart: {DIMENSION} (Values)",
        barChartLabel: "Bar Chart: {DIMENSION}",
        noDataForChart: "Not enough data to draw this chart.",
        pieChart: "Pie Chart",
        barChart: "Bar Chart",
        valuesSuffix: " (Values)",
    },
    'zh_TW': {
        // --- Main Menu ---
        mainMenuTitle: 'MasterDataAnalyzer',
        // --- Sub Menus ---
        importMenuTitle: '資料匯入工具',
        validationMenuTitle: '資料驗證工具',
        manageMenuTitle: '資料管理工具',
        guideMenuTitle: '範例生成與功能說明',
        // --- Items ---
        settingsItem: '⚙️ 資料匯入設定',
        manageSettingsItem: '⚙️ 資料監控管理設定',
        quickDeleteItem: '🗑️ 快速刪除分頁',
        reportSettingsItem: '📊 資料生成報表設定',
        runImportItem: '▶️ 資料匯入',
        stopImportItem: '⏹️ 終止執行匯入',
        resetImportItem: '🔄 清除匯入資料與進程',
        runCompareItem: '▶️ 執行資料比對',
        compareSettingsItem: '⚙️ 資料比對設定',
        verifySettingsItem: '⚙️ 資料驗證設定',
        runMsModeItem: '▶️ 執行驗證 (MS 累加項模式)',
        stopValidationItem: '⏹️ 終止執行驗證',
        verifySumsItem: '🔍 驗證總合項與累加項數值',
        cleanupItem: '🔄 清除驗證資料與進程',
        monitorMenuName: '資料變更監控',
        enableNotifyItem: '🟢 啟用自動通知',
        disableNotifyItem: '🔴 停用自動通知',
        checkNowItem: '✉️ 立即檢查並通知',
        privacyPolicyItem: '隱私權政策',
        // --- Guide Sub-menu Items (Placeholders) ---
        manufacturingGuide: '生產製造範例',
        manufacturingProductionTitle: '生產製造',
        businessGuide: '業務統計範例',
        // hrGuide: '人資管理範例',
        startBusinessGuide: '▶️ 啟動互動教學 (業務統計)',
        startManufacturingGuide: '▶️ 啟動互動教學 (生產製造)',
        // --- HTML UI Titles ---
        importSettingsTitle: '資料匯入設定',
        compareSettingsTitle: '資料比對設定',
        verifySettingsTitle: '資料驗證設定',
        manageSettingsTitle: '資料監控管理設定',
        quickDeleteTitle: '快速刪除分頁',
        reportSettingsTitle: '資料生成報表設定',
        sheetSelectionTitle: '選擇分頁',
        privacyPolicyTitle: '隱私權政策',
        // Privacy Policy Content
        privacyLastUpdated: '最後更新日期：2025年9月2日',
        privacyIntro: "感謝您使用 MasterDataAnalyzer (以下簡稱「本外掛」)。我們致力於保護您的隱私，並讓您清楚了解我們如何處理您的資料。本外掛的所有操作均在您的 Google 帳戶內部完成，我們不會收集、儲存或與任何第三方分享您的個人資訊或文件內容。",
        privacyDataCollectionTitle: '資料收集與使用',
        privacyDataCollectionP1: "本外掛是一個建立在 Google Apps Script 上的工具，旨在幫助您處理 Google Sheets 中的資料。",
        privacyDataCollectionL1: '**不收集個人資訊**：我們不會要求、收集或儲存任何您的個人身份資訊，例如姓名、電子郵件地址或聯絡資訊。',
        privacyDataCollectionL2: '**資料處理**：所有資料的讀取、處理與寫入操作，都直接在您授權的 Google Sheets、Google Docs 或 Google Drive 檔案中進行。資料不會被傳輸到我們的伺服器或任何外部服務。',
        privacyDataCollectionL3: "**設定儲存**：您為本外掛所做的設定（例如資料來源、欄位對應等）會使用 Google Apps Script 內建的 `PropertiesService` 儲存在與您的 Google 文件綁定的屬性中。這些設定僅供本外掛在您的帳戶中運作時讀取，我們無法從外部存取。",
        privacySecurityTitle: '安全性聲明與權限說明',
        privacySecurityP1: '為了提供完整的功能，本外掛在安裝時會向您請求以下 Google 服務的授權。這些權限僅用於實現所述之特定功能，絕不會用於其他目的。',
        privacySecurityP2: '以下是根據您 `appsscript.json` 檔案中 `oauthScopes` 的詳細說明：',
        privacyScopeSpreadsheetsTitle: '`https://www.googleapis.com/auth/spreadsheets`',
        privacyScopeSpreadsheetsPermission: '權限：查看、編輯、建立和刪除您的 Google Sheets 試算表。',
        privacyScopeSpreadsheetsPurpose: "用途：這是本外掛的核心權限。我們需要此權限來：",
        privacyScopeSpreadsheetsUse1: '讀取您指定的來源試算表資料。',
        privacyScopeSpreadsheetsUse2: '將處理後的資料寫入您的目標試算表。',
        privacyScopeSpreadsheetsUse3: '建立報表、範例資料等新分頁。',
        privacyScopeSpreadsheetsUse4: '執行「快速刪除分頁」功能。',
        privacyScopeDocumentsTitle: '`https://www.googleapis.com/auth/documents`',
        privacyScopeDocumentsPermission: '權限：查看、編輯、建立和刪除您的 Google Docs 文件。',
        privacyScopeDocumentsPurpose: '用途：此權限僅用於「匯出報表」功能。當您選擇將分析結果匯出為 Google 文件時，本外掛會建立一份新的 Google Doc 並將報表內容寫入其中。',
        privacyScopeDriveTitle: '`https://www.googleapis.com/auth/drive`',
        privacyScopeDrivePermission: '權限：查看、編輯、建立和刪除您 Google Drive 中的特定檔案。',
        privacyScopeDrivePurpose: '用途：此權限主要支援「匯出報表」功能：',
        privacyScopeDriveUse1: '當您選擇匯出為 Google 文件時，需要此權限在您的雲端硬碟中建立該文件。',
        privacyScopeDriveUse2: '當您選擇匯出為 PDF 時，本外掛會先建立一份 Google 文件，然後將其轉換為 PDF 檔案儲存至您的雲端硬碟，並可能刪除過程中的暫存文件。',
        privacyScopeGmailTitle: '`https://www.googleapis.com/auth/gmail.send`',
        privacyScopeGmailPermission: '權限：允許本外掛代表您傳送電子郵件。**（注意：本外掛無法讀取您的任何郵件）**',
        privacyScopeGmailPurpose: '用途：此權限僅用於「資料變更監控」功能。當您設定的儲存格範圍發生變更時，本外掛會依照您的設定，自動傳送通知郵件到您指定的信箱。',
        privacyScopeUITitle: '`https://www.googleapis.com/auth/script.container.ui`',
        privacyScopeUIPermission: '權限：在 Google Sheets 中顯示使用者介面。',
        privacyScopeUIPurpose: '用途：本外掛需要此權限才能在您的試算表上顯示所有操作介面，例如設定視窗、側邊欄、對話框與自訂選單。',
        privacyScopeScriptAppTitle: '`https://www.googleapis.com/auth/script.scriptapp`',
        privacyScopeScriptAppPermission: '權限：允許 Apps Script 建立及管理指令碼觸發器。',
        privacyScopeScriptAppPurpose: '用途：此權限用於建立 `onOpen` 觸發器，以便在您打開試算表時自動生成「MasterDataAnalyzer」選單。同時也用於「資料變更監控」功能，以建立偵測工作表變更的觸發器。',
        privacyScopeExternalRequestTitle: '`https://www.googleapis.com/auth/script.external_request`',
        privacyScopeExternalRequestPermission: '權限：允許 Apps Script 連接到外部網路服務。',
        privacyScopeExternalRequestPurpose: '用途：目前版本中，本外掛**不會**主動向任何非 Google 的外部伺服器發送請求或傳輸您的資料。此權限是為了保留未來可能的功能擴充性（例如，連接至公開的 API 服務），但現階段並未使用。',
        privacyScopeStorageTitle: '`https://www.googleapis.com/auth/script.storage`',
        privacyScopeStoragePermission: '權限：允許 Apps Script 儲存少量資料。',
        privacyScopeStoragePurpose: '用途：如前述「設定儲存」所述，我們使用此權限來保存您對各個功能的設定，以便您下次使用時無需重新輸入。',
        privacyChangesTitle: '政策變更',
        privacyChangesP1: '我們可能會不時更新本隱私權政策。任何變更都將發布在此頁面上，我們鼓勵您定期查看。',
        privacyContactTitle: '聯絡我們',
        privacyContactP1: '如果您對本隱私權政策有任何疑問，請透過 [tsengmercury@gmail.com] 與我們聯繫。',
        // --- NEW: Dashboard Generator UI ---
        sectionFieldMapping: '欄位對應',
        regionColumnLabel: '地區欄位',
        productColumnLabel: '產品/項目欄位',
        salesColumnLabel: '數值欄位 (銷售額)',
        generateReportButton: '生成報表',
        step1Title: '第一步：準備原始數據',
        step1Description: '這是一張典型的銷售流水帳，包含了日期、地區、產品和銷售額等資訊。',
        // sourceSpreadsheetUrlLabel: '來源資料表 URL',
        sourceDataSheetNameLabel: '來源資料分頁名稱',
        sourceDataRangeLabel: '資料範圍 (包含標頭)',
        sourceDataRangePlaceholder: '例如: A1:G100',
        generateDashboardButton: '一鍵生成儀表板',
        generatingDashboard: '儀表板生成中...',
        step2Title: '第二步：查看生成結果',
        overviewTab: '總覽',
        productAnalysisTab: '產品分析',
        regionAnalysisTab: '地區分析',
        rawDataTab: '原始數據',
        salesOverviewTitle: '銷售總覽',
        totalSalesLabel: '總銷售額',
        totalOrdersLabel: '總訂單數',
        regionSalesDistributionTitle: '各地區銷售佔比',
        productSalesAnalysisTitle: '各產品銷售分析',
        totalSalesAxisTitle: '總銷售額',
        productAxisTitle: '產品',
        regionSalesAnalysisTitle: '各地區銷售分析',
        regionAxisTitle: '地區',
        errorTitle: '錯誤',
        requiredFieldsError: '您選取的範圍中必須包含以下標頭：{HEADERS}。請檢查您的資料範圍。',
        // Tutorial Steps
        businessGuideTutorialTitle: '業務統計範例 - 互動教學',
        manufacturingGuideTutorialTitle: '生產製造範例 - 互動教學',
        // Business Tutorial
        businessWelcomeTitle: "歡迎！ (業務統計)",
        businessWelcomeInstruction: "歡迎來到「業務統計範例」的互動教學。\n\n我們的目標是將一份不完整的銷售紀錄，透過 MasterDataAnalyzer 的功能，變成一份完整的分析報表。\n\n點擊「下一步」開始我們的第一個任務！",
        businessTask1Step1Title: "任務一：資料匯入 (1/6)",
        businessTask1Step1Instruction: "首先，請確認您已啟用 <code>{SHEET_NAME}</code> 分頁。\n\n我們的目標是將 <code>{SOURCE_SHEET_NAME}</code> 的銷售紀錄，篩選後匯入到目前的儀表板中。",
        businessTask1Step2Title: "任務一：資料匯入 (2/6)",
        businessTask1Step2Instruction: "請點擊頂端選單的 <code>MasterDataAnalyzer > 資料匯入工具 > ⚙️ 資料匯入設定</code>。",
        businessTask1Step3Title: "任務一：資料匯入 (3/6)",
        businessTask1Step3Instruction: "在設定視窗中，請進行以下設定：\n1. **來源資料表 URL**: (貼上當前檔案的網址)\n2. **來源資料分頁名稱**: 選擇 <code>{SOURCE_SHEET_NAME}</code>\n3. **目標資料表分頁名稱**: 應會自動帶入 <code>{TARGET_SHEET_NAME}</code>",
        businessTask1Step4Title: "任務一：資料匯入 (4/6)",
        businessTask1Step4Instruction: "接著，設定資料範圍：\n1. **資料匯入的標頭之起始列數**: <code>1</code>\n2. **資料匯入的起始列數**: <code>2</code>\n3. **來源資料匯入範圍**: <code>A2:E9</code>",
        businessTask1Step5Title: "任務一：資料匯入 (5/6)",
        businessTask1Step5Instruction: "假設我們只想分析「{PRODUCT_NAME}」的業績，請新增一筆篩選條件：\n1. **標頭名稱**: <code>{HEADER_NAME}</code>\n2. **關鍵字**: <code>{PRODUCT_NAME}</code>",
        businessTask1Step6Title: "任務一：資料匯入 (6/6)",
        businessTask1Step6Instruction: "太棒了！所有設定都已完成。\n\n請點擊「儲存設定」並關閉視窗，然後從選單執行 <code>MasterDataAnalyzer > 資料匯入工具 > ▶️ 資料匯入</code>。",
        businessCheckpointTitle: "任務一完成！",
        businessCheckpointInstruction: "恭喜您！您已成功將原始數據匯入儀表板。\n\n您想繼續學習下一個核心功能「資料比對」嗎？",
        businessTask2Step1Title: "任務二：豐富客戶資料 (1/3)",
        businessTask2Step1Instruction: "太棒了！現在儀表板有了原始數據，但還缺少客戶的詳細資訊。\n\n接下來，我們將使用「資料比對」功能，從 \"[來源] 客戶主檔\" 中查找並填入客戶資料。\n\n打開 <code>資料匯入工具 > ⚙️ 資料比對設定</code>，並開始進行任務二的設定：",
        businessTask2Step2Title: "任務二：豐富客戶資料 (2/3)",
        businessTask2Step2Instruction: "請進行以下設定：\n1. **來源分頁**: 請選擇 \"來源資料分頁名稱\" 為 <code>{SOURCE_SHEET_NAME}</code>，而 \"目標資料表分頁名稱\" 為 <code>{TARGET_SHEET_NAME}</code>。\n2. \"目標資料起始列數\" 設定為 <code>2</code>，\"來源資料比對範圍\" 設定為 <code>A2:D6</code>。\n3. \"資料比對欄位對應\" 請依照下方順序逐一設定：\n  - **目標查找欄位**: <code>B</code> (客戶ID)\n  - **來源比對欄位**: <code>A</code> (客戶ID)\n  - **來源返回欄位**: <code>B</code> (客戶全名)\n  - **目標寫入欄位**: <code>C</code> (客戶全名)",
        businessTask2Step3Title: "任務二：豐富客戶資料 (3/3)",
        businessTask2Step3Instruction: "設定儲存後，請執行 <code>MasterDataAnalyzer > 資料匯入工具 > ▶️ 執行資料比對</code>。\n\n這時您將會看到客戶全名已被成功匯入。請重複此操作，將 \"所屬地區\" 和 \"負責業務員\" 也一併填入，您將可以得到完整的業務統計儀表板囉。",
        businessTask3Step1Title: "任務三：比對銷售目標 (1/2)",
        businessTask3Step1Instruction: "儀表板的資料越來越完整了！最後一步，讓我們來比對每項產品的銷售目標。\n\n請再次打開「資料比對設定」。",
        businessTask3Step2Title: "任務三：比對銷售目標 (2/2)",
        businessTask3Step2Instruction: "這次，請用 \"產品名稱\" 作為目標對比基準，並選擇 \"來源資料分頁名稱\" 為 <code>{SOURCE_SHEET_NAME}</code>。<br>\"來源資料比對範圍\" 請設定為 <code>A2:B5</code>，\"資料比對欄位對應\" 請依照下方順序逐一設定：<br>  - **目標查找欄位**: <code>F</code> (產品名稱)<br>  - **來源比對欄位**: <code>A</code> (產品名稱)<br>  - **來源返回欄位**: <code>B</code> (目標月銷售額)<br>  - **目標寫入欄位**: <code>J</code> (目標月銷售額)<br><br>接著，再次執行 \"資料比對\"。",
        businessFinalStepTitle: "恭喜您，分析完成！",
        businessFinalStepInstruction: "所有資料都已到位！您現在擁有了一份乾淨、完整的數據。\n\n接下來，您可以手動或點擊下方按鈕，在對應的欄位中填入公式來完成最後的計算：\n- **總銷售額**: <code>=H2*I2</code> (數量 * 單價)\n- **業績達成率**: <code>=G2/J2</code> (總銷售額 / 目標月銷售額)",
        // Manufacturing Tutorial
        mfgWelcomeTitle: "歡迎！ (生產製造)",
        mfgWelcomeInstruction: "歡迎來到「生產製造範例」的互動教學。\n\n此教學將引導您使用 MasterDataAnalyzer 的「資料匯入 (陣列模式)」與「資料驗證」功能，將複雜的物料清單轉換為標準化格式，並與主資料進行比對。",
        mfgTask1Step1Title: "任務一：資料匯入 (1/6)",
        mfgTask1Step1Instruction: "首先，請確認您已啟用 <code>{SHEET_NAME}</code> 分頁。\n\n我們的目標是將 <code>{SOURCE_SHEET_NAME}</code> 中非結構化的資料，轉換並匯入到目前的目標工作表中。",
        mfgTask1Step2Title: "任務一：資料匯入設定 (2/6)",
        mfgTask1Step2Instruction: "請點擊頂端選單的 <code>MasterDataAnalyzer > 資料匯入工具 > ⚙️ 資料匯入設定</code>。",
        mfgTask1Step3Title: "任務一：資料匯入設定 (3/6)",
        mfgTask1Step3Instruction: "在設定視窗中，請進行以下設定：\n1. **來源資料表 URL**: (貼上當前檔案的網址)\n2. **來源資料分頁名稱**: 選擇 <code>{SOURCE_SHEET_NAME}</code>\n3. **目標資料表分頁名稱**: 應會自動帶入 <code>{TARGET_SHEET_NAME}</code>",
        mfgTask1Step4Title: "任務一：資料匯入設定 (4/6)",
        mfgTask1Step4Instruction: "接下來是關鍵步驟，啟用「資料陣列比對」模式：\n1. **資料匯入的標頭之起始列數**: <code>3</code>\n2. **資料匯入的起始列數**: <code>4</code>\n3. **來源資料匯入範圍**: <code>A9:E14</code>\n4. **其他區塊的標頭的起始列數**: <code>F2:I2</code>\n5. **其他區塊的標頭內的數據範圍**: <code>F9:I14</code>",
        mfgTask1Step5Title: "任務一：資料篩選設定 (5/6)",
        mfgTask1Step5Instruction: "在篩選與驗證功能區塊裡，您可以自行挑選想匯入的標頭與關鍵字篩選。\n1. 在 \"來源標頭匯入範圍設定\" 輸入 <code>A8:E8</code>\n2. 在 \"關鍵字篩選條件 (AND)\" 選擇 <code>Owner</code>\n3. 點擊關鍵字右側 \"選取\" 方塊選擇 <code>Mark</code>、<code>Linda</code>、<code>Mary</code>",
        mfgTask1Step6Title: "任務一：執行資料匯入 (6/6)",
        mfgTask1Step6Instruction: "當資料匯入設定完成後請別忘記點擊 \"儲存設定\"。設定與儲存完成後，請繼續執行 <code>▶️ 資料匯入</code>。\n\n這時您應會看到 A - G 欄的匯入資料，代表著資料已成功篩選並匯入囉。",
        mfgCheckpointTitle: "任務一完成！",
        mfgCheckpointInstruction: "恭喜！您已成功將陣列資料轉換並匯入。\n\n接下來，您想學習如何使用「資料驗證」功能，來比對匯入的資料與主資料的差異嗎？",
        mfgTask2Step1Title: "任務二：資料驗證設定 (1/4)",
        mfgTask2Step1Instruction: "很好！現在我們來驗證匯入資料的正確性。\n\n請打開 <code>MasterDataAnalyzer > 資料驗證工具 > ⚙️ 資料驗證設定</code>。",
        mfgTask2Step2Title: "任務二：資料範圍設定 (2/4)",
        mfgTask2Step2Instruction: "在[資料驗證]設定視窗中，請進行以下設定：<br>1. **來源資料表 URL**: (貼上當前檔案的網址)<br>2. **來源資料分頁名稱**: 選擇 <code>{SOURCE_SHEET_NAME}</code><br>3. 請依照下列設定順序填妥資料範圍設定：<br>   - 資料驗證後匯入的起始列數請填入 <code>4</code><br>   - 目標資料標頭起始列請填入 <code>3</code><br>   - 來源資料標頭起始列請填入 <code>1</code><br><br>接下來繼續下一步，將開始設定 \"欄位驗證條件\" 與 \"驗證結果輸出\"。",
        mfgTask2Step3Title: "任務二：欄位驗證設定 (3/4)",
        mfgTask2Step3Instruction: "在欄位與驗證條件功能區塊裡，首先要設定的是 \"目標工作表欄位\" 與 \"來源工作表欄位\" 的對應關係。<br>建議您可以優先使用 <b>自動對應驗證欄位</b> 功能，此功能將會 <b>自動匹配與推薦</b> 合適的標頭給您設定。<br>在此項任務裡，欄位驗證要設定的條件如下：<br>目標工作表欄位 - 來源工作表欄位<br><code>B</code> - <code>B</code><br><code>C</code> - <code>C</code><br><code>D</code> - <code>D</code><br><code>E</code> - <code>E</code><br><code>G</code> - <code>A</code>",
        mfgTask2Step4Title: "任務二：驗證結果輸出設定 (4/4)",
        mfgTask2Step4Instruction: "接著設定當驗證成功時，要從來源回填的資料欄位。<br>同樣地，您也可以使用 <b>自動對應輸出欄位</b> 功能來加速設定。<br>驗證結果輸出要設定的欄位條件如下：<br>目標工作表欄位 - 來源工作表欄位<br><code>H</code> - <code>F</code><br><code>I</code> - <code>G</code><br><code>J</code> - <code>H</code><br><br>最後，我們將 \"不吻合資訊輸出\" 的欄位設為 <code>K</code>，以便腳本將錯誤訊息寫入此處。",
        mfgFinalStepTitle: "所有設定都已準備就緒！",
        mfgFinalStepInstruction: "請記得點擊 [儲存設定]，然後開始執行 <code>資料驗證工具 > ▶️ 執行驗證 (MS 累加項模式)</code>！<br><br>執行後，請查看 <code>{SHEET_NAME}</code> 的 K 欄，您會看到腳本已自動標示出所有不吻合的項目及其原因。",
        sectionSourceAndTarget: '來源與目標',
        sectionDataRanges: '資料範圍',
        sectionFilterAndValidate: '篩選與驗證',
        sectionValidationConditions: '欄位驗證條件',
        sectionValidationOutputs: '驗證結果輸出',
        sectionMismatchOutput: '不吻合資訊輸出',
        sectionDataManagement: '資料管理設定',
        sectionRangesAndConditions: '資料範圍與比對條件',
        sectionFieldMapping: '資料比對欄位對應',
        closeButton: '關閉',
        saveButton: '儲存設定',
        selectButton: '選取',
        okButton: '確定',
        defaultTemplateButton: '預設模板',
        removeAllButton: '移除所有欄位',
        checkButton: '檢查',
        pinWindowTooltip: '釘選視窗',
        unpinWindowTooltip: '取消釘選視窗',
        expandWindowTooltip: '擴展視窗',
        collapseWindowTooltip: '收合視窗',
        dragAndActionHelp: '點擊上方標題列右側空白區域來拖曳視窗\n點擊右側圖示可擴展或釘選。',
        // sourceSpreadsheetUrlLabel: '來源資料表 URL',
        sourceDataSheetNameLabel: '來源資料分頁名稱',
        // Import Settings UI
        currentTargetSheetNameLabel: '目標資料表分頁名稱',
        settingsForSheetHint: '當前儲存設定的頁面：{SHEET_NAME}',
        importHeaderStartRowLabel: '資料匯入的標頭之起始列數',
        importDataStartRowLabel: '資料匯入的起始列數',
        sourceDataRangeLabel: '來源資料匯入範圍',
        validationHeaderStartRowLabel: '其他區塊的標頭的起始列數',
        validationMatrixRangeLabel: '其他區塊的標頭內的數據範圍',
        headerImportFilterLabel: '來源標頭匯入範圍設定',
        keywordFiltersLabel: '關鍵字篩選條件 (AND)',
        addFilterConditionLabel: '+ 新增篩選條件',
        headerPlaceholder: '標頭名稱',
        keywordsPlaceholder: '關鍵字 (以逗號分隔)',
        dataStartRowLabel: '資料驗證後匯入的起始列數',
        verifySourceUrlLabel: '來源資料表 URL',
        verifySourceSheetNameLabel: '來源資料分頁名稱',
        mismatchColumnLabel: '不吻合描述輸出欄位',
        targetHeaderRowLabel: '目標資料標頭起始列',
        sourceHeaderRowLabel: '來源資料標頭起始列',
        targetColumnLabel: '目標工作表欄位',
        sourceColumnLabel: '來源工作表欄位',
        primaryValidationLabel: '設為主要驗證條件',
        selectAllLabel: '全選/取消所有驗證條件',
        fieldMappingJSONLabel: '欄位對應與驗證設定',
        checkEmptyValuesButton: '檢查來源空值',
        monitorRangeLabel: '文件內容變更自動通知',
        monitorEmailLabel: '通知接收者 Email',
        monitorSubjectLabel: '通知標題',
        monitorBodyLabel: '通知內文',
        addValidationMappingButton: '+ 新增驗證欄位',
        autoMapValidationButton: '自動對應驗證欄位',
        addOutputMappingButton: '+ 新增輸出欄位',
        autoMapOutputButton: '自動對應輸出欄位',
        targetStartRowLabel: '目標資料起始列數',
        targetStartRowHelp: '從此列號開始讀取目標資料並寫入比對結果。',
        sourceCompareRangeLabel: '來源資料比對範圍',
        sourceCompareRangeHelp: '在來源分頁中，包含「比對欄」與「返回欄」的範圍。',
        targetLookupColLabel: '目標查找欄位',
        targetLookupColHelp: '目標工作表中，要用來查找值的欄位。',
        targetWriteColLabel: '目標寫入欄位',
        targetWriteColHelp: '將比對結果寫入此欄位。',
        sourceLookupColLabel: '來源比對欄位',
        sourceLookupColHelp: '在「比對範圍」中，用來被比對的欄位。',
        sourceReturnColLabel: '來源返回欄位',
        sourceReturnColHelp: '在「比對範圍」中，找到後要返回的欄位。',
        sourceUrlHelp: '請貼上來源 Google Sheet 的完整網址。',
        importHeaderStartRowHelp: "數據最前方的欄位的分類名稱即為標頭。",
        importDataStartRowHelp: '資料從目標分頁的第幾列開始寫入。',
        sourceDataRangeHelp: "標頭之後的數據的行列範圍。",
        validationHeaderStartRowHelp: "若填寫此欄位，將啟用「資料陣列比對」模式。此功能可讓「資料匯入之標頭」對照「其他區塊的標頭的起始列數」並將項目數據統計的資料列出。",
        headerImportFilterHelp: '選填。指定要匯入的標頭，可使用逗號分隔的列表 (標頭1,標頭2) 或來源工作表上的範圍 (A1:D1)。若留空，則會匯入「來源資料匯入範圍」內的所有欄位。',
        verifyStartRowHelp: "資料驗證的起始列數將依據設定列數開始更新",
        verifySourceUrlHelp: '用來進行資料校對的試算表網址。',
        mismatchColumnHelp: '用來輸出不吻合資訊的欄位字母 (例如: K)。',
        targetHeaderRowHelp: '目標工作表中，標頭所在的列數。',
        sourceHeaderRowHelp: '來源工作表中，標頭所在的列數。',
        monitorRangeHelp: '請指定要監控變更的儲存格範圍 (例如 A2:E20)。',
        monitorEmailHelp: '請輸入要接收通知的 Email 地址。',
        monitorSubjectHelp: '自訂通知郵件的標題。可使用 {SHEET_NAME} 作為預留位置。',
        monitorBodyHelp: '自訂通知郵件的內文。可使用預留位置：{SHEET_NAME}, {RANGE}, {TIMESTAMP}, {CHANGES_COUNT}, {CHANGE_DETAILS}, {SHEET_URL}。',
        defaultSubjectTemplate: '[變更通知] 工作表 "{SHEET_NAME}" 內容已更新',
        defaultBodyTemplate: '您好，\n\n系統偵測到您監控的試算表內容已發生變更。\n\n工作表名稱: {SHEET_NAME}\n監控範圍: {RANGE}\n變更時間: {TIMESTAMP}\n\n變更詳情 ({CHANGES_COUNT} 項):\n{CHANGE_DETAILS}\n\n請點擊以下連結查看最新內容：\n{SHEET_URL}',
        savingMessage: '儲存中...',
        validatingMessage: '資料校驗中...',
        saveSuccess: '設定已成功儲存！',
        saveFailure: '儲存設定失敗',
        autoMappingMessage: '自動對應中...',
        autoMapSuccessBody: '成功找到並建立 {COUNT} 個欄位對應，請檢視後儲存。',
        autoMapNoMatchTitle: '找不到符合項目',
        autoMapNoMatchBody: '在來源與目標工作表中，找不到任何名稱相同的標頭。請檢查您的設定。',
        checkingMessage: '檢查中...',
        checkSuccess: '欄位檢查通過，沒有空值或重複值。',
        checkFailure: '欄位檢查提示：',
        importCancelled: '使用者已取消匯入。',
        settingsError: '設定錯誤',
        headerLessThanStartError: '「資料匯入的標頭之起始列數」必須小於「資料匯入的起始列數」。',
        preflightTitle: '設定確認',
        preflightWarning: '提醒：',
        preflightFilterWarning: '- 有一個篩選條件已指定標頭但未填寫關鍵字，該條件將被忽略。',
        preflightSuggestion: '這可能是非預期的結果，建議檢查您的設定。',
        preflightConfirmation: '您確定要繼續執行匯入嗎？',
        asymmetryWarningTitle: "設定不對稱警告",
        asymmetryWarningBody: "您設定的「來源資料匯入範圍」與「來源標頭匯入範圍設定」所涵蓋的欄位不一致。\n\n腳本將優先採用「來源標頭匯入範圍設定」的設定來匯入資料。\n\n範圍內的標頭：{RANGE_HEADERS}\n篩選器標頭：{FILTER_HEADERS}\n\n您確定要繼續嗎？",
        filterMismatchTitle: "找不到符合條件的資料",
        filterMismatchBody: "已找到來源資料，但沒有任何資料列符合您的篩選條件。\n\n請檢查您的「關鍵字篩選條件」並確保它們與來源資料正確對應。\n\n目標工作表將會被清空。",
        preCheckWarningTitle: '比對欄位警告',
        preCheckWarningBody: '在執行比對前，於「來源比對欄位」({COLUMN}) 中發現以下問題：\n\n{MESSAGE}\n\n繼續執行可能會導致非預期的比對結果。您確定要繼續嗎？',
        preCheckWarningBodyTarget: '在執行比對前，於「目標查找欄位」({COLUMN}) 中發現以下問題：\n\n{MESSAGE}\n\n繼續執行可能會導致非預期的比對結果。您確定要繼續嗎？',
        preCheckCancelled: '比對已由使用者取消。',
        errorUrlRequired: '「來源資料表 URL」不可為空。',
        errorSheetNameRequired: '「來源資料分頁名稱」不可為空。',
        errorTargetSheetRequired: '「現行資料表分頁名稱」不可為空。',
        errorVerifyUrlRequired: '「來源資料表 URL」不可為空。',
        errorVerifySheetNameRequired: '「來源資料分頁名稱」不可為空。',
        errorInvalidUrl: "無法存取您提供的 URL。請檢查網址是否正確，以及您是否擁有存取權限。",
        errorInvalidHeaderRange: "無法從指定範圍 '{RANGE}' 讀取標頭。請檢查該範圍與來源分頁名稱是否正確。",
        errorInvalidHeaderRow: "無法從列號 {ROW_NUM} 讀取標頭。請檢查列號與分頁名稱是否正確。",
        errorInvalidColumnFormat: "格式錯誤，請輸入有效的欄位字母 (例如: A, B, AA)。",
        errorInvalidColumnSave: "無法儲存，請修正標示為紅色的無效欄位字母。",
        noSheetsFound: "在試算表中找不到任何分頁。它可能是空的或無法存取。",
        emptyRowsFound: '發現空值於列: {ROWS}.',
        duplicateValuesFound: '發現重複值: \'{VALUE}\' 重複於列: {ROWS}',
        multipleDuplicateValuesFound: '發現重複值: {DETAILS}',
        sourceCompareFieldCheckError: '指定的來源比對欄位 ({COLUMN}) 不在來源比對範圍 ({RANGE}) 內。',
        targetLookupFieldCheckError: '在指定的目標查找欄位 ({COLUMN}) 中找不到資料。請檢查欄位字母與目標資料起始列數。',
        unsavedWarningTitle: '提醒',
        unsavedWarningBody: '選擇其他分頁時若有未儲存的設定將被清空，建議您對當前分頁的設定先儲存。',
        saveAndContinueButton: '儲存當前設定',
        continueAnywayButton: '仍然繼續',
        duplicateHeaderWarning: "在來源資料中偵測到重複的標頭：{HEADERS}。為確保篩選準確，請在「來源標頭匯入範圍設定」欄位中明確指定一個無重複的標頭範圍。",
        invalidHeaders: "找不到來源工作表以下標頭：{HEADERS}。",
        invalidKeywords: "在標頭 '{HEADER}' 下找不到以下關鍵字：{KEYWORDS}。",
        errorDialogTitle: '錯誤',
        cleanupError: '清理過程中發生錯誤：{MESSAGE}',
        sumVerificationError: '總和驗證過程中發生錯誤：{MESSAGE}',
        sheetNotFound: '找不到名為「{SHEET_NAME}」的工作表。',
        requiredMismatch: "必要 {FIELDS}_Mismatch",
        noDataTitle: "無資料可處理",
        noDataBody: "目標工作表中沒有找到任何資料列可供驗證。\n\nPlease ensure there is data below the configured \"Data Start Row\".",
        noProcessableRowsTitle: "找不到可處理的資料列",
        noProcessableRowsBody: "在目標工作表中，找不到任何包含「主要鍵」{KEY_NAME_DISPLAY}的資料列可供處理。\n\n請檢查：\n1. 您設定為必要(Y)的欄位中是否已填入資料。\n2. 「來源欄位驗證條件」中是否有至少一項被設為必要(Y)。",
        noSourceDataSuffix: '_無來源資料',
        perfectMatch: '匹配成功',
        matchFailedLabel: '匹配失敗 (Failed Matches)',
        unmatchedTarget: '不匹配的目標欄位',
        unmatchedSource: '不匹配的來源欄位',
        mappingSuccessFormat: '匹配 (Match)',
        mappingFailureFormat: '不匹配 (Not Match)',
        exampleTargetSheet: '目標資料表 (生產製造)',
        exampleImportSourceSheet: '來源_匯入 (生產製造)',
        exampleVerifySourceSheet: '來源_驗證 (生產製造)',
        exampleCompareSourceSheet: '來源_比對 (生產製造)',
        exampleDashboardSheet_Sales: '月報分析儀表板 (業務統計)',
        exampleSalesLogSheet_Sales: '[來源] 銷售紀錄',
        exampleCustomerMasterSheet_Sales: '[來源] 客戶主檔',
        exampleProductTargetsSheet_Sales: '[來源] 產品目標',
        exampleGenerationConfirmBody: '此操作將會在目前的檔案中新增三個工作分頁，並套用範例資料與格式。\n\n如果已存在同名分頁，其內容將會被【覆蓋】。\n\n您確定要繼續嗎？\n\n 若您不想覆蓋當前的分頁請按否，將會繼續開啟側邊欄引導教學。',
        generatingExampleProcess: '處理中',
        generatingExampleBody: '正在生成範例，請稍候...',
        generationSuccessTitle: '完成',
        generationSuccessBody: '生產製造範例已成功生成！',
        generationSuccessBodySales: '業務統計範例已成功生成！',
        operationCancelled: '已取消操作。',
        deleteExamplesItem: '🗑️ 刪除範例分頁',
        deleteExampleConfirmTitle: '確認刪除',
        deleteExampleConfirmBody: '此操作將永久刪除以下範例分頁：\n\n{SHEET_LIST}\n\n您確定要繼續嗎？',
        noExampleSheetsFound: '找不到可刪除的範例分頁。',
        deleteExampleSuccess: '範例分頁已成功刪除。',
        // Sales Example Data
        productLaptop: '高效能筆電',
        productKeyboard: '無線機械鍵盤',
        productMonitor: '27吋4K顯示器',
        productMouse: '無線滑鼠',
        customerA: '科技巨人股份有限公司',
        customerB: '創新設計工作室',
        customerC: '環球貿易有限公司',
        customerD: '數位潮流國際',
        customerE: '卓越製造工業',
        regionNorth: '北部',
        regionCentral: '中部',
        regionSouth: '南部',
        // Sales Example Headers
        headerOrderDate: '訂單日期',
        headerCustomerID: '客戶ID',
        headerCustomerName: '客戶全名',
        headerRegion: '所屬地區',
        headerSalesperson: '負責業務員',
        headerProductName: '產品名稱',
        headerTotalSales: '總銷售額',
        headerMonthlyTarget: '目標月銷售額',
        headerAchievementRate: '業績達成率',
        headerQuantity: '数量',
        headerUnitPrice: '單價',
        headerQuantity: '數量',
        headerUnitPrice: '單價',
        // Error Messages
        errorNoImportSettingsFound: '找不到工作表 "{SHEET_NAME}" 的資料匯入設定，請先設定。',
        compareFailedTitle: '資料比對失敗',
        errorNoCompareSettingsFound: '找不到此工作表的資料比對設定。請先透過「資料匯入工具 > 資料比對設定」進行設定。',
        // --- Report Generator ---
        reportSettingsTitle: '資料生成報表設定',
        sourceUrlPlaceholder: "若為外部檔案請貼上 URL，否則留空",
        exportReportButton: "匯出報表",
        cancelButton: "取消",
        confirmExportButton: "確認匯出",
        loadingMessage: "讀取中...",
        generatingMessage: "生成中...",
        exportingMessage: "報表匯出中，請稍候...",
        exportSuccess: "匯出成功！",
        exportFailure: "匯出失敗",
        exportSuccessSheet: "報表已成功匯出！您可以在試算表中名為「{SHEET_NAME}」的新分頁找到它。",
        exportSuccessLink: "報表已成功匯出！點此打開文件：",
        backendValidationError: "後端驗證錯誤",
        analysisError: "分析錯誤",
        executionError: "執行錯誤",
        errorMissingSheetName: "缺少工作表名稱，無法儲存設定。",
        errorSheetNotFound: "在指定的試算表中找不到名為「{SHEET_NAME}」的工作表。",
        errorInvalidRange: "無效的範圍「{RANGE}」。請檢查您的輸入。",
        errorAccessUrl: "無法存取提供的 URL。請檢查網址並確認您有權限。",
        errorNoHeaders: "在指定的範圍 {RANGE} 中找不到任何標頭資料。請檢查您的資料範圍是否正確，且第一行包含欄位名称。",
        errorNoAnalysisFields: "請設定至少一個分析欄位。",
        errorNoDimensions: "請設定至少一個「維度」。",
        errorNoMetrics: "請設定至少一個「指標」。",
        errorUnsupportedFormat: "不支援的匯出格式。",
        errorExportContent: "請選擇要匯出的內容",
        dataSourceTitle: "資料來源",
        fieldMappingTitle: "欄位對應",
        analysisResultsTitle: "分析結果",
        exportSettingsTitle: "匯出設定",
        exportFormatTitle: "1. 選擇匯出格式",
        exportContentTitle: "2. 選擇要匯出的內容",
        exportToSheetLabel: "新增至 Google Sheet 分頁",
        exportToDocLabel: "匯出至 Google 文件",
        exportToPdfLabel: "匯出為 PDF",
        addAnalysisFieldButton: "+ 新增分析欄位",
        dimensionOption: "維度 (分組依據)",
        metricOption: "指標 (計算數值)",
        distributionChartTitle: "{DIMENSION} 分佈",
        chartByTitle: "圖表：{METRIC} by {DIMENSION}",
        kpiCardTitle: "總 {METRIC}",
        overviewTab: "總覽",
        analysisTab: "{DIMENSION} 分析",
        rawDataTab: "原始數據",
        minimizeHint: "縮小",
        expandHint: "擴展視窗",
        pinHint: "釘選視窗",
        unminimizeHint: "展開",
        collapseHint: "收合視窗",
        unpinHint: "取消釘選",
        metricWarningHint: "此欄位可能包含非數值資料，不適合做為指標。",
        pieChartLegendLabel: "圖表：{DIMENSION} 圓餅圖 (圖例)",
        pieChartValuesLabel: "圖表：{DIMENSION} 圓餅圖 (數值)",
        barChartLabel: "圖表：{DIMENSION} 長條圖",
        noDataForChart: "此維度沒有足夠的數據可供繪圖。",
        pieChart: "圓餅圖",
        barChart: "長條圖",
        valuesSuffix: " (數值)",
    }
};

/**
 * Gets the appropriate translation object based on the user's locale.
 */
function getTranslations() {
    const locale = Session.getActiveUserLocale();
    return TRANSLATIONS[locale] || TRANSLATIONS.en;
}

/**
 * [REFACTORED] Adds all custom menus under a single "MasterDataAnalyzer" menu.
 */
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    const T = getTranslations();

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
        .addItem(T.disableNotifyItem, 'deleteOnChangeTrigger')
        .addSeparator()
        .addItem(T.checkNowItem, 'checkAndNotifyWrapper');

    const managementSubMenu = ui.createMenu(T.manageMenuTitle)
        .addItem(T.manageSettingsItem, 'showManageSettingsSidebar')
        .addSeparator()
        .addItem(T.quickDeleteItem, 'showQuickDeleteSheetUI')
        .addSeparator()
        .addSubMenu(monitorSubMenu)
        .addSeparator()
        .addItem(T.reportSettingsItem, 'showReportSettingsDialog');

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
    const T = getTranslations();
    const htmlTemplate = HtmlService.createTemplateFromFile('PrivacyPolicy.html');
    htmlTemplate.T = T;
    const htmlOutput = htmlTemplate.evaluate().setWidth(600).setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, T.privacyPolicyTitle);
}


/**
 * Shows the HTML settings for Data Import.
 */
function showImportSettingsSidebar() {
    const T = getTranslations();
    const htmlTemplate = HtmlService.createTemplateFromFile('SettingsPageImport');
    htmlTemplate.T = T;
    const htmlOutput = htmlTemplate.evaluate().setWidth(450).setHeight(600);
    SpreadsheetApp.getUi().showModelessDialog(htmlOutput, T.importSettingsTitle);
}

/**
 * Shows the HTML settings for Data Comparison.
 */
function showCompareSettingsSidebar() {
    const T = getTranslations();
    const htmlTemplate = HtmlService.createTemplateFromFile('SettingsPageCompare.html'); // Ensure this file exists
    htmlTemplate.T = T;
    const htmlOutput = htmlTemplate.evaluate().setWidth(450).setHeight(600);
    SpreadsheetApp.getUi().showModelessDialog(htmlOutput, T.compareSettingsTitle);
}

/**
 * Shows the HTML settings for Report Generation.
 */
function showReportSettingsDialog() {
    const T = getTranslations();
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
    const T = getTranslations();
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
    const T = getTranslations();
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
    const T = getTranslations();
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
    const T = getTranslations();
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
    const T = getTranslations();
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
    const T = getTranslations();
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
    const T = getTranslations();
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
    const T = getTranslations();

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
    const T = getTranslations();

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
    const T = getTranslations();
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
    const T = getTranslations();

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

    const monitorSettings = {};
    const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
    if (settingsSheet) {
        const values = settingsSheet.getDataRange().getValues();
        let currentSection = null;
        for (let i = 0; i < values.length; i++) {
            const row = values[i];
            const key = (row[0] || '').toString().trim();
            const key_lc = key.toLowerCase();

            if (key_lc.includes("data management settings") || key_lc.includes("資料管理設定")) {
                currentSection = 'manage';
                continue;
            }
            if (!key || currentSection !== 'manage') {
                continue;
            }

            const value = row[1] ? row[1].toString().trim() : null;
            if (key.includes("文件內容變更自動通知")) monitorSettings.monitorRange = value;
            else if (key.includes("通知接收者 Email")) monitorSettings.monitorRecipientEmail = value;
            else if (key.includes("通知標題")) monitorSettings.monitorSubject = value;
            else if (key.includes("通知內文")) monitorSettings.monitorBody = value;
        }
    }

    // Post-process header filter range
    if (importSettings.rawImportFilterHeaders && importSettings.sourceUrl && importSettings.sourceSheetName) {
        const T = getTranslations();
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

